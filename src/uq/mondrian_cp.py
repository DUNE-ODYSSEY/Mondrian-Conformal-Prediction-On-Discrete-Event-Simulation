"""
Week 11-12: Mondrian conformal prediction.

Standard CP (Week 9-10) calibrates a single pooled quantile across all
scenarios, which only guarantees *marginal* coverage - it says nothing about
whether coverage holds up within a specific staffing level or arrival regime.
That's exactly the first limitation Gopakumar et al. (2026) flag for CP in
physics domains. Mondrian CP addresses it by calibrating a separate quantile
per category, so each category gets its own (asymptotically valid) coverage
guarantee instead of relying on categories averaging out to the marginal rate.

Categories: our scenario space only has two real covariates (n_capacity,
arrival_rate_multiplier) - there's no "shift" field in the DES output, so we
don't invent one. Categories are the cross of staffing tercile x arrival-rate
tercile (Low/Med/High x Low/Med/High = 9 cells), using bin edges derived from
the calibration set's own quantiles (never the test set - that would leak
information into the taxonomy).

Same calibration/test data, same alpha=0.1, symmetric |y - yhat| nonconformity
measure as standard_cp.py's symmetric variant, so this is a clean, apples-to-
apples comparison of pooled vs. per-category calibration - not a different
setup dressed up as a comparison.

For each category we compute coverage two ways on the SAME test points:
  - "pooled" applies standard CP's single marginal quantile, evaluated
    per-category. This shows where the marginal guarantee breaks down.
  - "mondrian" applies that category's own quantile. This is the actual
    per-category coverage Mondrian CP delivers.
"""

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
DETAIL_PATH = "results/tables/mondrian_cp_detail.csv"
SUMMARY_PATH = "results/tables/mondrian_cp_summary.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1
LEVEL_NAMES = ["Low", "Med", "High"]


def conformal_quantile(scores, alpha):
    n = len(scores)
    level = min(np.ceil((n + 1) * (1 - alpha)) / n, 1.0)
    return np.quantile(scores, level, method="higher")


def assign_categories(df, cap_bounds, arr_bounds):
    cap_level = np.digitize(df["n_capacity"], cap_bounds)
    arr_level = np.digitize(df["arrival_rate_multiplier"], arr_bounds)
    labels = [f"staff={LEVEL_NAMES[c]}/arrival={LEVEL_NAMES[a]}" for c, a in zip(cap_level, arr_level)]
    return np.array(labels)


def main():
    df = pd.read_parquet(TRAIN_DATA_PATH)
    _, X_test, _, test_idx = train_test_split(df[FEATURES], df.index, test_size=0.2, random_state=42)
    test_df = df.loc[test_idx]

    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)

    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    test_cat = assign_categories(test_df, cap_bounds, arr_bounds)
    categories = sorted(set(cal_cat))

    detail_rows = []
    summary_rows = []
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")

        yhat_cal = model.predict(cal_df[FEATURES])
        y_cal = cal_df[target].values
        abs_resid_cal = np.abs(y_cal - yhat_cal)

        yhat_test = model.predict(X_test)
        y_test = test_df[target].values

        q_pooled = conformal_quantile(abs_resid_cal, ALPHA)

        pooled_coverages, mondrian_coverages = [], []
        for cat in categories:
            cal_mask = cal_cat == cat
            test_mask = test_cat == cat
            n_cal_cat = cal_mask.sum()
            n_test_cat = test_mask.sum()
            if n_test_cat == 0:
                continue

            q_cat = conformal_quantile(abs_resid_cal[cal_mask], ALPHA)

            y_c = y_test[test_mask]
            yhat_c = yhat_test[test_mask]

            covered_pooled = (np.abs(y_c - yhat_c) <= q_pooled).mean()
            covered_mondrian = (np.abs(y_c - yhat_c) <= q_cat).mean()

            pooled_coverages.append(covered_pooled)
            mondrian_coverages.append(covered_mondrian)

            detail_rows.append(
                {
                    "target": target,
                    "category": cat,
                    "n_cal": n_cal_cat,
                    "n_test": n_test_cat,
                    "pooled_coverage": covered_pooled,
                    "pooled_width": 2 * q_pooled,
                    "mondrian_coverage": covered_mondrian,
                    "mondrian_width": 2 * q_cat,
                }
            )

        pooled_coverages = np.array(pooled_coverages)
        mondrian_coverages = np.array(mondrian_coverages)
        summary_rows.append(
            {
                "target": target,
                "target_coverage": 1 - ALPHA,
                "pooled_coverage_min": pooled_coverages.min(),
                "pooled_coverage_max": pooled_coverages.max(),
                "pooled_coverage_range": pooled_coverages.max() - pooled_coverages.min(),
                "pooled_coverage_std": pooled_coverages.std(),
                "mondrian_coverage_min": mondrian_coverages.min(),
                "mondrian_coverage_max": mondrian_coverages.max(),
                "mondrian_coverage_range": mondrian_coverages.max() - mondrian_coverages.min(),
                "mondrian_coverage_std": mondrian_coverages.std(),
            }
        )

    pd.DataFrame(detail_rows).to_csv(DETAIL_PATH, index=False)
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(SUMMARY_PATH, index=False)
    print(summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
