"""
Publication-rigor upgrade, part 4 (continued): the actual generalizability
question - does Department A's core finding (Mondrian CP closes a real
marginal-vs-conditional coverage gap in the worst staffing/arrival category)
replicate at a second, structurally different site?

Department B is not just a scaled-down copy of A: real volume is roughly
half (133.4 vs. 258.2 visits/day) AND the real ESI acuity mix is
meaningfully different (23.1% ESI-2 vs. A's 37.9%, 26.8% ESI-4 vs. A's
16.4% - B skews toward lower-acuity visits, consistent with a community ED
vs. A's academic site). So this is a genuine second data point, not a
trivial rescaling of the same result.

Same standard CP (symmetric) + Mondrian CP logic as standard_cp.py /
mondrian_cp.py, same alpha=0.1, applied to B's own surrogate/calibration/
test data. Single split here (not the 30-repeat treatment from part 1) -
proportionate scope for a generalizability check, not a full re-run of
every rigor dimension for a second site.
"""

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

from src.uq.mondrian_cp import assign_categories, conformal_quantile

TRAIN_DATA_PATH = "data/processed/dept_b/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/dept_b/cp_calibration_data.parquet"
MODELS_DIR = "models/dept_b"
DETAIL_PATH = "results/tables/dept_b/mondrian_cp_detail.csv"
SUMMARY_PATH = "results/tables/dept_b/standard_vs_mondrian_summary.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1


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
        abs_resid_test = np.abs(y_test - yhat_test)

        q_pooled = conformal_quantile(abs_resid_cal, ALPHA)
        standard_covered_all = abs_resid_test <= q_pooled

        pooled_coverages, mondrian_coverages = [], []
        for cat in categories:
            cal_mask = cal_cat == cat
            test_mask = test_cat == cat
            n_cal_cat = cal_mask.sum()
            n_test_cat = test_mask.sum()
            if n_test_cat == 0:
                continue

            q_cat = conformal_quantile(abs_resid_cal[cal_mask], ALPHA)
            covered_pooled = (abs_resid_test[test_mask] <= q_pooled).mean()
            covered_mondrian = (abs_resid_test[test_mask] <= q_cat).mean()

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

        mondrian_covered_all = np.zeros(len(test_df), dtype=bool)
        for cat in categories:
            test_mask = test_cat == cat
            q_cat = conformal_quantile(abs_resid_cal[cal_cat == cat], ALPHA)
            mondrian_covered_all[test_mask] = abs_resid_test[test_mask] <= q_cat

        summary_rows.append(
            {
                "target": target,
                "target_coverage": 1 - ALPHA,
                "standard_overall_coverage": standard_covered_all.mean(),
                "standard_overall_width": 2 * q_pooled,
                "mondrian_overall_coverage": mondrian_covered_all.mean(),
                "pooled_coverage_min": pooled_coverages.min(),
                "pooled_coverage_max": pooled_coverages.max(),
                "pooled_coverage_range": pooled_coverages.max() - pooled_coverages.min(),
                "mondrian_coverage_min": mondrian_coverages.min(),
                "mondrian_coverage_max": mondrian_coverages.max(),
                "mondrian_coverage_range": mondrian_coverages.max() - mondrian_coverages.min(),
                "worst_pooled_category": categories[int(np.argmin(pooled_coverages))],
            }
        )

    pd.DataFrame(detail_rows).to_csv(DETAIL_PATH, index=False)
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(SUMMARY_PATH, index=False)
    print(summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
