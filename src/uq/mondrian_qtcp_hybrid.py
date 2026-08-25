"""
Mondrian + QT-CP hybrid: unlike SA-Mondrian CP (which replaces Mondrian's
2D staffing x arrival-rate grid with a 1D rho_hat partition, and loses to
it - see repeated_worst_category_evaluation.py), this keeps Mondrian's
existing 9-cell grid intact and asks a narrower question: even within one
Mondrian cell, rho_hat still varies - does normalizing residuals by QT-CP's
continuous sigma_hat(x) *before* calibrating each cell's own quantile
recover any of the residual within-cell heterogeneity Mondrian's flat
per-cell quantile misses?

rho_cap is selected once, globally, from calibration data only (same
nested 70/30 split QT-CP itself uses) - not re-selected per cell, since
each cell has too few calibration points (~130) to split further without
making the cap selection itself noisy.
"""
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
import joblib

from src.des.er_simulation import STUDY_PERIOD_DAYS
from src.uq.mondrian_cp import assign_categories, conformal_quantile
from src.uq.queueing_weighted_cp import estimate_rho, mean_service_time_minutes, sigma_hat, select_rho_cap

TABLES = "results/tables"
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1


def main():
    arrivals = pd.read_csv(f"{TABLES}/arrivals_by_hour_bin.csv", index_col=0)["count"]
    esi_mix = pd.read_csv(f"{TABLES}/esi_mix.csv", index_col=0)["proportion"].to_dict()
    baseline_visits_per_day = arrivals.sum() / STUDY_PERIOD_DAYS
    mean_service_time = mean_service_time_minutes(esi_mix)

    df = pd.read_parquet("data/processed/surrogate_training_data.parquet")
    _, X_test, _, test_idx = train_test_split(df[["n_capacity", "arrival_rate_multiplier"]], df.index,
                                               test_size=0.2, random_state=42)
    test_df = df.loc[test_idx]
    cal_df = pd.read_parquet("data/processed/cp_calibration_data.parquet")

    rho_cal = estimate_rho(cal_df["n_capacity"], cal_df["arrival_rate_multiplier"],
                            baseline_visits_per_day, mean_service_time).values
    rho_test = estimate_rho(test_df["n_capacity"], test_df["arrival_rate_multiplier"],
                             baseline_visits_per_day, mean_service_time).values

    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    test_cat = assign_categories(test_df, cap_bounds, arr_bounds)
    categories = sorted(set(cal_cat))

    rows = []
    for target in TARGETS:
        model = joblib.load(f"models/surrogate_{target}.joblib")
        yhat_cal = model.predict(cal_df[["n_capacity", "arrival_rate_multiplier"]])
        resid_cal = np.abs(cal_df[target].values - yhat_cal)
        yhat_test = model.predict(X_test)
        resid_test = np.abs(test_df[target].values - yhat_test)

        selected_cap = select_rho_cap(resid_cal, rho_cal)
        sigma_cal = sigma_hat(rho_cal, selected_cap)
        sigma_test = sigma_hat(rho_test, selected_cap)

        hybrid_coverages, mondrian_coverages = [], []
        for cat in categories:
            cal_mask = cal_cat == cat
            test_mask = test_cat == cat
            if cal_mask.sum() < 5 or test_mask.sum() == 0:
                continue

            # pure Mondrian (for a same-split sanity check against known numbers)
            q_flat = conformal_quantile(resid_cal[cal_mask], ALPHA)
            mondrian_coverages.append((resid_test[test_mask] <= q_flat).mean())

            # hybrid: normalize within-cell, calibrate on normalized scores
            q_norm = conformal_quantile(resid_cal[cal_mask] / sigma_cal[cal_mask], ALPHA)
            hybrid_covered = resid_test[test_mask] <= q_norm * sigma_test[test_mask]
            hybrid_coverages.append(hybrid_covered.mean())

        rows.append({
            "target": target, "selected_rho_cap": selected_cap,
            "mondrian_worst": min(mondrian_coverages),
            "hybrid_worst": min(hybrid_coverages),
        })
        print(f"{target:20s} cap={selected_cap}  mondrian_worst={min(mondrian_coverages):.3f}  "
              f"hybrid_worst={min(hybrid_coverages):.3f}  "
              f"{'HYBRID BETTER' if min(hybrid_coverages) > min(mondrian_coverages) else 'MONDRIAN BETTER'}")

    out = pd.DataFrame(rows)
    out.to_csv(f"{TABLES}/mondrian_qtcp_hybrid.csv", index=False)
    print(f"\nSaved {TABLES}/mondrian_qtcp_hybrid.csv")


if __name__ == "__main__":
    main()
