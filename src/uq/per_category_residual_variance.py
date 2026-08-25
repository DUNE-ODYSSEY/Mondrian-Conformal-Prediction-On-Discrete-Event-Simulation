"""
Per-category calibration residual variance, all four targets - the real
table behind the Discussion section's variance-asymmetry claim ("n_patients'
per-category residual std spans only 1.8x across the 9 categories, while
wait-time targets span 4.1-23.9x"). That claim was originally checked with
an ad-hoc inline command and never saved as a script or table, unlike every
other number in this paper - this closes that gap so the claim is traceable
the same way Table II/IV/etc. are, not just asserted.

Uses the same single calibration/test split (seed=42) and 2D staffing x
arrival-rate tercile categories every other single-split comparison in this
project uses (mondrian_qtcp_hybrid.py, queueing_weighted_cp.py).
"""
import numpy as np
import pandas as pd
import joblib

from src.uq.mondrian_cp import assign_categories

TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
TABLES = "results/tables"
DETAIL_PATH = f"{TABLES}/per_category_residual_std_detail.csv"
SUMMARY_PATH = f"{TABLES}/per_category_residual_std_summary.csv"


def main():
    cal_df = pd.read_parquet("data/processed/cp_calibration_data.parquet")

    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)

    detail_rows, summary_rows = [], []
    for target in TARGETS:
        model = joblib.load(f"models/surrogate_{target}.joblib")
        yhat_cal = model.predict(cal_df[["n_capacity", "arrival_rate_multiplier"]])
        resid_cal = np.abs(cal_df[target].values - yhat_cal)

        stds = {}
        for cat in sorted(set(cal_cat)):
            mask = cal_cat == cat
            std = resid_cal[mask].std()
            stds[cat] = std
            detail_rows.append({"target": target, "category": cat, "n_cal": mask.sum(), "resid_std": std})

        vals = np.array(list(stds.values()))
        summary_rows.append({
            "target": target, "min_std": vals.min(), "max_std": vals.max(),
            "max_over_min_ratio": vals.max() / vals.min(),
        })
        print(f"{target:20s} min={vals.min():.2f} max={vals.max():.2f} ratio={vals.max()/vals.min():.2f}x")

    pd.DataFrame(detail_rows).to_csv(DETAIL_PATH, index=False)
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(SUMMARY_PATH, index=False)
    print(f"\nSaved {DETAIL_PATH}")
    print(f"Saved {SUMMARY_PATH}")


if __name__ == "__main__":
    main()
