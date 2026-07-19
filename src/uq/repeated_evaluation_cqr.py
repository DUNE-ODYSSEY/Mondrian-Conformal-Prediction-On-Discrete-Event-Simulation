"""
Publication-rigor upgrade, part 2: CQR (Conformalized Quantile Regression)
and Mondrian-CQR, evaluated with the same 30-repeat rigor as
repeated_evaluation.py (GP / standard CP / Mondrian CP).

CQR's nonconformity score is max(qlo(x)-y, y-qhi(x)) instead of |y-yhat| -
the interval adapts to the quantile regressors' own estimate of local
spread (from train_quantile_surrogates.py), rather than a single fixed
width (standard CP) or fixed per-category width (Mondrian CP) built from a
mean predictor. Mondrian-CQR combines both ideas: per-category calibration
of an already-adaptive nonconformity score - the most sophisticated method
in the comparison.

Uses the exact same 30 (calibration, test) draws as repeated_evaluation.py
(identical seed formula: calibration sampling seed 1000+r / DES offset
500,000+r*10,000, test sampling seed 2000+r / DES offset 700,000+r*10,000).
generate() is deterministic given a seed, so this reproduces byte-identical
scenario data without needing to save/reload it - meaning CQR's results are
directly, validly paired with the GP/standard/Mondrian results from part 1
for significance testing, not a separately-drawn comparison.

No GP here, which is why this runs in ~15-20 min instead of ~82 min - GP
refitting was almost the entire cost of part 1, and CQR doesn't need it.
"""

import time

import joblib
import numpy as np
import pandas as pd

from src.surrogate.generate_training_data import generate
from src.uq.mondrian_cp import assign_categories, conformal_quantile

MODELS_DIR = "models"
DETAIL_PATH = "results/tables/repeated_evaluation_cqr_detail.csv"
SUMMARY_PATH = "results/tables/repeated_evaluation_cqr_summary.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1

N_REPEATS = 30
N_CALIBRATION = 1200
N_TEST = 1000


def run_repeats(n_repeats=N_REPEATS):
    qlo_models = {t: joblib.load(f"{MODELS_DIR}/surrogate_{t}_qlo.joblib") for t in TARGETS}
    qhi_models = {t: joblib.load(f"{MODELS_DIR}/surrogate_{t}_qhi.joblib") for t in TARGETS}

    rows = []
    t_start = time.time()

    for r in range(n_repeats):
        # identical seeds to repeated_evaluation.py -> identical (cal, test) draws
        cal_df = generate(n_scenarios=N_CALIBRATION, seed=1000 + r, seed_offset=500_000 + r * 10_000)
        test_df = generate(n_scenarios=N_TEST, seed=2000 + r, seed_offset=700_000 + r * 10_000)

        cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
        arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
        cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
        test_cat = assign_categories(test_df, cap_bounds, arr_bounds)

        for target in TARGETS:
            qlo_model, qhi_model = qlo_models[target], qhi_models[target]
            y_cal = cal_df[target].values
            y_test = test_df[target].values

            qlo_cal = qlo_model.predict(cal_df[FEATURES])
            qhi_cal = qhi_model.predict(cal_df[FEATURES])
            cqr_score_cal = np.maximum(qlo_cal - y_cal, y_cal - qhi_cal)

            qlo_test = qlo_model.predict(test_df[FEATURES])
            qhi_test = qhi_model.predict(test_df[FEATURES])

            # pooled CQR
            q_pooled = conformal_quantile(cqr_score_cal, ALPHA)
            lower = qlo_test - q_pooled
            upper = qhi_test + q_pooled
            covered = (y_test >= lower) & (y_test <= upper)
            rows.append({"repeat": r, "target": target, "method": "CQR",
                         "coverage": covered.mean(), "mean_width": (upper - lower).mean()})

            # Mondrian CQR (per-category calibration of the same adaptive score)
            mondrian_covered = np.zeros(len(test_df), dtype=bool)
            mondrian_widths = np.zeros(len(test_df))
            for cat in set(cal_cat):
                cal_mask = cal_cat == cat
                test_mask = test_cat == cat
                if test_mask.sum() == 0 or cal_mask.sum() == 0:
                    continue
                q_cat = conformal_quantile(cqr_score_cal[cal_mask], ALPHA)
                lo_c = qlo_test[test_mask] - q_cat
                hi_c = qhi_test[test_mask] + q_cat
                mondrian_covered[test_mask] = (y_test[test_mask] >= lo_c) & (y_test[test_mask] <= hi_c)
                mondrian_widths[test_mask] = hi_c - lo_c
            rows.append({"repeat": r, "target": target, "method": "Mondrian CQR",
                         "coverage": mondrian_covered.mean(), "mean_width": mondrian_widths.mean()})

        elapsed = time.time() - t_start
        print(f"repeat {r + 1}/{n_repeats} done ({elapsed:.0f}s elapsed)")

    return pd.DataFrame(rows)


def summarize(df):
    def ci95(x):
        n = len(x)
        return 1.96 * x.std(ddof=1) / np.sqrt(n)

    return df.groupby(["target", "method"]).agg(
        n_repeats=("coverage", "size"),
        coverage_mean=("coverage", "mean"),
        coverage_std=("coverage", "std"),
        coverage_ci95=("coverage", ci95),
        width_mean=("mean_width", "mean"),
        width_std=("mean_width", "std"),
        width_ci95=("mean_width", ci95),
    ).reset_index()


def main():
    detail = run_repeats()
    detail.to_csv(DETAIL_PATH, index=False)
    summary = summarize(detail)
    summary.to_csv(SUMMARY_PATH, index=False)
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
