"""
Week 9-10: standard (split) conformal prediction on the surrogate's residuals.

Uses the surrogate models already trained in Week 6-7 (models/surrogate_*.joblib)
as point predictors - CP wraps a point predictor with a calibrated interval,
it doesn't replace it. Calibration uses the separate set from
generate_calibration_data.py, never the surrogate's own training data.

Same test set (same random_state=42 split of surrogate_training_data.parquet)
and same alpha=0.1 as gp_baseline.py, so all three methods (GP, standard CP,
Mondrian CP later) are directly comparable.

Two nonconformity measures, per the roadmap's "test multiple nonconformity
measures":
  - symmetric: |y - yhat|, giving a fixed-width interval around the point
    prediction.
  - asymmetric: separate upper/lower residual quantiles (each calibrated at
    1-alpha/2, combined via a union bound for overall >=1-alpha coverage).
    Matters here because several targets (p95_wait_minutes especially) are
    right-skewed - a symmetric interval either overcovers on the left or
    undercovers on the right.
"""

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/standard_cp_metrics.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1  # must match gp_baseline.py


def conformal_quantile(scores, alpha):
    """Finite-sample-corrected (1-alpha) empirical quantile of calibration scores."""
    n = len(scores)
    level = min(np.ceil((n + 1) * (1 - alpha)) / n, 1.0)
    return np.quantile(scores, level, method="higher")


def main():
    df = pd.read_parquet(TRAIN_DATA_PATH)
    _, X_test, _, test_idx = train_test_split(df[FEATURES], df.index, test_size=0.2, random_state=42)

    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    X_cal = cal_df[FEATURES]

    metrics = []
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")

        yhat_cal = model.predict(X_cal)
        y_cal = cal_df[target].values
        residual_cal = y_cal - yhat_cal

        yhat_test = model.predict(X_test)
        y_test = df.loc[test_idx, target].values

        # symmetric
        q = conformal_quantile(np.abs(residual_cal), ALPHA)
        lower_sym = yhat_test - q
        upper_sym = yhat_test + q
        covered_sym = (y_test >= lower_sym) & (y_test <= upper_sym)

        # asymmetric
        q_hi = conformal_quantile(residual_cal, ALPHA / 2)
        q_lo = conformal_quantile(-residual_cal, ALPHA / 2)
        lower_asym = yhat_test - q_lo
        upper_asym = yhat_test + q_hi
        covered_asym = (y_test >= lower_asym) & (y_test <= upper_asym)

        metrics.append(
            {
                "target": target,
                "alpha": ALPHA,
                "target_coverage": 1 - ALPHA,
                "symmetric_coverage": covered_sym.mean(),
                "symmetric_mean_width": (upper_sym - lower_sym).mean(),
                "asymmetric_coverage": covered_asym.mean(),
                "asymmetric_mean_width": (upper_asym - lower_asym).mean(),
            }
        )

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    print(metrics_df)


if __name__ == "__main__":
    main()
