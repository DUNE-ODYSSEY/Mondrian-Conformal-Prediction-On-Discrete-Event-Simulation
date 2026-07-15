"""
Week 8: GP baseline UQ — the benchmark that standard CP and Mondrian CP get
compared against in Phase 2.

Fits a Gaussian Process per target on (n_capacity, arrival_rate_multiplier)
and uses its native predictive mean +/- z * std to build (1-alpha) intervals.
alpha=0.1 (90% target coverage) is fixed here and must stay the same alpha
used later for standard/Mondrian CP — the whole comparison is meaningless if
the methods target different coverage levels.

GP training is fit on a 1000-point subsample of the training set, not the
full ~4000, because sklearn's exact GP is O(n^3): benchmarked at ~8s for
n=1000 vs ~43s for n=2000, so full-scale training would take tens of minutes
across 4 targets. This is not just a shortcut — GP's poor scalability vs.
CP's near-constant-cost calibration is itself one of the points this project
is comparing (see Week 14-15's computation-time chart), so documenting GP's
cost here is directly relevant, not a detail to hide.

Uses the SAME train/test split (random_state=42, test_size=0.2) as
train_surrogate.py so every UQ method (GP here, standard/Mondrian CP later)
is evaluated on identical test points — required for a fair comparison.
"""

import time

import joblib
import numpy as np
import pandas as pd
from scipy.stats import norm
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, ConstantKernel, WhiteKernel
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

DATA_PATH = "data/processed/surrogate_training_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/gp_baseline_metrics.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1  # 90% target coverage — must match the CP methods used later
N_GP_TRAIN = 1000  # subsample size for GP fitting; see module docstring
Z = norm.ppf(1 - ALPHA / 2)


def main():
    df = pd.read_parquet(DATA_PATH)
    X_train_full, X_test, train_idx, test_idx = train_test_split(
        df[FEATURES], df.index, test_size=0.2, random_state=42
    )

    rng = np.random.default_rng(42)
    gp_train_idx = rng.choice(train_idx, size=N_GP_TRAIN, replace=False)

    scaler = StandardScaler().fit(df.loc[gp_train_idx, FEATURES])
    X_train_scaled = scaler.transform(df.loc[gp_train_idx, FEATURES])
    X_test_scaled = scaler.transform(X_test)

    metrics = []
    for target in TARGETS:
        y_train = df.loc[gp_train_idx, target].values
        y_test = df.loc[test_idx, target].values

        kernel = ConstantKernel(1.0) * RBF(length_scale=[1.0, 1.0]) + WhiteKernel(1.0)
        gp = GaussianProcessRegressor(kernel=kernel, normalize_y=True, n_restarts_optimizer=2, random_state=42)

        t0 = time.time()
        gp.fit(X_train_scaled, y_train)
        fit_seconds = time.time() - t0

        t0 = time.time()
        y_pred, y_std = gp.predict(X_test_scaled, return_std=True)
        predict_seconds = time.time() - t0

        lower = y_pred - Z * y_std
        upper = y_pred + Z * y_std
        covered = (y_test >= lower) & (y_test <= upper)

        metrics.append(
            {
                "target": target,
                "alpha": ALPHA,
                "target_coverage": 1 - ALPHA,
                "empirical_coverage": covered.mean(),
                "mean_interval_width": (upper - lower).mean(),
                "fit_seconds": fit_seconds,
                "predict_seconds": predict_seconds,
                "n_gp_train": N_GP_TRAIN,
            }
        )
        joblib.dump(gp, f"{MODELS_DIR}/gp_baseline_{target}.joblib")

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    joblib.dump(scaler, f"{MODELS_DIR}/gp_baseline_scaler.joblib")
    print(metrics_df)


if __name__ == "__main__":
    main()
