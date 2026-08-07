"""
Multi-architecture surrogate benchmark: Random Forest, XGBoost, and
LightGBM, trained on the identical data and train/test split as the
primary gradient-boosting surrogate (train_surrogate.py, random_state=42)
and the MLP robustness check (train_mlp_surrogate.py), so all five
architectures are directly comparable on the same held-out test set -
a fair architecture comparison, not a differently-trained one.

Adds three tree-ensemble variants beyond the primary HistGradientBoosting
model to check whether its accuracy is specific to that one implementation
or general to gradient-boosted/bagged tree ensembles on this problem.
Default hyperparameters throughout (matching train_surrogate.py's own
practice of not hand-tuning a 2-input, ~5000-row tabular problem) - the
comparison is about architecture family, not a hyperparameter search.
"""

import time

import joblib
import pandas as pd
from lightgbm import LGBMRegressor
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split
from xgboost import XGBRegressor

DATA_PATH = "data/processed/surrogate_training_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/rf_xgb_lgb_surrogate_metrics.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ARCHITECTURES = {
    "RandomForest": lambda: RandomForestRegressor(random_state=42, n_jobs=-1),
    "XGBoost": lambda: XGBRegressor(random_state=42, n_jobs=-1),
    "LightGBM": lambda: LGBMRegressor(random_state=42, n_jobs=-1, verbosity=-1),
}


def main():
    df = pd.read_parquet(DATA_PATH)
    X_train, X_test, train_idx, test_idx = train_test_split(
        df[FEATURES], df.index, test_size=0.2, random_state=42
    )

    metrics = []
    for arch_name, make_model in ARCHITECTURES.items():
        for target in TARGETS:
            y_train = df.loc[train_idx, target]
            y_test = df.loc[test_idx, target]

            model = make_model()
            t0 = time.perf_counter()
            model.fit(X_train, y_train)
            fit_seconds = time.perf_counter() - t0

            t0 = time.perf_counter()
            y_pred = model.predict(X_test)
            predict_seconds = time.perf_counter() - t0

            metrics.append({
                "architecture": arch_name,
                "target": target,
                "mae": mean_absolute_error(y_test, y_pred),
                "rmse": mean_squared_error(y_test, y_pred) ** 0.5,
                "r2": r2_score(y_test, y_pred),
                "fit_seconds": fit_seconds,
                "predict_seconds": predict_seconds,
            })
            joblib.dump(model, f"{MODELS_DIR}/surrogate_{target}_{arch_name.lower()}.joblib")
            print(f"{arch_name} / {target}: MAE={metrics[-1]['mae']:.2f} RMSE={metrics[-1]['rmse']:.2f} "
                  f"R2={metrics[-1]['r2']:.3f} fit={fit_seconds:.2f}s")

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    print(f"\nSaved {METRICS_PATH}")


if __name__ == "__main__":
    main()
