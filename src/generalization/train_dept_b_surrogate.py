"""
Publication-rigor upgrade, part 4 (continued): train Department B's
surrogate. Same architecture and train/test split logic as
train_surrogate.py (gradient boosting, random_state=42, 80/20 split) - only
the data source and output paths differ, so this is a fair architecture
comparison, not a differently-tuned model.
"""

import joblib
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split

DATA_PATH = "data/processed/dept_b/surrogate_training_data.parquet"
MODELS_DIR = "models/dept_b"
METRICS_PATH = "results/tables/dept_b/surrogate_metrics.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]


def main():
    df = pd.read_parquet(DATA_PATH)
    X_train, X_test, train_idx, test_idx = train_test_split(
        df[FEATURES], df.index, test_size=0.2, random_state=42
    )

    metrics = []
    for target in TARGETS:
        y_train = df.loc[train_idx, target]
        y_test = df.loc[test_idx, target]

        model = HistGradientBoostingRegressor(random_state=42)
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)

        metrics.append(
            {
                "target": target,
                "mae": mean_absolute_error(y_test, y_pred),
                "rmse": mean_squared_error(y_test, y_pred) ** 0.5,
                "r2": r2_score(y_test, y_pred),
            }
        )
        joblib.dump(model, f"{MODELS_DIR}/surrogate_{target}.joblib")

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    print(metrics_df)


if __name__ == "__main__":
    main()
