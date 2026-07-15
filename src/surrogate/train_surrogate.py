"""
Week 6-7: train a surrogate model on the DES scenario data.

Gradient boosting (sklearn's HistGradientBoostingRegressor), not a neural
net — this is a 2-input tabular regression problem with ~5000 rows, where
gradient boosting is the standard strong baseline and needs no architecture
tuning or scaling. A NN would add complexity without a corresponding benefit
for this problem size/shape.

One model per target metric (n_patients, mean_wait_minutes,
mean_total_minutes, p95_wait_minutes) trained independently — simpler than a
true multi-output model and fine here since sklearn doesn't support
multi-output HGBR natively.
"""

import joblib
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split

DATA_PATH = "data/processed/surrogate_training_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/surrogate_metrics.csv"

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
