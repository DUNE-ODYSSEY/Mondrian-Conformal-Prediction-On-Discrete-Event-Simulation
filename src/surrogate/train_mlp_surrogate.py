"""
Publication-rigor upgrade, part 3: a second surrogate architecture.

Week 13's exchangeability finding traced the coverage collapse to a specific
mechanism: HistGradientBoostingRegressor (tree-based) cannot extrapolate
past its training range, so its prediction literally freezes for any input
outside [0.8, 1.3]. That's a real, verified mechanism - but it's specific to
tree ensembles. A neural network doesn't have the same "frozen leaf value"
limitation; its predictions keep changing outside the training range
(usually roughly linearly in the final layer, though not necessarily
*correctly*). Training an MLP surrogate here to re-run the exchangeability
stress test against and see whether the coverage collapse is a property of
CP-wrapping-an-unreliable-model in general, or specific to gradient
boosting's exact failure mode.

Same train/test split (random_state=42) as train_surrogate.py, so this is
trained on identical data to the original surrogate - a fair architecture
comparison, not a differently-set-up experiment. MLPs are scale-sensitive
(unlike tree ensembles), so each target gets a Pipeline(StandardScaler,
MLPRegressor) - the scaler is fit only on training data and saved as part of
the pipeline, so calibration/test/stress-test data downstream doesn't need
its own separate scaling step.
"""

import joblib
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split
from sklearn.neural_network import MLPRegressor
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

DATA_PATH = "data/processed/surrogate_training_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/mlp_surrogate_metrics.csv"

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

        model = make_pipeline(
            StandardScaler(),
            MLPRegressor(
                hidden_layer_sizes=(64, 64),
                activation="relu",
                early_stopping=True,
                n_iter_no_change=20,
                max_iter=2000,
                random_state=42,
            ),
        )
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
        joblib.dump(model, f"{MODELS_DIR}/surrogate_{target}_mlp.joblib")

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    print(metrics_df)


if __name__ == "__main__":
    main()
