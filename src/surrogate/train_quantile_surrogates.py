"""
Publication-rigor upgrade: train quantile-regression surrogates for CQR
(Conformalized Quantile Regression), a stronger CP baseline than symmetric/
asymmetric split CP for skewed targets (p95_wait_minutes especially).

CQR calibrates a nonconformity score of max(qlo(x)-y, y-qhi(x)) instead of
|y-yhat| - the interval adapts to the quantile regressors' own estimate of
local spread, rather than using one fixed width (symmetric CP) or two fixed
offsets (asymmetric CP) everywhere.

Same train/test split (random_state=42) as train_surrogate.py, so the
quantile regressors see exactly the same training data as the original mean
regressor - a fair comparison of CP methods, not a different setup dressed
up as one. Trained once, like the mean surrogate; only calibration/test data
varies across the repeated-evaluation runs in repeated_evaluation.py.
"""

import joblib
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import mean_pinball_loss
from sklearn.model_selection import train_test_split

DATA_PATH = "data/processed/surrogate_training_data.parquet"
MODELS_DIR = "models"
METRICS_PATH = "results/tables/quantile_surrogate_metrics.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1
Q_LO, Q_HI = ALPHA / 2, 1 - ALPHA / 2


def main():
    df = pd.read_parquet(DATA_PATH)
    X_train, X_test, train_idx, test_idx = train_test_split(
        df[FEATURES], df.index, test_size=0.2, random_state=42
    )

    metrics = []
    for target in TARGETS:
        y_train = df.loc[train_idx, target]
        y_test = df.loc[test_idx, target]

        for label, q in [("qlo", Q_LO), ("qhi", Q_HI)]:
            model = HistGradientBoostingRegressor(loss="quantile", quantile=q, random_state=42)
            model.fit(X_train, y_train)
            y_pred = model.predict(X_test)
            pinball = mean_pinball_loss(y_test, y_pred, alpha=q)
            metrics.append({"target": target, "quantile_label": label, "quantile": q, "pinball_loss": pinball})
            joblib.dump(model, f"{MODELS_DIR}/surrogate_{target}_{label}.joblib")

        # sanity check: how often does the raw (uncalibrated) [qlo, qhi] interval cover y_test?
        qlo_model = joblib.load(f"{MODELS_DIR}/surrogate_{target}_qlo.joblib")
        qhi_model = joblib.load(f"{MODELS_DIR}/surrogate_{target}_qhi.joblib")
        qlo_pred = qlo_model.predict(X_test)
        qhi_pred = qhi_model.predict(X_test)
        raw_coverage = ((y_test >= qlo_pred) & (y_test <= qhi_pred)).mean()
        metrics.append({"target": target, "quantile_label": "raw_interval_coverage",
                         "quantile": None, "pinball_loss": raw_coverage})

    metrics_df = pd.DataFrame(metrics)
    metrics_df.to_csv(METRICS_PATH, index=False)
    print(metrics_df.to_string(index=False))


if __name__ == "__main__":
    main()
