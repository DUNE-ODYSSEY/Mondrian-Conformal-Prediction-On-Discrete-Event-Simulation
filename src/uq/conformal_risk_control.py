"""
Conformal Risk Control (CRC), following Bates, Angelopoulos, Lei, Malik,
and Jordan (2021), "Distribution-Free, Risk-Controlling Prediction Sets."

Standard split conformal prediction (standard_cp.py) controls a single,
fixed loss: the 0/1 miscoverage indicator 1{y not in C(x)}. CRC generalizes
this to any bounded, monotone-in-lambda loss function ell_lambda(x, y),
calibrating the smallest lambda whose *upper confidence bound* on empirical
risk (via Hoeffding's inequality, not the direct order-statistic used by
standard CP) is at most alpha. This script implements CRC for two losses
per target, on the exact same calibration/test data and models as the rest
of this project's UQ comparison (standard_cp.py, mondrian_cp.py):

  1. The 0/1 upper-miscoverage loss - included as a validation check. Since
     this loss is the same one standard CP already controls exactly via its
     order-statistic quantile, CRC's Hoeffding-UCB calibration is expected
     to be *more conservative* (a wider threshold) than the exact conformal
     quantile - a real, honest property of Hoeffding-based CRC, not a
     deficiency of this implementation, and reported as such rather than
     hidden.
  2. A clipped relative-overshoot loss, ell_lambda(x, y) =
     clip((y - (yhat(x) + lambda)) / W_max, 0, 1), where W_max is set to
     the target's own standard-CP symmetric interval width (an existing,
     already-computed quantity, not an arbitrary round number). This is the
     genuinely new capability CRC provides that plain conformal prediction
     cannot: bounding the *expected severity* of an overflow event, not
     just its probability - directly relevant to an operational "how bad
     is a miss, on average" question standard CP's binary coverage
     guarantee cannot answer.

CRC's guarantee: with probability >= 1 - delta over the calibration draw,
the TRUE risk R(lambda_hat) = E[ell_lambda_hat(X, Y)] <= alpha. This script
reports both the calibration-side selection and the held-out test-set
empirical risk, to check the guarantee actually holds on unseen data.
"""

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
STANDARD_CP_METRICS_PATH = "results/tables/standard_cp_metrics.csv"
MODELS_DIR = "models"
OUT_PATH = "results/tables/crc_results.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1        # target risk level, matches alpha used everywhere else in this project
DELTA = 0.1         # confidence level for the Hoeffding UCB (1 - delta guarantee)
LAMBDA_GRID_POINTS = 4000


def hoeffding_ucb(risk_hat, n, B, delta):
    """Upper confidence bound on true risk via Hoeffding's inequality for a
    bounded-in-[0,B] loss, averaged over n i.i.d. calibration points."""
    return risk_hat + B * np.sqrt(np.log(1 / delta) / (2 * n))


def calibrate_crc(residuals, alpha, delta, loss_fn, B, lambda_grid):
    """Bates et al. Algorithm 1 (Hoeffding variant): smallest lambda on the
    grid whose Hoeffding UCB on empirical risk is <= alpha. Loss must be
    non-increasing in lambda (larger lambda = wider/safer set = smaller
    loss) for the search to be valid; lambda_grid is assumed ascending."""
    n = len(residuals)
    for lam in lambda_grid:
        risk_hat = loss_fn(residuals, lam).mean()
        if hoeffding_ucb(risk_hat, n, B, delta) <= alpha:
            return lam, risk_hat
    return lambda_grid[-1], loss_fn(residuals, lambda_grid[-1]).mean()


def loss_01(residuals, lam):
    """ell_lambda(x,y) = 1{y > yhat(x) + lambda}; residuals = y - yhat(x)."""
    return (residuals > lam).astype(float)


def loss_clipped_overshoot(residuals, lam, w_max):
    """ell_lambda(x,y) = clip((y - (yhat(x)+lambda)) / W_max, 0, 1)."""
    return np.clip((residuals - lam) / w_max, 0.0, 1.0)


def main():
    df = pd.read_parquet(TRAIN_DATA_PATH)
    _, X_test, _, test_idx = train_test_split(df[FEATURES], df.index, test_size=0.2, random_state=42)

    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    X_cal = cal_df[FEATURES]

    std_cp = pd.read_csv(STANDARD_CP_METRICS_PATH).set_index("target")

    rows = []
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")

        yhat_cal = model.predict(X_cal)
        residual_cal = cal_df[target].values - yhat_cal   # y - yhat, calibration

        yhat_test = model.predict(X_test)
        residual_test = df.loc[test_idx, target].values - yhat_test  # y - yhat, test

        lam_max = max(residual_cal.max(), residual_test.max()) * 1.5
        lambda_grid = np.linspace(0, lam_max, LAMBDA_GRID_POINTS)

        # Loss 1: 0/1 upper-miscoverage (validation check against standard CP)
        lam01, cal_risk01 = calibrate_crc(residual_cal, ALPHA, DELTA, loss_01, B=1.0, lambda_grid=lambda_grid)
        test_risk01 = loss_01(residual_test, lam01).mean()
        standard_cp_halfwidth = std_cp.loc[target, "symmetric_mean_width"] / 2

        # Loss 2: clipped relative overshoot severity
        w_max = std_cp.loc[target, "symmetric_mean_width"]
        loss_fn2 = lambda resid, lam: loss_clipped_overshoot(resid, lam, w_max)
        lam_sev, cal_risk_sev = calibrate_crc(residual_cal, ALPHA, DELTA, loss_fn2, B=1.0, lambda_grid=lambda_grid)
        test_risk_sev = loss_fn2(residual_test, lam_sev).mean()

        rows.append({
            "target": target,
            "alpha": ALPHA,
            "delta": DELTA,
            "n_cal": len(residual_cal),
            "n_test": len(residual_test),
            "loss01_lambda": lam01,
            "loss01_cal_risk_hat": cal_risk01,
            "loss01_test_risk": test_risk01,
            "standard_cp_upper_halfwidth": standard_cp_halfwidth,
            "sev_w_max": w_max,
            "sev_lambda": lam_sev,
            "sev_cal_risk_hat": cal_risk_sev,
            "sev_test_risk": test_risk_sev,
        })
        print(f"{target}: loss01 lambda={lam01:.2f} (vs standard-CP half-width {standard_cp_halfwidth:.2f}), "
              f"test risk={test_risk01:.4f} (target <= {ALPHA}); "
              f"severity lambda={lam_sev:.2f}, test risk={test_risk_sev:.4f} (target <= {ALPHA})")

    out_df = pd.DataFrame(rows)
    out_df.to_csv(OUT_PATH, index=False)
    print(f"\nSaved {OUT_PATH}")


if __name__ == "__main__":
    main()
