"""
Adaptive Conformal Inference (ACI), following Gibbs and Candes (2021),
"Adaptive Conformal Inference Under Distribution Shift."

Standard split conformal prediction (standard_cp.py) fixes its miscoverage
target alpha and its calibration quantile once, then applies both
unchanged to every future test point - which is exactly why coverage
collapses under the exchangeability stress test (exchangeability_stress_
test.py) once the test distribution moves away from the fixed calibration
distribution: there is no mechanism in standard CP for the interval to
respond to what it is actually observing. ACI removes that rigidity: it
maintains a running miscoverage target alpha_t, using only the calibration
set's *fixed* nonconformity-score distribution as a fixed lookup table, but
re-querying that lookup at a shifted level 1 - alpha_t and, after each new
point is revealed, nudging alpha_t up or down depending on whether that
point was covered:

    alpha_{t+1} = alpha_t + gamma * (alpha - err_t),   err_t = 1{y_t not in C_t}

This script builds a single, genuinely time-ordered test stream - a demand
surge ramping from within-range (0.8x) up to 3.0x the calibrated arrival
rate, the same severity progression as exchangeability_stress_test.py, but
concatenated into one sequence rather than evaluated as separate i.i.d.
batches - and compares ACI's windowed coverage against static standard
CP's on the identical stream, same model, same initial calibration set.
"""

import joblib
import numpy as np
import pandas as pd

from src.des.er_simulation import run_scenario

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
OUT_PATH = "results/tables/aci_results.csv"
OUT_TRAJECTORY_PATH = "results/tables/aci_alpha_trajectory.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1
GAMMA = 0.05
N_CAPACITY_RANGE = (15, 45)
N_PER_LEVEL = 60
# Same severity progression as exchangeability_stress_test.py, concatenated
# into a single ordered stream (a demand surge ramping up over "time")
SEVERITY_LEVELS = [0.8, 1.0, 1.3, 1.5, 1.8, 2.0, 2.5, 3.0]


def build_stream():
    rng = np.random.default_rng(300_000)
    rows = []
    step = 0
    for level_idx, mult in enumerate(SEVERITY_LEVELS):
        for i in range(N_PER_LEVEL):
            cap = int(rng.integers(N_CAPACITY_RANGE[0], N_CAPACITY_RANGE[1] + 1))
            result = run_scenario(n_capacity=cap, arrival_rate_multiplier=mult, seed=300_000 + level_idx * 10_000 + i)
            rows.append({"step": step, "arrival_rate_multiplier": mult, "in_range": mult <= 1.3,
                         "n_capacity": cap, **result})
            step += 1
    return pd.DataFrame(rows)


def quantile_at_level(scores, level):
    """Empirical quantile of a fixed score set at an arbitrary (possibly
    time-varying) level, clipped to a valid probability - this is ACI's
    'fixed lookup table, shifted query level' mechanism."""
    level = float(np.clip(level, 0.0, 1.0))
    if level >= 1.0:
        return np.max(scores)
    if level <= 0.0:
        return np.min(np.abs(scores)) * 0  # degenerate, zero-width
    n = len(scores)
    order_level = min(np.ceil((n + 1) * level) / n, 1.0)
    return np.quantile(scores, order_level, method="higher")


def run_aci_vs_static(residual_cal_abs, residual_stream, alpha, gamma):
    """Returns per-step covered_static, covered_aci, alpha_t trajectory,
    interval half-width trajectory for ACI."""
    n = len(residual_stream)
    covered_static = np.zeros(n, dtype=bool)
    covered_aci = np.zeros(n, dtype=bool)
    alpha_t_trace = np.zeros(n)
    width_aci_trace = np.zeros(n)

    q_static = quantile_at_level(residual_cal_abs, 1 - alpha)

    alpha_t = alpha
    for t in range(n):
        covered_static[t] = residual_stream[t] <= q_static

        alpha_t_trace[t] = alpha_t
        q_aci = quantile_at_level(residual_cal_abs, 1 - alpha_t)
        width_aci_trace[t] = q_aci
        covered_aci[t] = residual_stream[t] <= q_aci

        err_t = 0.0 if covered_aci[t] else 1.0
        alpha_t = alpha_t + gamma * (alpha - err_t)

    return covered_static, covered_aci, alpha_t_trace, width_aci_trace


def rolling_coverage(covered, window=60):
    s = pd.Series(covered.astype(float))
    return s.rolling(window, min_periods=1).mean().values


def main():
    stream_df = build_stream()

    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    X_cal = cal_df[FEATURES]

    summary_rows = []
    trajectory_rows = []
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")

        yhat_cal = model.predict(X_cal)
        residual_cal_abs = np.abs(cal_df[target].values - yhat_cal)

        yhat_stream = model.predict(stream_df[FEATURES])
        residual_stream_abs = np.abs(stream_df[target].values - yhat_stream)

        covered_static, covered_aci, alpha_t_trace, width_aci_trace = run_aci_vs_static(
            residual_cal_abs, residual_stream_abs, ALPHA, GAMMA
        )

        roll_static = rolling_coverage(covered_static)
        roll_aci = rolling_coverage(covered_aci)

        for t in range(len(stream_df)):
            trajectory_rows.append({
                "target": target, "step": t,
                "arrival_rate_multiplier": stream_df["arrival_rate_multiplier"].iloc[t],
                "in_range": stream_df["in_range"].iloc[t],
                "alpha_t": alpha_t_trace[t],
                "aci_width": 2 * width_aci_trace[t],
                "rolling_coverage_static": roll_static[t],
                "rolling_coverage_aci": roll_aci[t],
            })

        # overall + out-of-range-only coverage
        oor_mask = ~stream_df["in_range"].values
        summary_rows.append({
            "target": target,
            "overall_coverage_static": covered_static.mean(),
            "overall_coverage_aci": covered_aci.mean(),
            "oor_coverage_static": covered_static[oor_mask].mean(),
            "oor_coverage_aci": covered_aci[oor_mask].mean(),
            "final_alpha_t": alpha_t_trace[-1],
            "mean_width_aci_oor": (2 * width_aci_trace)[oor_mask].mean(),
        })
        print(f"{target}: overall static={covered_static.mean():.3f} aci={covered_aci.mean():.3f} | "
              f"out-of-range static={covered_static[oor_mask].mean():.3f} aci={covered_aci[oor_mask].mean():.3f} "
              f"(target 0.90)")

    pd.DataFrame(summary_rows).to_csv(OUT_PATH, index=False)
    pd.DataFrame(trajectory_rows).to_csv(OUT_TRAJECTORY_PATH, index=False)
    print(f"\nSaved {OUT_PATH} and {OUT_TRAJECTORY_PATH}")


if __name__ == "__main__":
    main()
