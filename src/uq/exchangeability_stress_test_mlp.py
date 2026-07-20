"""
Publication-rigor upgrade, part 3 (continued): re-run the Week 13
exchangeability stress test using the MLP surrogate instead of gradient
boosting, to test whether the coverage collapse outside the training range
is specific to tree-based extrapolation (frozen leaf predictions, verified
in Week 13) or a more general property of wrapping CP around an
unreliable-outside-its-domain point predictor.

Identical setup to exchangeability_stress_test.py otherwise: same
calibration data, same severity levels, same seeds - only the model files
differ (`*_mlp.joblib`). Calibration is recomputed fresh here (not reused
from the gradient-boosting run) because the MLP's residuals on the
calibration set are different from gradient boosting's - each surrogate
needs its own conformal quantile.
"""

import joblib
import numpy as np
import pandas as pd

from src.des.er_simulation import run_scenario
from src.uq.mondrian_cp import assign_categories, conformal_quantile

CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
OUT_PATH = "results/tables/exchangeability_stress_test_mlp.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1
N_CAPACITY_RANGE = (15, 45)
N_PER_SEVERITY = 300
SEVERITY_LEVELS = [0.8, 1.0, 1.3, 1.5, 1.8, 2.0, 2.5, 3.0]


def generate_ood_data(arrival_multiplier, n_scenarios, seed_offset):
    rng = np.random.default_rng(seed_offset)
    n_capacity = rng.integers(N_CAPACITY_RANGE[0], N_CAPACITY_RANGE[1] + 1, size=n_scenarios)
    rows = []
    for i in range(n_scenarios):
        result = run_scenario(
            n_capacity=int(n_capacity[i]),
            arrival_rate_multiplier=arrival_multiplier,
            seed=seed_offset + i,
        )
        rows.append({"n_capacity": n_capacity[i], "arrival_rate_multiplier": arrival_multiplier, **result})
    return pd.DataFrame(rows)


def main():
    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)

    rows = []
    for level_idx, mult in enumerate(SEVERITY_LEVELS):
        # same seed_offset scheme as the gradient-boosting version -> identical OOD scenarios
        ood_df = generate_ood_data(mult, N_PER_SEVERITY, seed_offset=200_000 + level_idx * 10_000)
        ood_cat = assign_categories(ood_df, cap_bounds, arr_bounds)

        for target in TARGETS:
            model = joblib.load(f"{MODELS_DIR}/surrogate_{target}_mlp.joblib")

            yhat_cal = model.predict(cal_df[FEATURES])
            y_cal = cal_df[target].values
            abs_resid_cal = np.abs(y_cal - yhat_cal)
            q_pooled = conformal_quantile(abs_resid_cal, ALPHA)

            yhat_ood = model.predict(ood_df[FEATURES])
            y_ood = ood_df[target].values
            abs_resid_ood = np.abs(y_ood - yhat_ood)

            standard_covered = abs_resid_ood <= q_pooled

            mondrian_covered = np.zeros(len(ood_df), dtype=bool)
            mondrian_width = np.zeros(len(ood_df))
            for cat in set(cal_cat):
                q_cat = conformal_quantile(abs_resid_cal[cal_cat == cat], ALPHA)
                mask = ood_cat == cat
                mondrian_covered[mask] = abs_resid_ood[mask] <= q_cat
                mondrian_width[mask] = 2 * q_cat

            rows.append(
                {
                    "target": target,
                    "arrival_rate_multiplier": mult,
                    "in_training_range": mult <= 1.3,
                    "n_scenarios": len(ood_df),
                    "standard_coverage": standard_covered.mean(),
                    "standard_width": 2 * q_pooled,
                    "mondrian_coverage": mondrian_covered.mean(),
                    "mondrian_mean_width": mondrian_width.mean(),
                    "yhat_capacity30": float(model.predict(pd.DataFrame(
                        [[30, mult]], columns=FEATURES))[0]),
                }
            )

    out_df = pd.DataFrame(rows)
    out_df.to_csv(OUT_PATH, index=False)
    print(out_df.to_string(index=False))


if __name__ == "__main__":
    main()
