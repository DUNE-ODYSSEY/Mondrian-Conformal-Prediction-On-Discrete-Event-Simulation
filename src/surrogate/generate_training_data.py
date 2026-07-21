"""
Week 6-7: run the calibrated DES across varied staffing/arrival scenarios to
generate surrogate training data.

Each row = one independent simulated day at a randomly sampled
(n_capacity, arrival_rate_multiplier), with its own random seed so the DES's
intrinsic stochasticity (arrival randomness, ESI mix, service-time variance)
stays in the labels. That residual variability is exactly what the later UQ
step (GP baseline / conformal prediction) needs to characterize — smoothing
it out here would defeat the point.

n_capacity range (15-45) and arrival_rate_multiplier range (0.8-1.3) span
under- to over-provisioned staffing and below/above-average demand days,
while staying inside "normal operations" — the Week 13 surge-day stress
test intentionally goes outside this range later to test exchangeability.
"""

import numpy as np
import pandas as pd

from src.des.er_simulation import run_scenario

N_SCENARIOS = 5000
N_CAPACITY_RANGE = (15, 45)
ARRIVAL_MULTIPLIER_RANGE = (0.8, 1.3)
OUT_PATH = "data/processed/surrogate_training_data.parquet"


def generate(n_scenarios=N_SCENARIOS, seed=0, seed_offset=0,
             n_capacity_range=N_CAPACITY_RANGE, arrival_multiplier_range=ARRIVAL_MULTIPLIER_RANGE,
             tables_dir=None):
    """seed_offset shifts the per-scenario DES seeds so a second call (e.g. for
    CP calibration data) draws fully independent randomness from the first,
    on top of already using a different `seed` for the parameter sampling.

    n_capacity_range/arrival_multiplier_range/tables_dir default to department
    A's values (module constants / er_simulation.py's own default) so existing
    callers are unaffected; pass different values to calibrate against a
    different department (see src/generalization/)."""
    rng = np.random.default_rng(seed)
    n_capacity = rng.integers(n_capacity_range[0], n_capacity_range[1] + 1, size=n_scenarios)
    arrival_mult = rng.uniform(*arrival_multiplier_range, size=n_scenarios)

    run_kwargs = {} if tables_dir is None else {"tables_dir": tables_dir}
    rows = []
    for i in range(n_scenarios):
        result = run_scenario(
            n_capacity=int(n_capacity[i]),
            arrival_rate_multiplier=float(arrival_mult[i]),
            seed=seed_offset + i,
            **run_kwargs,
        )
        rows.append(
            {
                "n_capacity": n_capacity[i],
                "arrival_rate_multiplier": arrival_mult[i],
                **result,
            }
        )
    return pd.DataFrame(rows)


def main():
    df = generate()
    df.to_parquet(OUT_PATH, index=False)
    print(f"Generated {len(df)} scenarios -> {OUT_PATH}")
    print(df.describe())


if __name__ == "__main__":
    main()
