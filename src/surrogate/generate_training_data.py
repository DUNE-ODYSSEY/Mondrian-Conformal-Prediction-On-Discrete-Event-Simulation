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


def generate(n_scenarios=N_SCENARIOS, seed=0):
    rng = np.random.default_rng(seed)
    n_capacity = rng.integers(N_CAPACITY_RANGE[0], N_CAPACITY_RANGE[1] + 1, size=n_scenarios)
    arrival_mult = rng.uniform(*ARRIVAL_MULTIPLIER_RANGE, size=n_scenarios)

    rows = []
    for i in range(n_scenarios):
        result = run_scenario(
            n_capacity=int(n_capacity[i]),
            arrival_rate_multiplier=float(arrival_mult[i]),
            seed=i,
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
