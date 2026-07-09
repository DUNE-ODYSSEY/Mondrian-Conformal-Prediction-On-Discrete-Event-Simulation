"""
Week 4-5: validate the calibrated DES against real aggregated stats.

Runs many independent simulated days and compares mean simulated patients/day
to the real department-A daily average (calibration_meta.csv). Some
undershoot is expected: patients still queued when a 24h simulated day ends
are cut off (right-censoring at the day boundary) rather than carried over,
since each run models one independent day for later surrogate-training
sampling (Week 6-7), not a continuous multi-day rollout.
"""

import pandas as pd

from src.des.er_simulation import run_scenario

N_DAYS = 200
N_CAPACITY = 30


def main():
    results = [run_scenario(n_capacity=N_CAPACITY, seed=s) for s in range(N_DAYS)]
    df = pd.DataFrame(results)

    meta = pd.read_csv("results/tables/calibration_meta.csv", index_col=0)
    real_visits_per_day = float(meta.loc["visits_per_day", "value"])
    sim_visits_per_day = df["n_patients"].mean()

    summary = pd.Series(
        {
            "n_days_simulated": N_DAYS,
            "n_capacity": N_CAPACITY,
            "real_visits_per_day": real_visits_per_day,
            "sim_mean_visits_per_day": sim_visits_per_day,
            "sim_std_visits_per_day": df["n_patients"].std(),
            "ratio_sim_over_real": sim_visits_per_day / real_visits_per_day,
            "sim_mean_wait_minutes": df["mean_wait_minutes"].mean(),
            "sim_mean_total_minutes": df["mean_total_minutes"].mean(),
            "sim_mean_p95_wait_minutes": df["p95_wait_minutes"].mean(),
        }
    )
    summary.to_csv("results/tables/des_validation.csv", header=["value"])
    print(summary)


if __name__ == "__main__":
    main()
