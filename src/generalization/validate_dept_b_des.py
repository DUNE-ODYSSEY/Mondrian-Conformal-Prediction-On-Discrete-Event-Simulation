"""
Publication-rigor upgrade, part 4 (continued): validate Department B's DES
calibration against B's real aggregated stats, same approach as
des/validate.py for department A (200 simulated days, compare mean
simulated patients/day to the real rate).
"""

import pandas as pd

from src.des.er_simulation import run_scenario

N_DAYS = 200
N_CAPACITY = 14
TABLES_DIR = "results/tables/dept_b"


def main():
    results = [run_scenario(n_capacity=N_CAPACITY, seed=s, tables_dir=TABLES_DIR) for s in range(N_DAYS)]
    df = pd.DataFrame(results)

    meta = pd.read_csv(f"{TABLES_DIR}/calibration_meta.csv", index_col=0)
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
    summary.to_csv(f"{TABLES_DIR}/des_validation.csv", header=["value"])
    print(summary)


if __name__ == "__main__":
    main()
