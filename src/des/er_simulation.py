"""
Week 4-5: ER discrete-event simulation, calibrated on real arrival/ESI
distributions (results/tables/, department A, see extract_distributions.py)
and literature service-time params.

Single resource pool (`capacity`) represents a combined bed+provider slot for
the full visit duration — the standard simplification in ED queueing/DES
literature (M/G/c queue), used here because we have no real data to split
staffing into separate doctor/bed/nurse ratios; modeling them separately
would just add unverified guesses without adding insight for this project's
actual question (UQ methodology, not hospital operations realism).

`n_capacity` is a parameter, not fixed — Week 6-7 varies it (along with
arrival-rate scaling) to generate surrogate training data across scenarios.

Default n_capacity=30: offered load = arrival_rate * mean_service_time
(Erlangs) is ~22 erlangs on average and ~34 erlangs during the 11-14 peak
bin (department A: ~258 visits/day, ~121 min mean service time). 30 servers
sits between those, giving a stable but visibly congested system — useful
for a DES whose whole purpose is producing realistic queueing/wait
variation, not a system with near-zero wait.
"""

import numpy as np
import pandas as pd
import simpy

TABLES = "results/tables"

# Literature log-normal service time params by ESI, from extract_distributions.py
SERVICE_TIME_PARAMS_BY_ESI = {
    1: (180, 90),
    2: (150, 75),
    3: (120, 60),
    4: (75, 40),
    5: (45, 25),
}

STUDY_PERIOD_DAYS = 1248  # March 2014 - July 2017, per extract_distributions.py

HOUR_BINS = ["23-02", "03-06", "07-10", "11-14", "15-18", "19-22"]

# "23-02" wraps midnight (hours 23,0,1,2) — build the hour->bin map explicitly
# rather than hour // 4, which would wrongly put hours 0-3 in "23-02".
HOUR_TO_BIN = {}
for _bin_label, _hours in zip(
    HOUR_BINS, [[23, 0, 1, 2], [3, 4, 5, 6], [7, 8, 9, 10], [11, 12, 13, 14], [15, 16, 17, 18], [19, 20, 21, 22]]
):
    for _h in _hours:
        HOUR_TO_BIN[_h] = _bin_label


def lognormal_params_from_mean_sd(mean, sd):
    """Convert a desired (mean, sd) in minutes to the underlying normal's (mu, sigma)."""
    variance = sd**2
    mu = np.log(mean**2 / np.sqrt(variance + mean**2))
    sigma = np.sqrt(np.log(1 + variance / mean**2))
    return mu, sigma


def load_calibration():
    """Real arrival-by-hour-bin distribution and real ESI mix (single department)."""
    arrivals = pd.read_csv(f"{TABLES}/arrivals_by_hour_bin.csv", index_col=0)["count"]
    arrivals = arrivals.reindex(HOUR_BINS)
    esi_mix = pd.read_csv(f"{TABLES}/esi_mix.csv", index_col=0)["proportion"]
    return arrivals, esi_mix


class ERSimulation:
    """
    One 24h day of ER operation. Arrival rate varies by 4-hour bin (real data),
    ESI is sampled from the real acuity mix, service time is sampled per-ESI
    log-normal (literature params). Higher-acuity (lower ESI number) patients
    get priority for the shared capacity pool, matching real triage behavior.
    """

    def __init__(self, n_capacity, arrivals_by_bin, esi_mix, arrival_rate_multiplier=1.0, seed=None):
        self.n_capacity = n_capacity
        self.arrivals_by_bin = arrivals_by_bin
        self.esi_mix = esi_mix
        self.np_rng = np.random.default_rng(seed)

        total_visits_per_day = arrivals_by_bin.sum() / STUDY_PERIOD_DAYS
        self.rate_per_bin = (
            (arrivals_by_bin / arrivals_by_bin.sum()) * total_visits_per_day / 4 * arrival_rate_multiplier
        )  # per hour

        self.wait_times = []
        self.total_times = []
        self.n_patients = 0

    def patient(self, env, capacity):
        arrival_time = env.now
        esi = self.np_rng.choice(self.esi_mix.index.astype(int), p=self.esi_mix.values)
        mean, sd = SERVICE_TIME_PARAMS_BY_ESI[esi]
        mu, sigma = lognormal_params_from_mean_sd(mean, sd)
        service_time = self.np_rng.lognormal(mu, sigma)

        with capacity.request(priority=esi) as req:
            yield req
            wait = env.now - arrival_time
            self.wait_times.append(wait)
            yield env.timeout(service_time)
            self.total_times.append(env.now - arrival_time)
            self.n_patients += 1

    def arrival_process(self, env, capacity):
        for hour in range(24):
            bin_idx = HOUR_TO_BIN[hour]
            rate = self.rate_per_bin[bin_idx]
            for _ in range(60):  # simulate minute-by-minute within this hour
                if self.np_rng.random() < rate / 60:
                    env.process(self.patient(env, capacity))
                yield env.timeout(1)

    def run(self):
        env = simpy.Environment()
        capacity = simpy.PriorityResource(env, capacity=self.n_capacity)
        env.process(self.arrival_process(env, capacity))
        env.run(until=24 * 60)
        return {
            "n_patients": self.n_patients,
            "mean_wait_minutes": np.mean(self.wait_times) if self.wait_times else 0.0,
            "mean_total_minutes": np.mean(self.total_times) if self.total_times else 0.0,
            "p95_wait_minutes": np.percentile(self.wait_times, 95) if self.wait_times else 0.0,
        }


def run_scenario(n_capacity=30, arrival_rate_multiplier=1.0, seed=None):
    arrivals_by_bin, esi_mix = load_calibration()
    sim = ERSimulation(n_capacity, arrivals_by_bin, esi_mix, arrival_rate_multiplier, seed=seed)
    return sim.run()


if __name__ == "__main__":
    result = run_scenario(seed=42)
    print(result)
