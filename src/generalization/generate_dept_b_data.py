"""
Publication-rigor upgrade, part 4 (continued): generate Department B's
training and CP-calibration scenario sweeps.

n_capacity default/range recalibrated for B's own offered load, not copied
from department A's absolute numbers - B's real volume is roughly half A's
(133.4 vs. 258.2 visits/day) and its ESI mix skews toward lower acuity
(23.1% ESI-2 vs. A's 37.9%, 26.8% ESI-4 vs. A's 16.4%), so a "capacity of
30" would be wildly over-provisioned and produce uninteresting, barely-
congested dynamics. Computed the same way A's default was: offered load
(Erlangs) = arrival_rate * weighted mean service time. B: ~10.4 erlangs
average, ~16.1 peak (vs. A's ~22 / ~34) - roughly half, as expected given
roughly half the volume. Default n_capacity=14 sits at the same relative
position between average and peak load as A's 30 does for A's own numbers.
Range (7,22) mirrors A's ~0.5x-1.5x-of-default pattern (A: 15-45 around 30).

arrival_rate_multiplier range is unchanged (0.8-1.3) - it's already a
relative multiplier on top of whatever tables_dir's own calibrated rate is,
so it doesn't need department-specific rescaling.

Same sizes as department A (5000 training scenarios, 1200 calibration) for
a like-for-like comparison; same seed-range discipline (calibration draws
use a different sampling seed and a large DES-seed offset from training,
consistent with generate_calibration_data.py).
"""

from src.surrogate.generate_training_data import generate

TABLES_DIR = "results/tables/dept_b"
N_CAPACITY_RANGE = (7, 22)
ARRIVAL_MULTIPLIER_RANGE = (0.8, 1.3)

TRAINING_OUT = "data/processed/dept_b/surrogate_training_data.parquet"
CALIBRATION_OUT = "data/processed/dept_b/cp_calibration_data.parquet"


def main():
    train_df = generate(
        n_scenarios=5000, seed=0, seed_offset=0,
        n_capacity_range=N_CAPACITY_RANGE, arrival_multiplier_range=ARRIVAL_MULTIPLIER_RANGE,
        tables_dir=TABLES_DIR,
    )
    train_df.to_parquet(TRAINING_OUT, index=False)
    print(f"Generated {len(train_df)} training scenarios -> {TRAINING_OUT}")
    print(train_df.describe())

    # different sampling seed + large DES-seed offset -> independent of training data,
    # same discipline as generate_calibration_data.py for department A
    cal_df = generate(
        n_scenarios=1200, seed=1, seed_offset=100_000,
        n_capacity_range=N_CAPACITY_RANGE, arrival_multiplier_range=ARRIVAL_MULTIPLIER_RANGE,
        tables_dir=TABLES_DIR,
    )
    cal_df.to_parquet(CALIBRATION_OUT, index=False)
    print(f"\nGenerated {len(cal_df)} calibration scenarios -> {CALIBRATION_OUT}")


if __name__ == "__main__":
    main()
