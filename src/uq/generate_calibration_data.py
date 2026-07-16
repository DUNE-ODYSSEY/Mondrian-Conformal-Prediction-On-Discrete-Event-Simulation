"""
Week 9-10: generate a calibration set for conformal prediction.

CP calibration data must be separate from what the surrogate was trained on
- reusing training-set residuals would understate the true residual spread
(the surrogate fits those points), which would break coverage validity. This
uses a different parameter-sampling seed AND a large DES seed offset from
generate_training_data.py, so nothing here overlaps with the training/test
data used to fit and evaluate the surrogate and GP baseline.

Same scenario ranges as training data (n_capacity 15-45, arrival_rate_multiplier
0.8-1.3) - calibration data has to be exchangeable with the test set, which
means drawn from the same distribution, not a different one.
"""

from src.surrogate.generate_training_data import generate

N_CALIBRATION = 1200
OUT_PATH = "data/processed/cp_calibration_data.parquet"


def main():
    df = generate(n_scenarios=N_CALIBRATION, seed=1, seed_offset=100_000)
    df.to_parquet(OUT_PATH, index=False)
    print(f"Generated {len(df)} calibration scenarios -> {OUT_PATH}")
    print(df.describe())


if __name__ == "__main__":
    main()
