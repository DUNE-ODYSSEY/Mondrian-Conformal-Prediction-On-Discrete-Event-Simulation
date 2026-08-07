"""
Likelihood-ratio weighted conformal prediction under covariate shift,
following Tibshirani, Barber, Candes, and Ramdas (2019), "Conformal
Prediction Under Covariate Shift."

Standard split CP (standard_cp.py) assumes calibration and test covariates
are drawn from the *same* distribution. Weighted CP relaxes this to a
known, *bounded* covariate shift: if the test covariate density q(x) is
absolutely continuous with respect to the calibration density p(x) (i.e.
q(x) > 0 implies p(x) > 0 - test points must fall within calibration's
support), reweighting each calibration point by the likelihood ratio
w(x) = q(x)/p(x) and taking a weighted conformal quantile restores exact
coverage despite the shift.

This project's calibration distribution for arrival_rate_multiplier is
Uniform[0.8, 1.3] (generate_calibration_data.py) - a *known* distribution,
since this project controls the DES's own sampling code, so exact
likelihood ratios can be computed analytically rather than estimated. This
script evaluates weighted CP under a genuinely different, complementary
shift from the exchangeability stress test's up-to-3.0x severity sweep:
a *moderate* shift, arrival_rate_multiplier ~ Uniform[0.9, 1.6], chosen
specifically so that most of its mass overlaps calibration's support
([0.9, 1.3]) while a real tail ([1.3, 1.6]) extends outside it - which
lets this script show both what weighted CP *can* fix (the overlapping
region) and the specific theoretical reason it *cannot* fix an
unbounded-density-ratio region (the non-overlapping tail), rather than
picking a shift scenario engineered to make weighting look unconditionally
effective.
"""

import joblib
import numpy as np
import pandas as pd

from src.des.er_simulation import run_scenario

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
OUT_PATH = "results/tables/weighted_cp_results.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]

ALPHA = 0.1
N_CAPACITY_RANGE = (15, 45)
N_TEST_SCENARIOS = 600

CAL_ARR_LOW, CAL_ARR_HIGH = 0.8, 1.3          # calibration support (generate_calibration_data.py)
TEST_ARR_LOW, TEST_ARR_HIGH = 0.9, 1.6         # moderate shift: partial overlap, real out-of-support tail
DENSITY_FLOOR = 1e-6                            # numerically-safe stand-in for "zero calibration density"


def cal_density(arr_mult):
    """Known Uniform[0.8, 1.3] calibration density (generate_calibration_data.py)."""
    inside = (arr_mult >= CAL_ARR_LOW) & (arr_mult <= CAL_ARR_HIGH)
    d = np.where(inside, 1.0 / (CAL_ARR_HIGH - CAL_ARR_LOW), DENSITY_FLOOR)
    return d


def test_density(arr_mult):
    """Known Uniform[0.9, 1.6] shifted test density (this script's own sampler)."""
    inside = (arr_mult >= TEST_ARR_LOW) & (arr_mult <= TEST_ARR_HIGH)
    return np.where(inside, 1.0 / (TEST_ARR_HIGH - TEST_ARR_LOW), 0.0)


def likelihood_ratio(arr_mult):
    return test_density(arr_mult) / cal_density(arr_mult)


def weighted_quantile(cal_scores, cal_weights, test_weight, alpha):
    """Tibshirani et al. (2019) weighted conformal quantile: normalizes
    calibration weights together with the test point's own weight (the
    weighted analogue of standard CP's n+1 finite-sample correction), then
    finds the smallest calibration score whose cumulative normalized
    weight reaches 1 - alpha."""
    order = np.argsort(cal_scores)
    sorted_scores = cal_scores[order]
    sorted_weights = cal_weights[order]
    total = sorted_weights.sum() + test_weight
    if total <= 0:
        return np.max(cal_scores)
    cum = np.cumsum(sorted_weights) / total
    idx = np.searchsorted(cum, 1 - alpha)
    if idx >= len(sorted_scores):
        return np.inf  # test point's own weight is too large a share -> infinite interval
    return sorted_scores[idx]


def generate_test_stream(n, seed=400_000):
    rng = np.random.default_rng(seed)
    arr_mult = rng.uniform(TEST_ARR_LOW, TEST_ARR_HIGH, size=n)
    n_capacity = rng.integers(N_CAPACITY_RANGE[0], N_CAPACITY_RANGE[1] + 1, size=n)
    rows = []
    for i in range(n):
        result = run_scenario(n_capacity=int(n_capacity[i]), arrival_rate_multiplier=float(arr_mult[i]), seed=seed + i)
        rows.append({"n_capacity": n_capacity[i], "arrival_rate_multiplier": arr_mult[i], **result})
    return pd.DataFrame(rows)


def main():
    test_df = generate_test_stream(N_TEST_SCENARIOS)
    in_support = (test_df["arrival_rate_multiplier"] <= CAL_ARR_HIGH).values

    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    X_cal = cal_df[FEATURES]
    cal_weights = likelihood_ratio(cal_df["arrival_rate_multiplier"].values)
    test_weights = likelihood_ratio(test_df["arrival_rate_multiplier"].values)

    rows = []
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")

        yhat_cal = model.predict(X_cal)
        abs_resid_cal = np.abs(cal_df[target].values - yhat_cal)

        yhat_test = model.predict(test_df[FEATURES])
        abs_resid_test = np.abs(test_df[target].values - yhat_test)

        # unweighted (standard) split CP quantile, same calibration set
        n = len(abs_resid_cal)
        std_level = min(np.ceil((n + 1) * (1 - ALPHA)) / n, 1.0)
        q_unweighted = np.quantile(abs_resid_cal, std_level, method="higher")
        covered_unweighted = abs_resid_test <= q_unweighted

        # weighted CP: per-test-point weighted quantile (weight depends on that point's own arrival_rate_multiplier)
        covered_weighted = np.zeros(len(test_df), dtype=bool)
        width_weighted = np.zeros(len(test_df))
        for i in range(len(test_df)):
            q_i = weighted_quantile(abs_resid_cal, cal_weights, test_weights[i], ALPHA)
            width_weighted[i] = 2 * q_i if np.isfinite(q_i) else np.nan
            covered_weighted[i] = abs_resid_test[i] <= q_i

        for region_name, mask in [("overlap_region", in_support), ("out_of_support_tail", ~in_support), ("overall", np.ones(len(test_df), dtype=bool))]:
            rows.append({
                "target": target,
                "region": region_name,
                "n_test": int(mask.sum()),
                "unweighted_coverage": covered_unweighted[mask].mean(),
                "unweighted_width": 2 * q_unweighted,
                "weighted_coverage": covered_weighted[mask].mean(),
                "weighted_mean_width": np.nanmean(width_weighted[mask]),
            })

    out_df = pd.DataFrame(rows)
    out_df.to_csv(OUT_PATH, index=False)
    print(out_df.to_string(index=False))
    print(f"\nSaved {OUT_PATH}")


if __name__ == "__main__":
    main()
