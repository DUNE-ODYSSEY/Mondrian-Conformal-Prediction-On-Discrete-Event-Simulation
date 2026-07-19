"""
Publication-rigor upgrade: repeat the full GP / standard-CP / Mondrian-CP
pipeline across many independent (calibration, test) draws, not just the
single train_test_split used in Weeks 8-14. Reports mean/std/95% CI of
coverage and width per target per method, so results are backed by a
sampling distribution instead of a single point estimate - the main gap
between the Week 14-15 comparison and something defensible in a paper.

Both calibration AND test data are freshly generated DES scenarios on every
repeat (never reused across repeats or from the original training data) -
this captures both calibration-quantile variance and evaluation-sample
variance, i.e. it answers "does this coverage number replicate" rather than
"was this one split lucky." The underlying surrogate/GP models are fixed
per repeat's own GP refit (matching how gp_baseline.py subsamples fresh
training data each time) - only standard/Mondrian CP reuse the pre-trained
Week 6-7 mean surrogate, exactly as in Weeks 9-12.

Seed ranges are kept clearly separate from every other script in this repo
so nothing here silently overlaps with training/calibration/test/stress-test
data used elsewhere:
  calibration draws: sampling seed 1000+r, DES seed_offset 500,000 + r*10,000
  test draws:         sampling seed 2000+r, DES seed_offset 700,000 + r*10,000
"""

import time

import joblib
import numpy as np
import pandas as pd
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, ConstantKernel, WhiteKernel
from sklearn.preprocessing import StandardScaler

from src.surrogate.generate_training_data import generate
from src.uq.mondrian_cp import assign_categories, conformal_quantile

MODELS_DIR = "models"
DETAIL_PATH = "results/tables/repeated_evaluation_detail.csv"
SUMMARY_PATH = "results/tables/repeated_evaluation_summary.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1
Z = 1.6448536269514722  # norm.ppf(1 - ALPHA/2), avoids a scipy import for one constant

N_REPEATS = 30
N_CALIBRATION = 1200
N_TEST = 1000
N_GP_TRAIN = 1000
GP_N_RESTARTS = 1  # was 2 in Week 8's one-off gp_baseline.py; halved here since this
# refits GP fresh on every one of 30 repeats. N_GP_TRAIN is left unchanged (1000,
# same as Week 8) so GP's data budget stays directly comparable to the original
# baseline - only the optimizer's restart count (a speed/thoroughness knob, not
# the comparison basis) is reduced. One restart still escapes the worst local
# optima for a kernel this simple (2D RBF + white noise).


def gp_predict(target, cal_df, test_df, seed):
    rng = np.random.default_rng(seed)
    gp_idx = rng.choice(len(cal_df), size=N_GP_TRAIN, replace=False)
    scaler = StandardScaler().fit(cal_df.iloc[gp_idx][FEATURES])
    X_train_s = scaler.transform(cal_df.iloc[gp_idx][FEATURES])
    X_test_s = scaler.transform(test_df[FEATURES])
    y_train = cal_df.iloc[gp_idx][target].values

    kernel = ConstantKernel(1.0) * RBF(length_scale=[1.0, 1.0]) + WhiteKernel(1.0)
    gp = GaussianProcessRegressor(kernel=kernel, normalize_y=True, n_restarts_optimizer=GP_N_RESTARTS, random_state=seed)
    gp.fit(X_train_s, y_train)
    y_pred, y_std = gp.predict(X_test_s, return_std=True)
    return y_pred, y_std


def run_repeats(n_repeats=N_REPEATS):
    models = {t: joblib.load(f"{MODELS_DIR}/surrogate_{t}.joblib") for t in TARGETS}
    rows = []
    t_start = time.time()

    for r in range(n_repeats):
        cal_df = generate(n_scenarios=N_CALIBRATION, seed=1000 + r, seed_offset=500_000 + r * 10_000)
        test_df = generate(n_scenarios=N_TEST, seed=2000 + r, seed_offset=700_000 + r * 10_000)

        cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
        arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
        cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
        test_cat = assign_categories(test_df, cap_bounds, arr_bounds)

        for target in TARGETS:
            model = models[target]
            y_test = test_df[target].values

            # GP baseline
            gp_pred, gp_std = gp_predict(target, cal_df, test_df, seed=3000 + r)
            gp_lower, gp_upper = gp_pred - Z * gp_std, gp_pred + Z * gp_std
            gp_covered = (y_test >= gp_lower) & (y_test <= gp_upper)
            rows.append({"repeat": r, "target": target, "method": "GP baseline",
                         "coverage": gp_covered.mean(), "mean_width": (gp_upper - gp_lower).mean()})

            # standard CP
            yhat_cal = model.predict(cal_df[FEATURES])
            resid_cal = np.abs(cal_df[target].values - yhat_cal)
            yhat_test = model.predict(test_df[FEATURES])
            resid_test = np.abs(y_test - yhat_test)

            q_pooled = conformal_quantile(resid_cal, ALPHA)
            standard_covered = resid_test <= q_pooled
            rows.append({"repeat": r, "target": target, "method": "Standard CP",
                         "coverage": standard_covered.mean(), "mean_width": 2 * q_pooled})

            # Mondrian CP
            mondrian_covered = np.zeros(len(test_df), dtype=bool)
            mondrian_widths = np.zeros(len(test_df))
            for cat in set(cal_cat):
                cal_mask = cal_cat == cat
                test_mask = test_cat == cat
                if test_mask.sum() == 0 or cal_mask.sum() == 0:
                    continue
                q_cat = conformal_quantile(resid_cal[cal_mask], ALPHA)
                mondrian_covered[test_mask] = resid_test[test_mask] <= q_cat
                mondrian_widths[test_mask] = 2 * q_cat
            rows.append({"repeat": r, "target": target, "method": "Mondrian CP",
                         "coverage": mondrian_covered.mean(), "mean_width": mondrian_widths.mean()})

        elapsed = time.time() - t_start
        print(f"repeat {r + 1}/{n_repeats} done ({elapsed:.0f}s elapsed)")

    return pd.DataFrame(rows)


def summarize(df):
    def ci95(x):
        n = len(x)
        return 1.96 * x.std(ddof=1) / np.sqrt(n)

    summary = df.groupby(["target", "method"]).agg(
        n_repeats=("coverage", "size"),
        coverage_mean=("coverage", "mean"),
        coverage_std=("coverage", "std"),
        coverage_ci95=("coverage", ci95),
        width_mean=("mean_width", "mean"),
        width_std=("mean_width", "std"),
        width_ci95=("mean_width", ci95),
    ).reset_index()
    return summary


def main():
    detail = run_repeats()
    detail.to_csv(DETAIL_PATH, index=False)
    summary = summarize(detail)
    summary.to_csv(SUMMARY_PATH, index=False)
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
