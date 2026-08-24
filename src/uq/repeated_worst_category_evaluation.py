"""
Publication-rigor upgrade for SA-Mondrian CP, matching the same 30-repeat
discipline already applied to the paper's core marginal-coverage result
(repeated_evaluation.py): repeat the WORST-CATEGORY coverage comparison -
Mondrian CP (2D tercile grid) vs. QT-CP vs. SA-Mondrian CP (1D adaptive) -
across 30 independent (calibration, test) draws, rather than trusting the
single-split result reported in the paper's Section VII-C.

Reuses the EXACT same seed convention as repeated_evaluation.py (calibration
seed=1000+r/offset 500,000+r*10,000, test seed=2000+r/offset 700,000+r*10,000)
deliberately - not to avoid collision, but to get the SAME 30 draws so a
paired t-test between methods is valid (it needs to compare methods on
matched data, not independent samples).

SA-Mondrian's n_min and QT-CP's rho_cap are both reselected from calibration
data alone on EVERY repeat (never the test set, and never fixed across
repeats) - the same nested-calibration discipline used in the single-split
version, just run 30 times instead of once.
"""
import time

import joblib
import numpy as np
import pandas as pd
from scipy import stats

from src.surrogate.generate_training_data import generate
from src.uq.mondrian_cp import assign_categories, conformal_quantile
from src.uq.queueing_weighted_cp import (estimate_rho, mean_service_time_minutes,
                                          sigma_hat, select_rho_cap)
from src.uq.adaptive_geometric_mondrian_cp import (build_adaptive_bins, assign_from_edges,
                                                     select_n_min, worst_bin_coverage)
from src.des.er_simulation import STUDY_PERIOD_DAYS

MODELS_DIR = "models"
TABLES = "results/tables"
DETAIL_PATH = f"{TABLES}/repeated_worst_category_detail.csv"
SUMMARY_PATH = f"{TABLES}/repeated_worst_category_summary.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1
N_REPEATS = 30
N_CALIBRATION = 1200
N_TEST = 1000


def worst_category_mondrian(cal_df, test_df, resid_cal, resid_test):
    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    test_cat = assign_categories(test_df, cap_bounds, arr_bounds)
    coverages = []
    for cat in set(cal_cat):
        cal_mask = cal_cat == cat
        test_mask = test_cat == cat
        if cal_mask.sum() == 0 or test_mask.sum() == 0:
            continue
        q_cat = conformal_quantile(resid_cal[cal_mask], ALPHA)
        coverages.append((resid_test[test_mask] <= q_cat).mean())
    return min(coverages) if coverages else np.nan


def worst_category_qtcp(rho_cal, rho_test, resid_cal, resid_test, cal_df, test_df):
    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    test_cat = assign_categories(test_df, cap_bounds, arr_bounds)

    selected_cap = select_rho_cap(resid_cal, rho_cal)
    sigma_cal = sigma_hat(rho_cal, selected_cap)
    sigma_test = sigma_hat(rho_test, selected_cap)
    q_norm = conformal_quantile(resid_cal / sigma_cal, ALPHA)

    coverages = []
    for cat in set(cal_cat):
        test_mask = test_cat == cat
        if test_mask.sum() == 0:
            continue
        covered = resid_test[test_mask] <= q_norm * sigma_test[test_mask]
        coverages.append(covered.mean())
    return min(coverages) if coverages else np.nan


def run_repeats(n_repeats=N_REPEATS):
    models = {t: joblib.load(f"{MODELS_DIR}/surrogate_{t}.joblib") for t in TARGETS}

    arrivals = pd.read_csv(f"{TABLES}/arrivals_by_hour_bin.csv", index_col=0)["count"]
    esi_mix = pd.read_csv(f"{TABLES}/esi_mix.csv", index_col=0)["proportion"].to_dict()
    baseline_visits_per_day = arrivals.sum() / STUDY_PERIOD_DAYS
    mean_service_time = mean_service_time_minutes(esi_mix)

    rows = []
    t_start = time.time()
    for r in range(n_repeats):
        cal_df = generate(n_scenarios=N_CALIBRATION, seed=1000 + r, seed_offset=500_000 + r * 10_000)
        test_df = generate(n_scenarios=N_TEST, seed=2000 + r, seed_offset=700_000 + r * 10_000)

        rho_cal = estimate_rho(cal_df["n_capacity"], cal_df["arrival_rate_multiplier"],
                                baseline_visits_per_day, mean_service_time).values
        rho_test = estimate_rho(test_df["n_capacity"], test_df["arrival_rate_multiplier"],
                                 baseline_visits_per_day, mean_service_time).values

        for target in TARGETS:
            model = models[target]
            yhat_cal = model.predict(cal_df[FEATURES])
            resid_cal = np.abs(cal_df[target].values - yhat_cal)
            yhat_test = model.predict(test_df[FEATURES])
            resid_test = np.abs(test_df[target].values - yhat_test)

            mondrian_worst = worst_category_mondrian(cal_df, test_df, resid_cal, resid_test)
            qtcp_worst = worst_category_qtcp(rho_cal, rho_test, resid_cal, resid_test, cal_df, test_df)

            selected_n_min = select_n_min(rho_cal, resid_cal, seed=r)
            sa_worst, sa_n_bins = worst_bin_coverage(rho_cal, resid_cal, rho_test, resid_test, selected_n_min)

            rows.append({"repeat": r, "target": target,
                         "mondrian_worst": mondrian_worst, "qtcp_worst": qtcp_worst,
                         "sa_mondrian_worst": sa_worst, "sa_n_min": selected_n_min, "sa_n_bins": sa_n_bins})

        elapsed = time.time() - t_start
        print(f"repeat {r + 1}/{n_repeats} done ({elapsed:.0f}s elapsed, "
              f"~{elapsed / (r + 1) * (n_repeats - r - 1):.0f}s remaining)")

    return pd.DataFrame(rows)


def summarize(df):
    rows = []
    for target, g in df.groupby("target"):
        diff = g["sa_mondrian_worst"] - g["mondrian_worst"]
        t_stat, p_value = stats.ttest_rel(g["sa_mondrian_worst"], g["mondrian_worst"])
        rows.append({
            "target": target, "n_repeats": len(g),
            "mondrian_mean": g["mondrian_worst"].mean(), "mondrian_std": g["mondrian_worst"].std(),
            "qtcp_mean": g["qtcp_worst"].mean(), "qtcp_std": g["qtcp_worst"].std(),
            "sa_mondrian_mean": g["sa_mondrian_worst"].mean(), "sa_mondrian_std": g["sa_mondrian_worst"].std(),
            "sa_minus_mondrian_mean_diff": diff.mean(),
            "paired_ttest_t": t_stat, "paired_ttest_p": p_value,
        })
    return pd.DataFrame(rows)


def main():
    detail = run_repeats()
    detail.to_csv(DETAIL_PATH, index=False)
    summary = summarize(detail)
    summary.to_csv(SUMMARY_PATH, index=False)
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
