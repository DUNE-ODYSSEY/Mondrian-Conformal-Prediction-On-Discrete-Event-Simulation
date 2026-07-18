"""
Week 14-15: full comparison of GP baseline vs. standard CP vs. Mondrian CP -
coverage, interval width, computation time.

Coverage/width for GP and standard CP come straight from their existing
result tables (gp_baseline_metrics.csv, standard_cp_metrics.csv - symmetric
variant, for a fair like-for-like comparison against GP's single interval
and Mondrian's single per-category interval). Mondrian's overall (marginal)
coverage/width isn't in mondrian_cp_summary.csv - that file only has
per-category spread - so it's computed here by re-applying the per-category
quantiles to the full test set and aggregating, the same way GP/standard CP
report one marginal number.

Computation time is measured on a like-for-like basis: GP's fit_seconds is
its own model-fitting cost (already recorded in Week 8). Standard/Mondrian
CP don't fit a model - they wrap the already-trained surrogate - so their
fair "computation time" is their own calibration cost: computing the
empirical residual quantile(s) from calibration data. That's timed fresh
here since it wasn't recorded in Week 9-12.
"""

import time

import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split

from src.uq.mondrian_cp import assign_categories, conformal_quantile

TRAIN_DATA_PATH = "data/processed/surrogate_training_data.parquet"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
MODELS_DIR = "models"
OUT_TABLE = "results/tables/full_comparison.csv"
OUT_FIGURES = "results/figures"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1


def mondrian_overall(target, cal_df, test_df, cap_bounds, arr_bounds, cal_cat, test_cat):
    model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")
    model.predict(cal_df[FEATURES])  # warm-up: exclude first-call JIT/thread-pool overhead from the timing

    t0 = time.time()
    yhat_cal = model.predict(cal_df[FEATURES])
    resid_cal = np.abs(cal_df[target].values - yhat_cal)
    q_by_cat = {cat: conformal_quantile(resid_cal[cal_cat == cat], ALPHA) for cat in set(cal_cat)}
    calibration_seconds = time.time() - t0

    yhat_test = model.predict(test_df[FEATURES])
    y_test = test_df[target].values
    widths = np.array([2 * q_by_cat[c] for c in test_cat])
    covered = np.abs(y_test - yhat_test) <= widths / 2

    return covered.mean(), widths.mean(), calibration_seconds


def standard_calibration_time(target, cal_df):
    model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")
    model.predict(cal_df[FEATURES])  # warm-up: exclude first-call JIT/thread-pool overhead from the timing
    t0 = time.time()
    yhat_cal = model.predict(cal_df[FEATURES])
    resid_cal = np.abs(cal_df[target].values - yhat_cal)
    conformal_quantile(resid_cal, ALPHA)
    return time.time() - t0


def main():
    gp_df = pd.read_csv("results/tables/gp_baseline_metrics.csv").set_index("target")
    cp_df = pd.read_csv("results/tables/standard_cp_metrics.csv").set_index("target")

    df = pd.read_parquet(TRAIN_DATA_PATH)
    _, X_test, _, test_idx = train_test_split(df[FEATURES], df.index, test_size=0.2, random_state=42)
    test_df = df.loc[test_idx]
    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)

    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    test_cat = assign_categories(test_df, cap_bounds, arr_bounds)

    rows = []
    for target in TARGETS:
        rows.append(
            {
                "target": target,
                "method": "GP baseline",
                "coverage": gp_df.loc[target, "empirical_coverage"],
                "mean_width": gp_df.loc[target, "mean_interval_width"],
                "computation_seconds": gp_df.loc[target, "fit_seconds"] + gp_df.loc[target, "predict_seconds"],
            }
        )
        rows.append(
            {
                "target": target,
                "method": "Standard CP",
                "coverage": cp_df.loc[target, "symmetric_coverage"],
                "mean_width": cp_df.loc[target, "symmetric_mean_width"],
                "computation_seconds": standard_calibration_time(target, cal_df),
            }
        )
        mondrian_cov, mondrian_width, mondrian_seconds = mondrian_overall(
            target, cal_df, test_df, cap_bounds, arr_bounds, cal_cat, test_cat
        )
        rows.append(
            {
                "target": target,
                "method": "Mondrian CP",
                "coverage": mondrian_cov,
                "mean_width": mondrian_width,
                "computation_seconds": mondrian_seconds,
            }
        )

    out_df = pd.DataFrame(rows)
    out_df.to_csv(OUT_TABLE, index=False)
    print(out_df.to_string(index=False))
    make_charts(out_df)


def make_charts(df):
    import matplotlib.pyplot as plt

    methods = ["GP baseline", "Standard CP", "Mondrian CP"]
    colors = {"GP baseline": "#888888", "Standard CP": "#4C72B0", "Mondrian CP": "#55A868"}

    fig, axes = plt.subplots(1, 3, figsize=(15, 4.5))

    ax = axes[0]
    x = np.arange(len(TARGETS))
    width = 0.25
    for i, m in enumerate(methods):
        vals = [df[(df.target == t) & (df.method == m)]["coverage"].iloc[0] * 100 for t in TARGETS]
        ax.bar(x + (i - 1) * width, vals, width, label=m, color=colors[m])
    ax.axhline(90, color="black", linestyle="--", linewidth=1, label="90% target")
    ax.set_xticks(x)
    ax.set_xticklabels(TARGETS, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("Empirical coverage (%)")
    ax.set_title("Coverage")
    ax.legend(fontsize=8)

    ax = axes[1]
    for i, m in enumerate(methods):
        vals = [df[(df.target == t) & (df.method == m)]["mean_width"].iloc[0] for t in TARGETS]
        ax.bar(x + (i - 1) * width, vals, width, label=m, color=colors[m])
    ax.set_xticks(x)
    ax.set_xticklabels(TARGETS, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("Mean interval width")
    ax.set_title("Interval width")
    ax.legend(fontsize=8)

    ax = axes[2]
    comp_times = [df[df.method == m]["computation_seconds"].mean() for m in methods]
    ax.bar(methods, comp_times, color=[colors[m] for m in methods])
    ax.set_ylabel("Mean computation time (s, log scale)")
    ax.set_yscale("log")
    ax.set_title("Computation time")

    fig.tight_layout()
    fig.savefig(f"{OUT_FIGURES}/full_comparison.png", dpi=150)
    print(f"Saved {OUT_FIGURES}/full_comparison.png")


if __name__ == "__main__":
    main()
