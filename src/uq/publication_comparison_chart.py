"""
Publication-rigor upgrade, final: the authoritative 5-method comparison
chart, built from the 30-repeat statistical results (parts 1-2) instead of
the single-split point estimates full_comparison.py used. Combines
repeated_evaluation_summary.csv (GP, Standard CP, Mondrian CP) and
repeated_evaluation_cqr_summary.csv (CQR, Mondrian CQR) - same 30 seeds
underlie all five, so error bars are directly comparable across methods.

Error bars are 95% CI on the mean (coverage_ci95/width_ci95 columns) -
tight enough (a few tenths of a point on coverage) that they read almost as
single dots at this scale, which is itself the point: these numbers are not
single-split noise, they're stable estimates.
"""

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
METHODS = ["GP baseline", "Standard CP", "Mondrian CP", "CQR", "Mondrian CQR"]
COLORS = {
    "GP baseline": "#888888",
    "Standard CP": "#4C72B0",
    "Mondrian CP": "#55A868",
    "CQR": "#C44E52",
    "Mondrian CQR": "#8172B2",
}
OUT_PATH = "results/figures/publication_comparison.png"


def main():
    df1 = pd.read_csv("results/tables/repeated_evaluation_summary.csv")
    df2 = pd.read_csv("results/tables/repeated_evaluation_cqr_summary.csv")
    df = pd.concat([df1, df2], ignore_index=True)

    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    x = np.arange(len(TARGETS))
    width = 0.16

    ax = axes[0]
    for i, m in enumerate(METHODS):
        vals, errs = [], []
        for t in TARGETS:
            row = df[(df.target == t) & (df.method == m)].iloc[0]
            vals.append(row["coverage_mean"] * 100)
            errs.append(row["coverage_ci95"] * 100)
        ax.bar(x + (i - 2) * width, vals, width, yerr=errs, capsize=2, label=m, color=COLORS[m])
    ax.axhline(90, color="black", linestyle="--", linewidth=1, label="90% target")
    ax.set_xticks(x)
    ax.set_xticklabels(TARGETS, rotation=25, ha="right", fontsize=8)
    ax.set_ylabel("Empirical coverage (%), mean ± 95% CI")
    ax.set_title("Coverage (30 independent repeats)")
    ax.set_ylim(70, 100)
    ax.legend(fontsize=7, loc="lower right")

    ax = axes[1]
    for i, m in enumerate(METHODS):
        vals, errs = [], []
        for t in TARGETS:
            row = df[(df.target == t) & (df.method == m)].iloc[0]
            vals.append(row["width_mean"])
            errs.append(row["width_ci95"])
        ax.bar(x + (i - 2) * width, vals, width, yerr=errs, capsize=2, label=m, color=COLORS[m])
    ax.set_xticks(x)
    ax.set_xticklabels(TARGETS, rotation=25, ha="right", fontsize=8)
    ax.set_ylabel("Mean interval width, mean ± 95% CI")
    ax.set_title("Interval width (30 independent repeats)")
    ax.legend(fontsize=7)

    fig.tight_layout()
    fig.savefig(OUT_PATH, dpi=150)
    print(f"Saved {OUT_PATH}")


if __name__ == "__main__":
    main()
