"""
Two-panel figure for SA-Mondrian CP: (left) the theorem - worst-bin scale
ratio for equal-width vs. geometric-in-(1-rho) binning as the observed
utilization ceiling rho_max approaches 1 - and (right) the honest empirical
result - worst-category coverage across 30 independent repeats (mean +/-
std), Mondrian CP vs. QT-CP vs. SA-Mondrian CP, all four targets, with
paired-t-test significance markers on SA-Mondrian's loss to Mondrian CP.
Built from results/tables/repeated_worst_category_summary.csv - no
invented numbers; this replaced an earlier, misleadingly favorable
single-split comparison once 30-repeat testing showed SA-Mondrian CP
actually loses to Mondrian CP (significantly on 3/4 targets).
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

OUT = "reports/paper/figures"
os.makedirs(OUT, exist_ok=True)

BLUE, ORANGE, GREEN = "#2a78d6", "#eb6834", "#1b9e5a"
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
})

fig, axes = plt.subplots(1, 2, figsize=(9, 4))

# --- Left: the theorem, worst-bin scale ratio vs. rho_max ---
ax = axes[0]
rho_max_grid = np.array([0.5, 0.7, 0.8, 0.9, 0.95, 0.98, 0.99, 0.995, 0.999])
B = 9
equal_width_ratio = 1 + (rho_max_grid / B) / (1 - rho_max_grid)
geometric_ratio = np.full_like(rho_max_grid, 2.0)

ax.plot(rho_max_grid, equal_width_ratio, "o-", color=BLUE, label=f"Equal-width bins (B={B})")
ax.plot(rho_max_grid, geometric_ratio, "o-", color=ORANGE, label="Geometric bins (any B)")
ax.set_yscale("log")
ax.set_xlabel(r"Observed utilization ceiling $\rho_{max}$")
ax.set_ylabel("Worst-bin scale ratio (log scale)")
ax.set_title("Equal-width bins blow up near the\nheavy-traffic singularity; geometric bins don't", fontsize=9.5)
ax.legend(fontsize=8)

# --- Right: honest empirical result, 30 repeats, mean +/- std, with
# paired-t-test significance markers on SA-Mondrian's loss to Mondrian CP ---
ax = axes[1]
rep = pd.read_csv("results/tables/repeated_worst_category_summary.csv")
targets = rep["target"].tolist()
mondrian_mean = rep["mondrian_mean"].values * 100
mondrian_std = rep["mondrian_std"].values * 100
qtcp_mean = rep["qtcp_mean"].values * 100
qtcp_std = rep["qtcp_std"].values * 100
sa_mean = rep["sa_mondrian_mean"].values * 100
sa_std = rep["sa_mondrian_std"].values * 100
p_values = rep["paired_ttest_p"].values

x = np.arange(len(targets))
w = 0.26
ax.bar(x - w, mondrian_mean, w, yerr=mondrian_std, capsize=2, label="Mondrian CP (2D)", color=BLUE)
ax.bar(x, qtcp_mean, w, yerr=qtcp_std, capsize=2, label="QT-CP", color=ORANGE)
ax.bar(x + w, sa_mean, w, yerr=sa_std, capsize=2, label="SA-Mondrian CP (1D, this work)", color=GREEN)
ax.axhline(90, color="#c0392b", linestyle=":", linewidth=1)
for i, p in enumerate(p_values):
    if p < 0.05:
        star = "**" if p < 0.001 else "*"
        ax.text(i + w, sa_mean[i] + sa_std[i] + 2, star, ha="center", fontsize=11, color="#c0392b")
ax.set_xticks(x); ax.set_xticklabels(targets, rotation=20, fontsize=8)
ax.set_ylabel("Worst-category coverage (%)")
ax.set_title("30 repeats, mean ± std: SA-Mondrian loses to\nMondrian CP (* p<0.05) but still beats QT-CP", fontsize=9.5)
ax.legend(fontsize=7.5)

fig.tight_layout()
fig.savefig(f"{OUT}/sa_mondrian_comparison.png", dpi=200, bbox_inches="tight")
plt.close(fig)
print("Saved", f"{OUT}/sa_mondrian_comparison.png")
