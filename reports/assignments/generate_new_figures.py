"""
New supporting figures for the two book-format assignment reports, generated
from existing, already-computed results/tables/*.csv - no new experiments,
no invented numbers. Uses this project's established categorical color
scheme (grey=GP, blue=Standard CP, green=Mondrian CP) extended with a
validated colorblind-safe categorical order for the 5-method frontier plot.
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

OUT = "reports/assignments/figures"
os.makedirs(OUT, exist_ok=True)

# Validated categorical palette (dataviz skill reference palette, light mode)
BLUE = "#2a78d6"
ORANGE = "#eb6834"
AQUA = "#1baf7a"
YELLOW = "#eda100"
MAGENTA = "#e87ba4"
GREY = "#8a8a86"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.edgecolor": "#52514e",
    "axes.labelcolor": "#0b0b0b",
    "text.color": "#0b0b0b",
    "xtick.color": "#52514e",
    "ytick.color": "#52514e",
})

# ---------------------------------------------------------------------
# Figure A1: Per-category pooled vs Mondrian coverage heatmap
# (mean_wait_minutes, the target with the clearest core-finding story)
# ---------------------------------------------------------------------
df = pd.read_csv("results/tables/mondrian_cp_detail.csv")
target = "mean_wait_minutes"
sub = df[df["target"] == target].copy()
sub["staff"] = sub["category"].str.extract(r"staff=(\w+)")
sub["arrival"] = sub["category"].str.extract(r"arrival=(\w+)")
order = ["Low", "Med", "High"]
pooled_grid = sub.pivot(index="staff", columns="arrival", values="pooled_coverage").reindex(index=order, columns=order) * 100
mondrian_grid = sub.pivot(index="staff", columns="arrival", values="mondrian_coverage").reindex(index=order, columns=order) * 100

fig, axes = plt.subplots(1, 2, figsize=(9, 4))
for ax, grid, title in zip(axes, [pooled_grid, mondrian_grid], ["Pooled (Standard) CP", "Mondrian CP"]):
    im = ax.imshow(grid.values, cmap="RdYlGn", vmin=60, vmax=100, aspect="auto")
    ax.set_xticks(range(3)); ax.set_xticklabels(order)
    ax.set_yticks(range(3)); ax.set_yticklabels(order)
    ax.set_xlabel("Arrival-rate tercile")
    ax.set_ylabel("Staffing tercile")
    ax.set_title(title, fontsize=11)
    for i in range(3):
        for j in range(3):
            v = grid.values[i, j]
            ax.text(j, i, f"{v:.1f}%", ha="center", va="center",
                     color="white" if v < 82 else "#0b0b0b", fontsize=9, fontweight="bold")
fig.suptitle(f"Per-category coverage: {target} (target = 90%)", fontsize=12, y=1.02)
fig.tight_layout()
fig.savefig(f"{OUT}/per_category_coverage_heatmap.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure A2: 30-repeat coverage distribution (boxplot) per method
# ---------------------------------------------------------------------
rep = pd.read_csv("results/tables/repeated_evaluation_detail.csv")
methods = ["GP baseline", "Standard CP", "Mondrian CP"]
colors = [GREY, BLUE, AQUA]
targets = rep["target"].unique()
fig, axes = plt.subplots(1, len(targets), figsize=(13, 4), sharey=False)
for ax, t in zip(axes, targets):
    data = [rep[(rep["target"] == t) & (rep["method"] == m)]["coverage"].values * 100 for m in methods]
    bp = ax.boxplot(data, patch_artist=True, widths=0.6, medianprops=dict(color="#0b0b0b"))
    for patch, c in zip(bp["boxes"], colors):
        patch.set_facecolor(c)
        patch.set_alpha(0.75)
    ax.axhline(90, color="#e34948", linestyle="--", linewidth=1, label="90% target")
    ax.set_xticks([1, 2, 3])
    ax.set_xticklabels(["GP", "Std CP", "Mondrian"], rotation=20, fontsize=8)
    ax.set_title(t, fontsize=9)
    if ax is axes[0]:
        ax.set_ylabel("Coverage across 30 repeats (%)")
fig.suptitle("Distribution of empirical coverage across 30 independent (calibration, test) draws", fontsize=11, y=1.03)
fig.tight_layout()
fig.savefig(f"{OUT}/repeated_evaluation_boxplot.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure A3: Coverage/width frontier across 5 methods (30-repeat means)
# ---------------------------------------------------------------------
cqr = pd.read_csv("results/tables/repeated_evaluation_cqr_detail.csv")
method_colors = {
    "GP baseline": GREY, "Standard CP": BLUE, "Mondrian CP": AQUA,
    "CQR": ORANGE, "Mondrian CQR": MAGENTA,
}
fig, axes = plt.subplots(1, 4, figsize=(14, 3.6))
for ax, t in zip(axes, targets):
    for m, c in method_colors.items():
        if m in ("CQR", "Mondrian CQR"):
            d = cqr[(cqr["target"] == t) & (cqr["method"] == m)]
        else:
            d = rep[(rep["target"] == t) & (rep["method"] == m)]
        if len(d) == 0:
            continue
        cov = d["coverage"].mean() * 100
        wid = d["mean_width"].mean()
        ax.scatter(wid, cov, color=c, s=70, label=m, edgecolor="white", linewidth=0.8, zorder=3)
    ax.axhline(90, color="#c9c8c0", linestyle="--", linewidth=1, zorder=1)
    ax.set_title(t, fontsize=9)
    ax.set_xlabel("Mean width", fontsize=8)
    if ax is axes[0]:
        ax.set_ylabel("Mean coverage (%)")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=5, bbox_to_anchor=(0.5, -0.08), frameon=False, fontsize=9)
fig.suptitle("Coverage/width frontier across five UQ methods (30-repeat means)", fontsize=11, y=1.05)
fig.tight_layout()
fig.savefig(f"{OUT}/coverage_width_frontier.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure A4: Real arrival distribution by hour bin (department A)
# ---------------------------------------------------------------------
arr = pd.read_csv("results/tables/arrivals_by_hour_bin.csv")
fig, ax = plt.subplots(figsize=(6.5, 3.8))
ax.bar(arr["arrivalhour_bin"], arr["count"], color=BLUE, width=0.6)
ax.set_xlabel("Arrival hour bin")
ax.set_ylabel("Real visit count (Department A)")
ax.set_title("Real ED arrival volume by 4-hour bin (calibration source)", fontsize=11)
fig.tight_layout()
fig.savefig(f"{OUT}/real_arrivals_by_hour.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure B1 (Assignment 2): coverage vs severity, GBR vs MLP
# ---------------------------------------------------------------------
gbr = pd.read_csv("results/tables/exchangeability_stress_test.csv")
mlp = pd.read_csv("results/tables/exchangeability_stress_test_mlp.csv")
fig, axes = plt.subplots(1, 4, figsize=(14, 3.6), sharex=True)
for ax, t in zip(axes, targets):
    g = gbr[gbr["target"] == t].sort_values("arrival_rate_multiplier")
    m = mlp[mlp["target"] == t].sort_values("arrival_rate_multiplier")
    ax.plot(g["arrival_rate_multiplier"], g["standard_coverage"] * 100, "o-", color=BLUE, label="GBR surrogate")
    ax.plot(m["arrival_rate_multiplier"], m["standard_coverage"] * 100, "s-", color=ORANGE, label="MLP surrogate")
    ax.axvline(1.3, color="#c9c8c0", linestyle="--", linewidth=1)
    ax.axhline(90, color="#e34948", linestyle=":", linewidth=1)
    ax.set_title(t, fontsize=9)
    ax.set_xlabel("Arrival-rate multiplier", fontsize=8)
    if ax is axes[0]:
        ax.set_ylabel("Standard CP coverage (%)")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.1), frameon=False, fontsize=9)
fig.suptitle("Coverage collapse under exchangeability violation: GBR vs. MLP surrogate\n(dashed vertical = training boundary; dotted horizontal = 90% target)", fontsize=10, y=1.08)
fig.tight_layout()
fig.savefig(f"{OUT}/exchangeability_coverage_collapse.png", dpi=170, bbox_inches="tight")
plt.close(fig)

print("All figures generated in", OUT)
for f in sorted(os.listdir(OUT)):
    print(" -", f)
