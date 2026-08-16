"""
Comparison figure for the new queueing-theoretic normalized CP (QT-CP)
method: worst-category coverage and mean marginal width, standard CP vs.
Mondrian CP vs. QT-CP, all four targets. Built from
results/tables/queueing_weighted_cp_{detail,summary}.csv - no invented
numbers.
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

OUT = "reports/paper/figures"
os.makedirs(OUT, exist_ok=True)

BLUE, ORANGE, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
})

summary = pd.read_csv("results/tables/queueing_weighted_cp_summary.csv")
detail = pd.read_csv("results/tables/queueing_weighted_cp_detail.csv")

targets = summary["target"].tolist()
standard_worst = detail.groupby("target")["standard_coverage"].min().reindex(targets) * 100
mondrian_worst = summary["mondrian_worst_category_coverage"].values * 100
qtcp_worst = summary["qtcp_worst_category_coverage"].values * 100

fig, axes = plt.subplots(1, 2, figsize=(9, 4))
ax = axes[0]
x = np.arange(len(targets))
w = 0.26
ax.bar(x - w, standard_worst, w, label="Standard CP", color="#8a8a86")
ax.bar(x, mondrian_worst, w, label="Mondrian CP", color=BLUE)
ax.bar(x + w, qtcp_worst, w, label="QT-CP (this work)", color=ORANGE)
ax.axhline(90, color="#c0392b", linestyle=":", linewidth=1)
ax.set_xticks(x); ax.set_xticklabels(targets, rotation=20, fontsize=8)
ax.set_ylabel("Worst-category coverage (%)")
ax.set_title("Conditional coverage: worst category", fontsize=10)
ax.legend(fontsize=7.5)

ax = axes[1]
std_w = summary["standard_mean_width"].values
qtcp_w = summary["qtcp_mean_width"].values
ratio = qtcp_w / std_w
colors = [AQUA if r < 1 else "#c0392b" for r in ratio]
ax.bar(x, ratio, 0.5, color=colors)
ax.axhline(1.0, color="#40403c", linestyle="--", linewidth=1)
ax.set_xticks(x); ax.set_xticklabels(targets, rotation=20, fontsize=8)
ax.set_ylabel("QT-CP width / Standard CP width")
ax.set_title("Marginal width efficiency", fontsize=10)
for i, r in enumerate(ratio):
    ax.text(i, r + 0.03, f"{r:.2f}x", ha="center", fontsize=8)

fig.suptitle("Queueing-theoretic normalized CP (QT-CP): narrower on 3/4 targets,\n"
             "but does not match Mondrian's conditional-coverage repair", fontsize=10)
fig.tight_layout()
fig.savefig(f"{OUT}/qtcp_comparison.png", dpi=200, bbox_inches="tight")
plt.close(fig)
print("Saved", f"{OUT}/qtcp_comparison.png")
