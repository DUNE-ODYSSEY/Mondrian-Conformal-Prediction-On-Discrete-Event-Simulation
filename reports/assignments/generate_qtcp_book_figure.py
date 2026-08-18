"""
Book-style 3-way per-category heatmap for the new Section 6.5 (QT-CP):
Pooled/Standard vs. Mondrian vs. QT-CP coverage across the same 9
staffing x arrival-rate cells used throughout this report, for
mean_wait_minutes (the target with the clearest conditional gap). Matches
the visual language of generate_new_experiment_figures.py's
per_category_coverage_heatmap.png - built from
results/tables/queueing_weighted_cp_detail.csv, no invented numbers.
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

OUT = "reports/assignments/figures"
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
})

detail = pd.read_csv("results/tables/queueing_weighted_cp_detail.csv")
target = "mean_wait_minutes"
sub = detail[detail["target"] == target].copy()
sub["staff"] = sub["category"].str.extract(r"staff=(\w+)")
sub["arrival"] = sub["category"].str.extract(r"arrival=(\w+)")
order = ["Low", "Med", "High"]

grids = {}
for col, name in [("standard_coverage", "Pooled (Standard) CP"),
                   ("mondrian_coverage", "Mondrian CP"),
                   ("qtcp_coverage", "QT-CP (this work)")]:
    grids[name] = sub.pivot(index="staff", columns="arrival", values=col).reindex(index=order, columns=order) * 100

fig, axes = plt.subplots(1, 3, figsize=(11, 4))
for ax, (name, grid) in zip(axes, grids.items()):
    im = ax.imshow(grid.values, cmap="RdYlGn", vmin=60, vmax=100, aspect="auto")
    ax.set_xticks(range(3)); ax.set_xticklabels(order)
    ax.set_yticks(range(3)); ax.set_yticklabels(order)
    ax.set_xlabel("Arrival-rate tercile")
    if ax is axes[0]:
        ax.set_ylabel("Staffing tercile")
    ax.set_title(name, fontsize=10)
    for i in range(3):
        for j in range(3):
            v = grid.values[i, j]
            ax.text(j, i, f"{v:.1f}%", ha="center", va="center",
                     color="white" if v < 78 else "#0b0b0b", fontsize=9, fontweight="bold")
fig.suptitle(f"Per-category coverage, all three methods: {target} (target = 90%)", fontsize=12, y=1.03)
fig.tight_layout()
fig.savefig(f"{OUT}/qtcp_per_category_heatmap.png", dpi=170, bbox_inches="tight")
plt.close(fig)
print("Saved", f"{OUT}/qtcp_per_category_heatmap.png")
