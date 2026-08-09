"""
Further Chapter 10 figures: a full capacity-optimization sensitivity
heatmap across all 5 W_max thresholds x 6 arrival multipliers (built from
the existing results/tables/capacity_optimization.csv - no new
experiments), and a deployment-architecture dataflow diagram.
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import os

OUT = "reports/assignments/figures"
os.makedirs(OUT, exist_ok=True)

BLUE = "#2a78d6"
ORANGE = "#eb6834"
AQUA = "#1baf7a"
GREY = "#8a8a86"
RED = "#c0392b"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#52514e", "axes.labelcolor": "#0b0b0b",
    "text.color": "#0b0b0b", "xtick.color": "#52514e", "ytick.color": "#52514e",
})

# ---------------------------------------------------------------------
# Full sensitivity heatmap: extra capacity (CP-constrained - point-only),
# with infeasible (CP-constrained) cells marked explicitly
# ---------------------------------------------------------------------
opt = pd.read_csv("results/tables/capacity_optimization.csv")
w_maxes = sorted(opt["W_max_minutes"].unique())
arrivals = sorted(opt["arrival_rate_multiplier"].unique())

grid = np.full((len(w_maxes), len(arrivals)), np.nan)
infeasible = np.zeros_like(grid, dtype=bool)
for i, w in enumerate(w_maxes):
    for j, a in enumerate(arrivals):
        row = opt[(opt["W_max_minutes"] == w) & (opt["arrival_rate_multiplier"] == a)]
        if len(row) == 0:
            continue
        cp_cap = row["n_capacity_star_cp_constrained"].values[0]
        extra = row["extra_capacity_from_using_cp_bound"].values[0]
        if pd.isna(cp_cap):
            infeasible[i, j] = True
        else:
            grid[i, j] = extra

fig, ax = plt.subplots(figsize=(7, 4.5))
masked = np.ma.masked_invalid(grid)
im = ax.imshow(masked, cmap="YlOrRd", aspect="auto", vmin=0, vmax=26)
ax.set_xticks(range(len(arrivals))); ax.set_xticklabels(arrivals)
ax.set_yticks(range(len(w_maxes))); ax.set_yticklabels(w_maxes)
ax.set_xlabel("Arrival-rate multiplier")
ax.set_ylabel("Policy wait-time ceiling W_max (minutes)")
for i in range(len(w_maxes)):
    for j in range(len(arrivals)):
        if infeasible[i, j]:
            ax.text(j, i, "infeasible", ha="center", va="center", fontsize=7.5,
                     color="white", fontweight="bold",
                     bbox=dict(boxstyle="round", facecolor="#7a1f1f", edgecolor="none"))
        elif not np.isnan(grid[i, j]):
            ax.text(j, i, f"+{grid[i,j]:.0f}", ha="center", va="center", fontsize=9, fontweight="bold")
cbar = fig.colorbar(im, ax=ax, shrink=0.85)
cbar.set_label("Extra capacity, CP-constrained vs. point-only plan")
ax.set_title("Where a point-prediction-only plan understates required capacity\n(dark red = no capacity in [15,45] can honestly guarantee this ceiling)", fontsize=10)
fig.tight_layout()
fig.savefig(f"{OUT}/capacity_sensitivity_heatmap.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Deployment architecture / dataflow diagram
# ---------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(10, 4.2))
ax.set_xlim(0, 10)
ax.set_ylim(0, 4)
ax.axis("off")

boxes = [
    (0.3, 1.5, 1.7, 1.2, "Real ED data\n(arrivals, ESI mix)\nSection 5.1", "#e8f4fc"),
    (2.4, 1.5, 1.7, 1.2, "Calibrated DES\nSection 4.2", "#e8f4fc"),
    (4.5, 1.5, 1.7, 1.2, "Surrogate +\nMondrian CP\nSections 4.3-4.4", "#e6f7ef"),
    (6.6, 2.6, 1.7, 1.1, "Ops dashboard\nSection 10.1", "#fdeee6"),
    (6.6, 0.5, 1.7, 1.1, "Capacity optimizer\nSection 10.2", "#fdeee6"),
    (8.7, 1.5, 1.1, 1.2, "Staffing\ndecision", "#f5f0fa"),
]
for x, y, w, h, label, color in boxes:
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05,rounding_size=0.08",
                          linewidth=1.2, edgecolor="#52514e", facecolor=color)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, label, ha="center", va="center", fontsize=8.5)

arrows = [
    (2.0, 2.1, 2.4, 2.1),
    (4.1, 2.1, 4.5, 2.1),
    (6.2, 2.1, 6.6, 3.15),
    (6.2, 2.1, 6.6, 1.05),
    (8.3, 3.15, 8.7, 2.4),
    (8.3, 1.05, 8.7, 1.8),
]
for x0, y0, x1, y1 in arrows:
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=14,
                                   color="#52514e", linewidth=1.3))

ax.text(5, 3.6, "Every arrow is real: trained models and calibrated quantiles this report\nderives and evaluates, not a proposed or mocked pipeline (Section 5.5.1).",
        ha="center", fontsize=8.5, style="italic", color="#40403c")
ax.set_title("Chapter 10 deployment dataflow: from real data to a staffing decision", fontsize=11)
fig.tight_layout()
fig.savefig(f"{OUT}/deployment_architecture.png", dpi=170, bbox_inches="tight")
plt.close(fig)

print("Chapter 10 figures generated in", OUT)
for f in ["capacity_sensitivity_heatmap.png", "deployment_architecture.png"]:
    print(" -", f, "OK" if os.path.exists(f"{OUT}/{f}") else "MISSING")
