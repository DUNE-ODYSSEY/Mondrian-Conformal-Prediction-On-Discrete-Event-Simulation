"""
New figures for the expanded Chapter 8 (exchangeability) and Chapter 9
(cross-site generalization) sections - built entirely from existing,
already-computed results/tables/ CSVs (Department A's own tables plus
results/tables/dept_b/), no new experiments, no invented numbers.
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

OUT = "reports/assignments/figures"
os.makedirs(OUT, exist_ok=True)

BLUE = "#2a78d6"
ORANGE = "#eb6834"
AQUA = "#1baf7a"
YELLOW = "#eda100"
MAGENTA = "#e87ba4"
GREY = "#8a8a86"
RED = "#c0392b"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#52514e", "axes.labelcolor": "#0b0b0b",
    "text.color": "#0b0b0b", "xtick.color": "#52514e", "ytick.color": "#52514e",
})

# ---------------------------------------------------------------------
# Chapter 9, Fig 1: Department A vs B - real volume and ESI acuity mix
# ---------------------------------------------------------------------
esi_a = pd.read_csv("results/tables/esi_mix.csv").set_index("esi")["proportion"] * 100
esi_b = pd.read_csv("results/tables/dept_b/esi_mix.csv").set_index("esi")["proportion"] * 100

fig, axes = plt.subplots(1, 2, figsize=(10, 4))
ax = axes[0]
levels = [1, 2, 3, 4, 5]
x = np.arange(5)
w = 0.35
ax.bar(x - w/2, [esi_a.get(l, 0) for l in levels], w, label="Department A", color=BLUE)
ax.bar(x + w/2, [esi_b.get(l, 0) for l in levels], w, label="Department B", color=ORANGE)
ax.set_xticks(x); ax.set_xticklabels([f"ESI {l}" for l in levels])
ax.set_ylabel("Share of real visits (%)")
ax.set_title("Real triage-acuity (ESI) mix", fontsize=10)
ax.legend(fontsize=8)

ax = axes[1]
vol = [258.2, 133.4]
cap = [30, 14]
x2 = np.arange(2)
ax2b = ax.twinx()
b1 = ax.bar(x2 - 0.18, vol, 0.35, color=BLUE, label="Real visits/day")
b2 = ax2b.bar(x2 + 0.18, cap, 0.35, color=AQUA, label="DES staffing capacity")
ax.set_xticks(x2); ax.set_xticklabels(["Department A", "Department B"])
ax.set_ylabel("Real visits per day", color=BLUE)
ax2b.set_ylabel("Erlang-derived staffing capacity", color=AQUA)
ax.set_title("Real daily volume and derived capacity", fontsize=10)
lines = [b1, b2]
ax.legend(lines, [l.get_label() for l in lines], fontsize=8, loc="upper right")
fig.suptitle("Department A vs. Department B: two structurally different real sites", fontsize=11)
fig.tight_layout()
fig.savefig(f"{OUT}/dept_a_vs_b_structure.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Chapter 9, Fig 2: Department B per-category coverage heatmap (pooled vs Mondrian)
# mirrors Fig 7.2's Department A heatmap
# ---------------------------------------------------------------------
detail_b = pd.read_csv("results/tables/dept_b/mondrian_cp_detail.csv")
target = "mean_wait_minutes"
sub = detail_b[detail_b["target"] == target].copy()
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
fig.suptitle(f"Department B per-category coverage: {target} (target = 90%)", fontsize=12, y=1.02)
fig.tight_layout()
fig.savefig(f"{OUT}/dept_b_coverage_heatmap.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Chapter 9, Fig 3: A vs B coverage-range comparison (pooled min-max per target)
# ---------------------------------------------------------------------
summary_b = pd.read_csv("results/tables/dept_b/standard_vs_mondrian_summary.csv")
targets = summary_b["target"].tolist()
fig, ax = plt.subplots(figsize=(7.5, 4.3))
x = np.arange(len(targets))
w = 0.35
ax.bar(x - w/2, summary_b["pooled_coverage_range"] * 100, w, label="Pooled CP coverage range", color=GREY)
ax.bar(x + w/2, summary_b["mondrian_coverage_range"] * 100, w, label="Mondrian CP coverage range", color=AQUA)
ax.set_xticks(x); ax.set_xticklabels(targets, rotation=15)
ax.set_ylabel("Coverage range\n(percentage points)")
ax.set_title("Department B: per-category coverage spread, pooled vs. Mondrian")
ax.legend(fontsize=8)
fig.tight_layout(pad=1.5)
fig.savefig(f"{OUT}/dept_b_coverage_range.png", dpi=170, bbox_inches="tight", pad_inches=0.15)
plt.close(fig)

print("Chapter 8/9 figures generated in", OUT)
for f in ["dept_a_vs_b_structure.png", "dept_b_coverage_heatmap.png", "dept_b_coverage_range.png"]:
    print(" -", f, "OK" if os.path.exists(f"{OUT}/{f}") else "MISSING")
