"""
Further Chapter 8 figures: a conceptual calibration-vs-test-support
overlap diagram (explains the mechanism, not decorative), and a
Mondrian-vs-standard-CP coverage comparison under the severity sweep,
built from the existing exchangeability_stress_test*.csv tables - no new
experiments.
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
GREY = "#8a8a86"
RED = "#c0392b"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#52514e", "axes.labelcolor": "#0b0b0b",
    "text.color": "#0b0b0b", "xtick.color": "#52514e", "ytick.color": "#52514e",
})

# ---------------------------------------------------------------------
# Conceptual diagram: calibration support vs. severity-sweep test points
# ---------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(9, 3.2))

cal_lo, cal_hi = 0.8, 1.3
ax.axvspan(cal_lo, cal_hi, color=AQUA, alpha=0.25, label="Calibration support [0.8, 1.3]")
ax.plot([cal_lo, cal_hi], [0.5, 0.5], color=AQUA, linewidth=6, solid_capstyle="round")

severities = [0.8, 1.0, 1.3, 1.5, 1.8, 2.0, 2.5, 3.0]
in_range = [s <= 1.3 for s in severities]
colors = [AQUA if r else RED for r in in_range]
ax.scatter(severities, [0.5]*len(severities), c=colors, s=90, zorder=5, edgecolor="white", linewidth=1)
for s, r in zip(severities, in_range):
    ax.annotate(f"{s}x", (s, 0.5), textcoords="offset points", xytext=(0, 16),
                ha="center", fontsize=9, color=(AQUA if r else RED), fontweight="bold")

ax.axvline(1.3, color="#0b0b0b", linestyle="--", linewidth=1)
ax.text(1.3, 0.15, "training\nboundary", ha="center", fontsize=8, color="#0b0b0b")
ax.set_xlim(0.6, 3.15)
ax.set_ylim(0, 1)
ax.set_yticks([])
ax.set_xlabel("Arrival-rate multiplier")
ax.set_title("Exchangeability violated: test points sampled far outside calibration's own support", fontsize=11)
handles = [plt.Rectangle((0,0),1,1, color=AQUA, alpha=0.25), plt.scatter([],[],c=AQUA,s=60), plt.scatter([],[],c=RED,s=60)]
ax.legend(["Calibration support", "In-range test severity", "Out-of-range test severity"],
          loc="upper right", fontsize=8, frameon=False)
fig.tight_layout()
fig.savefig(f"{OUT}/exchangeability_support_diagram.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Mondrian vs standard CP coverage under the severity sweep (GBR)
# ---------------------------------------------------------------------
gbr = pd.read_csv("results/tables/exchangeability_stress_test.csv")
targets = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
fig, axes = plt.subplots(1, 4, figsize=(14, 3.6), sharex=True)
for ax, t in zip(axes, targets):
    sub = gbr[gbr["target"] == t].sort_values("arrival_rate_multiplier")
    ax.plot(sub["arrival_rate_multiplier"], sub["standard_coverage"] * 100, "o-", color=BLUE, label="Standard CP")
    ax.plot(sub["arrival_rate_multiplier"], sub["mondrian_coverage"] * 100, "s-", color=AQUA, label="Mondrian CP")
    ax.axvline(1.3, color="#c9c8c0", linestyle="--", linewidth=1)
    ax.axhline(90, color=RED, linestyle=":", linewidth=1)
    ax.set_title(t, fontsize=9)
    ax.set_xlabel("Arrival-rate multiplier", fontsize=8)
    if ax is axes[0]:
        ax.set_ylabel("Coverage (%)")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.1), frameon=False, fontsize=9)
fig.suptitle("Mondrian CP tracks standard CP's collapse - per-category structure offers no protection\n"
             "(dashed vertical = training boundary; dotted horizontal = 90% target)", fontsize=10, y=1.1)
fig.tight_layout()
fig.savefig(f"{OUT}/mondrian_vs_standard_under_shift.png", dpi=170, bbox_inches="tight")
plt.close(fig)

print("Chapter 8 extra figures generated in", OUT)
for f in ["exchangeability_support_diagram.png", "mondrian_vs_standard_under_shift.png"]:
    print(" -", f, "OK" if os.path.exists(f"{OUT}/{f}") else "MISSING")
