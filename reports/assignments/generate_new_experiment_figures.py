"""
Figures for the new empirical work (CRC, ACI, weighted CP, 5-architecture
surrogate benchmark, capacity optimization) added to the book report,
built from the real results tables those scripts produced under
results/tables/ - no new numbers invented here, purely visualization.
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
# Figure: CRC test risk vs target, both losses, against alpha
# ---------------------------------------------------------------------
crc = pd.read_csv("results/tables/crc_results.csv")
fig, ax = plt.subplots(figsize=(7, 4))
x = np.arange(len(crc))
w = 0.35
ax.bar(x - w/2, crc["loss01_test_risk"], w, label="0/1 upper-miscoverage loss", color=BLUE)
ax.bar(x + w/2, crc["sev_test_risk"], w, label="Clipped overshoot-severity loss", color=ORANGE)
ax.axhline(0.1, color=RED, linestyle="--", linewidth=1.2, label="alpha = 0.10 target")
ax.set_xticks(x); ax.set_xticklabels(crc["target"], rotation=15)
ax.set_ylabel("Held-out test risk")
ax.set_title("Conformal Risk Control: test-set risk stays at or below alpha")
ax.legend(fontsize=8)
fig.tight_layout()
fig.savefig(f"{OUT}/crc_test_risk.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure: ACI rolling coverage over the severity-ramp stream, static vs ACI
# ---------------------------------------------------------------------
traj = pd.read_csv("results/tables/aci_alpha_trajectory.csv")
targets = traj["target"].unique()
fig, axes = plt.subplots(1, len(targets), figsize=(15, 3.6), sharey=True)
for ax, t in zip(axes, targets):
    sub = traj[traj["target"] == t]
    ax.plot(sub["step"], sub["rolling_coverage_static"] * 100, color=GREY, label="Static CP", linewidth=1.3)
    ax.plot(sub["step"], sub["rolling_coverage_aci"] * 100, color=AQUA, label="ACI", linewidth=1.3)
    first_oor = sub[~sub["in_range"]]["step"].min()
    ax.axvline(first_oor, color="#c9c8c0", linestyle="--", linewidth=1)
    ax.axhline(90, color=RED, linestyle=":", linewidth=1)
    ax.set_title(t, fontsize=9)
    ax.set_xlabel("Stream step (severity ramping up)")
axes[0].set_ylabel("Rolling coverage, 60-step window (%)")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.08), frameon=False, fontsize=9)
fig.suptitle("Adaptive Conformal Inference vs. static CP under a live demand-surge stream\n"
             "(dashed vertical = leaves training range; dotted horizontal = 90% target)", fontsize=10, y=1.1)
fig.tight_layout()
fig.savefig(f"{OUT}/aci_rolling_coverage.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure: weighted CP coverage by region, unweighted vs weighted
# ---------------------------------------------------------------------
wcp = pd.read_csv("results/tables/weighted_cp_results.csv")
wcp_overall_regions = wcp[wcp["region"] != "overall"]
fig, ax = plt.subplots(figsize=(8, 4.2))
targets4 = wcp_overall_regions["target"].unique()
regions = ["overlap_region", "out_of_support_tail"]
x = np.arange(len(targets4))
w = 0.2
for i, (region, color_u, color_w) in enumerate(zip(regions, [GREY, GREY], [BLUE, MAGENTA])):
    sub = wcp_overall_regions[wcp_overall_regions["region"] == region].set_index("target").loc[targets4]
    offset = (i - 0.5) * 2 * w
    ax.bar(x + offset - w/2, sub["unweighted_coverage"] * 100, w, color=color_u, alpha=0.55,
           label="Unweighted CP" if i == 0 else None)
    ax.bar(x + offset + w/2, sub["weighted_coverage"] * 100, w, color=color_w,
           label=f"Weighted CP ({region.replace('_',' ')})")
ax.axhline(90, color=RED, linestyle="--", linewidth=1, label="90% target")
ax.set_xticks(x); ax.set_xticklabels(targets4, rotation=15)
ax.set_ylabel("Coverage (%)")
ax.set_title("Likelihood-ratio weighted CP under a moderate covariate shift, by region")
ax.legend(fontsize=7, loc="lower right")
fig.tight_layout()
fig.savefig(f"{OUT}/weighted_cp_coverage.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure: 5-architecture surrogate R2 comparison
# ---------------------------------------------------------------------
gbr = pd.read_csv("results/tables/surrogate_metrics.csv"); gbr["architecture"] = "GBR (primary)"
mlp = pd.read_csv("results/tables/mlp_surrogate_metrics.csv"); mlp["architecture"] = "MLP"
other = pd.read_csv("results/tables/rf_xgb_lgb_surrogate_metrics.csv")
allarch = pd.concat([gbr[["architecture", "target", "r2"]], mlp[["architecture", "target", "r2"]],
                      other[["architecture", "target", "r2"]]], ignore_index=True)
arch_order = ["GBR (primary)", "MLP", "RandomForest", "XGBoost", "LightGBM"]
colors5 = [BLUE, ORANGE, GREY, YELLOW, AQUA]
fig, ax = plt.subplots(figsize=(9, 4.5))
targets_order = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
x = np.arange(len(targets_order))
w = 0.16
for i, (arch, color) in enumerate(zip(arch_order, colors5)):
    vals = [allarch[(allarch.architecture == arch) & (allarch.target == t)]["r2"].values[0] for t in targets_order]
    ax.bar(x + (i - 2) * w, vals, w, label=arch, color=color)
ax.set_xticks(x); ax.set_xticklabels(targets_order, rotation=10)
ax.set_ylabel("R-squared (held-out test set)")
ax.set_title("Surrogate accuracy across five architectures, identical train/test split")
ax.legend(fontsize=8, ncol=5, loc="lower center", bbox_to_anchor=(0.5, -0.35))
fig.tight_layout()
fig.savefig(f"{OUT}/five_architecture_r2.png", dpi=170, bbox_inches="tight")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure: capacity optimization - CP-constrained vs point-prediction-only
# ---------------------------------------------------------------------
opt = pd.read_csv("results/tables/capacity_optimization.csv")
fig, axes = plt.subplots(1, 2, figsize=(11, 4))
for ax, w_max in zip(axes, [90, 180]):
    sub = opt[opt["W_max_minutes"] == w_max].sort_values("arrival_rate_multiplier")
    ax.plot(sub["arrival_rate_multiplier"], sub["n_capacity_star_cp_constrained"], "o-", color=BLUE, label="CP-constrained (safe)")
    ax.plot(sub["arrival_rate_multiplier"], sub["n_capacity_star_point_prediction_only"], "s--", color=ORANGE, label="Point-prediction only")
    ax.set_title(f"W_max = {w_max} min", fontsize=10)
    ax.set_xlabel("Arrival-rate multiplier")
    if ax is axes[0]:
        ax.set_ylabel("Minimum feasible staffing capacity")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.05), frameon=False, fontsize=9)
fig.suptitle("Prescriptive capacity allocation: CP-constrained vs. point-prediction-only planning", fontsize=11, y=1.03)
fig.tight_layout()
fig.savefig(f"{OUT}/capacity_optimization.png", dpi=170, bbox_inches="tight")
plt.close(fig)

print("All new-experiment figures generated in", OUT)
for f in sorted(os.listdir(OUT)):
    print(" -", f)
