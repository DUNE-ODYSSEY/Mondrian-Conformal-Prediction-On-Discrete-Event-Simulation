"""
One new figure for the paper: a methodology pipeline diagram (DES ->
surrogate -> standard/Mondrian CP -> evaluation). Every other figure in
the paper is reused directly from reports/assignments/figures/ (already
generated and verified for the book report) - this is the only genuinely
new visual, since no existing figure shows the end-to-end pipeline.
"""
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import os

OUT = "reports/paper/figures"
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9,
})

fig, ax = plt.subplots(figsize=(7, 3.4))
ax.set_xlim(0, 10)
ax.set_ylim(0, 4.6)
ax.axis("off")

boxes = [
    (0.2, 1.5, 1.9, 1.2, "Real ED data\n(560,486 visits)", "#e8f4fc"),
    (2.5, 1.5, 1.9, 1.2, "Calibrated DES\n(SimPy)", "#e8f4fc"),
    (4.8, 1.5, 1.9, 1.2, "Surrogate\n(gradient boosting)", "#e6f7ef"),
    (7.1, 2.6, 2.7, 1.1, "Standard CP\n(pooled quantile)", "#fdeee6"),
    (7.1, 0.5, 2.7, 1.1, "Mondrian CP\n(per-category quantile)", "#fdeee6"),
]
for x, y, w, h, label, color in boxes:
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05,rounding_size=0.08",
                          linewidth=1.1, edgecolor="#40403c", facecolor=color)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, label, ha="center", va="center", fontsize=8.5)

arrows = [
    (2.1, 2.1, 2.5, 2.1),
    (4.4, 2.1, 4.8, 2.1),
    (6.7, 2.1, 7.1, 3.15),
    (6.7, 2.1, 7.1, 1.05),
]
for x0, y0, x1, y1 in arrows:
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle="-|>", mutation_scale=13,
                                   color="#40403c", linewidth=1.2))

ax.text(8.45, 4.3, "Section V-A/C: marginal coverage", ha="center", fontsize=7.5, style="italic", color="#40403c")
ax.text(8.45, 4.05, "Section V-B/D: per-category coverage", ha="center", fontsize=7.5, style="italic", color="#40403c")

fig.tight_layout()
fig.savefig(f"{OUT}/methodology_pipeline.png", dpi=200, bbox_inches="tight")
plt.close(fig)

print("Paper figure generated:", f"{OUT}/methodology_pipeline.png")
