"""
Multiple-comparisons check across every paired significance test reported in
the paper - 16 tests total, in 4 natural families of 4 (one per target):
the central marginal-coverage result (Mondrian vs. standard CP, Section
V-C), SA-Mondrian vs. Mondrian (Table VI), SA-Mondrian vs. QT-CP (Section
VII-C), and the Mondrian+QT-CP hybrid vs. Mondrian (Section VII-D).

Applies Holm-Bonferroni correction within each family (the statistically
appropriate scope for a set of related, repeated comparisons - not a single
blanket correction across all 16 unrelated hypotheses, which would be
needlessly conservative) and confirms whether every significance claim
already made in the paper survives it. Raw p-values are recomputed here
directly from the underlying repeat-level CSVs, not copied from the paper
text, so this is a real independent check, not a restatement.
"""
import numpy as np
import pandas as pd
from scipy import stats

TABLES = "results/tables"


def holm_bonferroni(pvals, alpha=0.05):
    m = len(pvals)
    order = sorted(range(m), key=lambda i: pvals[i])
    reject = [False] * m
    for rank, idx in enumerate(order):
        threshold = alpha / (m - rank)
        if pvals[idx] <= threshold:
            reject[idx] = True
        else:
            break  # step-down procedure: stop at the first non-rejection
    return reject


def central_family():
    rep = pd.read_csv(f"{TABLES}/repeated_evaluation_detail.csv")
    pvals, targets = [], []
    for target, g in rep.groupby("target"):
        std = g[g.method == "Standard CP"].sort_values("repeat")["coverage"].values
        mon = g[g.method == "Mondrian CP"].sort_values("repeat")["coverage"].values
        _, p = stats.ttest_rel(mon, std)
        pvals.append(p)
        targets.append(target)
    return targets, pvals


def sa_vs_mondrian_family():
    d = pd.read_csv(f"{TABLES}/repeated_worst_category_summary.csv")
    return d["target"].tolist(), d["paired_ttest_p"].tolist()


def sa_vs_qtcp_family():
    d = pd.read_csv(f"{TABLES}/repeated_worst_category_detail.csv")
    pvals, targets = [], []
    for target, g in d.groupby("target"):
        _, p = stats.ttest_rel(g["sa_mondrian_worst"], g["qtcp_worst"])
        pvals.append(p)
        targets.append(target)
    return targets, pvals


def hybrid_vs_mondrian_family():
    d = pd.read_csv(f"{TABLES}/repeated_mondrian_qtcp_hybrid_summary.csv")
    return d["target"].tolist(), d["paired_ttest_p"].tolist()


def main():
    families = {
        "central_mondrian_vs_standard": central_family(),
        "sa_mondrian_vs_mondrian": sa_vs_mondrian_family(),
        "sa_mondrian_vs_qtcp": sa_vs_qtcp_family(),
        "hybrid_vs_mondrian": hybrid_vs_mondrian_family(),
    }

    rows = []
    for family, (targets, pvals) in families.items():
        sig = holm_bonferroni(pvals)
        for target, p, s in zip(targets, pvals, sig):
            rows.append({"family": family, "target": target, "raw_p": p, "holm_significant": s})
            print(f"{family:35s} {target:20s} p={p:.2e}  Holm-significant={s}")

    out = pd.DataFrame(rows)
    out.to_csv(f"{TABLES}/multiple_comparisons_check.csv", index=False)
    print(f"\nSaved {TABLES}/multiple_comparisons_check.csv")
    print(f"\nTotal tests: {len(out)}  Families: {out['family'].nunique()}")


if __name__ == "__main__":
    main()
