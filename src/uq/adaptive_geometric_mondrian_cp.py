"""
Sample-size-floored adaptive binning: the fix motivated by why plain
geometric Mondrian binning (geometric_mondrian_cp.py) lost to equal-width
binning on real data despite provably lower scale distortion - geometric
bins concentrate many bins in the sparse rho -> 1 tail, where this
project's real scenario sweep has few calibration points, so quantile
estimation noise swamps the scale-homogeneity benefit.

This scheme builds bins from the most extreme (highest rho_hat) end inward,
each guaranteed at least n_min calibration points before it closes - i.e.
bin WIDTH adapts to local calibration density (narrow where data is dense,
wide where it's sparse) rather than being fixed a priori by either equal
spacing or a geometric schedule. The resulting worst-case scale ratio per
bin is then measured, not assumed - an honest empirical check of whether
respecting the sample-size constraint still leaves the scale-ratio blowup
under control in practice, or merely trades one failure mode for another.
"""
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
import joblib

from src.des.er_simulation import STUDY_PERIOD_DAYS
from src.uq.mondrian_cp import conformal_quantile
from src.uq.queueing_weighted_cp import estimate_rho, mean_service_time_minutes

TABLES = "results/tables"
TARGETS = ["mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes", "n_patients"]
ALPHA = 0.1
N_MIN_GRID = [40, 60, 80, 100, 130]


def build_adaptive_bins(rho_cal, n_min):
    """Sort calibration points by rho descending, greedily close a bin the
    moment it reaches n_min points, working from the riskiest (highest rho)
    end inward. Returns bin edges (rho thresholds, descending)."""
    order = np.argsort(-rho_cal)
    sorted_rho = rho_cal[order]
    edges = [np.inf]
    count = 0
    for i, r in enumerate(sorted_rho):
        count += 1
        remaining = len(sorted_rho) - (i + 1)
        if count >= n_min and remaining >= n_min:
            edges.append(r)
            count = 0
    edges.append(-np.inf)
    return np.array(edges)  # descending: edges[0]=+inf > edges[1] > ... > edges[-1]=-inf


def assign_from_edges(rho, edges):
    # edges descending; bin b covers (edges[b+1], edges[b]]
    bins = np.zeros(len(rho), dtype=int)
    for b in range(len(edges) - 1):
        mask = (rho <= edges[b]) & (rho > edges[b + 1])
        bins[mask] = b
    return bins


def worst_bin_coverage(rho_fit, resid_fit, rho_eval, resid_eval, n_min):
    """One evaluation of the adaptive scheme's worst-bin coverage, edges
    built from *_fit only, scored on *_eval - the primitive both the
    calibration-only selector and the final honest test-set score reuse."""
    edges = build_adaptive_bins(rho_fit, n_min)
    fit_bins = assign_from_edges(rho_fit, edges)
    eval_bins = assign_from_edges(rho_eval, edges)
    coverages = []
    for b in range(len(edges) - 1):
        fit_mask = fit_bins == b
        eval_mask = eval_bins == b
        if fit_mask.sum() < 5 or eval_mask.sum() == 0:
            continue
        q_b = conformal_quantile(resid_fit[fit_mask], ALPHA)
        coverages.append((resid_eval[eval_mask] <= q_b).mean())
    return (min(coverages) if coverages else np.nan), len(edges) - 1


def select_n_min(rho_cal, abs_resid_cal, seed=0):
    """Choose n_min using only calibration data, split 70/30: fit bins +
    quantiles on the 70% slice, score worst-bin coverage on the held-out
    30%, pick the n_min with the best held-out worst-bin coverage - the
    same nested-calibration pattern already used to select rho_cap for
    QT-CP (src/uq/queueing_weighted_cp.py), so this method's own
    hyperparameter never sees the test set either."""
    rng = np.random.default_rng(seed)
    n = len(rho_cal)
    idx = rng.permutation(n)
    fit_idx, sel_idx = idx[: int(0.7 * n)], idx[int(0.7 * n):]

    best_n_min, best_score = N_MIN_GRID[0], -np.inf
    for n_min in N_MIN_GRID:
        score, _ = worst_bin_coverage(rho_cal[fit_idx], abs_resid_cal[fit_idx],
                                       rho_cal[sel_idx], abs_resid_cal[sel_idx], n_min)
        if not np.isnan(score) and score > best_score:
            best_n_min, best_score = n_min, score
    return best_n_min


def main():
    arrivals = pd.read_csv(f"{TABLES}/arrivals_by_hour_bin.csv", index_col=0)["count"]
    esi_mix = pd.read_csv(f"{TABLES}/esi_mix.csv", index_col=0)["proportion"].to_dict()
    baseline_visits_per_day = arrivals.sum() / STUDY_PERIOD_DAYS
    mean_service_time = mean_service_time_minutes(esi_mix)

    df = pd.read_parquet("data/processed/surrogate_training_data.parquet")
    _, X_test, _, test_idx = train_test_split(df[["n_capacity", "arrival_rate_multiplier"]], df.index,
                                               test_size=0.2, random_state=42)
    test_df = df.loc[test_idx]
    cal_df = pd.read_parquet("data/processed/cp_calibration_data.parquet")

    rho_cal = estimate_rho(cal_df["n_capacity"], cal_df["arrival_rate_multiplier"],
                            baseline_visits_per_day, mean_service_time).values
    rho_test = estimate_rho(test_df["n_capacity"], test_df["arrival_rate_multiplier"],
                             baseline_visits_per_day, mean_service_time).values

    diagnostic_rows, headline_rows = [], []
    for target in TARGETS:
        model = joblib.load(f"models/surrogate_{target}.joblib")
        yhat_cal = model.predict(cal_df[["n_capacity", "arrival_rate_multiplier"]])
        y_cal = cal_df[target].values
        abs_resid_cal = np.abs(y_cal - yhat_cal)

        yhat_test = model.predict(X_test)
        y_test = test_df[target].values
        abs_resid_test = np.abs(y_test - yhat_test)

        # Diagnostic-only sweep across n_min, test-set scored - kept purely
        # to show the bias/variance trade-off shape, NEVER used to pick the
        # method's actual hyperparameter (that would leak the test set into
        # a design choice, which this project avoids everywhere else).
        print(f"\n[diagnostic only - not used for selection] {target}")
        for n_min in N_MIN_GRID:
            cov, n_bins = worst_bin_coverage(rho_cal, abs_resid_cal, rho_test, abs_resid_test, n_min)
            diagnostic_rows.append({"target": target, "n_min": n_min, "n_bins_produced": n_bins,
                                     "test_worst_coverage_DIAGNOSTIC_ONLY": cov})
            print(f"  n_min={n_min:4d} -> {n_bins:2d} bins  worst_coverage={cov:.3f}")

        # Real, honest result: n_min selected from calibration data alone,
        # scored once on the test set with that single selected value.
        selected_n_min = select_n_min(rho_cal, abs_resid_cal)
        final_cov, final_n_bins = worst_bin_coverage(rho_cal, abs_resid_cal, rho_test, abs_resid_test,
                                                       selected_n_min)
        headline_rows.append({
            "target": target, "selected_n_min": selected_n_min, "n_bins": final_n_bins,
            "adaptive_worst_category_coverage": final_cov,
        })
        print(f"  SELECTED (calibration-only) n_min={selected_n_min} -> {final_n_bins} bins, "
              f"real test worst-category coverage = {final_cov:.3f}")

    pd.DataFrame(diagnostic_rows).to_csv(f"{TABLES}/adaptive_geometric_mondrian_cp_diagnostic.csv", index=False)
    headline_df = pd.DataFrame(headline_rows)
    headline_df.to_csv(f"{TABLES}/adaptive_geometric_mondrian_cp.csv", index=False)
    print(f"\nSaved {TABLES}/adaptive_geometric_mondrian_cp.csv (headline, calibration-selected)")
    print(f"Saved {TABLES}/adaptive_geometric_mondrian_cp_diagnostic.csv (diagnostic sweep, not for selection)")

    existing = pd.read_csv(f"{TABLES}/queueing_weighted_cp_summary.csv")
    comparison = headline_df.merge(
        existing[["target", "mondrian_worst_category_coverage", "qtcp_worst_category_coverage"]],
        on="target")
    print("\n=== Honest comparison, worst-category coverage (target 90%) ===")
    print(comparison.to_string(index=False))


if __name__ == "__main__":
    main()
