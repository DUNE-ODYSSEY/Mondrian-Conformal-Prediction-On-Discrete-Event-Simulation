"""
Geometric Mondrian CP: a Mondrian-CP variant whose bin edges are spaced
geometrically in (1 - rho_hat) rather than equally in rho_hat itself.

Motivation (proved in the book/paper's new theoretical section): under the
heavy-traffic scale model sigma(rho) = kappa / (1 - rho) implied by
Kingman's heavy-traffic limit theorem, an equal-width bin adjacent to the
rho -> 1 singularity has a within-bin scale ratio that grows without bound
as the bin approaches rho = 1 - for a FIXED bin count B, this ratio is
Theta(1 / (1 - rho_max)); to keep it below any fixed constant, the number
of equal-width bins needed grows the same way, Theta(1 / (1 - rho_max)).
Bins spaced geometrically in (1 - rho) (edges at 1 - rho = 2^-k) instead
hold a CONSTANT within-bin scale ratio of exactly 2 no matter how close a
bin sits to the singularity, needing only O(log(1/(1-rho_max))) bins to
cover the same range - exponentially fewer bins for the same worst-case
distortion bound.

This script tests whether that real, provable structural advantage
actually translates into better empirical worst-category coverage,
compared honestly against standard (equal-width) rho-binning and this
project's own existing 2D staffing x arrival-rate tercile Mondrian CP, on
the same real calibration/test split every other method here uses.
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
BIN_COUNTS = [3, 5, 9, 15]
RHO_CAP_FOR_BINNING = 0.98  # matches the cap already used elsewhere to keep 1/(1-rho) finite


def equal_width_bins(rho_capped, n_bins, rho_min, rho_max):
    edges = np.linspace(rho_min, rho_max, n_bins + 1)
    edges[-1] += 1e-9
    return np.digitize(rho_capped, edges[1:-1])


def geometric_bins(rho_capped, n_bins, rho_min, rho_max):
    # edges equally spaced in log(1-rho) between rho_min and rho_max
    lo, hi = np.log(1 - rho_max), np.log(1 - rho_min)  # note: rho_max -> smaller (1-rho) -> more negative log
    edges = np.exp(np.linspace(hi, lo, n_bins + 1))  # decreasing (1-rho) edges
    rho_edges = 1 - edges
    rho_edges[0] -= 1e-9
    rho_edges[-1] += 1e-9
    return np.digitize(rho_capped, np.sort(rho_edges)[1:-1])


def evaluate_binning(bin_fn, n_bins, rho_cal, rho_test, abs_resid_cal, abs_resid_test, rho_min, rho_max):
    rho_cal_capped = np.minimum(rho_cal, rho_max)
    rho_test_capped = np.minimum(rho_test, rho_max)
    cal_bins = bin_fn(rho_cal_capped, n_bins, rho_min, rho_max)
    test_bins = bin_fn(rho_test_capped, n_bins, rho_min, rho_max)

    coverages, ns = [], []
    for b in range(n_bins):
        cal_mask = cal_bins == b
        test_mask = test_bins == b
        if cal_mask.sum() < 5 or test_mask.sum() == 0:
            continue
        q_b = conformal_quantile(abs_resid_cal[cal_mask], ALPHA)
        cov_b = (abs_resid_test[test_mask] <= q_b).mean()
        coverages.append(cov_b)
        ns.append(test_mask.sum())
    coverages = np.array(coverages)
    return coverages.min() if len(coverages) else np.nan, coverages.max() if len(coverages) else np.nan


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

    rho_min = min(rho_cal.min(), rho_test.min())
    rho_max = RHO_CAP_FOR_BINNING

    rows = []
    for target in TARGETS:
        model = joblib.load(f"models/surrogate_{target}.joblib")
        yhat_cal = model.predict(cal_df[["n_capacity", "arrival_rate_multiplier"]])
        y_cal = cal_df[target].values
        abs_resid_cal = np.abs(y_cal - yhat_cal)

        yhat_test = model.predict(X_test)
        y_test = test_df[target].values
        abs_resid_test = np.abs(y_test - yhat_test)

        for n_bins in BIN_COUNTS:
            ew_worst, ew_best = evaluate_binning(equal_width_bins, n_bins, rho_cal, rho_test,
                                                  abs_resid_cal, abs_resid_test, rho_min, rho_max)
            geo_worst, geo_best = evaluate_binning(geometric_bins, n_bins, rho_cal, rho_test,
                                                    abs_resid_cal, abs_resid_test, rho_min, rho_max)
            rows.append({
                "target": target, "n_bins": n_bins,
                "equal_width_worst_coverage": ew_worst, "equal_width_best_coverage": ew_best,
                "geometric_worst_coverage": geo_worst, "geometric_best_coverage": geo_best,
            })
            print(f"{target:20s} B={n_bins:3d}  equal-width worst={ew_worst:.3f}  "
                  f"geometric worst={geo_worst:.3f}  "
                  f"{'GEO BETTER' if geo_worst > ew_worst else ('EQUAL BETTER' if geo_worst < ew_worst else 'TIE')}")

    out = pd.DataFrame(rows)
    out.to_csv(f"{TABLES}/geometric_mondrian_cp.csv", index=False)
    print(f"\nSaved {TABLES}/geometric_mondrian_cp.csv")


if __name__ == "__main__":
    main()
