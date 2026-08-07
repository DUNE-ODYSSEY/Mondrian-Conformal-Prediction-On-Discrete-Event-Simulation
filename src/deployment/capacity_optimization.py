"""
Prescriptive capacity allocation under uncertainty: formulates choosing
staffing capacity as a constrained optimization problem where the
constraint is expressed directly in terms of this project's own Mondrian
conformal prediction intervals, rather than the surrogate's raw point
prediction - a concrete, worked example of what "using a conformal
interval as an explicit constraint set" (rather than only a reported
number) looks like in an operational decision.

Problem, for a given arrival-rate scenario a and a policy wait-time
ceiling W_max:

    minimize    n_capacity                              (staffing cost proxy)
    subject to  yhat(n_capacity, a) + q_cat(n_capacity, a) <= W_max

i.e. choose the *cheapest* staffing level whose Mondrian conformal upper
bound - not its point prediction - stays under the ceiling. Using the
interval's upper bound rather than the point prediction is what makes this
a genuinely conservative, risk-aware plan: because Mondrian's interval
width varies by category (Section on the core finding), the safety margin
this constraint enforces is automatically larger in the categories this
project's own results show are least reliably predicted, and smaller
where they are not - the optimization inherits that structure directly
from the calibrated intervals rather than needing it encoded separately.

Solved by direct grid search over the small, integer n_capacity domain
this project's DES already uses (15-45) - exact for a monotone-ish
constraint on a small discrete domain, no need for a general-purpose
nonlinear solver.
"""

import numpy as np
import pandas as pd
import joblib

from src.uq.mondrian_cp import LEVEL_NAMES, assign_categories, conformal_quantile

MODELS_DIR = "models"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
OUT_PATH = "results/tables/capacity_optimization.csv"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGET = "p95_wait_minutes"
ALPHA = 0.1

CAPACITY_DOMAIN = np.arange(15, 46)
ARRIVAL_SCENARIOS = [0.8, 0.9, 1.0, 1.1, 1.2, 1.3]
W_MAX_SWEEP = [60, 90, 120, 180, 240]  # policy wait-time ceilings (minutes), swept rather than a single asserted value


def main():
    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)

    model = joblib.load(f"{MODELS_DIR}/surrogate_{TARGET}.joblib")
    yhat_cal = model.predict(cal_df[FEATURES])
    abs_resid_cal = np.abs(cal_df[TARGET].values - yhat_cal)
    cat_quantile = {cat: conformal_quantile(abs_resid_cal[cal_cat == cat], ALPHA) for cat in set(cal_cat)}

    grid_df = pd.DataFrame(
        [(c, a) for c in CAPACITY_DOMAIN for a in ARRIVAL_SCENARIOS],
        columns=FEATURES,
    )
    grid_df["predicted"] = model.predict(grid_df[FEATURES])
    grid_df["category"] = assign_categories(grid_df, cap_bounds, arr_bounds)
    grid_df["q"] = grid_df["category"].map(cat_quantile)
    grid_df["upper_bound"] = grid_df["predicted"] + grid_df["q"]

    rows = []
    for w_max in W_MAX_SWEEP:
        for arrival in ARRIVAL_SCENARIOS:
            sub = grid_df[grid_df["arrival_rate_multiplier"] == arrival].sort_values("n_capacity")
            feasible = sub[sub["upper_bound"] <= w_max]
            # for comparison: capacity needed if using the raw point prediction instead of the CP upper bound
            feasible_point_only = sub[sub["predicted"] <= w_max]

            n_star_cp = int(feasible["n_capacity"].min()) if len(feasible) else None
            n_star_point = int(feasible_point_only["n_capacity"].min()) if len(feasible_point_only) else None

            rows.append({
                "W_max_minutes": w_max,
                "arrival_rate_multiplier": arrival,
                "n_capacity_star_cp_constrained": n_star_cp,
                "n_capacity_star_point_prediction_only": n_star_point,
                "extra_capacity_from_using_cp_bound": (n_star_cp - n_star_point) if (n_star_cp is not None and n_star_point is not None) else None,
            })

    out_df = pd.DataFrame(rows)
    out_df.to_csv(OUT_PATH, index=False)
    print(out_df.to_string(index=False))
    print(f"\nSaved {OUT_PATH}")


if __name__ == "__main__":
    main()
