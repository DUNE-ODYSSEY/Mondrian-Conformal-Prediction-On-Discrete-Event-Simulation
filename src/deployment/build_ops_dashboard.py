"""
Prototype operational dashboard: a real, self-contained, interactive HTML
artifact (Plotly, no server required) demonstrating how this project's
surrogate + Mondrian conformal prediction pipeline could surface prediction
intervals to an ED operations manager, rather than only reporting them as
static tables in this document.

Two linked panels, both driven by a target-metric dropdown:
  1. A heatmap of the surrogate's point prediction across the real
     (staffing capacity, arrival-rate multiplier) scenario grid this
     project's DES actually samples from, with each cell's hover text
     giving the point prediction *and* its Mondrian conformal interval -
     computed live from the trained surrogate and the real per-category
     quantiles in results/tables/mondrian_cp_detail.csv, not mocked data.
  2. A bar chart of Mondrian interval width by category, directly
     visualizing this project's central finding (width concentrates where
     the true operating risk is - understaffed, high-arrival cells) in the
     form an operations dashboard would actually present it.

Static HTML export cannot call back into a live Python backend, so full
interactivity is achieved the standard way for a no-server Plotly
dashboard: every (target, grid-point) prediction is precomputed here and
embedded as a Plotly trace, with the dropdown just toggling which
precomputed trace set is visible - genuinely interactive in any browser,
not a screenshot.
"""

import numpy as np
import pandas as pd
import joblib
import plotly.graph_objects as go

from src.uq.mondrian_cp import LEVEL_NAMES, assign_categories, conformal_quantile

MODELS_DIR = "models"
CALIBRATION_DATA_PATH = "data/processed/cp_calibration_data.parquet"
OUT_PATH = "reports/assignments/figures/ops_dashboard.html"

FEATURES = ["n_capacity", "arrival_rate_multiplier"]
TARGETS = ["n_patients", "mean_wait_minutes", "mean_total_minutes", "p95_wait_minutes"]
ALPHA = 0.1

CAPACITY_GRID = np.arange(15, 46, 3)
ARRIVAL_GRID = np.round(np.arange(0.8, 1.31, 0.05), 2)


def build_category_quantiles(cal_df, cap_bounds, arr_bounds):
    cal_cat = assign_categories(cal_df, cap_bounds, arr_bounds)
    quantiles = {}
    for target in TARGETS:
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")
        yhat_cal = model.predict(cal_df[FEATURES])
        abs_resid = np.abs(cal_df[target].values - yhat_cal)
        quantiles[target] = {
            cat: conformal_quantile(abs_resid[cal_cat == cat], ALPHA)
            for cat in set(cal_cat)
        }
    return quantiles


def main():
    cal_df = pd.read_parquet(CALIBRATION_DATA_PATH)
    cap_bounds = np.quantile(cal_df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(cal_df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    cat_quantiles = build_category_quantiles(cal_df, cap_bounds, arr_bounds)

    grid_cap, grid_arr = np.meshgrid(CAPACITY_GRID, ARRIVAL_GRID)
    grid_df = pd.DataFrame({"n_capacity": grid_cap.ravel(), "arrival_rate_multiplier": grid_arr.ravel()})
    grid_cat = assign_categories(grid_df, cap_bounds, arr_bounds)

    detail = pd.read_csv("results/tables/mondrian_cp_detail.csv")

    heatmap_traces, bar_traces, buttons = [], [], []
    for i, target in enumerate(TARGETS):
        model = joblib.load(f"{MODELS_DIR}/surrogate_{target}.joblib")
        yhat = model.predict(grid_df[FEATURES])
        q = np.array([cat_quantiles[target][c] for c in grid_cat])
        lower, upper = yhat - q, yhat + q

        hover = np.array([
            f"n_capacity={c}<br>arrival_mult={a:.2f}<br>predicted={p:.1f}"
            f"<br>90% interval=[{lo:.1f}, {hi:.1f}]<br>category={cat}"
            for c, a, p, lo, hi, cat in zip(grid_df["n_capacity"], grid_df["arrival_rate_multiplier"], yhat, lower, upper, grid_cat)
        ]).reshape(grid_cap.shape)

        heatmap_traces.append(go.Heatmap(
            x=CAPACITY_GRID, y=ARRIVAL_GRID, z=yhat.reshape(grid_cap.shape),
            colorscale="RdYlGn_r", text=hover, hoverinfo="text",
            visible=(i == 0), colorbar=dict(title=target, x=0.44, len=0.9),
        ))

        sub = detail[detail["target"] == target].sort_values("category")
        bar_traces.append(go.Bar(
            x=sub["category"], y=sub["mondrian_width"], name=target,
            marker_color=["#c0392b" if "Low/arrival=High" in c else "#2a78d6" for c in sub["category"]],
            visible=(i == 0), xaxis="x2", yaxis="y2",
        ))

        buttons.append(dict(
            label=target, method="update",
            args=[{"visible": [j == i for j in range(len(TARGETS))] * 2}],
        ))

    fig = go.Figure(data=heatmap_traces + bar_traces)
    fig.update_layout(
        title="ER Operations Dashboard (Prototype) - Predicted Load and Mondrian CP Interval Width",
        updatemenus=[dict(active=0, buttons=buttons, x=1.15, y=1.15, xanchor="left")],
        grid=dict(rows=1, columns=2, pattern="independent"),
        xaxis=dict(domain=[0.0, 0.42], title="Staffing capacity"),
        yaxis=dict(domain=[0.0, 1.0], title="Arrival-rate multiplier"),
        xaxis2=dict(domain=[0.55, 1.0], title="Category", tickangle=30),
        yaxis2=dict(domain=[0.0, 1.0], title="Mondrian interval width"),
        width=1150, height=520, template="plotly_white",
    )
    fig.write_html(OUT_PATH, include_plotlyjs="cdn")
    print(f"Saved {OUT_PATH}")

    # static PNG snapshot (mean_wait_minutes view) for embedding in the PDF
    # report, which cannot embed the live interactive HTML directly - built
    # as an independent figure (no updatemenus/dropdown at all) rather than
    # copying the interactive figure's layout, since dropdown UI chrome has
    # no function in a static image and its "active" label does not track
    # which traces were actually made visible below.
    snap_heatmap = go.Heatmap(heatmap_traces[1])
    snap_heatmap.visible = True
    snap_bar = go.Bar(bar_traces[1])
    snap_bar.visible = True
    snapshot = go.Figure(data=[snap_heatmap, snap_bar])
    snapshot.update_layout(
        title="ER Operations Dashboard (Prototype) - mean_wait_minutes view",
        grid=dict(rows=1, columns=2, pattern="independent"),
        xaxis=dict(domain=[0.0, 0.42], title="Staffing capacity"),
        yaxis=dict(domain=[0.0, 1.0], title="Arrival-rate multiplier"),
        xaxis2=dict(domain=[0.55, 1.0], title="Category", tickangle=30),
        yaxis2=dict(domain=[0.0, 1.0], title="Mondrian interval width"),
        width=1150, height=520, template="plotly_white",
    )
    snapshot.write_image("reports/assignments/figures/ops_dashboard_snapshot.png", scale=2)
    print("Saved reports/assignments/figures/ops_dashboard_snapshot.png")


if __name__ == "__main__":
    main()
