# Mondrian Conformal Prediction on Discrete-Event Simulation

Uncertainty quantification for ER (emergency room) queueing surrogate models,
using conformal prediction — testing whether Gopakumar et al. (2026)'s
physics-domain CP results (and their stated limitations: marginal coverage,
exchangeability assumption) hold in a discrete-event/queueing domain.

## Pipeline

DES (SimPy, calibrated on real hospital data) → surrogate model (trained on
DES outputs) → uncertainty quantification (GP baseline vs. standard CP vs.
Mondrian CP).

## Repo layout

```
data/
  raw/            # Hospital Triage and Patient History Data (Kaggle) — not committed
  processed/      # extracted distributions, calibration outputs
src/
  des/            # SimPy ER discrete-event simulation
  surrogate/      # surrogate model training (NN / gradient boosting)
  uq/             # GP baseline, standard CP, Mondrian CP
  utils/          # shared helpers
notebooks/        # exploratory analysis
literature/       # bibliography, paper notes
reports/
  mid_sem/
  end_sem/
  assignments/    # final 200+ page book-format report (docx/pdf) and its build scripts
slides/           # PPT source/assets
results/
  figures/
  tables/
tests/            # unit tests for the CP calibration math and the DES core
```

## Setup

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Status: complete

Both phases below are done, plus a substantial extension beyond the
original scope. See `PROJECT_LOG.md` for the full session-by-session
history and `reports/assignments/` for the final 200+ page book-format
report.

### Phase 1 — Mid-Sem

- [x] Literature review (30 papers across 5 categories, critically assessed)
- [x] Explored Hospital Triage and Patient History Data (Kaggle) — extracted arrival
      patterns, service/treatment time, patient volume by hour/day
- [x] Built ER DES in SimPy calibrated on extracted distributions; validated
      against real aggregated stats (91.0% match, department A)
- [x] Ran calibrated DES across staffing/arrival scenarios to generate surrogate
      training data; trained surrogate (gradient boosting + MLP); evaluated MAE/RMSE/R²
- [x] Implemented GP baseline UQ; measured coverage and interval width
- [x] Mid-sem PPT

### Phase 2 — End-Sem

- [x] Standard conformal prediction on surrogate residuals
- [x] Mondrian CP — partitioned by staffing tercile x arrival-rate tercile (9 cells);
      per-category coverage
- [x] Stress-tested exchangeability — out-of-distribution demand-surge sweep, both
      GBR and MLP surrogates
- [x] Full comparison (GP vs. standard CP vs. Mondrian CP) — coverage, width,
      computation time, plots
- [x] End-sem PPT

### Beyond the original scope

- [x] Conformalized quantile regression (CQR) and Mondrian-CQR
- [x] Conformal risk control (CRC), adaptive conformal inference (ACI), and
      likelihood-ratio weighted CP under covariate shift
- [x] 5-architecture surrogate benchmark (GBR, MLP, RandomForest, XGBoost, LightGBM)
- [x] Independent cross-site replication at a second real department (Dept B)
- [x] Interactive ops dashboard and CP-constrained capacity-planning optimization
- [x] FDA AI/ML SaMD regulatory framing
- [x] Single 200+ page book-format report, restructured into 11 chapters

## Reference paper

Gopakumar et al. (2026) — validates CP for surrogate UQ in physics domains (PDEs, MHD,
weather, fusion). States two explicit limitations: marginal coverage, exchangeability
assumption. This project tests those limitations in a new, untested domain:
discrete-event/queueing systems (ER simulation).
