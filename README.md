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
slides/           # PPT source/assets
results/
  figures/
  tables/
tests/
```

## Setup

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Roadmap

### Phase 1 — Mid-Sem

- [ ] Week 1-2: Literature review (30 papers), 3 core papers in depth, environment setup
- [ ] Week 3: Explore Hospital Triage and Patient History Data (Kaggle) — extract arrival
      patterns, service/treatment time, patient volume by hour/day
- [ ] Week 4-5: Build ER DES in SimPy calibrated on extracted distributions; validate
      against real aggregated stats
- [ ] Week 6-7: Run calibrated DES across staffing/arrival scenarios to generate surrogate
      training data; train surrogate (NN or gradient boosting); evaluate MAE/RMSE/R²
- [ ] Week 8: Implement GP baseline UQ; measure coverage and interval width
- [ ] Week 8 (parallel): Mid-sem PPT (title, problem statement, research gap, bridging
      approach, methodology diagram, dataset used, work completed, lit review snapshot,
      future scope)

**Mid-sem deliverable:** DES + surrogate + GP baseline + literature review draft + PPT.

### Phase 2 — End-Sem

- [ ] Week 9-10: Standard conformal prediction on surrogate residuals; multiple
      nonconformity measures
- [ ] Week 11-12: Mondrian CP — partition by staffing level/shift/arrival-rate category;
      per-category coverage
- [ ] Week 13: Stress-test exchangeability — out-of-distribution surge-day scenario
- [ ] Week 14-15: Full comparison (GP vs. standard CP vs. Mondrian CP) — coverage, width,
      computation time, plots
- [ ] Week 16: End-sem PPT (recap, pipeline results, UQ comparison, per-category coverage,
      exchangeability stress test, key finding, contribution summary, limitations/future
      scope, references)

**End-sem deliverable:** Full comparative results, final report, PPT.

## Reference paper

Gopakumar et al. (2026) — validates CP for surrogate UQ in physics domains (PDEs, MHD,
weather, fusion). States two explicit limitations: marginal coverage, exchangeability
assumption. This project tests those limitations in a new, untested domain:
discrete-event/queueing systems (ER simulation).
