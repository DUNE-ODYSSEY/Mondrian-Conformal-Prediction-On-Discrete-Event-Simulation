# Project Log

Chronological record of what was done, what changed, and why. Update this
whenever a milestone completes or a decision is made that isn't obvious from
the code alone. See `README.md` for the roadmap/current structure.

---

## 2026-07-09 — Project scaffold + repo setup

- Created project structure at `Desktop/er-conformal-uq`: `src/{des,surrogate,uq,utils}`,
  `data/{raw,processed}`, `notebooks`, `literature`, `reports/{mid_sem,end_sem}`,
  `slides`, `results/{figures,tables}`, `tests`.
- Installed GitHub CLI, authenticated as `venu211007-source`.
- Created public repo **DUNE-ODYSSEY/Mondrian-Conformal-Prediction-On-Discrete-Event-Simulation**
  and pushed initial scaffold (commit `af93e44`).
- Created Python venv (`.venv`, Python 3.13). Installed and import-verified: numpy,
  pandas, scipy, scikit-learn, simpy, **mapie 1.4.1** (confirmed no Python 3.13
  compatibility issues), matplotlib, seaborn, jupyter, pytest, python-pptx.
- `data/`, `.venv/` are gitignored — only code and small result tables go to GitHub.
  Raw/processed data lives locally only; re-download via the `kaggle` CLI command
  in the README.

## 2026-07-09 — Kaggle dataset download (Week 3)

- Installed `kaggle` CLI into the venv. User generated a Kaggle API token (new
  `KGAT_` format) and saved it to `~/.kaggle/access_token`.
  - **Note:** this token was briefly visible in a screenshot pasted into chat.
    User was advised to regenerate if concerned; decided to keep the existing
    token (low blast radius — Kaggle account access only, no billing/write
    access to other systems).
- Identified and downloaded the correct dataset: `maalona/hospital-triage-and-patient-history-data`
  ("Hospital Triage and Patient History Data") into `data/raw/` (96MB, `.rdata` format).
- Converted to `data/processed/triage_data.parquet` via `pyreadr` for faster
  loading (added `pyreadr`, `pyarrow` to `requirements.txt`).
- **Dataset shape:** 560,486 ER visit records × 972 columns — vitals, ~250 diagnosis
  flags, ~300 lab columns, ~180 chief-complaint flags, ESI acuity, disposition, etc.
- **Key finding:** the dataset has **no ED length-of-stay / treatment-duration field.**
  Only `arrivalmonth`, `arrivalday`, `arrivalhour_bin` (4-hour arrival buckets) exist —
  no discharge/departure timestamp. Confirmed via exhaustive column-name search
  across all 972 columns. This is a property of the dataset itself (de-identified,
  MIMIC-style), not something we can recover from it.

## 2026-07-09 — Week 3 distribution extraction: hybrid calibration decision

Decision (user approved, told to keep it simple): since real treatment-time data
doesn't exist in this dataset, calibrate the DES with a **hybrid approach**:

- **Arrival process + patient volume** → real data, extracted directly:
  `results/tables/arrivals_by_hour_bin.csv`, `arrivals_by_day.csv`,
  `arrivals_by_month.csv`, `esi_mix.csv`.
- **Service/treatment time** → literature-standard log-normal parameters by ESI
  acuity level (not derived from the Kaggle data). See
  `results/tables/service_time_params.csv` and the docstring in
  `src/utils/extract_distributions.py` for the values and rationale.

**This distinction must be stated explicitly in the mid-sem/end-sem report and PPT** —
arrivals are real-data-calibrated, service times are literature-calibrated. Do not
present both as derived from the dataset.

Script: `src/utils/extract_distributions.py` — run via
`.venv\Scripts\python.exe src\utils\extract_distributions.py`. Outputs land in
`results/tables/`.

## 2026-07-10 — Week 4-5: DES build, caught and fixed a bad calibration assumption

First version of `src/des/er_simulation.py` (SimPy) produced obviously broken output:
54 patients/day with 650-min average waits, when real data implies ~1500+/day.
Root cause was an **unverified guess**, not a code bug: I'd assumed the Kaggle
dataset spans 365 days to convert its total record count into a daily arrival
rate. Checked the actual Kaggle dataset description instead of continuing to
guess — it states the data covers **March 2014 - July 2017 (1248 days)** across
**three EDs** (`dep_name` A/B/C: one academic, two community), not one year at
one site.

Fixes made:
- `extract_distributions.py` now filters to a single department (default `A`,
  the largest/most likely academic site: 322,283 of 560,486 visits) instead of
  pooling all three EDs, and uses `STUDY_PERIOD_DAYS = 1248` instead of the
  wrong 365-day guess. Real calibrated rate: **258.2 visits/day** for department A.
- Fixed a second latent bug found during this pass: the `"23-02"` arrival-hour
  bin wraps midnight (hours 23,0,1,2), but the original `hour // 4` indexing
  put it at hours 0-3 instead. Replaced with an explicit hour→bin map.
- Simplified the DES's resource model from two nested pools (`doctors` +
  `beds`) down to **one combined `capacity` resource** (bed+provider slot for
  the full visit). We have no real data to calibrate a doctor:bed ratio, so a
  second resource pool would only add more unverified numbers without adding
  insight for this project's actual question (UQ methodology, not hospital
  operations realism). This is also the standard M/G/c simplification used in
  ED queueing literature.
- Picked `n_capacity=30` from a quick offered-load (Erlang) calculation using
  the real arrival rate and literature service times (~22 erlangs average load,
  ~34 erlangs at the 11-14 peak bin) — not a guess, but not over-engineered
  either; documented inline in `er_simulation.py`.

**Validation** (`src/des/validate.py`, 200 simulated days): mean simulated
patients/day = 235.1 vs. real 258.2 (**91.0% match**). The ~9% shortfall is
expected and not a calibration flaw: patients still queued when a simulated
24h day ends are cut off (right-censoring at the day boundary) rather than
carried into a "next day," since each run is meant to be one independent
sample for Week 6-7's surrogate training data, not a continuous rollout.
Output: `results/tables/des_validation.csv`.

**Lesson for later weeks:** don't hardcode a numeric assumption (day counts,
site scope, etc.) without checking the source — the 365-day guess produced a
result wrong by ~28x and would have silently poisoned every downstream step
(surrogate training, UQ) if not caught here.

## 2026-07-10 — Week 6-7: surrogate training data + surrogate model

- `src/surrogate/generate_training_data.py`: ran the calibrated DES across
  5000 randomly sampled scenarios (`n_capacity` in [15,45], `arrival_rate_multiplier`
  in [0.8,1.3]), each with its own random seed so DES stochasticity (arrival
  randomness, ESI mix, service-time variance) stays in the labels — that
  residual noise is what the UQ step (Week 8+) needs to characterize.
  Output: `data/processed/surrogate_training_data.parquet` (~75s to generate).
- `src/surrogate/train_surrogate.py`: trained `HistGradientBoostingRegressor`
  (gradient boosting, not a NN — simpler and standard for this tabular
  2-input problem) as one model per target metric. 80/20 train/test split.

**Results** (`results/tables/surrogate_metrics.csv`):

| target | MAE | RMSE | R² |
|---|---|---|---|
| n_patients | 9.90 | 12.59 | 0.929 |
| mean_wait_minutes | 8.86 | 13.23 | 0.787 |
| mean_total_minutes | 9.88 | 13.61 | 0.762 |
| p95_wait_minutes | 66.94 | 102.47 | 0.647 |

`p95_wait_minutes` is the weakest fit (R²=0.65) — expected, since a tail
statistic from one noisy simulated day depends heavily on the specific
stochastic realization, not just the two scalar scenario parameters. This
is actually a useful property going forward: it's the metric most likely to
show CP intervals doing real work (wider, more informative uncertainty)
versus a well-fit target like `n_patients` where intervals should stay tight.

Models saved to `models/*.joblib` (gitignored — regenerable via the two
scripts above, not checked into GitHub).

## 2026-07-10 — Week 8: GP baseline UQ

`src/uq/gp_baseline.py`: Gaussian Process per target, predictive mean +/- z*std
as the (1-alpha) interval, alpha=0.1 (90% target coverage) fixed here since
**standard/Mondrian CP in Phase 2 must target the same alpha** for the
comparison to mean anything.

GP trained on a 1000-point subsample of the training set, not the full
~4000 — sklearn's exact GP is O(n^3) (benchmarked: ~8s at n=1000, ~43s at
n=2000, so full-scale would be tens of minutes across 4 targets). This is
also directly relevant to the project's own narrative: GP's poor scalability
vs. CP's near-constant calibration cost is one of the things Week 14-15's
computation-time comparison is meant to show, so it's worth having fit/predict
timings recorded now (`fit_seconds`/`predict_seconds` columns) rather than
just discarding them.

Same train/test split (`random_state=42, test_size=0.2`) as `train_surrogate.py`
— every UQ method needs to be evaluated on identical test points.

**Results** (`results/tables/gp_baseline_metrics.csv`):

| target | target coverage | empirical coverage | mean interval width |
|---|---|---|---|
| n_patients | 90% | 88.5% | 39.3 |
| mean_wait_minutes | 90% | 87.7% | 43.6 |
| mean_total_minutes | 90% | 88.8% | 43.4 |
| p95_wait_minutes | 90% | 90.1% | 350.3 |

Slight undercoverage on 3/4 targets (87.7-88.8% vs. 90% target) is expected
for GP — no finite-sample coverage guarantee, unlike CP. This is the actual
gap the project is testing whether Mondrian CP closes in this domain.
Interval width scales with target difficulty (tight for `n_patients`, very
wide for `p95_wait_minutes`), consistent with the Week 6-7 surrogate R² results.

## 2026-07-16 — Mid-sem PPT built

`slides/build_mid_sem_ppt.py` generates `slides/mid_sem_presentation.pptx`
(python-pptx) — 9 slides per the roadmap spec, populated with the real
numbers from `results/tables/` (DES validation, surrogate MAE/RMSE/R², GP
baseline coverage/width). Team: G Venugopalan (CB.AI.U4AID25115), Vipin
Sudhakar (CB.AI.U4AID25166), Rithvik Arulprakash (CB.AI.U4AID25148),
Harshith Kv (CB.AI.U4AID25119). Course: 23AID201.

Slide 8 (literature review snapshot) is a **clearly marked placeholder** —
category boxes with blank counts (`[ __ ] papers reviewed`) and an amber
"PLACEHOLDER — update before presenting" banner — since the literature
review is being done separately by the user and isn't finished. Not faked.

Verified visually by rendering all 9 slides to PNG via PowerPoint COM
automation (`Presentations.Open` + `.Export`) before calling this done —
no overlap/overflow issues, including on the data-dense Slide 7 (three
tables). Re-run the build script any time `results/tables/` changes;
edit `TEAM`/`COURSE_CODE` at the top of the script if roster changes.

## 2026-07-16 — Mid-sem PPT redesigned (less "AI pitch deck", more standard student deck)

The first version of the PPT looked like a generated pitch deck — dark hero
title slide, a repeated kicker/underline/page-number formula on every slide,
hand-typed bullet glyphs instead of native bullets, a curated Tailwind-blue
palette, and perfectly symmetric card grids. Rebuilt `slides/build_mid_sem_ppt.py`
to use PowerPoint's actual default theme and native layouts/placeholders
(Title Slide, Title and Content, Two Content) instead of custom-drawn boxes
for nearly everything, and loosened the bullet text so it reads more like
notes than marketing copy.

Caught three real layout bugs by rendering every slide to PNG via PowerPoint
COM automation before calling it done (`Presentations.Open` + `.Export` —
same verification approach used for the first PPT version):
- Title slide: title/subtitle placeholders sized for a short 1-line title
  overlapped when given the actual 2-line project title. Fixed by explicit
  positioning + smaller font.
- Slide 5 (pipeline diagram): default `Presentation()` template is 10x7.5in
  (4:3), not 13.333x7.5in widescreen — confirmed by inspecting
  `prs.slide_width`/layout placeholder positions directly rather than
  guessing. My box x-positions assumed the wider canvas and ran off the
  right edge. Also, PowerPoint's built-in autoshapes carry a theme "shape
  style" that defaults text to white (assuming a colored fill) — since
  these boxes use no fill, the labels rendered invisible white-on-white
  until font color was set explicitly.
- Slide 4 (two-content columns): setting only `.top`/`.height` on a
  placeholder that has no prior explicit transform creates a fresh xfrm
  with the *unset* fields (`.left`/`.width`) defaulting to zero — collapsed
  the column to zero width, wrapping every line to one character. Fixed by
  always setting all four of left/top/width/height together. Also had to
  explicitly shrink the level-0 bullet font (inherits a large ~28pt default
  meant for short headings), which was pushing sub-bullets off the bottom
  of the slide.

**Lesson for building any future slides/report programmatically:** never
trust python-pptx layout math without rendering and looking at it — several
of these bugs (invisible text, zero-width columns, off-canvas shapes) would
have shipped completely silently; nothing in the code raises an error.

## 2026-07-16 — Week 9-10: standard conformal prediction

`src/uq/generate_calibration_data.py`: 1200 fresh DES scenarios purely for
CP calibration, drawn with a different sampling seed and a large DES-seed
offset (100,000) from the training data, so nothing overlaps with what the
surrogate was trained on or with the test set. This matters methodologically:
calibrating on training-set residuals would understate the true residual
spread (the surrogate fits those points) and invalidate the coverage
guarantee. Same scenario ranges as training data - has to be exchangeable
with the test set, i.e. same distribution, not a different one.

`src/uq/standard_cp.py`: split conformal prediction wrapping the existing
surrogate models (point predictors, unchanged from Week 6-7). Same test set
(same `random_state=42` split) and same alpha=0.1 as the GP baseline, so
all three UQ methods stay directly comparable. Two nonconformity measures:
- symmetric: `|y - yhat|`, fixed-width interval.
- asymmetric: separate upper/lower residual quantiles, each calibrated at
  1-alpha/2 and combined via a union bound (valid, if slightly conservative,
  route to >=1-alpha coverage without needing a variance model).

**Results** (`results/tables/standard_cp_metrics.csv`, target coverage 90%):

| target | symmetric coverage | symmetric width | asymmetric coverage | asymmetric width |
|---|---|---|---|---|
| n_patients | 92.1% | 43.6 | 92.0% | 43.6 |
| mean_wait_minutes | 89.0% | 47.2 | 88.7% | 47.2 |
| mean_total_minutes | 90.2% | 46.4 | 89.6% | 47.2 |
| p95_wait_minutes | 90.8% | 362.9 | 89.7% | 355.6 |

Compare to the GP baseline (88.5%, 87.7%, 88.8%, 90.1%): CP lands closer to
the 90% nominal target across the board and doesn't show GP's systematic
undercoverage on 3/4 targets - the small fluctuations here (88.7-92.1%) look
like normal finite-sample noise around 90%, not a directional bias. For the
right-skewed `p95_wait_minutes`, the asymmetric measure gives a narrower
interval at comparable coverage (355.6 vs. 362.9) - real evidence the
skew-aware measure is doing something useful, not just an academic variant.

## 2026-07-17 — Week 11-12: Mondrian conformal prediction

`src/uq/mondrian_cp.py`: calibrates a separate quantile per category instead
of standard CP's single pooled quantile. Categories are the cross of
staffing tercile x arrival-rate tercile (Low/Med/High x Low/Med/High = 9
cells) - the DES output has no "shift" field, so category boundaries only
use the two real covariates we actually have, not an invented third one.
Bin edges come from the calibration set's own quantiles, never the test set.
Same calibration/test split and alpha=0.1 as standard_cp.py, symmetric
`|y - yhat|` nonconformity measure, so this is a clean pooled-vs-per-category
comparison and not a differently-set-up experiment dressed up as one.

For each category, coverage is evaluated two ways on the *same* test points:
applying standard CP's single pooled quantile (shows where the marginal
guarantee actually breaks down), vs. applying that category's own Mondrian
quantile (`results/tables/mondrian_cp_detail.csv`, summarized per-target in
`mondrian_cp_summary.csv`).

**This is the core result of the whole project.** The worst pooled-CP
category is consistently `staff=Low/arrival=High` - the most congested
scenario:

| target | pooled coverage (worst category) | Mondrian coverage (same category) | target |
|---|---|---|---|
| mean_wait_minutes | 68.2% | 90.9% | 90% |
| mean_total_minutes | 80.7% | 92.0% | 90% |
| p95_wait_minutes | 72.7% | 92.0% | 90% |

A single pooled quantile is calibrated on *average* difficulty, so it
silently fails exactly the operationally important scenarios (understaffed +
high demand) while overcovering the easy ones (`staff=High/arrival=Low` hits
100% pooled coverage - wasted interval width). Mondrian fixes the
undercoverage by giving the hard category an honestly wider interval instead
of pretending it's as easy as the rest.

`n_patients` is the one target where Mondrian does *not* help - its
pooled-CP per-category coverage was already fairly uniform (87.3-98.9%, no
real conditional miscalibration to begin with), so splitting the calibration
set into 9 smaller cells (~130 points each vs. 1200 pooled) just adds
finite-sample noise to the per-category quantile estimate without a
corresponding benefit (coverage range actually widens slightly, 11.5% ->
13.9%). This is a known, textbook-documented Mondrian CP tradeoff (smaller
per-cell calibration sets = noisier per-cell quantiles), not a bug - and
it's a more honest, useful finding than "Mondrian always wins everywhere"
would have been.

**Answer to the project's core question so far:** yes, Mondrian CP closes
the marginal-vs-conditional coverage gap in this discrete-event/queueing
domain, for targets where that gap is real (mean_wait, mean_total,
p95_wait) - Gopakumar et al.'s marginal-coverage limitation, demonstrated
outside physics simulation for the first time here. Week 13's exchangeability
stress test (OOD surge day) is still open and tests their *other* stated
limitation.

## 2026-07-17 — Week 13: exchangeability stress test

`src/uq/exchangeability_stress_test.py`: keeps the Week 9-12 calibration
fixed exactly as-is and pushes the *test* distribution's arrival_rate_multiplier
progressively past the training range (0.8-1.3) up to 3.0x, with n_capacity
still drawn normally from [15,45] - isolates the shift to arrival rate alone.
300 DES days per severity level, both standard CP and Mondrian CP evaluated
on the same test points. Output: `results/tables/exchangeability_stress_test.csv`.

**In range (0.8, 1.0, 1.3x) coverage matches Week 9-12** (82-95%, both
methods), confirming the test harness is consistent before looking at the
OOD region.

**Past the training boundary, coverage collapses fast:**

| arrival multiplier | n_patients | mean_wait | mean_total | p95_wait |
|---|---|---|---|---|
| 1.3 (in range) | 93.3% | 83.0% | 88.0% | 82.0% |
| 1.5 | 78.7% | 68.7% | 73.7% | 69.7% |
| 1.8 | 70.3% | 42.7% | 47.3% | 43.0% |
| 2.0 | 64.7% | 34.3% | 39.0% | 33.0% |
| 3.0 | 31.7% | 12.3% | 4.7% | 78.3%* |

(standard CP shown; Mondrian tracks closely, see below. *p95_wait's apparent
recovery at 2.5-3.0x is explained below - not a good sign.)

**Root cause, verified directly, not inferred:** the gradient-boosting
surrogate (`HistGradientBoostingRegressor`) is tree-based, and tree models
cannot extrapolate past the range of their training data - predictions for
any input outside that range clip to whatever leaf the boundary training
points fell into. Checked this directly: `mean_wait_minutes`'s prediction is
**frozen at 42.4** for every arrival multiplier from 1.3 through 3.0, and
`p95_wait_minutes`'s is **frozen at 221.2** across the same range, while the
DES's actual output keeps moving. As the true value drifts further from a
prediction that literally cannot move, the residual grows past whatever the
calibration quantile allows, and coverage fails - not because CP's math is
wrong, but because the point predictor it's wrapping stops being informative
outside its training domain. Both standard and Mondrian CP fail here since
neither can fix an uninformative underlying prediction; Mondrian's per-category
quantiles are also derived from in-range calibration data, so they're equally
blind to the shift.

**The `p95_wait_minutes` "recovery" at 2.5-3.0x is not the interval working -
it's a data-generating-process artifact.** Diagnostic run (n=40/point,
`n_capacity=30` fixed): true `p95_wait_minutes` goes 99.8 (1.0x) -> 249.5
(1.3x) -> 378.3 (1.5x) -> 420.4 (2.0x) -> 172.1 (2.5x) -> 297.9 (3.0x) -
non-monotonic, not a clean saturation curve. This connects back to the
Week 4-5 finding that patients still queued when a simulated day ends are
right-censored out of the stats: at extreme overload, *fewer* patients
finish service within 24h at all, shrinking and changing the composition of
the "completed visits" pool the p95 is computed over. The true value drifting
back down toward the frozen prediction is coincidental, driven by the
censoring artifact getting worse, not by the surrogate or the interval
becoming more reliable.

**Answer to Gopakumar et al.'s exchangeability limitation, in this domain:**
confirmed broken outside the training distribution, and both standard and
Mondrian CP degrade together once it does - Mondrian's per-category structure
doesn't protect against exchangeability violation, only against pooling
categories that are individually still in-distribution. One thing worth
noting for the report: the failure is *detectable*, not silent - residuals
balloon and coverage measurably craters rather than the surrogate confidently
reporting a wrong-but-narrow interval without any signal something's off.
That's a real practical property, even though it doesn't rescue coverage
here.

## Status vs. roadmap (as of 2026-07-17)

- **Week 1-2**: Environment setup ✅ done. Literature review (30 papers) and
  3 core papers in depth — **not done**, this is reading/analysis work only the
  user can do.
- **Week 3**: Done, with the hybrid-calibration caveat above.
- **Week 4-5** (SimPy DES build + validation against real stats): done —
  `src/des/er_simulation.py`, `src/des/validate.py`, 91.0% match to real
  daily volume (see 2026-07-10 entry above for the full story, including a
  caught calibration bug).
- **Week 6-7** (run DES across scenarios to generate surrogate training data,
  train surrogate, evaluate MAE/RMSE/R²): done — see entry above.
- **Week 8** (GP baseline UQ, coverage + interval width): done — see entry above.
  **This completes the mid-sem deliverable's technical work** (DES + surrogate +
  GP baseline). Remaining for mid-sem: literature review + the PPT itself
  (PPT is done, see entries above — just needs the lit review slide filled in).
- **Week 9-10** (standard CP, multiple nonconformity measures): done — see entry above.
- **Week 11-12** (Mondrian CP, per-category coverage): done — see entry above. Core
  project result now in hand: Mondrian CP closes the marginal/conditional coverage
  gap for 3 of 4 targets.
- **Week 13** (exchangeability stress test, OOD surge day): done — see entry
  above. Both stated Gopakumar et al. limitations now tested in this domain:
  marginal coverage (Week 11-12, Mondrian closes the gap where real) and
  exchangeability (this entry, breaks down as expected, root cause identified).
- **Week 14-15** (full GP vs. standard CP vs. Mondrian CP comparison — coverage,
  width, computation time): not started.
