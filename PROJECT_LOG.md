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

## Status vs. roadmap (as of 2026-07-09)

- **Week 1-2**: Environment setup ✅ done. Literature review (30 papers) and
  3 core papers in depth — **not done**, this is reading/analysis work only the
  user can do.
- **Week 3**: Done, with the hybrid-calibration caveat above.
- **Week 4-5** (SimPy DES build + validation against real stats): done —
  `src/des/er_simulation.py`, `src/des/validate.py`, 91.0% match to real
  daily volume (see 2026-07-10 entry above for the full story, including a
  caught calibration bug).
- **Week 6-7** (run DES across scenarios to generate surrogate training data,
  train surrogate, evaluate MAE/RMSE/R²): not started.
