# Literature Review — Candidate Papers (target: 30)

Status: **candidate list, sourced and citation-verified via web search on 2026-07-24.**
Every entry below is a real, existing paper — title/authors/venue confirmed against at
least one primary source (publisher page, DOI, or preprint server), not generated from
memory. What's still needed from the team: **actually reading each one and writing the
relevance/summary column** — that's the part only a human can do honestly, and it's
what the professor is actually grading.

Columns: **Read?** (☐/☑) — **PDF** (📄 = already downloaded to `literature/pdfs/`,
🔒 = paywalled/blocked, use the link and your institutional library access) —
**Relevance note** (why this paper matters to *our* project, 1-2 sentences, written
after reading — not a restatement of the abstract).

`literature/pdfs/` is gitignored — these are mirrored copies for the team's own reading,
not meant to be redistributed via the public GitHub repo.

A couple of entries (marked ⚠) have a citation detail worth double-checking against the
actual PDF/DOI before it goes in the final bibliography — search snippets occasionally
truncate page ranges or drop a middle author.

---

## A. Conformal Prediction — Foundations (8)

| # | Citation | PDF | Read? | Relevance note |
|---|---|---|---|---|
| 1 | Vovk, V., Gammerman, A., Shafer, G. (2005). *Algorithmic Learning in a Random World*. Springer. [publisher page](https://link.springer.com/book/10.1007/b106715) | 🔒 (commercial book, no legal free copy) | ☐ | The original CP book — foundational framework, exchangeability assumption. |
| 2 | Papadopoulos, H., Proedrou, K., Vovk, V., Gammerman, A. (2002). Inductive Confidence Machines for Regression. *ECML 2002*, LNCS 2430, pp. 345–356. [DOI](https://doi.org/10.1007/3-540-36755-1_29) | 🔒 | ☐ | Introduces split (inductive) conformal prediction for regression — the exact form `standard_cp.py` implements. |
| 3 | Shafer, G., Vovk, V. (2008). A Tutorial on Conformal Prediction. *JMLR*, 9, 371–421. [arXiv](https://arxiv.org/abs/0706.3188) · [JMLR](https://jmlr.org/papers/v9/shafer08a.html) | 📄 `03_shafer_vovk_2008_tutorial.pdf` | ☐ | *(already selected)* CP foundations. |
| 4 | Lei, J., G'Sell, M., Rinaldo, A., Tibshirani, R.J., Wasserman, L. (2018). Distribution-Free Predictive Inference for Regression. *JASA*, 113(523), 1094–1111. [DOI](https://doi.org/10.1080/01621459.2017.1307116) | 🔒 | ☐ | Establishes finite-sample marginal coverage guarantees for split conformal — the theoretical basis being tested against DES data. |
| 5 | Romano, Y., Patterson, E., Candès, E. (2019). Conformalized Quantile Regression. *NeurIPS* 32. [arXiv](https://arxiv.org/abs/1905.03222) | 📄 `05_romano_et_al_2019_cqr.pdf` | ☐ | Direct source for the CQR baseline implemented in `train_quantile_surrogates.py` / `repeated_evaluation_cqr.py`. |
| 6 | Angelopoulos, A., Bates, S. (2021). A Gentle Introduction to Conformal Prediction and Distribution-Free Uncertainty Quantification. arXiv:2107.07511. | 📄 `06_angelopoulos_bates_2021_gentle_intro.pdf` | ☐ | Widely-used accessible survey — good for the report's background section. |
| 7 | Tibshirani, R.J., Barber, R.F., Candès, E., Ramdas, A. (2019). Conformal Prediction Under Covariate Shift. *NeurIPS* 32. [arXiv](https://arxiv.org/abs/1904.06019) | 📄 `07_tibshirani_et_al_2019_covariate_shift.pdf` | ☐ | Directly relevant to the Week 13 exchangeability stress test — covariate shift is one specific way exchangeability breaks. |
| 8 | Barber, R.F., Candès, E., Ramdas, A., Tibshirani, R.J. (2023). Conformal Prediction Beyond Exchangeability. *Annals of Statistics*, 51(2), 816–845. [arXiv](https://arxiv.org/abs/2202.13415) | 📄 `08_barber_et_al_2023_beyond_exchangeability.pdf` | ☐ | Formalizes what happens when exchangeability fails and how to (partially) correct for it — squarely the project's Week 13 topic. |

## B. Mondrian / Conditional Coverage (4)

| # | Citation | PDF | Read? | Relevance note |
|---|---|---|---|---|
| 9 | ⚠ Vovk, V., Lindsay, D., Nouretdinov, I., Gammerman, A. (2003). *Mondrian Confidence Machine*. Technical Report, Royal Holloway, University of London. [PDF](http://alrw.net/old/04.pdf) | 📄 `09_vovk_et_al_2003_mondrian_confidence_machine.pdf` | ☐ | Origin of the "Mondrian" name/concept — group-conditional coverage via partitioning, the core idea `mondrian_cp.py` implements. |
| 10 | Boström, H., Johansson, U. (2020). Mondrian Conformal Regressors. *PMLR* 128 (COPA 2020), pp. 114–133. [PMLR](https://proceedings.mlr.press/v128/bostrom20a.html) | 📄 `10_bostrom_johansson_2020_mondrian_regressors.pdf` | ☐ | *(already selected)* Direct methodological basis for `mondrian_cp.py`. |
| 11 | Boström, H., Johansson, U., Löfström, T. (2021). Mondrian Conformal Predictive Distributions. *PMLR* 152 (COPA 2021). [PMLR](https://proceedings.mlr.press/v152/bostrom21a.html) | 📄 `11_bostrom_et_al_2021_mondrian_predictive_dist.pdf` | ☐ | Follow-up extending Mondrian CP to full predictive distributions, not just intervals — relevant future-work angle. |
| 12 | Toccaceli, P., Gammerman, A. (2019). Combination of Inductive Mondrian Conformal Predictors. *Machine Learning*, 108, 489–510. [DOI](https://doi.org/10.1007/s10994-018-5754-9) | 🔒 | ☐ | Addresses combining multiple Mondrian category predictions — relevant to the 9-cell staffing×arrival partition used here. |

*(Note: Mondrian CP is a narrow subfield — only a handful of dedicated papers exist. That thinness is itself worth a sentence in the research-gap section: it supports the claim that applying it to a new domain, discrete-event/queueing simulation, is a genuine gap rather than a crowded space.)*

## C. Surrogate Modeling & Uncertainty Quantification (6)

| # | Citation | PDF | Read? | Relevance note |
|---|---|---|---|---|
| 13 | ⚠ Gopakumar, V. et al. (2026). Uncertainty Quantification of Surrogate Models Using Conformal Prediction. *Machine Learning: Science and Technology*. [IOPscience](https://iopscience.iop.org/article/10.1088/2632-2153/ae2e7b) (verify exact vol/issue/page against your own copy — search returned a slightly different article ID than the 7(1) 015025 you gave) | 🔒 (bot-protected download, use browser) | ☐ | *(already selected — base paper)*. |
| 14 | Kennedy, M.C., O'Hagan, A. (2001). Bayesian Calibration of Computer Models. *JRSS-B*, 63(3), 425–464. [DOI](https://doi.org/10.1111/1467-9868.00294) | 📄 `14_kennedy_ohagan_2001_bayesian_calibration.pdf` | ☐ | Classic computer-model/surrogate calibration paper — relevant background for the DES→surrogate pipeline generally. |
| 15 | Rasmussen, C.E., Williams, C.K.I. (2006). *Gaussian Processes for Machine Learning*. MIT Press. [free book site](https://gaussianprocess.org/gpml/) | 🔒 (site blocks scripted download, browse manually) | ☐ | Theoretical basis for the GP baseline in `gp_baseline.py`. |
| 16 | Friedman, J.H. (2001). Greedy Function Approximation: A Gradient Boosting Machine. *Annals of Statistics*, 29(5), 1189–1232. [Project Euclid](https://projecteuclid.org/journals/annals-of-statistics/volume-29/issue-5/Greedy-function-approximation-A-gradient-boosting-machine/10.1214/aos/1013203451.full) | 🔒 | ☐ | Theoretical basis for `HistGradientBoostingRegressor`, the primary surrogate architecture used throughout. |
| 17 | Lakshminarayanan, B., Pritzel, A., Blundell, C. (2017). Simple and Scalable Predictive Uncertainty Estimation using Deep Ensembles. *NeurIPS* 30, 6402–6413. [NeurIPS](https://proceedings.neurips.cc/paper/2017/hash/9ef2ed4b7fd2c810847ffa5fa85bce38-Abstract.html) | 📄 `17_lakshminarayanan_et_al_2017_deep_ensembles.pdf` | ☐ | Alternative (non-conformal) UQ baseline — good contrast case for the report's "why CP over other UQ methods" framing. |
| 18 | Abdar, M. et al. (2021). A Review of Uncertainty Quantification in Deep Learning: Techniques, Applications and Challenges. *Information Fusion*, 76, 243–297. [DOI](https://doi.org/10.1016/j.inffus.2021.05.008) | 📄 `18_abdar_et_al_2021_uq_deep_learning_review.pdf` | ☐ | Broad UQ survey — useful for situating CP among the wider UQ landscape in the report's intro. |

## D. Queueing Theory & ED Operations Research (5)

| # | Citation | PDF | Read? | Relevance note |
|---|---|---|---|---|
| 19 | Green, L.V., Soares, J., Giglio, J.F., Green, R.A. (2006). Using Queueing Theory to Increase the Effectiveness of Emergency Department Provider Staffing. *Academic Emergency Medicine*, 13(1), 61–68. [DOI](https://doi.org/10.1197/j.aem.2005.07.034) · [PubMed](https://pubmed.ncbi.nlm.nih.gov/16365329/) | 🔒 | ☐ | Real-world queueing-theory application to ED staffing — directly parallels this project's staffing/arrival-rate scenario design. |
| 20 | Green, L.V. Queueing Analysis in Healthcare. Book chapter, Columbia Business School. [PDF](https://business.columbia.edu/sites/default/files-efs/pubfiles/4386/chapter%2011%20QueueingAnalysis.pdf) | 📄 `20_green_queueing_analysis_healthcare_chapter.pdf` | ☐ | Background on M/M/s, M/G/1-type models used to justify the DES's queueing structure and the Erlang-load capacity calculation. |
| 21 | Hu, X. et al. (2018). Applying Queueing Theory to the Study of Emergency Department Operations: A Survey and a Discussion of Comparable Simulation Studies. *International Transactions in Operational Research*. [DOI](https://doi.org/10.1111/itor.12400) | 🔒 | ☐ | Survey comparing queueing-theory vs. simulation approaches to ED modeling — directly relevant to justifying the DES-over-analytic-queueing choice. |
| 22 | Performance Evaluation of a M/G/1 Queue Model for Patient Flow in a Health Care System. *Mathematical Modelling of Engineering Problems* (IIETA). [link](https://www.iieta.org/journals/mmep/paper/10.18280/mmep.110403) | 🔒 | ☐ | M/G/1 analytic queueing model for patient flow — comparison point for the DES's M/G/c-style resource model. |
| 23 | Decision Support for the Optimization of Provider Staffing for Hospital Emergency Departments with a Queue-Based Approach. [PMC6947400](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC6947400/) | 🔒 (PMC blocks scripted download, browse manually — it's free) | ☐ | Staffing optimization via queueing — relevant to the staffing-tercile dimension of the Mondrian CP partition. |

## E. Discrete-Event Simulation & ED-specific ML (7)

| # | Citation | PDF | Read? | Relevance note |
|---|---|---|---|---|
| 24 | A Simulation-Based Optimization Approach for the Calibration of a Discrete Event Simulation Model of an Emergency Department. [arXiv:2102.00945](https://arxiv.org/abs/2102.00945) (2021). | 📄 `24_des_calibration_ed_arxiv2102.00945.pdf` | ☐ | Directly parallel project: DES calibration methodology for an ED — good comparison to this project's hybrid (real arrivals + literature service-time) calibration approach. |
| 25 | Discrete Event Simulation for Emergency Department Modelling: A Systematic Review of Validation Methods. (2022). [ScienceDirect](https://www.sciencedirect.com/science/article/abs/pii/S2211692322000029) | 🔒 | ☐ | Systematic review of how ED DES models are validated — directly relevant to justifying the 91% daily-volume validation approach used here. |
| 26 | Discrete Event Simulation Modelling for an Improved Patient Flow at the Emergency Department, Sygehus Lillebælt, Kolding. [PMC3327033](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC3327033/) | 🔒 (PMC blocks scripted download, browse manually — it's free) | ☐ | Case study of ED DES for patient-flow improvement — real-world precedent for SimPy-style ED modeling. |
| 27 | A Simulation-Based Optimization Approach for Analyzing the Ambulance Diversion Phenomenon in an Emergency Department Network. [arXiv:2108.04162](https://arxiv.org/abs/2108.04162) | 📄 `27_ambulance_diversion_arxiv2108.04162.pdf` | ☐ | Related DES-based ED capacity/overload analysis — relevant to the exchangeability stress-test's surge-scenario framing. |
| 28 | Machine Learning-Based Prediction of Hospital Prolonged Length of Stay Admission at Emergency Department: A Gradient Boosting Algorithm Analysis. *Frontiers in Artificial Intelligence* (2023). [link](https://www.frontiersin.org/journals/artificial-intelligence/articles/10.3389/frai.2023.1179226/full) | 📄 `28_frontiers_2023_gbr_los_prediction.pdf` | ☐ | Gradient boosting for ED outcome prediction — same model family as this project's surrogate, applied to a related prediction task. |
| 29 | An Artificial Intelligence-Based Framework for Predicting Emergency Department Overcrowding: Development and Evaluation Study. [arXiv:2504.18578](https://arxiv.org/abs/2504.18578) (2025). | 📄 `29_ed_overcrowding_ai_arxiv2504.18578.pdf` | ☐ | Recent ML-for-ED-operations work — good for positioning this project within current (2025-26) literature, not just older queueing-theory work. |
| 30 | Machine Learning-Based Triage to Identify Low-Severity Patients with a Short Discharge Length of Stay in Emergency Department. [PMC9123815](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9123815/) | 🔒 (PMC blocks scripted download, browse manually — it's free) | ☐ | ML applied to ESI/triage-adjacent prediction — relevant given this project's DES also models ESI acuity mix. |

---

## F. DES Service-Time / Arrival Calibration Cross-Check (10, added 2026-08-05)

Separate from the 30-paper literature review above (which addresses the professor's
review requirement) - this set was gathered specifically to replace the book report's
previously vague, unverified "service-time parameters are literature-typical, e.g.
Ahalt et al. (2018) and similar" attribution (`src/utils/extract_distributions.py`)
with real, checkable numbers. Each entry below reports actual ED service-time,
length-of-stay, or arrival-rate statistics from its own real hospital dataset, used in
`reports/assignments/book_common.py` Section 4.2.3.1 to cross-check this project's
literature-calibrated service-time parameters quantitatively (not just cite them as
generically supportive). All ten verified against a primary source (journal page, DOI,
or arXiv abstract) before inclusion, same practice as sections A-E above.

| # | Citation | Real dataset | What it's used for |
|---|---|---|---|
| 31 | Hoot, N.R., LeBlanc, L.J., Jones, I., et al. (2008). Forecasting Emergency Department Crowding: A Discrete Event Simulation. *Annals of Emergency Medicine*, 52(2), 116-125. | 1 US academic ED, DES model | Closest structural match: log-normal per-ESI treatment duration, ESI mix, nonstationary Poisson arrivals. Direct numeric comparison table against this project's own ESI parameters. |
| 32 | Otto, R., Blaschke, S., Schirrmeister, W., et al. (2022). Length of Stay as Quality Indicator in Emergency Departments: AKTIN Registry. *Internal and Emergency Medicine*, 17(4), 1199-1209. | AKTIN registry, 12 German EDs, n=304,606 | Mean +/- SD LOS by triage level - used to check this project's SD/mean ratio against real-world variability. |
| 33 | Theiling, B.J., Kennedy, K.V., Limkakeng, A.T. Jr., et al. (2020). A Method for Grouping Emergency Department Visits by Severity and Complexity. *Western Journal of Emergency Medicine*, 21(5), 1147-1155. | US NHAMCS, ~805.7M weighted visits | Confirms duration-decreases-with-acuity shape at national scale. |
| 34 | Karaca, Z., Wong, H.S., Mutter, R.L. (2012). Duration of Patients' Visits to the Hospital Emergency Department. *BMC Emergency Medicine*, 12, 15. | AZ/MA/UT state databases, n=4,955,590 | Confirms right-skewed duration distribution at very large scale. |
| 35 | Kim, T.Y., Ohmart, C., Khan, Z., Lance, M., Kim, S. (2021). The Effect on Length of Stay After Implementation of Discharging Low Acuity Patients From Triage. *Cureus*, 13(9), e17640. | 1 US ED, n=2,107 | Independent ESI-4/5 mean LOS data point. |
| 36 | Mahmoodian, F., Eqtesadi, R., Ghareghani, A. (2014). Waiting Times in Emergency Department After Using the Emergency Severity Index Triage Tool. *Archives of Trauma Research*, 3(4), e19507. | 2 hospitals, n=900 | Independent full-ESI-range time-to-physician data point. |
| 37 | Laskowski, M., McLeod, R.D., Friesen, M.R., Podaima, B.W., Alfa, A.S. (2009). Models of Emergency Departments for Reducing Patient Waiting Times. *PLoS ONE*, 4(7), e6127. | 6 Winnipeg hospitals, n=185,659 | Cross-site support for priority-queue/simulation methodology calibrated on real CTAS data. |
| 38 | Locker, T.E., Mason, S.M. (2005). Analysis of the Distribution of Time That Patients Spend in Emergency Departments. *BMJ*, 330(7501), 1188-1189. | UK NHS EDs | Confirms right-skewed ED duration shape outside the US-heavy sample of the other nine. |
| 39 | De Santis, A., Giovannelli, T., Lucidi, S., Messedaglia, M., Roma, M. (2021). Determining the Optimal Piecewise Constant Approximation for the Nonhomogeneous Poisson Process Rate of ED Patient Arrivals. arXiv:2101.11138. | 1 large Italian ED | Validates this project's nonhomogeneous-Poisson hourly arrival-rate calibration as an established methodology. |
| 40 | Kramer, A., Dosi, C., Iori, M., Vignoli, M. (2020). Successful Implementation of Discrete Event Simulation: The Case of an Italian Emergency Department. arXiv:2006.13062. | 1 Italian ED, ~7,000 visits/month | Cross-site precedent for DES-based ED modeling reaching real operational use. |

---

## PDF status summary

**16 of 30 already downloaded** to `literature/pdfs/` (arXiv/PMLR/NeurIPS preprints and a
few author-hosted copies — all genuinely open access, fetched automatically).

**14 remaining are 🔒** — not paywalled in the "you can't get them" sense for the most
part, just blocked from scripted/automated download:
- **4 are on PubMed Central** (#23, #26, #30) and **truly free** — PMC just blocks
  non-browser downloads (bot protection). Open the link in a normal browser and
  save-as; takes seconds each.
- **The base paper (#13)** is on IOPscience, which sits behind the same kind of
  bot-protection page — same fix, open in a browser.
- **The GP book (#15)** is legitimately free at gaussianprocess.org but the site
  blocks scripted fetches too — browser download works fine.
- **The rest (#1, #2, #4, #12, #16, #19, #21, #22, #25)** are genuinely paywalled
  (Springer/Wiley/Elsevier/Annals of Statistics) — check whether your college library
  provides access (most do, via institutional login or a proxy/VPN), or search for an
  author's self-archived copy (many researchers post PDFs on their own homepage or
  ResearchGate).

## Next steps

1. Grab the remaining 14 PDFs (see above — mostly a 2-minute manual browser step,
   not a real blocker) or use your college library's access for the truly paywalled
   ones.
2. Read and check off; write the "why this matters to *our* project" note yourself —
   that's the part that needs to be genuine, not sourced.
3. Two entries flagged ⚠ need a quick citation double-check against the primary source
   before they go in a final bibliography.
4. If the professor wants "3 core papers in depth" (per the original roadmap), the
   strongest candidates given this project's actual content are: **#13 (Gopakumar et
   al., base paper)**, **#10 (Boström & Johansson, Mondrian CP)**, and **#5 (Romano et
   al., CQR)** — these three map directly onto code that exists in this repo
   (`gp_baseline.py`/base paper motivation, `mondrian_cp.py`, `repeated_evaluation_cqr.py`),
   so an in-depth treatment of them can lean on results you've already produced rather
   than needing fresh analysis.
