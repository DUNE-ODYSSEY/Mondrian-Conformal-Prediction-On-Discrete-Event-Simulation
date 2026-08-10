"""Sanity tests for the ER discrete-event simulation. The full scenario
sweep used to generate this project's actual results takes far too long to
run in a test, but a single deterministic 24h day is fast (well under a
second) and exercises the real calibration data committed to
results/tables/, not a mock - so a broken calibration file or a broken
queueing/priority mechanism would fail these."""

import numpy as np
import pytest

from src.des.er_simulation import lognormal_params_from_mean_sd, run_scenario


def test_lognormal_params_round_trip_the_target_mean():
    # Converting a desired (mean, sd) to the underlying normal's (mu, sigma)
    # should give back a log-normal whose analytic mean matches the target -
    # this is the calibration this project's service-time sampling relies on.
    for mean, sd in [(120, 60), (45, 25), (180, 90)]:
        mu, sigma = lognormal_params_from_mean_sd(mean, sd)
        implied_mean = np.exp(mu + sigma**2 / 2)
        implied_var = (np.exp(sigma**2) - 1) * np.exp(2 * mu + sigma**2)
        assert implied_mean == pytest.approx(mean, rel=1e-6)
        assert implied_var == pytest.approx(sd**2, rel=1e-6)


def test_run_scenario_is_deterministic_given_a_seed():
    result_a = run_scenario(n_capacity=30, arrival_rate_multiplier=1.0, seed=42)
    result_b = run_scenario(n_capacity=30, arrival_rate_multiplier=1.0, seed=42)
    assert result_a == result_b


def test_run_scenario_returns_sane_queueing_output():
    result = run_scenario(n_capacity=30, arrival_rate_multiplier=1.0, seed=42)
    assert result["n_patients"] > 0
    assert result["mean_wait_minutes"] >= 0
    assert result["mean_total_minutes"] >= result["mean_wait_minutes"]
    assert result["p95_wait_minutes"] >= result["mean_wait_minutes"]


def test_higher_arrival_multiplier_does_not_decrease_load():
    # More demand at fixed capacity should not, on average, mean fewer
    # patients served - a basic queueing sanity check, not a tight bound
    # (the surrogate/CP chapters separately document real single-seed
    # noise at sparse corners, so this compares generous multipliers).
    low = run_scenario(n_capacity=30, arrival_rate_multiplier=0.7, seed=7)
    high = run_scenario(n_capacity=30, arrival_rate_multiplier=1.5, seed=7)
    assert high["n_patients"] >= low["n_patients"]
