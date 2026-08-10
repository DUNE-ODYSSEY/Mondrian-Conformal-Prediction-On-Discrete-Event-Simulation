"""Unit tests for the pure calibration math in src/uq - the split conformal
quantile correction and the Mondrian category-assignment logic. These don't
touch any generated data files, so they run the same in any clone."""

import numpy as np
import pytest

from src.uq.standard_cp import conformal_quantile as standard_conformal_quantile
from src.uq.mondrian_cp import assign_categories, conformal_quantile


def test_standard_and_mondrian_quantile_functions_agree():
    # Both modules define the same finite-sample-corrected quantile; they
    # should never be allowed to silently drift apart.
    scores = np.array([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0])
    assert conformal_quantile(scores, 0.1) == standard_conformal_quantile(scores, 0.1)


def test_conformal_quantile_is_a_finite_sample_corrected_level():
    # For n calibration scores and miscoverage alpha, split conformal takes
    # the ceil((n+1)(1-alpha))/n empirical quantile, not the naive (1-alpha)
    # quantile - this is what gives the finite-sample coverage guarantee.
    scores = np.arange(1, 101, dtype=float)  # 1..100
    alpha = 0.1
    n = len(scores)
    expected_level = min(np.ceil((n + 1) * (1 - alpha)) / n, 1.0)
    q = conformal_quantile(scores, alpha)
    assert q == np.quantile(scores, expected_level, method="higher")


def test_conformal_quantile_never_exceeds_the_max_score():
    # The level is capped at 1.0, so the returned quantile can never fall
    # outside the observed calibration scores even for small n / large alpha.
    scores = np.array([1.0, 2.0, 3.0])
    q = conformal_quantile(scores, alpha=0.5)
    assert q <= scores.max()
    assert q >= scores.min()


def test_conformal_quantile_shrinks_as_alpha_grows():
    # Looser miscoverage tolerance (larger alpha) should never require a
    # *wider* calibrated quantile than a stricter one, on the same scores.
    scores = np.random.default_rng(0).exponential(size=200)
    q_strict = conformal_quantile(scores, alpha=0.05)
    q_loose = conformal_quantile(scores, alpha=0.2)
    assert q_loose <= q_strict


def test_assign_categories_produces_nine_cells():
    rng = np.random.default_rng(0)
    import pandas as pd

    df = pd.DataFrame({
        "n_capacity": rng.integers(15, 46, size=300),
        "arrival_rate_multiplier": rng.uniform(0.7, 1.4, size=300),
    })
    cap_bounds = np.quantile(df["n_capacity"], [1 / 3, 2 / 3])
    arr_bounds = np.quantile(df["arrival_rate_multiplier"], [1 / 3, 2 / 3])
    labels = assign_categories(df, cap_bounds, arr_bounds)

    assert set(labels) <= {
        f"staff={s}/arrival={a}" for s in ["Low", "Med", "High"] for a in ["Low", "Med", "High"]
    }
    # With bounds derived from this same data's own terciles, all 9 cells
    # should be populated for a reasonably sized, well-spread sample.
    assert len(set(labels)) == 9


def test_assign_categories_is_deterministic_given_bounds():
    import pandas as pd

    df = pd.DataFrame({"n_capacity": [10, 20, 30], "arrival_rate_multiplier": [0.5, 1.0, 1.5]})
    cap_bounds = [15, 25]
    arr_bounds = [0.8, 1.2]
    labels_1 = assign_categories(df, cap_bounds, arr_bounds)
    labels_2 = assign_categories(df, cap_bounds, arr_bounds)
    assert list(labels_1) == list(labels_2) == [
        "staff=Low/arrival=Low",
        "staff=Med/arrival=Med",
        "staff=High/arrival=High",
    ]
