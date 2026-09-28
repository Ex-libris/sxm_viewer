"""Tests for sxm_viewer.gui.dialogs.profile_measure.

    python -m pytest sxm_viewer/tests/test_profile_measure.py
"""
from __future__ import annotations

import numpy as np
import pytest

from sxm_viewer.gui.dialogs.profile_measure import (
    band_statistics,
    find_peaks,
    fmt_tol,
    peaks_in_range,
    snap_edges,
)


class TestFmtTol:
    def test_value_and_tolerance_share_one_si_prefix(self):
        # 139.3 nm +/- 0.46 nm, expressed in metres in - the exact scenario
        # from the redesign brief ("139.3 +/- 0.46 nm").
        text = fmt_tol(139.3e-9, 0.46e-9, "m", digits=4)
        assert text.endswith("nm")
        assert "±" in text
        assert "139.3" in text
        assert "0.46" in text

    def test_never_claims_more_precision_than_a_bare_format(self):
        # The reported bug: "79.328 nm" on ~10 nm sampling claims 3 decimal
        # digits the data can't support. fmt_tol's digits param controls
        # significant figures on the value; the tolerance is always shown
        # alongside it, so the display is self-documenting about precision.
        text = fmt_tol(79.328e-9, 5e-9, "m", digits=3)
        assert "±" in text

    def test_small_current_value_gets_pico_prefix(self):
        text = fmt_tol(52e-12, 0.5e-12, "A", digits=3)
        assert "pA" in text


class TestFindPeaks:
    def test_finds_single_clean_peak(self):
        x = np.linspace(0, 10, 200)
        y = np.exp(-0.5 * ((x - 5.0) / 0.3) ** 2)
        peaks = find_peaks(y)
        assert peaks.size >= 1
        assert abs(x[peaks[np.argmax(y[peaks])]] - 5.0) < 0.2

    def test_flat_signal_has_no_peaks_above_threshold(self):
        y = np.ones(50)
        peaks = find_peaks(y)
        assert peaks.size == 0

    def test_too_few_points_returns_empty(self):
        assert find_peaks([1.0, 2.0]).size == 0

    def test_ignores_peaks_below_fraction_threshold(self):
        y = np.zeros(100)
        y[50] = 100.0  # one dominant peak
        y[10] = 1.0    # tiny bump, well under 10% of the range
        peaks = find_peaks(y, frac=0.10)
        assert 50 in peaks
        assert 10 not in peaks


class TestBandStatistics:
    def test_basic_stats_over_a_linear_ramp(self):
        x = np.linspace(0.0, 10.0, 11)
        y = 2.0 * x + 1.0
        stats = band_statistics(x, y, 2.0, 8.0)
        assert stats["n"] == 7
        assert stats["slope"] == pytest.approx(2.0, abs=1e-6)
        assert stats["y_left"] == pytest.approx(5.0)
        assert stats["y_right"] == pytest.approx(17.0)

    def test_handles_reversed_edges(self):
        x = np.linspace(0.0, 10.0, 11)
        y = x.copy()
        stats_fwd = band_statistics(x, y, 2.0, 8.0)
        stats_rev = band_statistics(x, y, 8.0, 2.0)
        assert stats_fwd["n"] == stats_rev["n"]
        assert stats_fwd["mean"] == pytest.approx(stats_rev["mean"])

    def test_single_point_band_omits_derived_stats(self):
        x = np.linspace(0.0, 10.0, 11)
        y = x.copy()
        stats = band_statistics(x, y, 4.9, 5.0)
        assert stats["n"] <= 1
        assert "mean" not in stats
        assert "slope" not in stats

    def test_constant_signal_has_zero_slope_and_rms(self):
        x = np.linspace(0.0, 10.0, 11)
        y = np.full_like(x, 3.0)
        stats = band_statistics(x, y, 0.0, 10.0)
        assert stats["slope"] == pytest.approx(0.0, abs=1e-9)
        assert stats["rms"] == pytest.approx(0.0, abs=1e-9)
        assert stats["mean"] == pytest.approx(3.0)


class TestPeaksInRange:
    def test_counts_only_peaks_inside_band(self):
        x = np.linspace(0.0, 10.0, 11)
        peak_indices = [1, 5, 9]  # x = 1, 5, 9
        assert peaks_in_range(x, peak_indices, 4.0, 9.5) == 2

    def test_empty_peaks_returns_zero(self):
        x = np.linspace(0.0, 10.0, 11)
        assert peaks_in_range(x, [], 0.0, 10.0) == 0

    def test_handles_reversed_band(self):
        x = np.linspace(0.0, 10.0, 11)
        assert peaks_in_range(x, [5], 9.0, 1.0) == 1


class TestSnapEdges:
    def test_snaps_each_edge_independently_within_tolerance(self):
        peak_x = np.array([1.0, 5.0, 9.0])
        lo, hi = snap_edges((1.2, 8.8), peak_x, tol_px=6, px_size=0.1)
        assert lo == pytest.approx(1.0)
        assert hi == pytest.approx(9.0)

    def test_edge_far_from_any_peak_is_unchanged(self):
        peak_x = np.array([1.0, 5.0, 9.0])
        lo, hi = snap_edges((3.0, 8.8), peak_x, tol_px=2, px_size=0.1)
        assert lo == pytest.approx(3.0)  # too far from peak at 1.0 or 5.0
        assert hi == pytest.approx(9.0)

    def test_no_peaks_returns_edges_unchanged(self):
        lo, hi = snap_edges((3.0, 8.0), np.array([]), tol_px=5, px_size=0.1)
        assert (lo, hi) == (3.0, 8.0)

    def test_result_is_always_ordered_low_high(self):
        peak_x = np.array([1.0, 9.0])
        lo, hi = snap_edges((9.1, 0.9), peak_x, tol_px=6, px_size=0.1)
        assert lo <= hi
