"""Tests for sxm_viewer.gui.dialogs.profile_units.

Qt-free by design (see that module's docstring) - run with plain pytest,
no QApplication required:

    python -m pytest sxm_viewer/tests/test_profile_units.py
"""
from __future__ import annotations

import numpy as np
import pytest

from sxm_viewer.gui.dialogs.profile_units import (
    compose_axis_label,
    from_si_base,
    si_format,
    si_scale,
    to_si_base,
)


class TestToSiBase:
    def test_nm_converts_to_metres(self):
        values, unit, auto_prefix = to_si_base([1.0, 10.0, 232.211], "nm")
        assert unit == "m"
        assert auto_prefix is True
        np.testing.assert_allclose(values, [1e-9, 1e-8, 232.211e-9])

    @pytest.mark.parametrize("unit", ["A", "V", "Hz", "m"])
    def test_already_base_units_pass_through_unscaled(self, unit):
        values, out_unit, auto_prefix = to_si_base([1.0, -2.5, 3.0], unit)
        assert out_unit == unit
        assert auto_prefix is True
        np.testing.assert_allclose(values, [1.0, -2.5, 3.0])

    def test_unit_is_case_insensitive_for_recognition_but_preserves_symbol(self):
        # normalize_unit_and_data always yields canonical-cased symbols
        # ("A", "V", "Hz", "nm"); this just guards against a stray case
        # mismatch silently falling through to "unrecognized".
        values, unit, auto_prefix = to_si_base([5.0], "hz")
        assert unit == "hz"
        assert auto_prefix is True
        np.testing.assert_allclose(values, [5.0])

    def test_unrecognized_unit_passes_through_without_auto_prefix(self):
        values, unit, auto_prefix = to_si_base([1.0, 2.0], "deg")
        assert unit == "deg"
        assert auto_prefix is False
        np.testing.assert_allclose(values, [1.0, 2.0])

    def test_blank_unit(self):
        values, unit, auto_prefix = to_si_base([1.0], "")
        assert unit == ""
        assert auto_prefix is False
        np.testing.assert_allclose(values, [1.0])

    def test_none_values_returns_empty_array(self):
        values, unit, auto_prefix = to_si_base(None, "A")
        assert values.size == 0
        assert unit == "A"

    def test_never_produces_the_1e_minus_15_symptom(self):
        # The reported bug: a Hz-labelled axis with values sitting at ~1e-15
        # because a distance-style factor was applied to a channel already
        # in physical units. to_si_base for a genuine "Hz" unit must be a
        # pure pass-through - no factor of any kind.
        raw = np.array([-42.3, 0.0, 17.8])
        values, unit, _ = to_si_base(raw, "Hz")
        assert unit == "Hz"
        np.testing.assert_array_equal(values, raw)


class TestFromSiBase:
    def test_inverts_nm_conversion(self):
        original_nm = 138.8
        converted_m = original_nm * 1e-9
        assert from_si_base(converted_m, "m") == pytest.approx(original_nm)

    def test_no_op_for_units_never_converted(self):
        assert from_si_base(5.2, "A") == pytest.approx(5.2)


class TestComposeAxisLabel:
    def test_plain_name_and_unit(self):
        label, unit = compose_axis_label("Current", "A")
        assert (label, unit) == ("Current", "A")

    def test_strips_bracketed_unit_already_in_name(self):
        # The exact reported bug: channel caption already contains "[A]".
        label, unit = compose_axis_label("Current [A]", "A")
        assert label == "Current"
        assert unit == "A"

    def test_strips_parenthesized_unit_already_in_name(self):
        label, unit = compose_axis_label("Frequency shift (Hz)", "Hz")
        assert label == "Frequency shift"
        assert unit == "Hz"

    def test_never_produces_doubled_unit_rendering(self):
        label, unit = compose_axis_label("Current [A]", "A")
        rendered = f"{label} ({unit})" if unit else label
        assert rendered == "Current (A)"
        assert "[A] (A)" not in rendered

    def test_leaves_unrelated_brackets_alone(self):
        label, unit = compose_axis_label("Current [raw]", "A")
        assert label == "Current [raw]"
        assert unit == "A"

    def test_blank_unit_leaves_name_untouched(self):
        label, unit = compose_axis_label("Current", "")
        assert label == "Current"
        assert unit == ""


class TestSiScale:
    def test_order_one_value_gets_no_prefix(self):
        scale, prefix = si_scale(52.0)
        assert prefix == ""
        assert scale == pytest.approx(1.0)

    def test_pico_range_current(self):
        scale, prefix = si_scale(52e-12)
        assert prefix == "p"
        assert 52e-12 * scale == pytest.approx(52.0)

    def test_nano_range_distance(self):
        scale, prefix = si_scale(139.3e-9)
        assert prefix == "n"
        assert 139.3e-9 * scale == pytest.approx(139.3)

    def test_kilo_range(self):
        scale, prefix = si_scale(2500.0)
        assert prefix == "k"
        assert 2500.0 * scale == pytest.approx(2.5)

    def test_zero_is_handled_without_raising(self):
        scale, prefix = si_scale(0.0)
        assert prefix == ""
        assert scale == pytest.approx(1.0)

    def test_negative_values_use_the_same_prefix_as_positive(self):
        pos_scale, pos_prefix = si_scale(52e-12)
        neg_scale, neg_prefix = si_scale(-52e-12)
        assert pos_prefix == neg_prefix
        assert pos_scale == pytest.approx(neg_scale)


class TestSiFormat:
    def test_formats_with_chosen_prefix_and_unit(self):
        text = si_format(52e-12, unit="A", precision=3)
        assert "pA" in text
        assert "52" in text

    def test_never_produces_a_bare_multiplier(self):
        # The reported bug: a Hz axis with a bare "1e-15" multiplier
        # floating above it. A formatted scalar must never look like that.
        text = si_format(-42.3, unit="Hz", precision=4)
        assert "e-" not in text
        assert "e+" not in text

    def test_blank_unit_omits_trailing_space(self):
        text = si_format(1.5, unit="", precision=3)
        assert not text.endswith(" ")

    def test_blank_name(self):
        label, unit = compose_axis_label("", "A")
        assert label == ""
        assert unit == "A"
