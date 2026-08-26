"""
Tests for logic that does not require hardware or Qwidgets.
"""

from pathlib import Path

import numpy as np
import pytest

import NIRvana_HS_PyQt5_GUI as gui


def test_default_params_do_not_prepopulate_temp_fragment():
    """
    Temperature filename must be empty unless enabled.
    """
    params = gui.CollapsibleWidgetInternalGUI._default_params()

    assert params["name_temp"] == ""
    assert params["include_temp"] is False


class TestExposureParser:
    """
    Tests for restricted exposure array parser.
    """

    @staticmethod
    def parse(text):
        """
        Wrapper around parser.
        """
        return gui.CollapsibleWidgetInternalGUI._parse_exposure_array(text)

    def test_simple_list(self):
        """
        Python list -> parse directly.
        """
        assert self.parse("[1, 2.5]") == [1.0, 2.5]

    def test_nested_list_is_flattened(self):
        """
        Nested list -> flat exposure array.
        """
        assert self.parse("[[1, 2], [3]]") == [1.0, 2.0, 3.0]

    def test_tuple_is_accepted(self):
        """
        Tuples -> lists.
        """
        assert self.parse("(1, 2)") == [1.0, 2.0]

    def test_unary_plus_is_accepted(self):
        """
        Unary plus -> positive numbers.
        """
        assert self.parse("[+1]") == [1.0]

    def test_np_repeat(self):
        """
        np.repeat -> list.
        """
        assert self.parse("np.repeat([1], 3)") == [1.0, 1.0, 1.0]

    def test_np_arange(self):
        """
        np.arange -> list.
        """
        assert self.parse("np.arange(1, 4)") == [1.0, 2.0, 3.0]

    def test_np_linspace(self):
        """
        np.linspace -> list.
        """
        assert self.parse("np.linspace(1, 3, 3)") == [1.0, 2.0, 3.0]

    def test_np_ones(self):
        """
        np.ones -> list.
        """
        assert self.parse("np.ones(3)") == [1.0, 1.0, 1.0]

    def test_np_full(self):
        """
        np.full -> list.
        """
        assert self.parse("np.full(2, 5)") == [5.0, 5.0]

    def test_np_array(self):
        """
        np.array -> list.
        """
        assert self.parse("np.array([1, 2])") == [1.0, 2.0]

    def test_empty_array_rejected(self):
        """
        Empty exposure arrays are invalid.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[]")

    def test_zero_exposure_rejected(self):
        """
        Zero-second exposures are invalid.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[0]")

    def test_negative_exposure_rejected(self):
        """
        Negative exposures are invalid.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[-1]")

    def test_nonfinite_exposure_rejected(self):
        """
        Infinite exposures are invalid.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[1e400]")

    def test_boolean_rejected(self):
        """
        Boolean literals should not be accepted as numbers.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[True]")

    def test_unknown_name_rejected(self):
        """
        Only np.* names should be allowed.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("[x]")

    def test_unwhitelisted_numpy_function_rejected(self):
        """
        Only explicitly whitelisted numpy functions should be allowed.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("np.sum([1])")

    def test_keyword_arguments_rejected(self):
        """
        Keyword arguments are intentionally unsupported.
        """
        with pytest.raises(gui.ExposureParserError):
            self.parse("np.repeat([1], repeats=2)")

    def test_max_length_enforced(self, monkeypatch):
        """
        Arrays longer than MAX_EXPOSURE_ARRAY_LEN should be rejected.
        """
        monkeypatch.setattr(gui, "MAX_EXPOSURE_ARRAY_LEN", 5)

        with pytest.raises(gui.ExposureParserError):
            self.parse("[1, 2, 3, 4, 5, 6]")


class TestFilenameSanitisation:
    """
    Tests for filesystem-safe filename generation.
    """

    def test_plain_string_unchanged(self):
        """
        Safe strings should be unchanged.
        """
        assert gui.SaverThread._sanitize_filename("abc") == "abc"

    def test_invalid_characters_replaced(self):
        """
        Unsafe characters should be replaced.
        """
        assert gui.SaverThread._sanitize_filename("a:b") == "a_b"
        assert gui.SaverThread._sanitize_filename("a/b") == "a_b"

    def test_empty_becomes_image(self):
        """
        Empty strings should fall back to image.
        """
        assert gui.SaverThread._sanitize_filename("") == "image"

    def test_whitespace_only_becomes_image(self):
        """
        Whitespace-only strings should fall back to image.
        """
        assert gui.SaverThread._sanitize_filename("   ") == "image"

    def test_parent_directory_traversal_neutralised(self):
        """
        Path sequences should not remain unchanged.
        """
        result = gui.SaverThread._sanitize_filename("..")
        assert ".." not in result


class TestJsonDefault:
    """
    Tests for the json fallback.
    """

    def test_numpy_integer(self):
        """
        Numpy integers -> Python ints.
        """
        result = gui.MainWindow._json_default(np.int64(5))
        assert result == 5
        assert isinstance(result, int)

    def test_numpy_float(self):
        """
        Numpy floats -> Python floats.
        """
        result = gui.MainWindow._json_default(np.float64(1.5))
        assert result == 1.5
        assert isinstance(result, float)

    def test_numpy_array(self):
        """
        Numpy arrays should become lists.
        """
        result = gui.MainWindow._json_default(np.array([1, 2, 3]))
        assert result == [1, 2, 3]

    def test_pathlike(self):
        """
        Path objects should become native strings.
        """
        result = gui.MainWindow._json_default(Path("x"))
        assert result == "x"

    def test_unknown_type_raises(self):
        """
        Unknown types should raise TypeError.
        """
        with pytest.raises(TypeError):
            gui.MainWindow._json_default(object())
