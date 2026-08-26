"""
Tests for fits path/header generation and saver behaviour.
"""

from datetime import datetime
from pathlib import Path

import numpy as np
import pytest
from astropy.io import fits
from astropy.time import Time

import NIRvana_HS_PyQt5_GUI as gui


def make_meta(params=None, **overrides):
    """
    Create an ExposureMeta for tests.
    """
    base_params = gui.CollapsibleWidgetInternalGUI._default_params()
    if params is not None:
        base_params.update(params)

    values: dict = {
        "exp_idx": 0,
        "img_idx": 0,
        "rep_idx": 0,
        "exposure_s": 1.0,
        "params": base_params,
        "temp": -40.0,
        "target_temp": -40.0,
        "ts_utc": Time("2024-01-02T03:04:05", format="isot", scale="ut1"),
        "ts_ltc": datetime(2024, 1, 2, 3, 4, 5),
        "filter": None,
        "ra_target": None,
        "dec_target": None,
        "ra_mount": None,
        "dec_mount": None,
        "alt_mount": None,
        "az_mount": None,
        "focuser_pos": None,
    }

    values.update(overrides)
    return gui.ExposureMeta(**values)


@pytest.fixture
def saver():
    """
    Create a SaverThread instance without starting it.
    """
    return gui.SaverThread()


def test_build_path_basic(saver):
    """
    Basic path construction should include experiment/repeat/image indices.
    """
    params = {
        "nameprefix": "test",
        "savedirectory": "/tmp",
    }

    meta = make_meta(params=params, exp_idx=0, img_idx=1, rep_idx=2)
    path = saver._build_path(meta)

    assert Path(path).name == "test_exp1_rep3_img2.fits"


def test_build_path_exp_time_unit_conversion(saver):
    """
    Exposure suffix should convert seconds back into chosen units.
    """
    params = {
        "nameprefix": "exp",
        "savedirectory": "/tmp",
        "include_exp_time": True,
        "time_unit": "ms",
    }

    meta = make_meta(params=params, exposure_s=0.001)
    name = Path(saver._build_path(meta)).name

    assert "_1ms_exp" in name


def test_build_path_date_time(saver):
    """
    Date/time suffix should be present when enabled.
    """
    params = {
        "nameprefix": "dt",
        "savedirectory": "/tmp",
        "include_date_time": True,
    }

    meta = make_meta(
        params=params,
        ts_ltc=datetime(2024, 1, 2, 13, 4, 0),
    )
    name = Path(saver._build_path(meta)).name

    assert "_02_01_2024_" in name
    assert "pm" in name


def test_build_path_temperature(saver):
    """
    Temperature suffix should use the target temperature.
    """
    params = {
        "nameprefix": "temp",
        "savedirectory": "/tmp",
        "include_temp": True,
    }

    meta = make_meta(params=params, target_temp=-39.6)
    name = Path(saver._build_path(meta)).name

    assert "_-40C_exp" in name


def test_build_path_sanitizes_prefix(saver):
    """
    Unsafe filename characters should not persist.
    """
    params = {
        "nameprefix": "a/b:c*d",
        "savedirectory": "/tmp",
    }

    meta = make_meta(params=params)
    name = Path(saver._build_path(meta)).name

    invalid_characters = set('<>:"/\\|?*')
    assert not (set(name) & invalid_characters)
    assert name.endswith(".fits")


def test_build_header_core(saver):
    """
    Core fits header keywords should be present and correct.
    """
    params = {
        "width": 640,
        "height": 512,
        "bin_x": 1,
        "bin_y": 1,
    }

    meta = make_meta(params=params, exposure_s=2.5, focuser_pos=None)
    hdr = saver._build_header(meta)

    assert hdr["EXPTIME"] == 2.5
    assert hdr["BZERO"] == 32768.0
    assert hdr["BSCALE"] == 1.0
    assert "FOC-POS" not in hdr

    assert hdr["DATE-OBS"].startswith("2024-01-02T03:04:05")

    meta_with_focuser = make_meta(params=params, focuser_pos=123)
    hdr_with_focuser = saver._build_header(meta_with_focuser)

    assert hdr_with_focuser["FOC-POS"] == 123


def test_build_header_subframe_coordinates(saver):
    """
    Subframe origin keywords should be binned coordinates.
    """
    params = {
        "start_x": 4,
        "start_y": 6,
        "width": 64,
        "height": 64,
        "bin_x": 2,
        "bin_y": 2,
    }

    meta = make_meta(params=params)
    hdr = saver._build_header(meta)

    assert hdr["XORGSUBF"] == 2
    assert hdr["YORGSUBF"] == 3


def test_header_crpix_one_based(saver):
    """
    CRPIX should be (NAXIS + 1) / 2 for WCS.
    """
    params = {
        "width": 640,
        "height": 512,
        "bin_x": 1,
        "bin_y": 1,
    }

    meta = make_meta(params=params)
    hdr = saver._build_header(meta)

    assert hdr["CRPIX1"] == 320.5
    assert hdr["CRPIX2"] == 256.5
