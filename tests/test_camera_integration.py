"""
Tests for ASCOM simulator camera.

Skipped unless RUN_CAMERA_INTEGRATION=1.
"""

import os
import time

import pytest

import NIRvana_HS_PyQt5_GUI as gui

pytestmark = [
    pytest.mark.camera_integration,
    pytest.mark.skipif(
        os.environ.get("RUN_CAMERA_INTEGRATION", "0") != "1",
        reason="Requires a running camera simulator (set RUN_CAMERA_INTEGRATION=1)",
    ),
]


def test_simulated_camera_connects():
    """
    Connect to ASCOM alpaca camera simulator and read basic properties.
    """
    cam = gui.Camera(gui.ALPACA_ADDRESS, gui.ALPACA_CAMERA_DEVICE)
    connected = False

    try:
        cam.Connected = True
        connected = True

        assert bool(cam.Connected) is True
        assert int(cam.CameraXSize) > 0
        assert int(cam.CameraYSize) > 0

    finally:
        if connected:
            try:
                cam.Connected = False
            except Exception:
                pass


def test_simulated_camera_short_exposure():
    """
    Verify short exposure completes and returns an image.
    """
    cam = gui.Camera(gui.ALPACA_ADDRESS, gui.ALPACA_CAMERA_DEVICE)
    connected = False

    try:
        cam.Connected = True
        connected = True

        # Use a full-frame ROI before exposing.
        cam.BinX = 1
        cam.BinY = 1
        cam.StartX = 0
        cam.StartY = 0
        cam.NumX = int(cam.CameraXSize)
        cam.NumY = int(cam.CameraYSize)

        cam.StartExposure(0.1, True)

        deadline = time.monotonic() + 10.0
        while not cam.ImageReady and time.monotonic() < deadline:
            time.sleep(0.01)

        assert bool(cam.ImageReady) is True

        image = cam.ImageArray
        assert image is not None

    finally:
        if connected:
            try:
                cam.Connected = False
            except Exception:
                pass
