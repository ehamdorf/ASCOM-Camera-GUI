"""
Tests for Threads.
"""

import time
from unittest.mock import MagicMock, patch

import numpy as np
import pytest

import NIRvana_HS_PyQt5_GUI as gui


@pytest.fixture
def mock_camera():
    """
    Create a mock ASCOM camera.
    """
    cam = MagicMock()
    cam.Connected = True
    cam.CameraState = 0  # Idle
    cam.CCDTemperature = -40.0
    cam.SetCCDTemperature = -40.0
    cam.ImageReady = False
    
    # Simulate a 10x10 raw frame
    cam.ImageArray = np.zeros((10, 10), dtype=np.uint16)
    
    def start_exposure(duration, light):
        cam.CameraState = 2  # Exposing
        # Simulate instant exposure completion for testing
        cam.ImageReady = True
        cam.CameraState = 0  # Idle
        
    cam.StartExposure.side_effect = start_exposure
    cam.AbortExposure.return_value = None
    return cam


@pytest.fixture
def experiment_spec():
    """
    Create a minimal experiment for thread testing.
    """
    params = gui.CollapsibleWidgetInternalGUI._default_params()
    params["exp_array"] = [0.1]
    params["time_unit"] = "s"
    return gui.ExperimentSpec(
        index=0,
        title="Test Exp",
        params=params
    )


def test_experiment_thread_emits_save_requested(qapp, mock_camera, experiment_spec):
    """
    ExperimentThread should emit save_requested with correct metadata 
    after a successful exposure.
    """
    thread = gui.ExperimentThread(mock_camera, [experiment_spec])
    
    save_mock = MagicMock()
    thread.save_requested.connect(save_mock)
    
    # Run the thread logic synchronously for testing
    thread.run()
    
    assert save_mock.called
    args, kwargs = save_mock.call_args
    arr, meta = args
    
    assert isinstance(arr, np.ndarray)
    assert isinstance(meta, gui.ExposureMeta)
    assert meta.exposure_s == 0.1
    assert meta.exp_idx == 0
    assert meta.img_idx == 0


def test_experiment_thread_respects_abort(qapp, mock_camera, experiment_spec):
    """
    ExperimentThread should halt execution when run_flag is set to False.
    """
    thread = gui.ExperimentThread(mock_camera, [experiment_spec])
    thread.run_flag = False
    
    save_mock = MagicMock()
    thread.save_requested.connect(save_mock)
    
    thread.run()
    
    assert not save_mock.called


def test_preview_thread_aborts_cleanly(qapp, mock_camera):
    """
    PreviewThread abort should set run_flag to False and call AbortExposure.
    """
    params = gui.CollapsibleWidgetInternalGUI._default_params()
    thread = gui.PreviewThread(mock_camera, params)
    
    thread.abort(timeout=1.0)
    
    assert thread.run_flag is False
    mock_camera.AbortExposure.assert_called_once()


def test_saver_thread_processes_queue(qapp):
    """
    SaverThread should process queued items and emit progress signals.
    """
    saver = gui.SaverThread()
    saver._stopped = False
    
    progress_mock = MagicMock()
    saver.progress.connect(progress_mock)
    
    # Mock fits.writeto to prevent actual disk I/O
    with patch("astropy.io.fits.writeto") as mock_writeto:
        # Manually push a dummy job
        arr = np.zeros((10, 10), dtype=np.uint16)
        hdr = {"EXPTIME": 1.0}
        path = "/tmp/dummy_test_file.fits"
        
        saver.queue.put((arr, hdr, path))
        saver.queue.put(saver._SENTINEL)
        
        # Run the loop synchronously
        saver.run()
        
        assert mock_writeto.called
        assert progress_mock.called
