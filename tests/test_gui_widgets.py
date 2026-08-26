"""
Tests for Qwidget logic and UI.
"""

from unittest.mock import MagicMock

import pytest

import NIRvana_HS_PyQt5_GUI as gui


@pytest.fixture
def internal_gui(qapp):
    """
    Create an internal experiment settings widget.
    """
    widget = gui.CollapsibleWidgetInternalGUI()
    yield widget
    widget.deleteLater()


@pytest.fixture
def main_window(qapp):
    """
    Create the main window and ensure poll timer is stopped afterwards.
    """
    window = gui.MainWindow()
    yield window
    window._poll_timer.stop()
    window.deleteLater()


def test_example_filename_initial(internal_gui):
    """
    Example filename should not contain old elements.
    """
    assert internal_gui._file_name_label.text() == "Example: image_1.fits"


def test_example_filename_prefix(internal_gui):
    """
    Changing prefix should update example filename.
    """
    internal_gui._image_name_line_edit.setText("obs")
    assert internal_gui._file_name_label.text().startswith("Example: obs")


def test_example_filename_temp_checkbox(internal_gui):
    """
    Temperature element should appear only when the checkbox is enabled.
    """
    internal_gui.set_temp_value(-40.0)

    internal_gui._file_temp_checkbox.setChecked(True)
    assert "_-40C" in internal_gui._file_name_label.text()

    internal_gui._file_temp_checkbox.setChecked(False)
    assert "_-40C" not in internal_gui._file_name_label.text()


def test_roi_apply_enforces_alignment(qapp):
    """
    ROI should enforce binning and hardware alignment constraints.
    """
    window = gui.ROIWindow()

    window._edits[0].setText("2")
    window._edits[1].setText("0")
    window._edits[2].setText("100")
    window._edits[3].setText("64")

    window._bins[0].setCurrentText("1x")
    window._bins[1].setCurrentText("1x")

    window._on_apply()

    data = window.data
    assert data is not None

    end_x = data["start_x"] + data["width"]

    assert end_x % gui.CAMERA_X_ALIGNMENT_DIVISOR == 0
    assert data["width"] % data["bin_x"] == 0
    assert data["height"] % data["bin_y"] == 0

    window.deleteLater()


def test_main_window_teardown_no_threads(main_window):
    """
    Teardown should succeed when no worker threads are active.
    """
    assert main_window._teardown_threads() is True


def test_run_blocked_when_preview_active(main_window):
    """
    Experiments should not start while preview is active.
    """
    main_window._preview_thread = object()

    main_window._run_experiments_pressed()

    assert main_window._exp_thread is None


def test_preview_blocked_when_experiment_active(main_window):
    """
    Preview should not start while experiments are active.
    """
    main_window._exp_thread = object()

    main_window._preview_pressed()

    assert main_window._preview_thread is None


def test_teardown_success_cleans_preview(main_window):
    """
    Teardown should delete a preview thread that stops cleanly.
    """
    thread = MagicMock()
    thread.wait.return_value = True
    thread.isFinished.return_value = True

    main_window._preview_thread = thread

    assert main_window._teardown_threads() is True
    assert main_window._preview_thread is None

    thread.abort.assert_called_once()
    thread.quit.assert_called_once()
    thread.deleteLater.assert_called_once()


def test_teardown_failure_keeps_thread(main_window):
    """
    Teardown should not delete a preview thread that fails to stop.
    """
    thread = MagicMock()
    thread.wait.return_value = False
    thread.isFinished.return_value = False

    main_window._preview_thread = thread

    assert main_window._teardown_threads() is False
    assert main_window._preview_thread is thread

    thread.deleteLater.assert_not_called()


def test_open_camera_settings_no_camera(main_window):
    """
    Camera settings should fail gracefully when no camera exists.
    """
    main_window.camera = None

    # Should not raise.
    main_window._open_camera_settings()
