import ast
import copy
import html
import json
import os
import re
import sys
import threading
import time
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from queue import Full, Queue

import astropy.units as u
import numpy as np
import serial

# import win32com.client
# import pythoncom
from alpaca.camera import Camera
from alpaca.focuser import Focuser
from alpaca.telescope import Telescope
from astropy.coordinates import Angle, EarthLocation
from astropy.io import fits
from astropy.time import Time
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5 import QtCore, QtGui, QtWidgets


def _env_int(name, default):
    """
    Return an integer environment variable or a fallback default.
    """
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _env_float(name, default):
    """
    Return a float environment variable or a fallback default.
    """
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


#  Device/driver addresses
ALPACA_ADDRESS = os.environ.get("ALPACA_ADDRESS", "localhost:32323") # ASCOM.PI.Camera.1 : NIRvanaHS     localhost:32323 : ASCOM Simulator
ALPACA_CAMERA_DEVICE = _env_int("ALPACA_CAMERA_DEVICE", 0)
ALPACA_MOUNT_DEVICE = _env_int("ALPACA_MOUNT_DEVICE", 0)
ALPACA_FOCUSER_DEVICE = _env_int("ALPACA_FOCUSER_DEVICE", 0)
CFW10_PORT = os.environ.get("CFW10_PORT", "/dev/ttyUSB0")

# Observatory config
OBSERVATORY_LATITUDE = _env_float("OBSERVATORY_LATITUDE", -42.4311)
OBSERVATORY_LONGITUDE = _env_float("OBSERVATORY_LONGITUDE", 147.2878)
OBSERVATORY_ELEVATION = _env_float("OBSERVATORY_ELEVATION", 646.0)
OBSERVATORY_NAME = os.environ.get("OBSERVATORY_NAME", "X")
DEFAULT_OBSERVER = os.environ.get("DEFAULT_OBSERVER", "obs")

# Camera specific values
CAMERA_MIN_TEMP_C = -40.0
CAMERA_MAX_TEMP_C = 20.0
CAMERA_THERMAL_SHOCK_THRESHOLD_C = -15.0
CAMERA_SENSOR_WIDTH = 640
CAMERA_SENSOR_HEIGHT = 512
CAMERA_PIXEL_SIZE_UM = 20.0
CAMERA_BIT_DEPTH = 16
UINT16_MAX = (2**CAMERA_BIT_DEPTH) - 1  # 65535
CAMERA_MIN_ROI_SIZE = 16
CAMERA_X_ALIGNMENT_DIVISOR = 4  # NIRvana HS ROI hardware constraint
CAMERA_GAIN_OPTIONS = ["High", "Low"]
CAMERA_ROSPEED_OPTIONS = ["3.125MHz", "12.5MHz", "25MHz"]
DEFAULT_GAIN = "High"
DEFAULT_ROSPEED = "3.125MHz"
TIME_UNIT_MULTIPLIERS = {"s": 1.0, "ms": 1e-3, "μs": 1e-6}

# CFW10 (filter-wheel) specific values
CFW10_BAUDRATE = 9600
CFW10_MOVE_TIMEOUT_S = 5.0
CFW10_STATUS_TIMEOUT_S = 10.0
CFW10_MOTOR_OFF_TIMEOUT_S = 12.0
CFW10_FILTER_NAMES_SHORT = [
    "U",
    "B",
    "V",
    "R",
    "I",
    "H",
    "Clear",
    "-",
    "-",
    "-",
]
CFW10_FILTER_NAMES_LONG = [
    "1/10 - 'U' filter",
    "2/10 - 'B' filter",
    "3/10 - 'V' filter",
    "4/10 - 'R' filter",
    "5/10 - 'I' filter",
    "6/10 - 'H' filter",
    "7/10 - 'Clear' filter",
    "8/10 - '-' filter",
    "9/10 - '-' filter",
    "10/10 - '-' filter",
]

# Threading/queue limits
POLL_INTERVAL_MS = 500
CAMERA_ABORT_TIMEOUT_S = 5.0
SAVER_QUEUE_MAXSIZE = 100
SAVER_QUEUE_PUT_TIMEOUT_S = 2.0
MAX_EXPOSURE_ARRAY_LEN = 10_000


class CFW10(serial.Serial):
    """
    SBIG CFW-10 filterwheel class.
    Derived from the serial class (inherits serial.Serial)

    Instantiation parameters can be used to setup the serial port. 
    """

    def __init__(self, *args, **kwargs):
        self._lock = threading.RLock()

        # 6 byte communication packet always begins 0x5A,0x03
        self.pkt = [0xA5, 0x03, 0x00, 0x00, 0x00, 0x00]
        serial.Serial.__init__(self, *args, **kwargs)
        self.fname = CFW10_FILTER_NAMES_SHORT
        self._choose_filter(1)

    def _cmd(self, c, data):
        """
        Setup a command packet for the cfw-10.
        """
        self.pkt[2] = c
        self.pkt[3] = data % 256
        self.pkt[4] = data // 256
        self.pkt[5] = 0
        for i in self.pkt[0:5]:
            self.pkt[5] += i

    def _choose_filter(self, f):
        """
        Setup a CFW packet to change to filter f (1..10).
        """
        if f > 10 or f < 1:
            f = 1
        self._cmd(17, f)

    def _sendpkt(self):
        """
        Send the command in self.pkt to the CFW-10 on the rs232 port (ser).
        """
        while self.out_waiting > 0:
            time.sleep(self.out_waiting * 0.001 + 0.01)
            if self.out_waiting > 0:
                self.reset_output_buffer()
        time.sleep(0.01)
        self.reset_input_buffer()  # Flush the input so we get the ack packet
        for i in self.pkt:
            self.write(bytes([i]))

    def _get_status(self, n):
        r"""
        Read status word n from the cfw10, returned as chr or '\0xff' on error.
        """
        with self._lock:
            self._cmd(2, n)  # Command 2 is the status command
            self._sendpkt()
            timeout = time.time()
            pkt = b""
            while (
                time.time() - timeout
            ) < CFW10_STATUS_TIMEOUT_S:
                if self.in_waiting >= 6:
                    pkt = self.read(1)  # Read header byte
                    if pkt[0] == 0xA5:
                        pkt = self.read(5)  # Read the rest of the packet
                        if pkt[3] == 64: 
                            return pkt[2]  # Return byte directly
                        else:
                            print("cfw10 Status - Bad packet", pkt[3])
                            return 0xFE
                time.sleep(0.01)
            print("cfw10 Status - Timeout talking to CF-W10")
            return 0xFF

    def _wait_motor_off(self):
        """
        Read CFW10 status and return when motor is off or timeout.

        Returns
        -------
            0 for success, -1 for failed (timeout), -2 comm. error.
        """
        with self._lock:
            timeOut = time.time()
            state = self._get_status(0)
            while state & 0xF0 != 0:
                if (state & 0xFE) == 0xFE:  # comm. error
                    return -2
                time.sleep(0.02)
                if (time.time() - timeOut) > CFW10_MOTOR_OFF_TIMEOUT_S:
                    print("_wait_motor_off - timeout", state)
                    return -1
                if state & 0x40 != 0:
                    print("_wait_motor_off - timeout", state)
                    return -1
                state = self._get_status(0)
            return 0

    def change_filter(self, f):
        """
        Changes to filter f (1..10)

        Returns
        -------
            Number of filter in place when the motor stops, or -(ve) number for error.
        """
        with self._lock:
            self._choose_filter(f)
            self._sendpkt()
            time.sleep(0.01)
            if self.in_waiting == 0:
                return (-1, "cfw10 - No command ack.")
            ack = self.read(1)
            if not ack or ack[0] != 6:
                return (-2, "cfw10 - Wrong ack. code")  # Wrong command confirmation

            if self._wait_motor_off() < 0:
                return (
                    -3,
                    "cfw10 - Motor-off timeout",
                )  # Filter change probably failed

            # Return the filer number that the cfw10 actually moved to
            state = self._get_status(0)
            return (state, self.fname[state - 1])

    def query_state(self):
        """
        Check the current position with a threading lock.
        """
        with self._lock:
            state = self._get_status(0)
            return state


class ExposureParserError(ValueError):
    """
    Raised when an exposure expression cannot be parsed safely.
    """


class CollapsibleWidgetInternalGUI(QtWidgets.QWidget):
    """
    Create internal GUI elements for a custom collapsible window widget.

    Called and placed like any other QWidget. Added to the internal
    containers of the collapsiblewidget's layout.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        # Instantiate the dict of params
        self.window_params = self._default_params()
        self._temp_value = CAMERA_MIN_TEMP_C

        # Window master layout
        self._layout = QtWidgets.QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)

        # (Vertically & sequentially) Place all section frames within the master layout
        self._layout.addWidget(self._create_fits_settings_section())
        self._layout.addWidget(self._create_roi_settings_section())
        self._layout.addWidget(self._create_file_settings_section())
        self._layout.addWidget(self._create_exposure_settings_section())
        self._layout.addWidget(self._create_delay_setting_section())
        self._layout.addWidget(self._create_experiment_repeat_setting_section())
        self._layout.addStretch()

    @staticmethod
    def _default_params():
        """
        The default experiment parameter dict.
        """
        return {
            "gain": DEFAULT_GAIN,
            "rospeed": DEFAULT_ROSPEED,
            "start_x": 0,
            "start_y": 0,
            "width": CAMERA_SENSOR_WIDTH,
            "height": CAMERA_SENSOR_HEIGHT,
            "bin_x": 1,
            "bin_y": 1,
            "exp_array": [1],
            "time_unit": "s",
            "nameprefix": "image",
            "savedirectory": str(Path.home() / "Desktop"),
            "name_date_time": "",
            "name_gain_rospeed": "",
            "name_temp": "",
            "name_exp_time": "",
            "include_date_time": False,
            "include_gain_rospeed": False,
            "include_temp": False,
            "include_exp_time": False,
            "delay": 1.000,
            "repeats": 1,
        }

    def refresh_gui_from_params(self):
        """
        Synchronise every GUI widget to reflect the current window_params.

        Called after replacing window_params (e.g. when loading a saved 
        configuration).  Signals are blocked during the update so that 
        signal handlers do not overwrite the freshly loaded values with 
        regenerated ones (e.g. a new timestamp).
        """
        # Collect all widgets whose signals must be suppressed
        text_widgets = (
            self._gain_combobox,
            self._rospeed_combobox,
            self._image_name_line_edit,
            self._time_unit_combobox,
            self._delay_line_edit,
            self._experiment_repeat_line_edit,
            self._exposure_times_line_edit,
        )
        check_widgets = (
            self._file_date_checkbox,
            self._file_gain_rospeed_checkbox,
            self._file_temp_checkbox,
            self._file_exp_time_checkbox,
        )

        for w in text_widgets:
            w.blockSignals(True)
        for w in check_widgets:
            w.blockSignals(True)

        try:
            p = self.window_params

            self._gain_combobox.setCurrentText(p["gain"])
            self._rospeed_combobox.setCurrentText(p["rospeed"])
            self._image_name_line_edit.setText(p["nameprefix"])
            self._time_unit_combobox.setCurrentText(p["time_unit"])
            self._delay_line_edit.setText(str(p["delay"]))
            self._experiment_repeat_line_edit.setText(str(p["repeats"]))
            self._exposure_times_line_edit.setText(str(p["exp_array"]))

            self._directory_label.setText("Folder: " + repr(p["savedirectory"])[1:-1])

            self._file_date_checkbox.setChecked(p["include_date_time"])
            self._file_gain_rospeed_checkbox.setChecked(p["include_gain_rospeed"])
            self._file_temp_checkbox.setChecked(p["include_temp"])
            self._file_exp_time_checkbox.setChecked(p["include_exp_time"])

            bx, by = p["bin_x"], p["bin_y"]
            self._roi_label.setText(
                f"ROI: X={p['start_x'] // bx}"
                f"->{(p['start_x'] + p['width']) // bx}, "
                f"Y={p['start_y'] // by}"
                f"->{(p['start_y'] + p['height']) // by}, "
                f"{p['width'] // bx}x{p['height'] // by} "
                f"(Bin {bx}x{by})"
            )

            self._update_exposure_labels()

            self._update_example_filename()

        finally:
            for w in text_widgets:
                w.blockSignals(False)
            for w in check_widgets:
                w.blockSignals(False)

    # ===========================================================================
    # GUI Frame Sections
    # ===========================================================================

    def _create_fits_settings_section(self):
        """
        Frame containing Gain & R/O speed settings for .fits header book-keeping.

        Returns
        -------
            QtWidgets.QFrame
        """

        self._gain_combobox = QtWidgets.QComboBox()
        self._gain_combobox.addItems(CAMERA_GAIN_OPTIONS)

        self._gain_widget = self._create_row(
            "Gain Setting (for FITS header):", self._gain_combobox, 10
        )
        self._gain_widget.setToolTip(
            f"Set gain setting for saving to FITS header (Default {DEFAULT_GAIN})."
        )
        self._gain_combobox.currentTextChanged.connect(self._gain_changed)

        # Read-out speed
        self._rospeed_combobox = QtWidgets.QComboBox()
        self._rospeed_combobox.addItems(CAMERA_ROSPEED_OPTIONS)

        self._rospeed_widget = self._create_row(
            "R/O Speed setting (for FITS header):", self._rospeed_combobox, 10
        )
        self._rospeed_widget.setToolTip(
            f"Set current read-out speed for saving to FITS header (Default {DEFAULT_ROSPEED})."
        )
        self._rospeed_combobox.currentTextChanged.connect(self._rospeed_changed)

        return self._create_frame(
            [self._gain_widget, self._create_separator(), self._rospeed_widget]
        )

    def _create_roi_settings_section(self):
        """
        Frame containing ROI settings.

        Returns
        -------
            QtWidgets.QFrame
        """

        # Region-of-interest (ROI) settings:
        # For defining custom imaging sub-regions on the sensor
        self._roi_button = QtWidgets.QPushButton("Set ROI")
        self._roi_button.clicked.connect(self._open_roiwindow)

        self._roi_widget = self._create_row("ROI settings:", self._roi_button, 10)
        self._roi_widget.setToolTip(
            "Open the ROI settings window. Allows changing:\n"
            "- binning (not implemented)\n"
            "- imaging sub-region (ROI)"
        )

        # Current ROI label
        self._roi_label = QtWidgets.QLabel(
            f"ROI: X=0->{CAMERA_SENSOR_WIDTH}, Y=0->{CAMERA_SENSOR_HEIGHT}, {CAMERA_SENSOR_WIDTH}x{CAMERA_SENSOR_HEIGHT} (Bin 1x1)"
        )

        self._roi_label_row = self._create_row(widget=self._roi_label)

        return self._create_frame(
            [self._roi_widget, self._create_separator(), self._roi_label_row]
        )

    def _create_file_settings_section(self):
        """
        Frame containing file naming and save location settings.

        Returns
        -------
            QtWidgets.QFrame
        """

        # Image name prefix:
        # For defining the image file name prefix
        self._image_name_line_edit = QtWidgets.QLineEdit("image")
        self._image_name_line_edit.textChanged.connect(self._file_prefix_changed)

        self._image_name_widget = self._create_row(
            "Image name prefix:", self._image_name_line_edit, 10
        )
        self._image_name_widget.setToolTip("Set the FITS image file name prefix")

        # File save directory dialog:
        # For setting the file directory to save to
        self._file_dialog_button = QtWidgets.QPushButton("File Directory")
        self._file_dialog_button.clicked.connect(self._open_file_dialog)

        self._file_directory_widget = self._create_row(
            "Save to:", self._file_dialog_button, 10
        )
        self._file_directory_widget.setToolTip(
            "Open file explorer and set directory to save FITS images to."
        )

        # Various file name settings
        self._file_date_checkbox = QtWidgets.QCheckBox("Append Date && Time")
        self._file_date_checkbox.stateChanged.connect(self._append_date_time)
        self._file_gain_rospeed_checkbox = QtWidgets.QCheckBox("Append Gain && R/O")
        self._file_gain_rospeed_checkbox.stateChanged.connect(self._append_gain_rospeed)
        self._file_temp_checkbox = QtWidgets.QCheckBox("Append Temp.")
        self._file_temp_checkbox.stateChanged.connect(self._append_temp)
        self._file_exp_time_checkbox = QtWidgets.QCheckBox("Append Exp. time")
        self._file_exp_time_checkbox.stateChanged.connect(self._append_exp_time)

        self._file_checkboxes_row_one_widget = self._create_row(
            widget=[self._file_date_checkbox, self._file_gain_rospeed_checkbox],
            spacing=10,
        )
        self._file_checkboxes_row_two_widget = self._create_row(
            widget=[self._file_temp_checkbox, self._file_exp_time_checkbox], spacing=10
        )

        # Current directory label
        self._directory_label = QtWidgets.QLabel(
            "Folder: " + repr(self.window_params["savedirectory"])[1:-1]
        )

        self._directory_row = self._create_row(widget=self._directory_label)

        # Example file name label
        self._file_name_label = QtWidgets.QLabel(
            "Example: " + self.window_params["nameprefix"] + "_1.fits"
        )

        self._file_name_row = self._create_row(widget=self._file_name_label)

        return self._create_frame(
            [
                self._image_name_widget,
                self._create_separator(),
                self._file_directory_widget,
                self._create_separator(),
                self._file_checkboxes_row_one_widget,
                self._file_checkboxes_row_two_widget,
                self._create_separator(),
                self._directory_row,
                self._file_name_row,
            ]
        )

    def _create_exposure_settings_section(self):
        """
        Frame containing exposure time unit and array settings.

        Returns
        -------
            QtWidgets.QFrame
        """

        # Exposure times:
        # For defining the array of exposures
        self._exposure_times_line_edit = QtWidgets.QLineEdit("[1]")
        self._exposure_times_widget = self._create_row(
            "Array of exposures:", self._exposure_times_line_edit, 10
        )
        self._exposure_times_widget.setToolTip(
            "Define the exposure array: \n"
            "Evaluates the input as python code i.e.:\n"
            "- [1] is a single 1 time unit exposure.\n"
            "- np.repeat([1],20) is 20 1 time unit exposures.\n"
            "- np.repeat(np.arange(1,4),10) is 10 sets of [1,2,3] time unit exposures."
        )

        # Time unit:
        # For defining the time unit for exposures (i.e. s/ms/μs)
        # Also for setting the currently defined exposure array
        self._time_unit_combobox = QtWidgets.QComboBox()
        self._time_unit_combobox.addItems(list(TIME_UNIT_MULTIPLIERS.keys()))
        self._set_exposure_button = QtWidgets.QPushButton("Set Exposure")
        self._set_exposure_button.clicked.connect(self._set_exposure)
        self._time_unit_widget = self._create_row(
            "Time unit:", [self._time_unit_combobox, self._set_exposure_button], 10
        )
        self._time_unit_widget.setToolTip("Set the time unit for exposures (s/ms/μs).")
        self._time_unit_combobox.currentTextChanged.connect(self._time_unit_changed)

        # Current exposure array label
        self._num_exposures_label = QtWidgets.QLabel("1 exposure:")

        self._num_exposures_row = self._create_row(
            widget=self._num_exposures_label, spacing=10
        )

        self._exposure_array_label = QtWidgets.QLabel("[1]s")

        self._exposure_array_row = self._create_row(
            widget=self._exposure_array_label, spacing=10
        )

        return self._create_frame(
            [
                self._exposure_times_widget,
                self._time_unit_widget,
                self._create_separator(),
                self._num_exposures_row,
                self._exposure_array_row,
            ]
        )

    def _create_delay_setting_section(self):
        """
        Frame containing setting for delay b/w experiments.

        Returns
        -------
            QtWidgets.QFrame
        """

        # Experiment delay:
        # For setting the delay before an experiment begins
        # Useful for spacing subsequent experiments
        self._delay_line_edit = QtWidgets.QLineEdit("1.000")
        self._delay_line_edit.setValidator(
            QtGui.QDoubleValidator(bottom=0.0, decimals=3)
        )
        self._delay_line_edit.textChanged.connect(self._delay_changed)

        self._delay_widget = self._create_row(
            "Delay before start (s):", self._delay_line_edit, 10
        )
        self._delay_widget.setToolTip(
            "Delay in seconds before beginning this experiment block."
        )

        return self._create_frame([self._delay_widget])

    def _create_experiment_repeat_setting_section(self):
        """
        Frame containing experiment repeat setting.

        Returns
        -------
            QtWidgets.QFrame
        """

        # Experiment repeats:
        # For defining the number of times to repeat this experiment set
        self._experiment_repeat_line_edit = QtWidgets.QLineEdit("1")
        self._experiment_repeat_line_edit.setValidator(QtGui.QIntValidator(bottom=1))
        self._experiment_repeat_line_edit.textChanged.connect(
            self._experiment_repeat_changed
        )

        self._experiment_repeat_widget = self._create_row(
            "Experiment repeats:", self._experiment_repeat_line_edit, 10
        )
        self._experiment_repeat_widget.setToolTip(
            "Specify how many times to repeat the experiment block."
        )

        return self._create_frame([self._experiment_repeat_widget])

    # ===========================================================================
    # Widget Logic Functions
    # ===========================================================================

    def _gain_changed(self, text):
        self.window_params["gain"] = text
        self._append_gain_rospeed()

    def _rospeed_changed(self, text):
        self.window_params["rospeed"] = text
        self._append_gain_rospeed()

    def _set_exposure(self):
        """
        Parse and apply the exposure array without using eval().
        """
        try:
            exposures = self._parse_exposure_array(
                self._exposure_times_line_edit.text()
            )
        except ExposureParserError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid exposure array",
                str(exc),
            )
            return

        self.window_params["exp_array"] = exposures
        self._update_exposure_labels()
        self._append_exp_time()

    def _time_unit_changed(self, text):
        self.window_params["time_unit"] = text

    def _delay_changed(self, text):
        try:
            self.window_params["delay"] = float(text)
        except ValueError:
            pass

    def _experiment_repeat_changed(self, text):
        try:
            self.window_params["repeats"] = int(text)
        except ValueError:
            pass

    def _file_prefix_changed(self, text):
        self.window_params["nameprefix"] = text
        self._update_example_filename()

    def _open_file_dialog(self):
        path = QtWidgets.QFileDialog.getExistingDirectory()
        if path:
            self.window_params["savedirectory"] = path
            self._directory_label.setText(f"Folder: {path}")

    def _append_date_time(self):
        state = self._file_date_checkbox.checkState()
        if state == 2:
            ts = datetime.now()
            if int(ts.strftime("%H")) >= 12:
                self.window_params["name_date_time"] = ts.strftime("_%d_%m_%Y_%H-%Mpm")
            else:
                self.window_params["name_date_time"] = ts.strftime("_%d_%m_%Y_%H-%Mam")
            self.window_params["include_date_time"] = True
        else:
            self.window_params["name_date_time"] = ""
            self.window_params["include_date_time"] = False
        self._update_example_filename()

    def _append_gain_rospeed(self):
        state = self._file_gain_rospeed_checkbox.checkState()
        if state == 2:
            self.window_params["name_gain_rospeed"] = (
                "_" + self.window_params["gain"] + "_" + self.window_params["rospeed"]
            )
            self.window_params["include_gain_rospeed"] = True
        else:
            self.window_params["name_gain_rospeed"] = ""
            self.window_params["include_gain_rospeed"] = False
        self._update_example_filename()

    def _append_temp(self):
        state = self._file_temp_checkbox.checkState()
        if state == 2:
            target = getattr(self, "_temp_value", None)
            if target is not None:
                self.window_params["name_temp"] = f"_{target:.0f}C"
                self.window_params["include_temp"] = True
        else:
            self.window_params["name_temp"] = ""
            self.window_params["include_temp"] = False
        self._update_example_filename()

    def set_temp_value(self, temp_c):
        self._temp_value = float(temp_c)
        state = self._file_temp_checkbox.checkState()
        # Refresh the filename if the checkbox is already ticked
        if state == 2:
            self._append_temp()

    def _append_exp_time(self):
        state = self._file_exp_time_checkbox.checkState()
        if state == 2:
            self.window_params["name_exp_time"] = (
                "_"
                + str(self.window_params["exp_array"][0])
                + self.window_params["time_unit"]
            )
            self.window_params["include_exp_time"] = True
        else:
            self.window_params["name_exp_time"] = ""
            self.window_params["include_exp_time"] = False
        self._update_example_filename()

    def _update_example_filename(self):
        p = self.window_params

        self._example_file_name_value = (
            "Example: "
            + p["nameprefix"]
            + (p["name_date_time"] if p["include_date_time"] else "")
            + (p["name_gain_rospeed"] if p["include_gain_rospeed"] else "")
            + (p["name_temp"] if p["include_temp"] else "")
            + (p["name_exp_time"] if p["include_exp_time"] else "")
            + "_1.fits"
        )

        self._file_name_label.setText(self._example_file_name_value)

    def _open_roiwindow(self):
        dialog = ROIWindow(self)
        dialog.exec_()
        if dialog.data:
            d = dialog.data
            self.window_params["start_x"] = d["start_x"]
            self.window_params["start_y"] = d["start_y"]
            self.window_params["width"] = d["width"]
            self.window_params["height"] = d["height"]
            self.window_params["bin_x"] = d["bin_x"]
            self.window_params["bin_y"] = d["bin_y"]
            self._roi_label.setText(
                f"ROI: X={int(d['start_x'] / d['bin_x'])}->{int((d['start_x'] + d['width']) / d['bin_x'])},"
                f" Y={int(d['start_y'] / d['bin_y'])}->{int((d['start_y'] + d['height']) / d['bin_y'])},"
                f"  {int(d['width'] / d['bin_x'])}x{int(d['height'] / d['bin_y'])} (Bin {d['bin_x']}x{d['bin_y']})"
            )

    # ===========================================================================
    # Helper Functions
    # ===========================================================================

    @staticmethod
    def _parse_exposure_array(text):
        """
        Validator for parsing exposure expressions into flat lists of positive floats.

        Supports list/tuple, constants and whitelist of NumPy functions.

        NOTE - Will possibly replace with a simple float input + no. of times to repeat
               to avoid using eval, even if validated against whitelist.

        Parameters
        -------
        text : str
            expression, e.g. "[1]", "np.repeat([1], 20)".

        Returns
        -------
            Flat list of positive floats (i.e. exposure values).

        Raises
        ------
        ExposureParserError
            If the expression isn't whitelisted or unsafe.
        """
        allowed_functions = {
            np.repeat,
            np.arange,
            np.linspace,
            np.ones,
            np.full,
            np.array,
        }

        try:
            tree = ast.parse(text.strip(), mode="eval")
        except SyntaxError as exc:
            raise ExposureParserError(f"Invalid syntax: {exc.msg}") from exc

        def fail(node):
            raise ExposureParserError(
                f"Unsupported expression element: {type(node).__name__}"
            )

        def eval_node(node):
            """
            Recursively evaluate a whitelisted AST node.
            """
            if isinstance(node, ast.Expression):
                return eval_node(node.body)

            if isinstance(node, ast.Constant):
                if isinstance(node.value, (int, float)) and not isinstance(
                    node.value, bool
                ):
                    return node.value
                fail(node)

            if isinstance(node, (ast.List, ast.Tuple)):
                return [eval_node(element) for element in node.elts]

            if isinstance(node, ast.UnaryOp) and isinstance(
                node.op, (ast.UAdd, ast.USub)
            ):
                value = eval_node(node.operand)
                return value if isinstance(node.op, ast.UAdd) else -value

            if isinstance(node, ast.Attribute):
                if (
                    isinstance(node.value, ast.Name)
                    and node.value.id == "np"
                    and getattr(np, node.attr, None) in allowed_functions
                ):
                    return getattr(np, node.attr)
                fail(node)

            if isinstance(node, ast.Call):
                if node.keywords:
                    fail(node)

                func = eval_node(node.func)
                if func not in allowed_functions:
                    fail(node)

                args = [eval_node(arg) for arg in node.args]
                return func(*args)

            fail(node)

        allowed_node_types = (
            ast.Expression,
            ast.Constant,
            ast.List,
            ast.Tuple,
            ast.UnaryOp,
            ast.UAdd,
            ast.USub,
            ast.Call,
            ast.Attribute,
            ast.Name,
            ast.Load,
        )

        for node in ast.walk(tree):
            if not isinstance(node, allowed_node_types):
                raise ExposureParserError(f"Unsupported token: {type(node).__name__}")

            if isinstance(node, ast.Name) and node.id != "np":
                raise ExposureParserError("Only np.* functions are allowed.")

        result = eval_node(tree)
        
        # Flatten nested lists/tuples
        def _flatten(item):
            if isinstance(item, (list, tuple)):
                for sub in item:
                    yield from _flatten(sub)
            else:
                yield item
                
        flat_result = list(_flatten(result)) if isinstance(result, (list, tuple)) else result
        arr = np.asarray(flat_result, dtype=float).ravel()

        if arr.size == 0:
            raise ExposureParserError("Exposure array is empty.")

        if arr.size > MAX_EXPOSURE_ARRAY_LEN:
            raise ExposureParserError(
                f"Exposure array larger than max length: {MAX_EXPOSURE_ARRAY_LEN}"
            )

        if not np.all(np.isfinite(arr)):
            raise ExposureParserError("Exposure values must be finite.")

        if np.any(arr <= 0):
            raise ExposureParserError("Exposure values must be strictly positive.")

        return [float(value) for value in arr]

    def _update_exposure_labels(self):
        """
        Refresh the exposure-count and exposure-array summary labels.
        """
        arr = self.window_params["exp_array"]
        unit = self.window_params["time_unit"]
        n = len(arr)

        self._num_exposures_label.setText(f"{n} exposure{'s' if n != 1 else ''}:")

        if n > 35:
            head = ", ".join(str(v) for v in arr[:16])
            tail = ", ".join(str(v) for v in arr[-16:])
            self._exposure_array_label.setText(f"[{head}, ...., {tail}]{unit}")
        else:
            self._exposure_array_label.setText(f"{arr}{unit}")

    @staticmethod
    def _create_row(label_text=None, widget=None, spacing=0):
        """
        Add widgets side-by-side to a QtWidgets QHBoxLayout.

        Parameters
        ----------
        label_text : str
            A row label widget preceeding the
            functional widgets.

        widget : QtWidgets.QWidget
            The list of widgets to add.

        spacing : float
            Defines the vertical spacing of the
            row.

        Returns
        -------
            row : QtWidgets.QWidget
                The full QHBoxLayout as a
                placable QWidget.
        """
        row = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(row)
        if label_text:
            layout.addWidget(QtWidgets.QLabel(label_text))
        if widget:
            widgets = widget if isinstance(widget, (list, tuple)) else (widget,)
            for w in widgets:
                layout.addWidget(w)
        layout.addStretch()
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(spacing)
        return row

    @staticmethod
    def _create_frame(widgets, spacing=0):
        """
        Add widgets vertically & sequentially to a QtWidgets.QFrame
        with a QVBoxLayout.

        Parameters
        ----------
        widgets : list
            The list of widgets to add.

        spacing : float
            Defines the vertical spacing of the frame.

        Returns
        -------
        frame : QtWidgets.QFrame
            The full QVBoxLayout as a placable
            QFrame.
        """
        frame = QtWidgets.QFrame()
        frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        frame.setLineWidth(1)
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setSpacing(spacing)
        layout.setContentsMargins(2, 2, 2, 2)
        for widget in widgets:
            layout.addWidget(widget)
        return frame

    @staticmethod
    def _create_separator():
        """
        Create a stylised visual separating line from a placable QFrame.

        Returns
        -------
        sep : QtWidgets.QFrame
            The stylised separator as a placable
            QFrame.
        """
        sep = QtWidgets.QFrame()
        sep.setFrameShape(QtWidgets.QFrame.HLine)
        sep.setFrameShadow(QtWidgets.QFrame.Sunken)
        return sep


class CollapsibleWidget(QtWidgets.QWidget):
    """
    Create the container for a custom collapsible window widget.

    An instance of CollapsibleWidgetInternalGUI is added to the container layout
    as a QWidget.
    """

    def __init__(self, title="", parent=None):
        super().__init__(parent)

        self._layout = QtWidgets.QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)

        self._header = QtWidgets.QWidget()
        self._header.setObjectName("collapsibleHeader")
        self._header.setStyleSheet(
            "#collapsibleHeader {background-color: lightgrey; border: 1px solid grey; "
            "border-radius: 1px;}"
        )
        header_layout = QtWidgets.QHBoxLayout(self._header)
        header_layout.setContentsMargins(5, 5, 5, 5)

        self._icon = QtWidgets.QLabel("▼")
        self._icon.setFixedWidth(16)
        header_layout.addWidget(self._icon)

        # (Double click) Editable title
        self._title_label = QtWidgets.QLabel(title)
        self._title_label.setObjectName("titleLabel")
        self._title_label.setStyleSheet("#titleLabel {background-color: transparent;}")
        self._title_label.mouseDoubleClickEvent = lambda e: self._start_edit()

        self._title_edit = QtWidgets.QLineEdit(title)
        self._title_edit.hide()
        self._title_edit.returnPressed.connect(self._finish_edit)
        self._title_edit.editingFinished.connect(self._finish_edit)

        header_layout.addWidget(self._title_label)
        header_layout.addWidget(self._title_edit)
        header_layout.addStretch()

        self._header.mousePressEvent = lambda e: self.toggle()

        self._content = QtWidgets.QWidget()
        self._content.setObjectName("collapsibleContent")
        self._content.setStyleSheet(
            "#collapsibleContent {background-color: lightgrey; border: 1px solid grey; "
            "margin: 0 0px 0px 0px;}"
        )

        self._layout.addWidget(self._header)
        self._layout.addWidget(self._content)

    def _start_edit(self):
        self._title_label.hide()
        self._title_edit.setText(self._title_label.text())
        self._title_edit.show()
        self._title_edit.setFocus()
        self._title_edit.selectAll()

    def _finish_edit(self):
        self._title_edit.hide()
        self._title_label.setText(self._title_edit.text())
        self._title_label.show()

    def toggle(self):
        visible = not self._content.isVisible()
        self._content.setVisible(visible)
        self._icon.setText("▼" if visible else "▶")

    @property
    def contentWidget(self):
        return self._content

    @property
    def title(self):
        """
        Current title text.
        """
        return self._title_label.text()

    @title.setter
    def title(self, text):
        """
        Set the displayed title text.

        Parameters
        ----------
        text : str
            New title string.
        """
        self._title_label.setText(text)

    @property
    def internalGui(self):
        """
        Access the embedded CollapsibleWidgetInternalGUI instance.

        Returns
        -------
            CollapsibleWidgetInternalGUI
        """
        layout = self._content.layout()
        return layout.itemAt(0).widget() if layout and layout.count() else None


class ROIWidget(QtWidgets.QGraphicsView):
    """
    Interactive graphics display for the ROI selection.
    
    Added to the ROI settings window in the same manner as a QWidget.
    """

    roi_changed = QtCore.pyqtSignal(int, int, int, int)

    def __init__(self, width=CAMERA_SENSOR_WIDTH, height=CAMERA_SENSOR_HEIGHT):
        super().__init__()
        self._w, self._h = width, height
        left, top, right, bottom = 30, 5, 35, 25

        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        # Initialise the graphics scene
        self._scene = QtWidgets.QGraphicsScene(
            -left, -top, width + left + right, height + top + bottom
        )
        self.setScene(self._scene)

        buffer = 4
        self.setFixedSize(width + left + right + buffer, height + top + bottom + buffer)
        self.setRenderHint(QtGui.QPainter.Antialiasing)

        self._rect = None
        self._handles = []
        self._state = None
        self._offset = QtCore.QPointF()

        self._draw_grid()
        self._create_roi(0, 0, width, height)

    def _draw_grid(self):
        """
        Create grid-lines for the plot area of the ROI window
        using QtGui elements.
        """
        pen_minor = QtGui.QPen(QtCore.Qt.gray, 0.5)
        pen_minor.setStyle(QtCore.Qt.DashLine)
        pen_major = QtGui.QPen(QtCore.Qt.black, 1)
        pen_major.setStyle(QtCore.Qt.DashLine)
        pen_solid = QtGui.QPen(QtCore.Qt.black, 1)

        # Dashed internal grid lines
        for x in range(0, self._w + 1, 16):
            self._scene.addLine(
                x, 0, x, self._h, pen_major if x % 32 == 0 else pen_minor
            )
        for y in range(0, self._h + 1, 16):
            self._scene.addLine(
                0, y, self._w, y, pen_major if y % 32 == 0 else pen_minor
            )

        # Solid outer boundary
        self._scene.addRect(0, 0, self._w, self._h, pen_solid)

        # Label ticks
        font = QtGui.QFont("Arial", 8)
        for i in range(0, self._w + 1, 32):
            self._scene.addLine(i, self._h, i, self._h + 3, pen_solid)
            t = self._scene.addText(str(i), font)
            t.setPos(i - 10, self._h + 5)
        for i in range(0, self._h + 1, 32):
            self._scene.addLine(0, i, -3, i, pen_solid)
            t = self._scene.addText(str(i), font)
            t.setPos(-25, i - 5)

    def _create_roi(self, x, y, w, h):
        """
        Create a new ROI setting (removes and recreates the ROI region indicator on the graphics
        scene).

        Parameters
        ----------
        x : int
            Top left x coord of the selection region.
        y : int
            Top left y coord of the selection region.
        w : int
            Width of the selection region.
        h : int
            Height of the selection region.
        """
        # Clear the graphics scene of the previous region indicator
        if self._rect:
            self._scene.removeItem(self._rect)
            for handle in self._handles:
                self._scene.removeItem(handle)
        # Add a new region indicator corresponding to the new ROI.
        self._rect = self._scene.addRect(
            x,
            y,
            w,
            h,
            QtGui.QPen(QtCore.Qt.green, 2),
            QtGui.QBrush(QtGui.QColor(128, 128, 128, 51)),
        )
        self._rect.setZValue(10)

        # Create the resize handles on the indicator.
        self._handles = []
        pos = [
            (x, y),
            (x + w / 2, y),
            (x + w, y),
            (x, y + h / 2),
            (x + w, y + h / 2),
            (x, y + h),
            (x + w / 2, y + h),
            (x + w, y + h),
        ]
        for px, py in pos:
            self._handles.append(
                self._scene.addRect(
                    px - 3,
                    py - 3,
                    6,
                    6,
                    QtGui.QPen(QtCore.Qt.black),
                    QtGui.QBrush(QtCore.Qt.white),
                )
            )
            self._handles[-1].setZValue(11)

    def mousePressEvent(self, e):
        """
        Define the mouse click event for the indicator region & handles.
        """
        p = self.mapToScene(e.pos())
        self._state = next(
            (i for i, h in enumerate(self._handles) if h.rect().contains(p)), None
        )
        # If the mouse position, on click, is within the indicator region, start the drag logic.
        if self._state is not None and self._rect is not None:
            self._offset = p - self._rect.rect().topLeft()
        elif self._rect.rect().contains(p):
            self._state = "drag"
            self._offset = p - self._rect.rect().topLeft()

    def mouseMoveEvent(self, e):
        """
        Define the mouse move event for the indicator region & handles.
        """
        if self._state is None:
            return
        p = self.mapToScene(e.pos())
        r = self._rect.rect()
        # On mouse move, drag the indicator region (if state is drag), or resize the region.
        if self._state == "drag":
            self.set_roi(
                int(max(0, min(p.x() - self._offset.x(), self._w - r.width()))),
                int(max(0, min(p.y() - self._offset.y(), self._h - r.height()))),
                int(r.width()),
                int(r.height()),
            )
        else:
            self._resize(p)

    def mouseReleaseEvent(self):
        self._state = None

    def _resize(self, p):
        """
        Indicator region resizing logic.
        """
        r = self._rect.rect()
        x, y, w, h = r.x(), r.y(), r.width(), r.height()
        a = self._state
        # NIRvana HS hardware specific ROI coordinate restrictions applied here.
        if a in [0, 3, 5]:
            nw = max(CAMERA_MIN_ROI_SIZE, min(self._w - x, w + (x - p.x())))
            x, w = max(0, min(p.x(), self._w - nw)), nw
        if a in [2, 4, 7]:
            w = max(CAMERA_MIN_ROI_SIZE, min(self._w - x, p.x() - x))
        if a in [0, 1, 2]:
            nh = max(CAMERA_MIN_ROI_SIZE, min(self._h - y, h + (y - p.y())))
            y, h = max(0, min(p.y(), self._h - nh)), nh
        if a in [5, 6, 7]:
            h = max(CAMERA_MIN_ROI_SIZE, min(self._h - y, p.y() - y))
        self.set_roi(int(x), int(y), int(w), int(h))

    def get_roi(self):
        r = self._rect.rect()
        return int(r.x()), int(r.y()), int(r.width()), int(r.height())

    def set_roi(self, x, y, w, h):
        self._create_roi(x, y, w, h)
        self.roi_changed.emit(x, y, w, h)


class ROIWindow(QtWidgets.QDialog):
    """
    Create the full ROI settings window. 
    
    Add the ROI graphics view in the same manner 
    as a QWidget to the window.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ROI Settings")
        self.setFixedSize(1030, 580)
        self._main = QtWidgets.QHBoxLayout(self)
        self._main.setContentsMargins(5, 5, 5, 5)
        self._main.setSpacing(10)

        # Split the main layout into a left layout and
        # a right layout (containing only the ROIWidget).
        self._left = QtWidgets.QWidget()
        self._left.setFixedWidth(280)
        self._left_layout = QtWidgets.QVBoxLayout(self._left)
        self._left_layout.setContentsMargins(0, 0, 0, 0)
        self._left_layout.setSpacing(5)

        # Add all settings controls.
        self._bins = [QtWidgets.QComboBox() for _ in range(2)]
        [c.addItems(["1x", "2x", "4x", "8x"]) for c in self._bins]

        self._edits = [QtWidgets.QLineEdit() for _ in range(4)]
        vals = [
            (0, CAMERA_SENSOR_WIDTH),
            (0, CAMERA_SENSOR_HEIGHT),
            (0, CAMERA_SENSOR_WIDTH),
            (0, CAMERA_SENSOR_HEIGHT),
        ]
        [self._edits[i].setValidator(QtGui.QIntValidator(*vals[i])) for i in range(4)]

        self._full = QtWidgets.QPushButton("Full Sensor")
        self._centre = QtWidgets.QPushButton("Centre ROI")
        self._ROI_label = QtWidgets.QLabel(
            f"{CAMERA_SENSOR_WIDTH}x{CAMERA_SENSOR_HEIGHT}"
        )
        self._apply = QtWidgets.QPushButton("Apply Current ROI")

        self._last_data = None

        self._setup_ui()
        self._connect()
        self._view.set_roi(0, 0, CAMERA_SENSOR_WIDTH, CAMERA_SENSOR_HEIGHT)

    def _setup_ui(self):
        """
        Add elements to the main ROI window layout.
        """
        bin_row = self._row(
            [
                QtWidgets.QLabel("X:"),
                self._bins[0],
                QtWidgets.QLabel("Y:"),
                self._bins[1],
            ],
            10,
        )
        self._left_layout.addWidget(self._frame([bin_row], "Binning"))

        labels = ["Start X:", "Start Y:", "Width:", "Height:"]

        # Use _form_row for aligned labels.
        rows = [self._form_row(labels[i], self._edits[i]) for i in range(4)]
        rows.extend(
            [self._row([self._full, self._centre, self._ROI_label], 10), self._apply]
        )

        self._left_layout.addWidget(self._frame(rows, "Region-of-interest (ROI)"))
        self._left_layout.addStretch()

        self._view = ROIWidget()
        self._main.addWidget(self._left)
        self._main.addWidget(self._view)

    def _connect(self):
        """
        Connect logic functions to settings widgets.
        """
        [e.textChanged.connect(self._sync_edits) for e in self._edits]
        self._view.roi_changed.connect(self._sync_view)
        self._full.clicked.connect(
            lambda: self._view.set_roi(0, 0, CAMERA_SENSOR_WIDTH, CAMERA_SENSOR_HEIGHT)
        )
        self._centre.clicked.connect(self._on_centre)
        self._apply.clicked.connect(self._on_apply)

    def _sync_edits(self):
        """Sync all ROI related settings widgets."""
        try:
            x, y, w, h = [int(e.text() or 0) for e in self._edits]
            w, h = min(w, CAMERA_SENSOR_WIDTH - x), min(h, CAMERA_SENSOR_HEIGHT - y)
            self._view.set_roi(x, y, w, h)
        except Exception:
            pass

    def _sync_view(self, x, y, w, h):
        """
        Sync the ROI related labels.
        """
        [e.setText(str(v)) for e, v in zip(self._edits, [x, y, w, h])]
        self._ROI_label.setText(f"{w}x{h} unbinned")

    def _on_centre(self):
        x, y, w, h = self._view.get_roi()
        self._view.set_roi(
            (CAMERA_SENSOR_WIDTH - w) // 2, (CAMERA_SENSOR_HEIGHT - h) // 2, w, h
        )

    def _on_apply(self):
        """
        Validate, correct, and store the current ROI settings.
        """
        max_x = CAMERA_SENSOR_WIDTH
        max_y = CAMERA_SENSOR_HEIGHT

        try:
            start_x = int(self._edits[0].text() or 0)
            start_y = int(self._edits[1].text() or 0)
            width = int(self._edits[2].text() or 0)
            height = int(self._edits[3].text() or 0)
        except ValueError:
            QtWidgets.QMessageBox.warning(
                self,
                "ROI Error",
                "ROI values must be integers.",
            )
            return

        bin_x = int(self._bins[0].currentText()[0])
        bin_y = int(self._bins[1].currentText()[0])

        start_x = max(0, min(start_x, max_x - 1))
        start_y = max(0, min(start_y, max_y - 1))
        width = max(0, min(width, max_x - start_x))
        height = max(0, min(height, max_y - start_y))

        # Make origin and size divisible by binning.
        start_x -= start_x % bin_x
        start_y -= start_y % bin_y
        width -= width % bin_x
        height -= height % bin_y

        # NIRvana HS specific constraint: 
        # for 1x/2x binning, end X should be divisible by 4 (CAMERA_X_ALIGNMENT_DIVISOR).
        if bin_x in (1, 2):
            end_x = start_x + width
            end_x -= end_x % CAMERA_X_ALIGNMENT_DIVISOR
            width = max(0, end_x - start_x)
            width -= width % bin_x

        # Final bounds correction.
        if start_x + width > max_x:
            width = max_x - start_x
            width -= width % bin_x

        if start_y + height > max_y:
            height = max_y - start_y
            height -= height % bin_y

        if (
            width < max(CAMERA_MIN_ROI_SIZE, bin_x)
            or height < max(CAMERA_MIN_ROI_SIZE, bin_y)
            or width % bin_x != 0
            or height % bin_y != 0
        ):
            QtWidgets.QMessageBox.warning(
                self,
                "ROI Error",
                "ROI is too small or not compatible with the selected binning.",
            )
            return

        self._view.set_roi(start_x, start_y, width, height)

        self._last_data = {
            "start_x": start_x,
            "start_y": start_y,
            "width": width,
            "height": height,
            "bin_x": bin_x,
            "bin_y": bin_y,
        }

    @staticmethod
    def _row(widgets, spacing=0):
        """
        Add widgets side-by-side to a QtWidgets QHBoxLayout.

        Parameters
        ----------
        widgets : QtWidgets.QWidget
            The list of widgets to add.

        spacing : float
            Defines the vertical spacing of the row.

        Returns
        -------
        row : QtWidgets.QWidget
            The full QHBoxLayout as a placable QWidget.
        """
        row = QtWidgets.QFrame()
        row.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        row.setLineWidth(1)
        row.setStyleSheet("QFrame {background-color: lightgrey}")
        layout = QtWidgets.QHBoxLayout(row)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(spacing)
        [layout.addWidget(w) for w in widgets]
        layout.addStretch()
        return row

    @staticmethod
    def _form_row(label_text, widget):
        """Add widgets side-by-side to a QtWidgets QHBoxLayout.
            This sspecific method is used to ensure fixed label
            sizing.

        Args:
            label_text (str): A row label widget preceeding the
                              functional widgets.

            widget (QtWidgets.QWidget): The list of widgets to
                                        add.

        Returns
        -------
            row (QtWidgets.QWidget): The full QHBoxLayout as a
                                     placable QWidget.
        """
        row = QtWidgets.QFrame()
        row.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        row.setLineWidth(1)
        row.setStyleSheet("QFrame {background-color: lightgrey}")
        layout = QtWidgets.QHBoxLayout(row)
        layout.setContentsMargins(2, 2, 2, 2)
        label = QtWidgets.QLabel(label_text)
        label.setFixedWidth(60)  # Consistent label width
        layout.addWidget(label)
        layout.addWidget(widget)
        layout.addStretch()
        return row

    @staticmethod
    def _frame(widgets, title=""):
        """
        Add widgets vertically & sequentially to a QtWidgets.QFrame
        with a QVBoxLayout.

        Parameters
        ----------
        widgets : list
            The list of widgets to add.

        title : str
            The title of the frame to be displayed.

        Returns
        -------
        frame : QtWidgets.QFrame
            The full QVBoxLayout as a placable
            QFrame.
        """
        frame = QtWidgets.QFrame()
        frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        frame.setLineWidth(1)
        frame.setStyleSheet("QFrame {background-color: lightgrey}")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)
        if title:
            label = QtWidgets.QLabel(title)
            layout.addWidget(label)
        [layout.addWidget(w) for w in widgets]
        return frame

    @property
    def data(self):
        return self._last_data


class LogWidget(QtWidgets.QWidget):
    """
    A message logging window.

    Added in the same manner as a QWidget.
    """

    def __init__(self):
        super().__init__()
        self.setFixedSize(640, 256)
        self.te = QtWidgets.QTextEdit(
            readOnly=True, styleSheet="background-color: black;"
        )
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.te)
        layout.setContentsMargins(0, 0, 0, 0)

    def log(self, msg, colour):
        """
        Log a message with timestamp and color styling.

        Parameters
        ----------
        msg : str
            The message string to log.

        color : str
            The log message colour.
        """
        safe_msg = html.escape(str(msg))
        ts = datetime.now().strftime("[%d/%m/%y: %H:%M:%S]")
        self.te.append(
            f'<span style="color:{colour}; font-family:Arial; font-size:10pt; '
            f'font-weight:bold;">{ts} {safe_msg}</span>'
        )


class ImageViewerWidget(QtWidgets.QWidget):
    """
    Embeddable camera image viewer widget with histogram and contrast control.

    Added in the same manner as a QWidget.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(640, 722)
        self.zoom = 1.0
        self.item = self.data = self.rgb = None
        self.contrast_low = 0
        self.contrast_high = UINT16_MAX
        self.dragging_bar = None
        self.auto_contrast_enabled = True

        # Histogram setup
        self.fig = Figure(figsize=(4.8, 1.3), dpi=100, facecolor="black")
        self.ax = self.fig.add_subplot(111, facecolor="black")
        self.ax.set_xlim(0, UINT16_MAX)
        self.ax.set_ylabel("")
        self.ax.set_xlabel(
            "Pixel Value (DN)", fontsize=8, fontweight="medium", color="#e0e0e0"
        )
        self.ax.set_yticks([])
        self.ax.tick_params(axis="x", labelsize=8, colors="#e0e0e0")
        for spine in self.ax.spines.values():
            visible = spine.spine_type == "bottom"
            spine.set_visible(visible)
            if visible:
                spine.set_color("white")
        self.fig.subplots_adjust(left=0.0625, right=0.95, top=1, bottom=0.25)
        self.hist_canvas = FigureCanvas(self.fig)
        self.hist_canvas.setFixedSize(480, 150)
        self.fig.canvas.mpl_connect("button_press_event", self._on_bar_press)
        self.fig.canvas.mpl_connect("motion_notify_event", self._on_bar_motion)
        self.fig.canvas.mpl_connect("button_release_event", self._on_bar_release)
        self.bar_low = self.ax.axvline(0, color="purple", linewidth=3)
        self.bar_high = self.ax.axvline(UINT16_MAX, color="purple", linewidth=3)

        # Stats box setup
        self.stats_box = QtWidgets.QFrame()
        self.stats_box.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        self.stats_box.setLineWidth(1)
        self.stats_box.setStyleSheet("QFrame {background-color: lightgrey}")
        self.stats_box.setFixedSize(160, 150)
        layout = QtWidgets.QVBoxLayout(self.stats_box)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(3)
        self.auto_btn = QtWidgets.QPushButton("Auto Contrast")
        self.auto_btn.clicked.connect(self._auto_contrast)
        layout.addWidget(self.auto_btn)
        self.stats_labels = {}
        for name in ("Mean", "Median", "Max", "Min", "Std. Dev.", "No. Pix."):
            self.stats_labels[name] = QtWidgets.QLabel(f"{name}: ")
            layout.addWidget(self.stats_labels[name])

        # Image viewer setup
        self.scene = QtWidgets.QGraphicsScene()
        self.view = QtWidgets.QGraphicsView(self.scene)
        self.view.setFixedSize(640, 512)
        self.view.setDragMode(QtWidgets.QGraphicsView.ScrollHandDrag)
        self.view.wheelEvent = self._wheel_event
        self.view.mouseMoveEvent = self._mouse_move_event
        self.view.setMouseTracking(True)
        self.view.setFrameStyle(QtWidgets.QGraphicsView.NoFrame)
        self.view.setViewportMargins(0, 0, 0, 0)
        self.view.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop)

        # Info label and button
        self.info_row = QtWidgets.QFrame()
        self.info_row.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        self.info_row.setLineWidth(1)
        self.info_row.setStyleSheet("QFrame {background-color: lightgrey}")
        self.info_row_layout = QtWidgets.QVBoxLayout(self.info_row)
        self.info = QtWidgets.QLabel()
        self.info.setFixedSize(640, 20)

        # Button that generates a dummy image for testing the image display
        # self.btn = QtWidgets.QPushButton('Open Image')
        # self.btn.setFixedSize(640, 40)
        # self.btn.clicked.connect(self._open_image)

        self.info_row_layout.addWidget(self.info)
        # self.info_row_layout.addWidget(self.btn)

        # Main layout assembly
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        top_row = QtWidgets.QWidget()
        top_row.setFixedSize(640, 150)
        top_layout = QtWidgets.QHBoxLayout(top_row)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(0)
        top_layout.addWidget(self.hist_canvas)
        top_layout.addWidget(self.stats_box)
        main_layout.addWidget(top_row)
        main_layout.addWidget(self.view)
        main_layout.addWidget(self.info_row)

    def display_array(self, arr):
        """
        Call relevant functions to display array.
        """
        self.data = arr
        self._update_stats()
        # Always update histogram (xlims) for new data
        if self.dragging_bar is None:
            self._update_histogram()
        self._redraw_image()

    def _update_histogram(self):
        """
        Update xlims and redraw histogram plot.
        """
        if self.data is None:
            return
        counts = np.bincount(self.data.ravel(), minlength=UINT16_MAX + 1)
        self.ax.clear()
        self.ax.set_facecolor("black")

        mean, std = np.mean(self.data), np.std(self.data)

        # Set xlims based on incoming array
        xlim_low = max(0, mean - 5 * std)
        xlim_high = min(UINT16_MAX, mean + 5 * std)
        self.ax.set_xlim(xlim_low, xlim_high)

        # Auto-calculate contrast (pm 3 std dev.)
        if self.auto_contrast_enabled:
            self.contrast_low = max(0, mean - 3 * std)
            self.contrast_high = min(UINT16_MAX, mean + 3 * std)

        # Plot histogram and contrast bars
        self.ax.set_ylabel("")
        self.ax.set_xlabel(
            "Pixel Value (DN)", fontsize=8, fontweight="medium", color="#e0e0e0"
        )
        self.ax.set_yticks([])
        self.ax.tick_params(axis="x", labelsize=8, colors="#e0e0e0")
        for spine in self.ax.spines.values():
            if spine.spine_type == "bottom":
                spine.set_visible(True)
                spine.set_color("#808080")
                spine.set_linewidth(1.2)
            else:
                spine.set_visible(False)
        self.ax.plot(counts, color="#FFD700")
        self.bar_low = self.ax.axvline(self.contrast_low, color="purple", linewidth=3)
        self.bar_high = self.ax.axvline(self.contrast_high, color="purple", linewidth=3)
        self.hist_canvas.draw()

    def _auto_contrast(self):
        """
        Enable auto contrast flag and redraw.
        """
        if self.data is None:
            return
        self.auto_contrast_enabled = True
        self._update_histogram()
        self._redraw_image()

    def _on_bar_press(self, event):
        """
        Select a contrast bar if the click is within 6 pixels of it.
        """
        if event.inaxes != self.ax or self.data is None:
            return
        # Convert bar data-coords to pixel-coords for a scale-independent hit test.
        low_px = self.ax.transData.transform([self.contrast_low, 0])[0]
        high_px = self.ax.transData.transform([self.contrast_high, 0])[0]
        if abs(event.x - low_px) < 6:
            self.dragging_bar = "low"
        elif abs(event.x - high_px) < 6:
            self.dragging_bar = "high"

    def _on_bar_motion(self, event):
        if self.dragging_bar is None or event.inaxes != self.ax:
            return
        xdata = int(np.clip(event.xdata, 0, UINT16_MAX))
        (self.bar_low if self.dragging_bar == "low" else self.bar_high).set_xdata(
            [xdata, xdata]
        )
        self.hist_canvas.draw()

    def _on_bar_release(self, event):
        if self.dragging_bar is None or event.inaxes != self.ax:
            return
        xdata = int(np.clip(event.xdata, 0, UINT16_MAX))
        if self.dragging_bar == "low":
            self.contrast_low = xdata
        else:
            self.contrast_high = xdata
        self.dragging_bar = None
        self.auto_contrast_enabled = False  # Disable auto on manual adjust
        self._redraw_image()

    def _redraw_image(self):
        """
        Transform from raw data to rgb and draw as an image.
        """
        if self.data is None:
            return
        h, w = self.data.shape
        stretched = self._stretch_contrast(self.data)
        self.rgb = np.zeros((h, w, 3), np.uint8)
        # Saturation point
        sat = self.data == UINT16_MAX
        # 16 -> 8 bit
        self.rgb[..., 0] = np.where(sat, 255, stretched)
        self.rgb[..., 1] = np.where(sat, 0, stretched)
        self.rgb[..., 2] = np.where(sat, 0, stretched)
        # Draw the image
        qimg = QtGui.QImage(self.rgb.data, w, h, w * 3, QtGui.QImage.Format_RGB888)
        pixmap = QtGui.QPixmap.fromImage(qimg)
        if self.item:
            self.item.setPixmap(pixmap)
        else:
            self.item = self.scene.addPixmap(pixmap)
            self.item.setTransformationMode(QtCore.Qt.FastTransformation)
        self.scene.setSceneRect(0, 0, w, h)
        if self.zoom > 1.0:
            self.view.centerOn(self.item)
        self._update_transform()

    def _stretch_contrast(self, arr):
        """
        Apply contrast from the contrast control bars.
        """
        if self.contrast_high == self.contrast_low:
            return np.zeros(arr.shape, np.uint8)
        scaled = (
            (arr.astype(np.float32) - self.contrast_low)
            * 255.0
            / (self.contrast_high - self.contrast_low)
        )
        return np.clip(scaled, 0, 255).astype(np.uint8)

    def _update_stats(self):
        if self.data is None:
            return
        arr = self.data
        self.stats_labels["Mean"].setText(f"Mean: {np.mean(arr):.1f} DN")
        self.stats_labels["Median"].setText(f"Median: {np.median(arr):.1f} DN")
        self.stats_labels["Max"].setText(f"Max: {np.max(arr)} DN")
        self.stats_labels["Min"].setText(f"Min: {np.min(arr)} DN")
        self.stats_labels["Std. Dev."].setText(f"Std. Dev.: {np.std(arr):.1f} DN")
        self.stats_labels["No. Pix."].setText(f"No. Pix.: {arr.size}")

    def _open_image(self):
        """
        Generate a dummy data array.
        """
        arr = np.random.randint(0, UINT16_MAX + 1, (512, 640), np.uint16)
        arr[100:110, 100:110] = UINT16_MAX
        self.display_array(arr)

    def _wheel_event(self, event):
        """
        Determine image zoom based on scrolling the mouse-wheel.
        """
        delta = 1.2 if event.angleDelta().y() > 0 else 1 / 1.2
        self.zoom = max(1, min(self.zoom * delta, 20))
        self._update_transform()

    def _update_transform(self):
        """
        Apply a given zoom level and add/remove scrollbars if necessary.
        """
        self.view.setTransform(QtGui.QTransform().scale(self.zoom, self.zoom))
        sb = (
            QtCore.Qt.ScrollBarAsNeeded
            if self.zoom > 1.0
            else QtCore.Qt.ScrollBarAlwaysOff
        )
        self.view.setVerticalScrollBarPolicy(sb)
        self.view.setHorizontalScrollBarPolicy(sb)

    def _mouse_move_event(self, event):
        """
        Get the current pixel and associated value at the mouse position after moving.
        """
        if self.data is None:
            QtWidgets.QGraphicsView.mouseMoveEvent(self.view, event)
            return
        pos = self.view.mapToScene(event.pos())
        x, y = int(pos.x()), int(pos.y())
        if 0 <= x < self.data.shape[1] and 0 <= y < self.data.shape[0]:
            self.info.setText(
                f"x:{x} y:{y} Signal:{self.data[y, x]} DN zoom:{self.zoom:.1f}x"
            )
        else:
            self.info.clear()
        QtWidgets.QGraphicsView.mouseMoveEvent(self.view, event)


class PreviewThread(QtCore.QThread):
    """
    PyQt thread for previewing images from the camera.
    """

    # Define the inter-thread signals and states dict
    frame = QtCore.pyqtSignal(np.ndarray)
    status = QtCore.pyqtSignal(float, str)
    error = QtCore.pyqtSignal(str)
    _states = {
        0: "Idle",
        1: "Waiting",
        2: "Exposing",
        3: "Reading",
        4: "Download",
        5: "Error",
    }

    def __init__(self, cam, params):
        super().__init__()
        # Get and assign params
        self.cam = cam
        self.run_flag = True
        self.p = copy.deepcopy(params)

    def run(self):
        """
        Loop for acquiring images from the camera to preview.
        """
        # pythoncom.CoInitialize()
        try:
            cam = self.cam
            p = self.p

            # ROI setup
            cam.BinX, cam.BinY = p["bin_x"], p["bin_y"]
            cam.StartX, cam.StartY = p["start_x"], p["start_y"]
            cam.NumX, cam.NumY = p["width"], p["height"]

            # Pre-calculate exposure time
            unit = p["time_unit"]
            mult = TIME_UNIT_MULTIPLIERS[unit]
            exp_time = p["exp_array"][0] * mult

            states = self._states
            while self.run_flag:
                try:
                    self.status.emit(cam.CCDTemperature, "Exposing")
                    cam.StartExposure(exp_time, True)

                    poll_interval = min(0.01, max(0.0001, exp_time * 0.1))

                    while not cam.ImageReady and self.run_flag:
                        time.sleep(poll_interval)

                    if not self.run_flag:
                        break

                    img = cam.ImageArray
                    arr = np.transpose(np.array(img, dtype=np.int16).astype(np.uint16))
                    self.frame.emit(arr)
                    self.status.emit(
                        cam.CCDTemperature, states.get(cam.CameraState, "Unknown")
                    )
                except Exception as e:
                    self.error.emit(str(e))
                    break
        except Exception as e:
            self.error.emit(str(e))

        finally:
            # pythoncom.CoUninitialize()
            pass

    def abort(self, timeout=CAMERA_ABORT_TIMEOUT_S):
        """
        Abort preview.
        """
        self.run_flag = False
        self.cam.AbortExposure()
        start = time.time()
        while self.cam.CameraState not in (0, 5) and time.time() - start < timeout:
            time.sleep(0.1)
        if time.time() - start >= timeout:
            self.error.emit("Camera abort timeout!")


@dataclass(frozen=True, slots=True)
class ExperimentSpec:
    """
    Immutable description of one experiment block.

    Exists so worker threads never need to touch the Qwidgets.
    """

    index: int
    title: str
    params: dict[str, object]


@dataclass(slots=True)
class ExposureMeta:
    """
    Exposure metadata for file-saving.
    """
    exp_idx: int              # current experiment index
    img_idx: int              # current image index
    rep_idx: int              # current experiment repeat index
    exposure_s: float         # exposure array (sec)
    params: dict              # camera params dict (gain, rospeed etc.)
    temp: float               # current camera temp
    target_temp: float        # camera target temp
    ts_utc: object            # astropy.time.Time
    ts_ltc: datetime          # Local time at start of exposure
    filter: str | None        # current filter-wheel pos (None when not connected)
    ra_target: float | None   # mount target ra (None when not connected)
    dec_target: float | None  # mount target dec (None when not connected)
    ra_mount: float | None    # mount current ra (None when not connected)
    dec_mount: float | None   # mount current ra (None when not connected)
    alt_mount: float | None   # mount current alt (None when not connected)
    az_mount: float | None    # mount currrent az (None when not connected)
    focuser_pos: int | None   # focuser current pos (None when not connected)


class ExperimentThread(QtCore.QThread):
    """
    PyQt thread for acquiring images from the camera.
    """

    frame = QtCore.pyqtSignal(np.ndarray)
    status = QtCore.pyqtSignal(float, str)
    progress = QtCore.pyqtSignal(int, str, float, dict, int, int, int, int)
    error = QtCore.pyqtSignal(str)
    log_requested = QtCore.pyqtSignal(str, str)
    save_requested = QtCore.pyqtSignal(np.ndarray, object)
    sequence_finished = QtCore.pyqtSignal()

    _states = {
        0: "Idle",
        1: "Waiting",
        2: "Exposing",
        3: "Reading",
        4: "Download",
        5: "Error",
    }

    def __init__(
        self,
        camera,
        experiments,
        filterwheel=None,
        telescope=None,
        focuser=None):
        super().__init__()

        self.cam = camera
        self.experiments = copy.deepcopy(list(experiments))
        self.fw = filterwheel
        self.telescope = telescope
        self.focuser = focuser
        self.run_flag = True
        self.filter_names = CFW10_FILTER_NAMES_SHORT

    def _interruptible_sleep(self, seconds):
        end = time.monotonic() + max(0.0, float(seconds))

        while self.run_flag and time.monotonic() < end:
            time.sleep(min(0.1, max(0.0, end - time.monotonic())))

    def _current_filter_name(self):
        if self.fw is None:
            return None

        try:
            state = self.fw.query_state()
        except Exception as exc:
            self.log_requested.emit(f"Filter-wheel query failed: {exc}", "red")
            return None

        if 1 <= state <= len(self.filter_names):
            return self.filter_names[state - 1]

        return None

    def _read_mount_state(self):
        ra_mount = None
        dec_mount = None
        ra_target = None
        dec_target = None
        alt_mount = None
        az_mount = None

        if self.telescope is not None:
            try:
                ra_mount = float(self.telescope.RightAscension)
                dec_mount = float(self.telescope.Declination)
                ra_target = float(self.telescope.TargetRightAscension)
                dec_target = float(self.telescope.TargetDeclination)
                alt_mount = float(self.telescope.Altitude)
                az_mount = float(self.telescope.Azimuth)
            except Exception as exc:
                self.log_requested.emit(f"Mount metadata unavailable: {exc}", "orange")

        return ra_mount, dec_mount, ra_target, dec_target, alt_mount, az_mount

    def _read_focuser_position(self):
        if self.focuser is None:
            return -1

        try:
            return int(self.focuser.Position)
        except Exception as exc:
            self.log_requested.emit(f"Focuser metadata unavailable: {exc}", "orange")
            return -1

    def run(self):
        """
        Execute all experiment blocks.
        """
        try:
            cam = self.cam
            states = self._states

            for exp_idx, spec in enumerate(self.experiments):
                if not self.run_flag:
                    break

                params = copy.deepcopy(spec.params)
                name = spec.title

                try:
                    bin_x = int(params["bin_x"])
                    bin_y = int(params["bin_y"])
                    start_x = int(params["start_x"])
                    start_y = int(params["start_y"])
                    width = int(params["width"])
                    height = int(params["height"])

                    cam.BinX = bin_x
                    cam.BinY = bin_y
                    cam.StartX = start_x
                    cam.StartY = start_y
                    cam.NumX = width
                    cam.NumY = height
                except Exception as exc:
                    self.error.emit(
                        f"Experiment {exp_idx + 1} failed ROI configuration: {exc}"
                    )
                    continue

                unit = str(params.get("time_unit", "s"))
                mult = TIME_UNIT_MULTIPLIERS.get(unit)

                if mult is None:
                    self.error.emit(
                        f"Experiment {exp_idx + 1} has unknown time unit: {unit}"
                    )
                    continue

                try:
                    exp_array = np.asarray(params["exp_array"], dtype=float) * mult
                except Exception as exc:
                    self.error.emit(
                        f"Experiment {exp_idx + 1} has invalid exposure array: {exc}"
                    )
                    continue

                if exp_array.size == 0:
                    self.error.emit(f"Experiment {exp_idx + 1} has no exposures.")
                    continue

                if not np.all(np.isfinite(exp_array)) or np.any(exp_array <= 0):
                    self.error.emit(
                        f"Experiment {exp_idx + 1} contains non-positive or non-finite exposures."
                    )
                    continue

                try:
                    experiment_delay = float(params.get("delay", 0.0))
                except Exception:
                    experiment_delay = 0.0

                try:
                    experiment_repeats = max(1, int(params.get("repeats", 1)))
                except Exception:
                    experiment_repeats = 1

                total_exp = int(exp_array.size)
                filter_name = self._current_filter_name()

                for rep_idx in range(experiment_repeats):
                    if not self.run_flag:
                        break

                    if experiment_delay > 0:
                        self.log_requested.emit(
                            f"Waiting {experiment_delay:.3f}s",
                            "orange",
                        )
                        self._interruptible_sleep(experiment_delay)

                    if not self.run_flag:
                        break

                    for img_idx, exp in enumerate(exp_array):
                        if not self.run_flag:
                            break

                        try:
                            temp = float(cam.CCDTemperature)
                            target_temp = float(cam.SetCCDTemperature)
                            ts_ltc = datetime.now()
                            ts_utc = Time.now()

                            (
                                ra_mount,
                                dec_mount,
                                ra_target,
                                dec_target,
                                alt_mount,
                                az_mount,
                            ) = self._read_mount_state()

                            focuser_pos = self._read_focuser_position()

                            self.status.emit(cam.CCDTemperature, "Exposing")
                            cam.StartExposure(float(exp), True)

                            poll_interval = min(0.01, max(0.0001, float(exp) * 0.1))

                            while not cam.ImageReady and self.run_flag:
                                time.sleep(poll_interval)

                            if not self.run_flag:
                                break

                            img = cam.ImageArray
                            arr = np.transpose(
                                np.asarray(img, dtype=np.int16).astype(np.uint16)
                            )

                            self.frame.emit(arr)

                            meta = ExposureMeta(
                                exp_idx=exp_idx,
                                img_idx=img_idx,
                                rep_idx=rep_idx,
                                exposure_s=float(exp),
                                params=params,
                                temp=temp,
                                target_temp=target_temp,
                                ts_utc=ts_utc,
                                ts_ltc=ts_ltc,
                                filter=filter_name,
                                ra_target=ra_target,
                                dec_target=dec_target,
                                ra_mount=ra_mount,
                                dec_mount=dec_mount,
                                alt_mount=alt_mount,
                                az_mount=az_mount,
                                focuser_pos=focuser_pos,
                            )

                            self.save_requested.emit(arr, meta)

                            self.progress.emit(
                                exp_idx,
                                name,
                                float(params["exp_array"][img_idx]),
                                params,
                                img_idx + 1,
                                rep_idx + 1,
                                experiment_repeats,
                                total_exp,
                            )

                            self.status.emit(
                                cam.CCDTemperature,
                                states.get(cam.CameraState, "Unknown"),
                            )

                        except Exception as e:
                            self.error.emit(
                                f"Frame {img_idx + 1}/{total_exp} failed "
                                f"(exp {exp_idx + 1}, rep {rep_idx + 1}): {e}, skipping..."
                            )

                            try:
                                cam.AbortExposure()
                            except Exception:
                                pass

                            continue

                    if not self.run_flag:
                        break
        except Exception as e:
            self.error.emit(str(e))
        finally:
            # pythoncom.CoUninitialize()
            if self.run_flag:
                self.sequence_finished.emit()

    def abort(self, timeout = CAMERA_ABORT_TIMEOUT_S):
        self.run_flag = False

        try:
            self.cam.AbortExposure()
        except Exception:
            pass

        start = time.time()
        while self.cam.CameraState not in (0, 5) and time.time() - start < timeout:
            time.sleep(0.1)

        if time.time() - start >= timeout:
            self.error.emit("Camera abort timeout!")


class SaverThread(QtCore.QThread):
    """
    PyQt thread for saving a queue of images.
    """

    progress = QtCore.pyqtSignal(str)
    log_error = QtCore.pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self.queue = Queue(maxsize=SAVER_QUEUE_MAXSIZE)

        self._stopped: bool = False

    _SENTINEL = object()  # sentinel for queue shutdown

    def run(self):
        """
        File saving loop.
        """
        # pythoncom.CoInitialize()
        try:
            q = self.queue
            while True:
                job = q.get()
                if job is self._SENTINEL:
                    break
                arr, hdr, path = job

                try:
                    Path(path).parent.mkdir(parents=True, exist_ok=True)
                    safe_path = self._unique_path(path)
                    fits.writeto(safe_path, arr, hdr, overwrite=False)
                    if safe_path != path:
                        self.progress.emit(
                            f"{path} (saved as {os.path.basename(safe_path)} due to name collision)"
                        )
                    else:
                        self.progress.emit(safe_path)
                    self.progress.emit(path)
                except Exception as e:
                    self.log_error.emit(f"Save failed ({path}): {e}", "red")
        finally:
            # pythoncom.CoUninitialize()
            pass

    def save(self, arr, meta):
        """
        Queue a frame for saving if the saver is still active.
        """
        if self._stopped:
            self.log_error.emit(
                f"Saver stopped; dropping frame {meta.img_idx + 1}, "
                f"exp {meta.exp_idx + 1}, rep {meta.rep_idx + 1}.",
                "orange",
            )
            return

        try:
            path = self._build_path(meta)
            hdr = self._build_header(meta)
        except Exception as exc:
            self.log_error.emit(
                f"Failed to prepare frame {meta.img_idx + 1}, "
                f"exp {meta.exp_idx + 1}, rep {meta.rep_idx + 1}: {exc}",
                "red",
            )
            return

        try:
            self.queue.put((arr, hdr, path), timeout=SAVER_QUEUE_PUT_TIMEOUT_S)
        except Full:
            self.log_error.emit(
                f"Save queue full, dropping frame {meta.img_idx + 1}, "
                f"exp {meta.exp_idx + 1}, rep {meta.rep_idx + 1}!",
                "red",
            )
        except Exception as exc:
            self.log_error.emit(
                f"Save queue error for frame {meta.img_idx + 1}, "
                f"exp {meta.exp_idx + 1}, rep {meta.rep_idx + 1}: {exc}",
                "red",
            )

    def _build_path(self, meta):
        """
        Assemble the full output file path from metadata.
        """
        p = meta.params
        parts: list[str] = [str(p["nameprefix"])]

        if p["include_date_time"]:
            fmt = "_%d_%m_%Y_%H-%M" + ("pm" if meta.ts_ltc.hour >= 12 else "am")
            parts.append(meta.ts_ltc.strftime(fmt))

        if p["include_gain_rospeed"]:
            parts.append(f"_{p['gain']}_{p['rospeed']}")

        if p["include_temp"]:
            parts.append(f"_{meta.target_temp:.0f}C")

        if p["include_exp_time"]:
            scale = TIME_UNIT_MULTIPLIERS[p["time_unit"]]
            parts.append(f"_{meta.exposure_s / scale:g}{p['time_unit']}")

        base = self._sanitize_filename("".join(parts))

        fname = self._sanitize_filename(
            f"{base}_exp{meta.exp_idx + 1}_rep{meta.rep_idx + 1}_img{meta.img_idx + 1}.fits"
        )

        directory = Path(str(p["savedirectory"])).expanduser()

        if str(directory).strip() in ("", "."):
            directory = Path.home() / "Desktop"

        return str(directory / fname)

    def _build_header(self, meta):
        """
        Helper function to build the fits header.
        """
        obs_loc = EarthLocation(
            lat=OBSERVATORY_LATITUDE * u.deg, lon=OBSERVATORY_LONGITUDE * u.deg
        )
        p = meta.params
        ts_utc = meta.ts_utc
        ts_ltc = meta.ts_ltc
        lst = ts_utc.sidereal_time("mean", longitude=obs_loc.lon)
        exp = meta.exposure_s
        temp = meta.temp
        target_temp = meta.target_temp

        filter_name = meta.filter
        ra_target = meta.ra_target
        dec_target = meta.dec_target
        ra_mount = meta.ra_mount
        dec_mount = meta.dec_mount
        alt_mount = meta.alt_mount
        az_mount = meta.az_mount
        focuser_pos = meta.focuser_pos

        hdr = fits.Header()

        # precalc here to use in crpix
        naxis1 = int(p["width"] / p["bin_x"])
        naxis2 = int(p["height"] / p["bin_y"])

        # hdr['SIMPLE']  = (True, 'File conforms to FITS standard')                                  # <----- NOTE autopopulated by astropy
        # hdr['BITPIX']  = (16, 'Number of bits per data pixel')                                     # <----- NOTE autopopulated by astropy
        # hdr['NAXIS']   = (2, 'Number of data axes')                                                # <----- NOTE autopopulated by astropy
        # hdr['NAXIS1']  = (int(p['width']/p['bin_x']), 'Length of data axis 1 (X)')                 # <----- NOTE autopopulated by astropy
        # hdr['NAXIS2']  = (int(p['height']/p['bin_y']), 'Length of data axis 2 (Y)')                # <----- NOTE autopopulated by astropy
        hdr["EXTEND"] = (True, "FITS dataset may contain extensions")
        hdr["BZERO"] = (32768.0, "Offset data range (for unsigned 16-bit)")
        hdr["BSCALE"] = (1.0, "Scale factor for the data")
        # hdr['OBJECT']  = ('M42', 'Target name')                                                     # <----- NOTE unimplemented
        hdr["IMAGETYP"] = (
            "Light Frame",
            "Type of image (Light, Dark, Bias, Flat)",
        )  # <----- TODO (need to add ui element to specify type)
        hdr["OBSERVER"] = (DEFAULT_OBSERVER, "Name of observer")
        hdr["EXPTIME"] = (exp, "Exposure time (seconds)")
        hdr["DARKTIME"] = (exp, "Total dark current time (seconds)")
        hdr["DATE-OBS"] = (ts_utc.isot, "UTC at start of exposure")
        hdr["MJD-OBS"] = (ts_utc.mjd, "Modified Julian Date at start of exposure")
        hdr["JD-OBS"] = (ts_utc.jd, "Julian Date at start of exposure")
        hdr["LOCALTIM"] = (ts_ltc.isoformat(), "Local time at start of exposure")
        hdr["LST"] = (
            lst.to_string(unit=u.hourangle, sep=":", precision=2, pad=True),
            "Local Sidereal Time at exposure start",
        )
        # hdr['TELESCOP']= ('X', 'Name of telescope')                                                 # <----- TODO
        # hdr['FOCALLEN']= (42, 'Telescope focal length in mm')                                       # <----- TODO
        # hdr['APTDIA']  = (42, 'Telescope aperture diameter in mm')                                  # <----- TODO
        if focuser_pos is not None:
            hdr["FOC-POS"] = (focuser_pos, "Focuser position (steps)")
        if ra_target is not None:
            # ALPYCA TargetRightAscension returns RA in hours, convert to deg with Astropy
            ra_target_hour = Angle(ra_target, unit=u.hourangle)
            ra_target_deg = ra_target_hour.degree
            hdr["RA"] = (ra_target_deg, "Right Ascension of target (J2000, degrees)")
        if dec_target is not None:
            hdr["DEC"] = (dec_target, "Declination of target (J2000, degrees)")
        # hdr['HA']      = (mount_ha, 'Hour angle of target')                                         # <----- NOTE unimplemented
        # hdr['AIRMASS'] = (airmass, 'Airmass at start of exposure')                                  # <----- NOTE unimplemented
        if alt_mount is not None:
            hdr["ALTITUDE"] = (alt_mount, "Altitude of target (degrees)")
        if az_mount is not None:
            hdr["AZIMUTH"] = (az_mount, "Azimuth of target (degrees)")
        hdr["RADECSYS"] = ("ICRS", "Coordinate reference frame (e.g., ICRS, FK5)")
        hdr["INSTRUME"] = ("NIRvana HS", "The model of the camera")
        hdr["DETECTOR"] = ("InGaAs", "Detector")
        if filter_name is not None:
            hdr["FILTER"] = (filter_name, "Filter used for this observation")
        hdr["XBINNING"] = (p["bin_x"], "Binning factor in width")
        hdr["YBINNING"] = (p["bin_y"], "Binning factor in height")
        hdr["XPIXSZ"] = (
            CAMERA_PIXEL_SIZE_UM * p["bin_x"],
            "Pixel width in microns, after binning",
        )
        hdr["YPIXSZ"] = (
            CAMERA_PIXEL_SIZE_UM * p["bin_y"],
            "Pixel height in microns, after binning",
        )
        hdr["XORGSUBF"] = (
            int(p["start_x"] / p["bin_x"]),
            "Subframe upper-left X in pixels",
        )
        hdr["YORGSUBF"] = (
            int(p["start_y"] / p["bin_y"]),
            "Subframe upper-left Y in pixels",
        )
        hdr["CCD-TEMP"] = (temp, "Temperature of the sensor (degrees C)")
        hdr["SET-TEMP"] = (target_temp, "Cooler setpoint (degrees C)")
        hdr["GAIN"] = (p["gain"], "Gain mode of the camera (high/low)")
        hdr["ROSPEED"] = (p["rospeed"], "Read-out speed of the camera (MHz)")
        hdr["OBSERVAT"] = (OBSERVATORY_NAME, "Name of observatory/site")
        hdr["SITELAT"] = (OBSERVATORY_LATITUDE, "Site latitude (degrees, +North)")
        hdr["SITELONG"] = (OBSERVATORY_LONGITUDE, "Site longitude (degrees, +East)")
        hdr["SITEELEV"] = (OBSERVATORY_ELEVATION, "Site elevation (meters)")
        # hdr['TAMBIENT']= (ambient_temp, 'Ambient temperature (degrees C)')                           # <----- NOTE unimplemented
        # hdr['PRESSURE']= (pressure, 'Atmospheric pressure (hPa)')                                    # <----- NOTE unimplemented
        # hdr['HUMIDITY']= (humidity, 'Relative humidity in %')                                        # <----- NOTE unimplemented
        hdr["CTYPE1"] = ("RA---TAN", "Coordinate type for Axis 1")
        hdr["CTYPE2"] = ("DEC--TAN", "Coordinate type for Axis 2")
        if ra_mount is not None:
            # ALPYCA RightAscension returns RA in hours, convert to deg with Astropy
            ra_mount_hour = Angle(ra_mount, unit=u.hourangle)
            ra_mount_deg = ra_mount_hour.degree
            hdr["CRVAL1"] = (ra_mount_deg, "RA at reference pixel (degrees)")
        if dec_mount is not None:
            hdr["CRVAL2"] = (dec_mount, "DEC at reference pixel (degrees)")
        hdr["CRPIX1"] = ((naxis1 + 1) / 2.0, "Reference pixel on X axis")
        hdr["CRPIX2"] = ((naxis2 + 1) / 2.0, "Reference pixel on Y axis")
        hdr["COMMENT"] = "FITS (Flexible Image Transport System) format"
        hdr["HISTORY"] = "Created using Python ASCOM client"
        hdr["HISTORY"] = "Raw frame captured without flat field applied"

        return hdr

    def _unique_path(self, base_path):
        """
        Return a unique file path if the base path already exists.
        """
        path = Path(base_path)
        if not path.exists():
            return str(path)
        counter = 1
        while True:
            new_path = path.parent / f"{path.stem}_{counter}{path.suffix}"
            if not new_path.exists():
                return str(new_path)
            counter += 1

    _INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

    @classmethod
    def _sanitize_filename(cls, value):
        """
        Sanitize a string for safe use as a filename component.
        """
        value = str(value)
        value = cls._INVALID_FILENAME_CHARS.sub("_", value)
        value = value.replace("..", "_")
        value = value.strip(" .")
        return value or "image"

    def stop(self, timeout_ms = 10000):
        self._stopped = True

        try:
            self.queue.put(self._SENTINEL, timeout=SAVER_QUEUE_PUT_TIMEOUT_S)
        except Full:
            self.log_error.emit(
                "Could not enqueue saver shutdown sentinel (queue full).",
                "orange",
            )

        stopped = self.wait(timeout_ms)
        if not stopped:
            self.log_error.emit("Save thread did not stop within timeout.", "orange")

        return stopped


class FilterwheelThread(QtCore.QThread):
    """
    PyQt thread for changing filter-wheel position.
    """

    error = QtCore.pyqtSignal(str)
    move_finished = QtCore.pyqtSignal(int, str)

    def __init__(self, fw, selected_index, fname):
        super().__init__()
        self.fw, self.selected_index, self.fname = fw, selected_index, fname

    def run(self):
        """
        Move the wheel, then confirm the requested position was reached.
        """
        try:
            result = self.fw.change_filter(int(self.selected_index))
            if isinstance(result, tuple) and result[0] < 0:
                self.error.emit(result[1])
                return

            # Bounded confirmation poll (max CFW10_MOVE_TIMEOUT_S).
            state = self.fw.query_state()
            timeOut = time.time()
            while state != self.selected_index:
                if time.time() - timeOut > CFW10_MOVE_TIMEOUT_S:
                    self.error.emit("Filter-wheel did not reach requested position")
                    return
                time.sleep(0.02)
                state = self.fw.query_state()

            self.move_finished.emit(state, self.fname[state - 1])
        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QtWidgets.QWidget):
    """
    The main GUI window defined and usable as a QWidget.
    """

    GREEN_STYLE = """
        background-color: #45a049;
        color: white;
        border-radius: 5px;
        padding: 10px;
        font: bold 16px;
        """

    RED_STYLE = """
        background-color: #cc0000;
        color: white;
        border-radius: 5px;
        padding: 10px;
        font: bold 16px;
        """

    status_update = QtCore.pyqtSignal(float, str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ASCOM Camera Controller")

        # Initialise poll values
        self._temp_target = CAMERA_MIN_TEMP_C
        self.filterwheel_moving = False

        # Timer for idle polling (main thread only)
        self._poll_timer = QtCore.QTimer()
        self._poll_timer.setInterval(POLL_INTERVAL_MS)
        self._poll_timer.timeout.connect(self._idle_poll)
        self._poll_timer.start()

        self._active_experiment_count = 0

        # Thread placeholders
        self._preview_thread = None
        self._exp_thread = None
        self._saver_thread = None
        self._fw_thread = None

        # Wire signal to UI slot
        self.status_update.connect(self._update_status)

        # Master window layout
        self._master_layout = QtWidgets.QHBoxLayout(self)
        self.setLayout(self._master_layout)

        self.camera = None
        self.is_connected = False
        self.filterwheel_is_connected = False
        self.mount_is_connected = False
        self.fw = None
        self.telescope = None
        self.focuser = None

        # Build UI elements
        self._build_left_side()
        self._build_right_side()

    def _update_status(self, temp, state):
        """
        Update the current camera temperature label.
        """
        if abs(temp - self._temp_target) <= 0.5:
            self._current_temp_label.setText(
                f"Temperature: {temp:.1f}°C     State: {state}"
            )
        else:
            self._current_temp_label.setText(
                f"Temperature: ⚠{temp:.1f}°C ({self._temp_target}°C target)     State: {state}"
            )

    def _idle_poll(self):
        """
        Main thread polling when camera is idle.
        """
        if self.camera and self.camera.Connected:
            try:
                _state = {
                    0: "Idle",
                    1: "Waiting",
                    2: "Exposing",
                    3: "Reading",
                    4: "Download",
                    5: "Error",
                }.get(self.camera.CameraState, "Unknown")
                _temp = self.camera.CCDTemperature

                self._update_status(_temp, _state)
            except Exception as e:
                self.log.log(f"Error reading temperature/state: {e}", "red")

        if self.filterwheel_is_connected and not self.filterwheel_moving:
            try:
                state = self.fw.query_state()
                filterName = (
                    self.fname[state - 1] if 1 <= state <= len(self.fname) else None
                )
                self._current_filter_label.setText(
                    f"Position: {filterName}     State: Idle"
                )
            except Exception as e:
                self.log.log(f"Error reading Filter-wheel state: {e}", "red")

    def _update_containers(self):
        """
        Update the number of experiment containers.
        """
        target = self._spinbox.value()

        # Remove excess containers.
        while len(self._containers) > target:
            widget = self._containers.pop()
            self._container_layout.removeWidget(widget)
            widget.deleteLater()

        # Add new containers.
        while len(self._containers) < target:
            i = len(self._containers)
            container = CollapsibleWidget(f"Experiment {i + 1}")
            self._container_layout.insertWidget(i, container)
            self._containers.append(container)

            contents_layout = QtWidgets.QVBoxLayout(container.contentWidget)
            contents_layout.addWidget(CollapsibleWidgetInternalGUI())

    def _get_experiments(self):
        """
        Extract experiment index, title, and parameters from GUI containers.

        Returns
        -------
        experiments : list[tuple[int, str, dict[str, object]]]
            List of tuples containing experiment index, title, and params.
        """
        experiments = []

        for i, container in enumerate(self._containers):
            gui = container.internalGui
            if gui is not None:
                experiments.append((i, container.title, gui.window_params))

        return experiments

    def _save_config(self):
        """
        Save all experiment containers to a json file.

        Format is:
            {
              "version": 1,
              "experiments": [
                {"title": "...", "params": { ... window_params ... }},
                ...
              ]
            }
        """
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Experiment Configuration",
            "",
            "JSON Files (*.json)",
        )
        if not path:
            return
        config = {
            "version": 1,
            "experiments": [
                {
                    "title": c.title,
                    "params": c.internalGui.window_params,
                }
                for c in self._containers
                if c.internalGui is not None
            ],
        }

        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(
                    config,
                    fh,
                    indent=2,
                    ensure_ascii=False,
                    default=self._json_default,
                )
            self.log.log(
                f"Configuration saved: {os.path.basename(path)}",
                "#1dff35",
            )
        except Exception as e:
            self.log.log(f"Error saving configuration: {e}", "red")

    @staticmethod
    def _json_default(obj):
        """
        Fallback for types not handled by json:
        - numpy scalars
        - pathlib Paths

        Parameters
        ----------
        obj
            The object that json.dump could not serialise.

        Returns
        -------
            json-compatible representation of obj.

        Raises
        ------
            TypeError: If obj is not a recognised type.
        """
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        if isinstance(obj, os.PathLike):
            return str(obj)
        raise TypeError(f"Object of type {type(obj).__name__} is not JSON serialisable")

    def _load_config(self):
        """
        Load experiment containers from a json file.

        The experiment number widget is set to the saved experiment count, 
        triggering _update_containers to create or remove widgets, and each
        container's window_params is overwritten by saved values.
        Missing keys are filled from _default_params().
        """
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Experiment Configuration",
            "",
            "JSON Files (*.json)",
        )
        if not path:
            return

        # Warn if acquisitions are in progress.
        if self._exp_thread is not None or self._preview_thread is not None:
            btn = QtWidgets.QMessageBox.warning(
                self,
                "Acquisitions running",
                "<b>Loading a configuration will not affect the "
                "currently running acquisition.</b><br><br>"
                "Continue?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No,
            )
            if btn == QtWidgets.QMessageBox.No:
                return

        # Read and parse
        try:
            with open(path, encoding="utf-8") as fh:
                config = json.load(fh)
        except (json.JSONDecodeError, OSError) as e:
            self.log.log(f"Error loading configuration: {e}", "red")
            return

        # Validation
        if not isinstance(config, dict):
            self.log.log(
                "Invalid configuration: root element is not a JSON object",
                "red",
            )
            return

        experiments = config.get("experiments", [])
        if not isinstance(experiments, list) or not experiments:
            self.log.log(
                "Configuration file contains no experiments",
                "orange",
            )
            return

        # Validate every entry
        for i, entry in enumerate(experiments):
            if not isinstance(entry, dict):
                self.log.log(
                    f"Invalid configuration: experiment {i + 1} is not a JSON object",
                    "red",
                )
                return

        # Resize container list to match file
        self._spinbox.setValue(len(experiments))

        # Populate each container
        for i, exp_data in enumerate(experiments):
            container = self._containers[i]

            # Restore assigned title
            container.title = exp_data.get("title", f"Experiment {i + 1}")

            # Merge params over fresh defaults
            merged = CollapsibleWidgetInternalGUI._default_params()
            saved_params = exp_data.get("params", {})
            if isinstance(saved_params, dict):
                merged.update(saved_params)
            if merged["time_unit"] not in set(TIME_UNIT_MULTIPLIERS):
                merged["time_unit"] = "s"
            try:
                merged["exp_array"] = (
                    CollapsibleWidgetInternalGUI._parse_exposure_array(
                        str(merged.get("exp_array", "[1]"))
                    )
                )
            except ExposureParserError:
                merged["exp_array"] = [1.0]
                self.log.log(
                    f"Experiment {i + 1}: invalid saved exposure array; reset to [1.0]",
                    "orange",
                )
            container.internalGui.window_params = merged

            # Push loaded values into visible widgets.
            container.internalGui.refresh_gui_from_params()

        self.log.log(
            f"Loaded {len(experiments)} experiment(s) from {os.path.basename(path)}",
            "#1dff35",
        )

    # Dummy function for testing retrieval of experiment settings:
    # Dumps all experiment settings to the log widget.
    def _grab_contents_fun(self):
        """
        Dummy function for testing retrieval of experiment settings.

        Dumps all experiment settings to the log widget.
        """
        for idx, title, params in self._get_experiments():
            self.log.log(f"Running experiment {idx}: {title}", "#1dff35")
            self.log.log(f"Params: {params}", "#1dff35")

    def _connect_mount_pressed(self):
        """
        Connect/disconnect mount and focuser.
        """
        if not self.mount_is_connected:
            telescope = None
            focuser = None

            try:
                self.log.log("Connecting mount", "orange")
                telescope = Telescope(ALPACA_ADDRESS, ALPACA_MOUNT_DEVICE)
                telescope.Connected = True
                self.log.log(f"Connected mount: {telescope.Name}", "#1dff35")

                self.log.log("Connecting focuser", "orange")
                focuser = Focuser(ALPACA_ADDRESS, ALPACA_FOCUSER_DEVICE)
                focuser.Connected = True
                self.log.log(f"Connected focuser: {focuser.Name}", "#1dff35")

            except Exception as e:
                for device in (telescope, focuser):
                    if device is None:
                        continue

                    try:
                        device.Connected = False
                    except Exception:
                        pass

                self.telescope = None
                self.focuser = None
                self.mount_is_connected = False
                self.log.log(f"Error connecting mount/focuser: {e}", "red")
                return

            self.telescope = telescope
            self.focuser = focuser
            self.mount_is_connected = True

            self._mount_connect_button.setStyleSheet(self.RED_STYLE)
            self._mount_connect_button.setText("Disconnect Mount/Focuser")
        else:
            try:
                # Filterwheel disconnection logic
                if self.mount_is_connected:
                    try:
                        if self._exp_thread:
                            msg = (
                                "<b>Camera previewing/running experiments.</b><br><br>"
                                "(Consider halting preview/aborting experiments before disconnecting Mount/Focuser)"
                            )
                            btn = QtWidgets.QMessageBox.critical(
                                self,
                                "Preview/Experiments running",
                                msg,
                                QtWidgets.QMessageBox.Cancel
                                | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel,
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return
                    except Exception:
                        pass

                    self.log.log("Disconnecting Mount", "orange")
                    self.telescope.Connected = False
                    self.log.log(
                        f"Disconnected Mount: {self.telescope.Name}", "#1dff35"
                    )
                    self.telescope = None

                    self.log.log("Disconnecting Focuser", "orange")
                    self.focuser.Connected = False
                    self.log.log(f"Disconnected Mount: {self.focuser.Name}", "#1dff35")
                    self.focuser = None

                    self.mount_is_connected = False

                    # Switch from disconnect -> connect
                    self._mount_connect_button.setStyleSheet(self.GREEN_STYLE)
                    self._mount_connect_button.setText("Connect Mount/Focuser")

            except Exception as e:
                self.log.log(f"Error disconnecting Mount/Focuser: {e}", "red")

    def _connect_filterwheel_pressed(self):
        """
        Connect/disconnect filter-wheel.
        """
        if not self.filterwheel_is_connected:
            try:
                # Connect to filterwheel serial port
                self.log.log("Connecting Filter-wheel", "orange")

                self.fw = CFW10(
                    port=CFW10_PORT,
                    baudrate=CFW10_BAUDRATE,
                    timeout=0,
                    parity=serial.PARITY_NONE,
                )
                self.fw.reset_input_buffer()
                self.fw.reset_output_buffer()

                self.fname = CFW10_FILTER_NAMES_LONG

                # Switch from connect -> disconnect
                self._filterwheel_connect_button.setStyleSheet(self.RED_STYLE)
                self._filterwheel_connect_button.setText("Disconnect Filter-wheel")
                self._filter_settings_row.setVisible(True)
                self._current_filter_label.setVisible(True)
                state = self.fw.query_state()

                if 1 <= state <= len(self.fname):
                    filterName = self.fname[state - 1]
                    self._current_filter_label.setText(
                        f"Position: {filterName}     State: Idle"
                    )
                else:
                    self._current_filter_label.setText(
                        "Position: Unknown     State: Error"
                    )
                    self.log.log(
                        f"Filter-wheel reported invalid state {state}",
                        "orange",
                    )
                self.filterwheel_is_connected = True
                self.filterwheel_moving = False
                self.log.log("Connected Filter-wheel", "#1dff35")
            except Exception as e:
                self.log.log(f"Error connecting Filter-wheel: {e}", "red")
                self.filterwheel_is_connected = False
                self.filterwheel_moving = False
                if self.fw is not None:
                    self.fw.close()
        else:
            try:
                # Filterwheel disconnection logic
                if self.filterwheel_is_connected:
                    try:
                        if self._exp_thread:
                            msg = (
                                "<b>Camera previewing/running experiments.</b><br><br>"
                                "(Consider halting preview/aborting experiments before disconnecting Filter-wheel)"
                            )
                            btn = QtWidgets.QMessageBox.critical(
                                self,
                                "Preview/Experiments running",
                                msg,
                                QtWidgets.QMessageBox.Cancel
                                | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel,
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return
                        if self._fw_thread is not None:
                            msg = "<b>Cannot disconnect while Filter-wheel changing position.</b><br><br>"
                            btn = QtWidgets.QMessageBox.critical(
                                self,
                                "Filter-wheel changing position",
                                msg,
                                QtWidgets.QMessageBox.Cancel,
                            )
                            return
                    except Exception:
                        pass

                    self.log.log("Disconnecting Filter-wheel", "orange")
                    self.filterwheel_is_connected = False
                    self.filterwheel_moving = False
                    self.fw.close()
                    self.log.log("Filter-wheel disconnected", "#1dff35")

                    # Switch from disconnect -> connect
                    self._filterwheel_connect_button.setStyleSheet(self.GREEN_STYLE)
                    self._filterwheel_connect_button.setText("Connect Filter-wheel")
                    self._filter_settings_row.setVisible(False)
                    self._current_filter_label.setVisible(False)

            except Exception as e:
                self.log.log(f"Error disconnecting Filter-wheel: {e}", "red")

    def _connect_pressed(self):
        """
        Connect/disconnect camera.
        """
        if not self.is_connected:
            # Camera connection logic
            try:
                self.camera_running = True

                # Connect to NIRvana HS
                # self.camera = win32com.client.Dispatch('ASCOM.PI.Camera.1')

                # Connect to ASCOM SIM
                self.camera = Camera(ALPACA_ADDRESS, ALPACA_CAMERA_DEVICE)

                self.log.log("Connecting Camera", "orange")
                self.camera.Connected = True
                self.log.log(f"Connected Camera: {self.camera.Name}", "#1dff35")

                # Initialize ROI settings
                self.camera.BinX = 1
                self.camera.BinY = 1
                self.camera.StartX = 0
                self.camera.StartY = 0
                self.camera.NumX = self.camera.CameraXSize // self.camera.BinX
                self.camera.NumY = self.camera.CameraYSize // self.camera.BinY
                self.log.log("Initialised ROI settings", "#1dff35")

                # Set target temperature
                self.camera.CoolerOn = True
                self._temp_target = CAMERA_MIN_TEMP_C
                self.camera.SetCCDTemperature = self._temp_target
                for container in self._containers:
                    if container.internalGui is not None:
                        container.internalGui.set_temp_value(self._temp_target)
                self.log.log(f"Set target temp. {self._temp_target}°C", "#1dff35")

                # Start temperature monitoring poll
                # self._poll_timer.start()
                # self.log.log("Started temperature monitor", "#1dff35")

                # Switch from connect -> disconnect
                self._connect_button.setStyleSheet(self.RED_STYLE)
                self._connect_button.setText("Disconnect Camera")
                self._preview_run_buttons_row.setVisible(True)
                self._cam_settings_frame.setVisible(True)

                self.is_connected = True

            except Exception as e:
                self.camera_running = False
                self.log.log(f"Error connecting to camera: {e}", "red")

        else:
            try:
                # Camera disconnection logic
                if self.camera:
                    try:
                        if (
                            self._exp_thread is not None
                            or self._preview_thread is not None
                        ):
                            msg = (
                                "<b>Camera previewing/running experiments.</b><br><br>"
                                "(Consider halting preview/aborting experiments before disconnecting camera)"
                            )
                            btn = QtWidgets.QMessageBox.critical(
                                self,
                                "Preview/Experiments running",
                                msg,
                                QtWidgets.QMessageBox.Cancel
                                | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel,
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return

                        temp = self.camera.CCDTemperature
                        if temp < CAMERA_THERMAL_SHOCK_THRESHOLD_C:
                            msg = (
                                f"<b>Camera at {temp:.1f}°C. Disconnecting now may cause thermal shock.</b><br><br>"
                                f"(Consider warming the sensor above {CAMERA_THERMAL_SHOCK_THRESHOLD_C}°C before disconnecting camera)"
                            )
                            btn = QtWidgets.QMessageBox.warning(
                                self,
                                "Thermal Warning",
                                msg,
                                QtWidgets.QMessageBox.Cancel
                                | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel,
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return
                    except Exception:
                        pass

                    self._teardown_threads()

                    # self._poll_timer.stop()
                    self.log.log("Stopped temperature monitor", "#1dff35")

                    self.log.log(f"Disconnecting Camera: {self.camera.Name}", "orange")
                    self.camera.Connected = False
                    del self.camera
                    self.camera = None
                    self.log.log("Camera disconnected", "#1dff35")

                    self._current_temp_label.setText(
                        "Current Temperature: XX.x°C     State: X"
                    )

                    # Switch from disconnect -> connect
                    self._connect_button.setStyleSheet(self.GREEN_STYLE)
                    self._connect_button.setText("Connect Camera")
                    self._preview_run_buttons_row.setVisible(False)
                    self._cam_settings_frame.setVisible(False)

                    self.is_connected = False

            except Exception as e:
                self.log.log(f"Error disconnecting camera: {e}", "red")

    def _run_experiments_pressed(self):
        """
        Start or abort the experiment acquisition thread.
        """
        if self._preview_thread is not None:
            self.log.log("Run blocked: halt preview first.", "red")
            return
        btn = self._run_experiments_button

        if self._exp_thread is None:
            if not (self.camera and self.camera.Connected):
                self.log.log("Run failed: camera not connected", "red")
                return

            if not self._containers:
                self.log.log("Run failed: no experiments", "red")
                return

            if self._saver_thread is not None:
                self.log.log("Run blocked: save thread is still active", "red")
                return

            msg = (
                "<b>Have you set FITS header options as appropriate?</b><br><br>"
                "(Note: setting gain/read-speed etc. via camera settings does not "
                "update the FITS headers; these need to be updated by the user "
                "within each experiment settings block as appropriate.)"
            )

            btn_msg_box = QtWidgets.QMessageBox.warning(
                self,
                "FITS Header reminder",
                msg,
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.Yes,
            )

            if btn_msg_box == QtWidgets.QMessageBox.No:
                return

            experiments = [
                ExperimentSpec(
                    index=i,
                    title=title,
                    params=copy.deepcopy(params),
                )
                for i, title, params in self._get_experiments()
            ]

            if not experiments:
                self.log.log("Run failed: no valid experiments", "red")
                return

            self._active_experiment_count = len(experiments)

            self._saver_thread = SaverThread()
            self._saver_thread.progress.connect(self._log_saved_progress)
            self._saver_thread.log_error.connect(self.log.log)

            kwargs = {}

            if self.filterwheel_is_connected:
                kwargs["filterwheel"] = self.fw

            if self.mount_is_connected:
                kwargs["telescope"] = self.telescope
                kwargs["focuser"] = self.focuser

            self._exp_thread = ExperimentThread(
                self.camera,
                experiments,
                **kwargs,
            )

            self._exp_thread.frame.connect(self.imageviewer.display_array)
            self._exp_thread.status.connect(self._update_status)
            self._exp_thread.progress.connect(self._log_progress)
            self._exp_thread.log_requested.connect(self.log.log)
            self._exp_thread.save_requested.connect(self._saver_thread.save)
            self._exp_thread.error.connect(
                lambda e: self.log.log(f"Experiment error: {e}", "red")
            )
            self._exp_thread.sequence_finished.connect(self._experiments_finished)

            self._poll_timer.stop()
            self._saver_thread.start()
            self._exp_thread.start()

            btn.setStyleSheet(self.RED_STYLE)
            btn.setText("Abort Experiments")
            self.log.log("Experiments started", "#1dff35")

        else:
            self._active_experiment_count = 0
            self.log.log("Aborting experiments", "orange")

            if self._teardown_threads():
                self._poll_timer.start()
                btn.setStyleSheet(self.GREEN_STYLE)
                btn.setText("Run Experiments")
                self.log.log("Experiments aborted", "#1dff35")
            else:
                self.log.log(
                    "Abort incomplete: worker threads remain active.",
                    "red",
                )

    def _log_progress(
        self,
        exp_idx,
        exp_name,
        exp,
        params,
        exp_num,
        exp_rep_num,
        exp_reps,
        total_exp
    ):
        """
        Log the current image acquisition.

        Parameters
        ----------
        exp_idx : int
            Experiment index.
        exp_name : str 
            Experiment title.
        exp : float
            Exposure value in selected units.
        params : dict[str, object]
            Experiment parameter dict.
        exp_num : int
            Image index.
        exp_rep_num : int
            Experiment repeat index.
        exp_reps : int
            Total repeats.
        total_exp : int
            Total exposures per repeat.
        """
        total_experiments = self._active_experiment_count or len(self._containers)

        self.log.log(
            f"[Exp {exp_idx + 1}/{total_experiments}]"
            f"[{exp_rep_num}/{exp_reps}] {exp_name}: "
            f"Image {exp_num}/{total_exp}: {exp}{params.get('time_unit', '')}",
            "#1dff35",
        )

    def _log_saved_progress(self, path):
        self.log.log(f"Image saved to → {os.path.basename(path)}", "#1dff35")

    def _experiments_finished(self):
        """
        Handle normal completion of an experiment.
        """
        self._active_experiment_count = 0

        self._run_experiments_button.setStyleSheet(self.GREEN_STYLE)
        self._run_experiments_button.setText("Run Experiments")

        if self._exp_thread is not None:
            self._exp_thread.deleteLater()
            self._exp_thread = None

        if self._saver_thread is not None:
            if self._saver_thread.stop(15000):
                self._saver_thread.deleteLater()
                self._saver_thread = None
            else:
                self.log.log(
                    "Save thread did not stop cleanly (it remains active).",
                    "red",
                )

        self._poll_timer.start()
        self.log.log("All experiments complete", "#1dff35")

    def _preview_pressed(self):
        """
        Start the image preview thread.
        """
        if self._exp_thread is not None:
            self.log.log("Preview blocked: experiments are running.", "red")
            return
        if self._preview_button.styleSheet() == self.GREEN_STYLE:
            if not self.camera or not self.camera.Connected:
                self.log.log("Preview failed: camera not connected", "red")
                return
            if not self._containers:
                self.log.log("Preview failed: no experiments", "red")
                return
            p = self._containers[0].internalGui.window_params
            if not list(p.get("exp_array")):
                self.log.log("Preview failed: no exposures", "red")
                return
            self._preview_thread = PreviewThread(self.camera, p)
            self._preview_thread.frame.connect(self.imageviewer.display_array)
            self._preview_thread.status.connect(self._update_status)
            self._preview_thread.error.connect(
                lambda e: self.log.log(f"Preview error: {e}", "red")
            )
            self._poll_timer.stop()
            self._preview_thread.start()
            self._preview_button.setStyleSheet(self.RED_STYLE)
            self._preview_button.setText("Halt Preview")
            self.log.log("Preview started", "#1dff35")
        else:
            if self._preview_thread is not None:
                thread = self._preview_thread

                if not thread.isFinished():
                    thread.abort()
                    thread.quit()

                if thread.wait(2000):
                    thread.deleteLater()
                    self._preview_thread = None
                else:
                    self.log.log(
                        "Preview thread did not stop cleanly (it remains active).",
                        "red",
                    )
                    return

            if self.camera and self.camera.Connected:
                self._poll_timer.start()

            self._preview_button.setStyleSheet(self.GREEN_STYLE)
            self._preview_button.setText("Preview")
            self.log.log("Preview halted", "#1dff35")

    def _set_filter_target(self):
        """
        Set and log the filter target position.
        """
        if self.filterwheel_is_connected and not self.filterwheel_moving:
            if self._exp_thread is not None or self._preview_thread is not None:
                msg = (
                    "<b>Camera previewing/running experiments.</b><br><br>"
                    "(Consider halting preview/aborting experiments before changing filter)"
                )
                btn = QtWidgets.QMessageBox.critical(
                    self,
                    "Preview/Experiments running",
                    msg,
                    QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                    QtWidgets.QMessageBox.Cancel,
                )
                if btn == QtWidgets.QMessageBox.Ignore:
                    pass
                else:
                    return

            selected_position = self._filterwheel_combobox.currentText()
            selected_index = self._filterwheel_combobox.currentIndex() + 1
            self.log.log(f"Changing filter to {selected_position}", "orange")
            self._current_filter_label.setText("Position: Na     State: Moving")

            try:
                self.filterwheel_moving = True

                self._fw_thread = FilterwheelThread(self.fw, selected_index, self.fname)
                self._fw_thread.move_finished.connect(self._filter_move_finished)
                self._fw_thread.error.connect(self._filter_move_error)
                self._fw_thread.start()
            except Exception as e:
                self.log.log(
                    f"Error changing filter to {selected_position}: {e}", "red"
                )
                self.filterwheel_moving = False

    def _filter_move_finished(self, filterName):
        """
        Handle successful filter-wheel move completion.
        """
        self.filterwheel_moving = False
        selected_position = self._filterwheel_combobox.currentText()

        self.log.log(f"Filter changed to {selected_position}", "#1dff35")
        self._current_filter_label.setText(f"Position: {filterName}     State: Idle")

        if self._fw_thread is not None:
            self._fw_thread.deleteLater()
            self._fw_thread = None

    def _filter_move_error(self, err_msg):
        """
        Handle filter-wheel move failure.
        """
        self.filterwheel_moving = False
        self.log.log(f"Error changing filter: {err_msg}", "red")

        if self._fw_thread is not None:
            self._fw_thread.deleteLater()
            self._fw_thread = None

    def _wait_fw_thread(self, timeout_ms = 5000):
        """
        Wait for the filter-wheel thread to finish and clean it up.
        """
        thread = self._fw_thread

        if thread is None:
            self.filterwheel_moving = False
            return True

        if thread.wait(timeout_ms):
            try:
                thread.move_finished.disconnect(self._filter_move_finished)
                thread.error.disconnect(self._filter_move_error)
            except TypeError:
                pass

            thread.deleteLater()
            self._fw_thread = None
            self.filterwheel_moving = False
            return True

        self.log.log("Filter-wheel move did not finish before timeout", "orange")
        return False

    def _set_temp_target(self):
        """
        Set the cooler target temperature after validation.
        """
        raw = self._temp_target_line_edit.text().strip().replace(",", ".")

        try:
            target = float(raw)
        except ValueError:
            self.log.log("Invalid target temperature", "red")
            return

        if not CAMERA_MIN_TEMP_C <= target <= CAMERA_MAX_TEMP_C:
            self.log.log(
                f"Target temperature must be between {CAMERA_MIN_TEMP_C} and {CAMERA_MAX_TEMP_C} C",
                "red",
            )
            return

        self._temp_target = target

        try:
            if self.camera is not None and self.camera.Connected:
                self.camera.SetCCDTemperature = target

            self.log.log(f"Setting target temp. to {target:.1f}°C", "#1dff35")

            for container in self._containers:
                if container.internalGui is not None:
                    container.internalGui.set_temp_value(target)

        except Exception as exc:
            self.log.log(f"Error setting target temp.: {exc}", "red")

    def _open_camera_settings(self):
        """
        Open the camera ASCOM setup dialog and restore connection state.
        """
        if not self.camera:
            self.log.log("Cannot open camera settings, connect camera first", "red")
            return

        was_connected = False

        try:
            was_connected = bool(self.camera.Connected)

            if was_connected:
                self.log.log("Disconnecting camera while modifying settings", "orange")
                self.camera.Connected = False

            self.camera.SetupDialog()

        except Exception as e:
            self.log.log(f"Error opening camera setup: {e}", "red")

        finally:
            if self.camera is not None and was_connected:
                try:
                    self.log.log(
                        "Reconnecting camera after modifying settings", "orange"
                    )
                    self.camera.Connected = True
                except Exception as exc:
                    self.log.log(f"Error reconnecting camera: {exc}", "red")

    def _build_left_side(self):
        """
        Build the entire left half of the GUI (experiment settings windows, camera settings etc.).
        """
        # Master left half layout
        self._left_frame = QtWidgets.QFrame()
        self._left_layout = QtWidgets.QVBoxLayout(self._left_frame)
        self._left_layout.setSpacing(0)

        self._master_layout.addWidget(self._left_frame)

        # Experiment window number control panel.
        _control = QtWidgets.QFrame()
        _control.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        _control.setLineWidth(1)
        _control.setStyleSheet("QFrame {background-color: lightgrey}")
        _control.setToolTip(
            "Create the specified number of experiments"
            "\nand their associated settings windows below."
            "\nWindows can be renamed by double clicking the title."
        )
        _control_layout = QtWidgets.QHBoxLayout(_control)
        _control_layout.setContentsMargins(5, 5, 5, 5)
        _control_layout.addWidget(QtWidgets.QLabel("Experiments:"))

        self._spinbox = QtWidgets.QSpinBox()
        self._spinbox.setMinimum(1)
        self._spinbox.valueChanged.connect(self._update_containers)
        _control_layout.addWidget(self._spinbox)

        # Save / Load configuration buttons.
        self._save_config_button = QtWidgets.QPushButton("Save")
        self._save_config_button.setToolTip(
            "Save the current experiment profile to a JSON file."
        )
        self._save_config_button.clicked.connect(self._save_config)
        _control_layout.addWidget(self._save_config_button)

        self._load_config_button = QtWidgets.QPushButton("Load")
        self._load_config_button.setToolTip(
            "Load a saved experiment profile from a JSON file."
        )
        self._load_config_button.clicked.connect(self._load_config)
        _control_layout.addWidget(self._load_config_button)

        self._left_layout.addWidget(_control)

        # Scrollable container area.
        self._scroll_area = QtWidgets.QScrollArea()
        self._scroll_area.setFixedSize(400, 600)
        self._scroll_area.setWidgetResizable(True)
        self._left_layout.addWidget(self._scroll_area)

        self._container_widget = QtWidgets.QWidget()
        self._scroll_area.setWidget(self._container_widget)
        self._container_layout = QtWidgets.QVBoxLayout(self._container_widget)
        self._container_layout.setContentsMargins(0, 0, 0, 0)
        self._container_layout.setSpacing(0)
        self._container_layout.addStretch()

        # Dummy button for testing retrieval of experiment settings:
        # Dumps all experiment settings to the log widget.
        # _grab_contents = QtWidgets.QPushButton('Grab Contents')
        # _grab_contents.clicked.connect(self._grab_contents_fun)
        # self._left_layout.addWidget(_grab_contents)

        self._containers = []
        self._update_containers()

        # BUILD TEMPERATURE SECTION
        self._current_temp_label = QtWidgets.QLabel(
            "Current Temperature: XX.x°C     State: X"
        )
        self._temp_target_line_edit = QtWidgets.QLineEdit(f"{CAMERA_MIN_TEMP_C}")
        self._temp_target_line_edit.setValidator(
            QtGui.QDoubleValidator(
                bottom=CAMERA_MIN_TEMP_C, top=CAMERA_MAX_TEMP_C, decimals=1
            )
        )
        self._set_temp_target_button = QtWidgets.QPushButton("Set Target")
        self._set_temp_target_button.clicked.connect(self._set_temp_target)
        self._camera_settings_button = QtWidgets.QPushButton("Camera Settings")
        self._camera_settings_button.clicked.connect(self._open_camera_settings)
        self._cam_settings_row = CollapsibleWidgetInternalGUI._create_row(
            widget=[
                self._temp_target_line_edit,
                self._set_temp_target_button,
                self._camera_settings_button,
            ]
        )
        self._cam_settings_frame = CollapsibleWidgetInternalGUI._create_frame(
            widgets=[self._current_temp_label, self._cam_settings_row]
        )
        self._cam_settings_frame.setFrameStyle(
            QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised
        )
        self._cam_settings_frame.setLineWidth(1)
        self._cam_settings_frame.setStyleSheet("QFrame {background-color: lightgrey}")

        # self._left_layout.addWidget(self._cam_settings_frame)

        # BUILD CAMERA CONNECT/ACTIONS SECTION
        self._connect_button = QtWidgets.QPushButton("Connect Camera")
        self._connect_button.setStyleSheet(self.GREEN_STYLE)
        self._connect_button.clicked.connect(self._connect_pressed)

        self._run_experiments_button = QtWidgets.QPushButton("Run Experiments")
        self._run_experiments_button.setStyleSheet(self.GREEN_STYLE)
        self._run_experiments_button.clicked.connect(self._run_experiments_pressed)

        self._preview_button = QtWidgets.QPushButton("Preview")
        self._preview_button.setStyleSheet(self.GREEN_STYLE)
        self._preview_button.clicked.connect(self._preview_pressed)

        self._preview_run_buttons_row = CollapsibleWidgetInternalGUI._create_row(
            widget=[self._run_experiments_button, self._preview_button], spacing=5
        )

        self._cam_actions_frame = CollapsibleWidgetInternalGUI._create_frame(
            widgets=[
                self._connect_button,
                self._cam_settings_frame,
                self._preview_run_buttons_row,
            ]
        )
        self._cam_settings_frame.setVisible(False)
        self._preview_run_buttons_row.setVisible(False)
        self._cam_actions_frame.setFrameStyle(
            QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised
        )
        self._cam_actions_frame.setLineWidth(1)
        self._cam_actions_frame.setStyleSheet("QFrame {background-color: lightgrey}")

        self._left_layout.addWidget(self._cam_actions_frame)
        # self._left_layout.addWidget(self._cam_settings_frame)

        # BUILD MOUNT CONNECT SECTION
        self._mount_connect_button = QtWidgets.QPushButton("Connect Mount/Focuser")
        self._mount_connect_button.setStyleSheet(self.GREEN_STYLE)
        self._mount_connect_button.clicked.connect(self._connect_mount_pressed)

        self.mount_frame = CollapsibleWidgetInternalGUI._create_frame(
            widgets=[self._mount_connect_button]
        )
        self.mount_frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        self.mount_frame.setLineWidth(1)
        self.mount_frame.setStyleSheet("QFrame {background-color: lightgrey}")

        self._left_layout.addWidget(self.mount_frame)

        # BUILD FILTER WHEEL SECTION
        self._filterwheel_connect_button = QtWidgets.QPushButton("Connect Filter-wheel")
        self._filterwheel_connect_button.setStyleSheet(self.GREEN_STYLE)
        self._filterwheel_connect_button.clicked.connect(
            self._connect_filterwheel_pressed
        )

        self._current_filter_label = QtWidgets.QLabel("Current Filter: X     State: X")
        self._filterwheel_combobox = QtWidgets.QComboBox()
        self._filterwheel_combobox.addItems(CFW10_FILTER_NAMES_LONG)
        self._set_filter_target_button = QtWidgets.QPushButton("Set Target")
        self._set_filter_target_button.clicked.connect(self._set_filter_target)
        self._filter_settings_row = CollapsibleWidgetInternalGUI._create_row(
            widget=[self._filterwheel_combobox, self._set_filter_target_button]
        )

        self.filterwheel_frame = CollapsibleWidgetInternalGUI._create_frame(
            widgets=[
                self._filterwheel_connect_button,
                self._current_filter_label,
                self._filter_settings_row,
            ]
        )
        self.filterwheel_frame.setFrameStyle(
            QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised
        )
        self.filterwheel_frame.setLineWidth(1)
        self.filterwheel_frame.setStyleSheet("QFrame {background-color: lightgrey}")
        self._filter_settings_row.setVisible(False)
        self._current_filter_label.setVisible(False)

        self._left_layout.addWidget(self.filterwheel_frame)

        self._left_layout.addStretch()

    def _build_right_side(self):
        """
        Build the entire right half of the GUI (image display/statistics and log).
        """
        # Master right half layout
        self._right_frame = QtWidgets.QFrame()
        self._right_layout = QtWidgets.QVBoxLayout(self._right_frame)
        self._master_layout.addWidget(self._right_frame)

        self.log = LogWidget()
        self.imageviewer = ImageViewerWidget()

        # Add widgets to the layout
        self._right_layout.addWidget(self.imageviewer)
        self._right_layout.addWidget(self.log)

    def _teardown_threads(self):
        """
        Abort and clean up worker threads.
        """
        ok = True

        if self._preview_thread is not None:
            self._preview_thread.abort()
            self._preview_thread.quit()

            if self._preview_thread.wait(5000):
                self._preview_thread.deleteLater()
                self._preview_thread = None
            else:
                self.log.log(
                    "Preview thread did not stop (it remains active).",
                    "orange",
                )
                ok = False

        if self._exp_thread is not None:
            self._exp_thread.abort()
            self._exp_thread.quit()

            if self._exp_thread.wait(5000):
                self._exp_thread.deleteLater()
                self._exp_thread = None
            else:
                self.log.log(
                    "Experiment thread did not stop (it remains active).",
                    "orange",
                )
                ok = False

        if self._saver_thread is not None:
            if self._saver_thread.stop(5000):
                self._saver_thread.deleteLater()
                self._saver_thread = None
            else:
                self.log.log(
                    "Save thread did not stop (it remains active).",
                    "orange",
                )
                ok = False

        if self._fw_thread is not None:
            if not self._wait_fw_thread(2000):
                self.log.log(
                    "Filter-wheel thread did not stop (it remains active).",
                    "orange",
                )
                ok = False

        self._poll_timer.stop()
        return ok

    def closeEvent(self, event):
        """
        Close the application safely, refusing to destroy active threads.
        """
        if self._fw_thread is not None:
            msg = (
                "<b>Filter wheel is moving.</b><br><br>"
                "Closing now can corrupt serial communication."
            )
            btn = QtWidgets.QMessageBox.critical(
                self,
                "Filter-wheel moving",
                msg,
                QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                QtWidgets.QMessageBox.Cancel,
            )

            if btn == QtWidgets.QMessageBox.Cancel:
                event.ignore()
                return

            if not self._wait_fw_thread(15000):
                event.ignore()
                return

        cam = self.camera

        if cam and cam.Connected:
            try:
                if self._exp_thread is not None or self._preview_thread is not None:
                    msg = (
                        "<b>Camera previewing/running experiments.</b><br><br>"
                        "Consider halting preview/aborting experiments before closing."
                    )
                    btn = QtWidgets.QMessageBox.critical(
                        self,
                        "Preview/Experiments running",
                        msg,
                        QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                        QtWidgets.QMessageBox.Cancel,
                    )

                    if btn == QtWidgets.QMessageBox.Ignore:
                        self._teardown_threads()
                    else:
                        event.ignore()
                        return

                temp = cam.CCDTemperature
                if temp < CAMERA_THERMAL_SHOCK_THRESHOLD_C:
                    msg = (
                        f"<b>Camera at {temp:.1f}°C. Closing now may cause thermal shock.</b><br><br>"
                        f"(Consider warming the sensor above {CAMERA_THERMAL_SHOCK_THRESHOLD_C}°C before closing)"
                    )
                    btn = QtWidgets.QMessageBox.warning(
                        self,
                        "Thermal Warning",
                        msg,
                        QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                        QtWidgets.QMessageBox.Cancel,
                    )

                    if btn == QtWidgets.QMessageBox.Ignore:
                        self._teardown_threads()
                    else:
                        event.ignore()
                        return

            except Exception as exc:
                self.log.log(f"Error checking temperature on close: {exc}", "red")

        if self._saver_thread is not None:
            if self._saver_thread.stop(5000):
                self._saver_thread.deleteLater()
                self._saver_thread = None
            else:
                self.log.log("Cannot close: save thread is still active.", "red")
                event.ignore()
                return

        if (
            self._preview_thread is not None
            or self._exp_thread is not None
            or self._fw_thread is not None
            or self._saver_thread is not None
        ):
            self.log.log("Cannot close: worker threads are still active.", "red")
            event.ignore()
            return

        event.accept()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
