from PyQt5 import QtWidgets, QtGui, QtCore
from astropy.io import fits
from queue import Queue
import sys
import os
import numpy as np
import time
import copy
from datetime import datetime, timezone
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
#import win32com.client
#import pythoncom
from alpaca.camera import *

class CollapsibleWidgetInternalGUI(QtWidgets.QWidget):
    """ Create internal GUI elements for a custom collapsible window widget.
        Called and placed like any other QWidget. Added to the internal
        containers of the collapsiblewidget's layout.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)

        # Instantiate the dict of params.
        self.window_params = {'gain': 'High',
                              'rospeed': '3.125MHz',
                              'start_x': 0,
                              'start_y': 0,
                              'width': 640,
                              'height': 512,
                              'bin_x': 1,
                              'bin_y': 1,
                              'exp_array': [1],
                              'time_unit': 's',
                              'nameprefix': 'image',
                              'savedirectory': os.path.expanduser('~\Desktop'),
                              'name_date_time': '',
                              'name_gain_rospeed': '',
                              'name_temp': '',
                              'name_exp_time': '',
                              'include_date_time': False,
                              'include_gain_rospeed': False,
                              'include_temp': False,
                              'include_exp_time': False,
                              'delay': 1.000,
                              'repeats': 1}
        
        # Window master layout.
        self._layout = QtWidgets.QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)

        # (Vertically & sequentially) Place all section frames within the master layout.
        self._layout.addWidget(self._create_fits_settings_section())
        self._layout.addWidget(self._create_roi_settings_section())
        self._layout.addWidget(self._create_file_settings_section())
        self._layout.addWidget(self._create_exposure_settings_section())
        self._layout.addWidget(self._create_delay_setting_section())
        self._layout.addWidget(self._create_experiment_repeat_setting_section())
        self._layout.addStretch()
        

    ######################
    # GUI FRAME SECTIONS #
    ######################

    def _create_fits_settings_section(self):
        """ Frame containing Gain & R/O speed settings.
        """
        # Gain Setting (High/Low):
        # For .fits header book-keeping information.
        self._gain_combobox = QtWidgets.QComboBox()
        self._gain_combobox.addItems(['High','Low'])
        # Use the custom create row helper function
        self._gain_widget = self._create_row('Gain Setting (for FITS header):', self._gain_combobox, 10)
        self._gain_widget.setToolTip('Set gain setting for saving to FITS header (Default High).')
        self._gain_combobox.currentTextChanged.connect(self._gain_changed)
        
        # Read-out (R/O) speed Setting (3.125MHz/12.5MHz/25MHz):
        # For .fits header book-keeping information.
        self._rospeed_combobox = QtWidgets.QComboBox()
        self._rospeed_combobox.addItems(['3.125MHz','12.5MHz','25MHz'])
        # Use the custom create row helper function.
        self._rospeed_widget = self._create_row('R/O Speed setting (for FITS header):', self._rospeed_combobox, 10)
        self._rospeed_widget.setToolTip('Set current read-out speed for saving to FITS header (Default 3.125MHz).')
        self._rospeed_combobox.currentTextChanged.connect(self._rospeed_changed)

        # Return the custom create frame helper function.
        return self._create_frame([self._gain_widget, self._create_separator(), self._rospeed_widget])
        
    def _create_roi_settings_section(self):
        """ Frame containing ROI settings.
        """
        # Region-of-interest (ROI) settings: 
        # For defining custom imaging sub-regions on the sensor.
        self._roi_button = QtWidgets.QPushButton('Set ROI')
        self._roi_button.clicked.connect(self._open_ROIWindow)
        # Use the custom create row helper function.
        self._roi_widget = self._create_row('ROI settings:', self._roi_button, 10)
        self._roi_widget.setToolTip('Open the ROI settings window. Allows changing:\n'
                                    '- binning (not implemented)\n'
                                    '- imaging sub-region (ROI)')
        
        # Current ROI label:
        self._roi_label = QtWidgets.QLabel('ROI: X=0->640, Y=0->512, 640x512 (Bin 1x1)')
        # Use the custom create row helper function.
        self._roi_label_row = self._create_row(widget=self._roi_label)

        # Return the custom create frame helper function
        return self._create_frame([self._roi_widget, self._create_separator(), self._roi_label_row])
        
    def _create_file_settings_section(self):
        """ Frame containing file naming & save location settings.
        """
        # Image name prefix: 
        # For defining the image file name prefix.
        self._image_name_line_edit = QtWidgets.QLineEdit('image')
        self._image_name_line_edit.textChanged.connect(self._file_prefix_changed)
        # Use the custom create row helper function.
        self._image_name_widget = self._create_row('Image name prefix:', self._image_name_line_edit, 10)
        self._image_name_widget.setToolTip('Set the FITS image file name prefix')

        
        # File save directory dialog: 
        # For setting the file directory to save to.
        self._file_dialog_button = QtWidgets.QPushButton('File Directory')
        self._file_dialog_button.clicked.connect(self._open_file_dialog)
        # Use the custom create row helper function.
        self._file_directory_widget = self._create_row('Save to:', self._file_dialog_button, 10)
        self._file_directory_widget.setToolTip('Open file explorer and set directory to save FITS images to.')
        
        # Various file name settings:
        self._file_date_checkbox = QtWidgets.QCheckBox('Append Date && Time')
        self._file_date_checkbox.stateChanged.connect(self._append_date_time)
        self._file_gain_rospeed_checkbox = QtWidgets.QCheckBox('Append Gain && R/O')
        self._file_gain_rospeed_checkbox.stateChanged.connect(self._append_gain_rospeed)
        self._file_temp_checkbox = QtWidgets.QCheckBox('Append Temp.')
        self._file_temp_checkbox.stateChanged.connect(self._append_temp)
        self._file_exp_time_checkbox = QtWidgets.QCheckBox('Append Exp. time')
        self._file_exp_time_checkbox.stateChanged.connect(self._append_exp_time)
        # Use the custom create row helper function.
        self._file_checkboxes_row_one_widget = self._create_row(widget=[self._file_date_checkbox,self._file_gain_rospeed_checkbox], spacing=10)
        self._file_checkboxes_row_two_widget = self._create_row(widget=[self._file_temp_checkbox,self._file_exp_time_checkbox], spacing=10)
        
         
        # Current directory label:
        self._directory_label = QtWidgets.QLabel('Folder: ' + repr(self.window_params['savedirectory'])[1:-1])
        # Use the custom create row helper function.
        self._directory_row = self._create_row(widget=self._directory_label)
            
        
        # Example file name label:
        self._file_name_label = QtWidgets.QLabel('Example: ' + self.window_params['nameprefix'] + '_1.fits')
        # Use the custom create row helper function.
        self._file_name_row = self._create_row(widget=self._file_name_label)

        # Return the custom create frame helper function.
        return self._create_frame([self._image_name_widget, self._create_separator(), 
                                                        self._file_directory_widget, self._create_separator(),
                                                        self._file_checkboxes_row_one_widget, 
                                                        self._file_checkboxes_row_two_widget,
                                                        self._create_separator(), self._directory_row, 
                                                        self._file_name_row])
    
    def _create_exposure_settings_section(self):
        """ Frame containing exposure time unit & array settings.
        """
        # Exposure times: 
        # For defining the array of exposures.
        self._exposure_times_line_edit = QtWidgets.QLineEdit('[1]')
        self._exposure_times_line_edit.textChanged.connect(self._exposure_string_changed)
        self._exposure_times_widget = self._create_row('Array of exposures:', self._exposure_times_line_edit, 10)
        self._exposure_times_widget.setToolTip('Define the exposure array: \n'
                                               'Evaluates the input as python code i.e.:\n'
                                               '- [1] is a single 1 time unit exposure.\n'
                                               '- np.repeat([1],20) is 20 1 time unit exposures.\n'
                                               '- np.repeat(np.arange(1,4),10) is 10 sets of [1,2,3] time unit exposures.')
        
        # Time unit: 
        # For defining the time unit for exposures (i.e. s/ms/μs).
        # Also for setting the currently defined exposure array.
        self._time_unit_combobox = QtWidgets.QComboBox()
        self._time_unit_combobox.addItems(['s','ms','μs'])
        self._set_exposure_button = QtWidgets.QPushButton('Set Exposure')
        self._set_exposure_button.clicked.connect(self._set_exposure)
        self._time_unit_widget = self._create_row('Time unit:', [self._time_unit_combobox, self._set_exposure_button], 10)
        self._time_unit_widget.setToolTip('Set the time unit for exposures (s/ms/μs).')
        self._time_unit_combobox.currentTextChanged.connect(self._time_unit_changed)
    
        
        
        # Current exposure array label:
        self._num_exposures_label = QtWidgets.QLabel('1 exposure:')
        # Use the custom create row helper function.
        self._num_exposures_row = self._create_row(widget=self._num_exposures_label, spacing=10)

        self._exposure_array_label = QtWidgets.QLabel('[1]s')
        # Use the custom create row helper function.
        self._exposure_array_row = self._create_row(widget=self._exposure_array_label, spacing=10)
        
        # Return the custom create frame helper function.
        return self._create_frame([self._exposure_times_widget, self._time_unit_widget, self._create_separator(),
                                                            self._num_exposures_row, self._exposure_array_row])
        
    def _create_delay_setting_section(self):
        """ Frame containing delay b/w experiments setting.
        """
        # Experiment delay: 
        # For setting the delay before an experiment begins.
        # Useful for spacing subsequent experiments.
        self._delay_line_edit = QtWidgets.QLineEdit('1.000')
        self._delay_line_edit.setValidator(QtGui.QDoubleValidator(bottom=0.0, decimals=3))
        self._delay_line_edit.textChanged.connect(self._delay_changed)
        # Use the custom create row helper function.
        self._delay_widget = self._create_row('Delay before start (s):', self._delay_line_edit, 10)
        self._delay_widget.setToolTip('Delay in seconds before beginning this experiment block.')
        
        # Return the custom create frame helper function.
        return self._create_frame([self._delay_widget])

    def _create_experiment_repeat_setting_section(self):
        """ Frame containing experiment repeat setting.
        """
        # Experiment repeats: 
        # For defining the number of times to repeat this experiment set.
        self._experiment_repeat_line_edit = QtWidgets.QLineEdit('1')
        self._experiment_repeat_line_edit.setValidator(QtGui.QIntValidator(bottom=1))
        self._experiment_repeat_line_edit.textChanged.connect(self._experiment_repeat_changed)
        # Use the custom create row helper function.
        self._experiment_repeat_widget = self._create_row('Experiment repeats:', self._experiment_repeat_line_edit, 10)
        self._experiment_repeat_widget.setToolTip('Specify how many times to repeat the experiment block.')

        # Return the custom create frame helper function.
        return self._create_frame([self._experiment_repeat_widget])
    

    ##########################
    # WIDGET LOGIC FUNCTIONS #
    ##########################

    def _gain_changed(self, text):
        """ Logic function for the gain dropdown.
        """
        self.window_params['gain'] = text
        self._append_gain_rospeed()

    def _rospeed_changed(self, text):
        """ Logic function for the read-out speed dropdown.
        """
        self.window_params['rospeed'] = text
        self._append_gain_rospeed()
    
    def _exposure_string_changed(self, text):
        """ Logic function for the exposure array line-edit.
        """
        self._exposure_string = text
    
    def _set_exposure(self):
        """ Logic function for setting the exposure array using eval (potential
            for injection of dangerous code here, currently restricted to just
            base python modules + numpy as np).
        """
        exposure_string = self._exposure_times_line_edit.text()
        # Define and apply the safe module whitelist.
        safe_dict = {'np': np}
        self.window_params['exp_array'] = eval(exposure_string, {"__builtins__": {}}, safe_dict)
        # Ensure correct grammar for the text!
        if len(self.window_params['exp_array']) == 1:
            self._num_exposures_label.setText(str(len(self.window_params['exp_array'])) + ' exposure:')
            self._exposure_array_label.setText(str(self.window_params['exp_array'])+str(self.window_params['time_unit']))
        else:
            # More than 35 exposures in the array can overflow the text label. 
            # Therefore, restrict the label to 35 max elements plus a "..."
            if len(self.window_params['exp_array']) > 35:
                _str = '['
                for i in range(16):
                    _str += ' '+str(self.window_params['exp_array'][i])
                _str += '  .... '
                for i in range(len(self.window_params['exp_array'])-16, len(self.window_params['exp_array'])):
                    _str += ' '+str(self.window_params['exp_array'][i])
                _str += ']'
                self._exposure_array_label.setText(_str+str(self.window_params['time_unit']))
            else:
                self._exposure_array_label.setText(str(self.window_params['exp_array'])+str(self.window_params['time_unit']))
            self._num_exposures_label.setText(str(len(self.window_params['exp_array'])) + ' exposures:')
        self._append_exp_time()

    def _time_unit_changed(self, text):
        """ Logic function for the time unit dropdown.
        """
        self.window_params['time_unit'] = text

    def _delay_changed(self, text):
        """ Logic function for the experiment delay line-edit.
        """
        self.window_params['delay'] = text

    def _experiment_repeat_changed(self, text):
        """ Logic function for the experiment repeat line-edit.
        """
        self.window_params['repeats'] = text
    
    def _file_prefix_changed(self, text):
        """ Logic function for the file name prefix line-edit.
        """
        self.window_params['nameprefix'] = text
        self._update_example_filename()
        
    def _open_file_dialog(self):
        """ Logic function for the file directory dialog.
        """
        self.window_params['savedirectory'] = QtWidgets.QFileDialog.getExistingDirectory()
        self._directory_label.setText('Folder: ' + repr(self.window_params['savedirectory'])[1:-1])
        
    def _append_date_time(self):
        """ Logic function for the append date/time checkbox.
        """
        state = self._file_date_checkbox.checkState()
        if state == 2:
            ts = datetime.now()
            if int(ts.strftime("%H")) >= 12:
                self.window_params['name_date_time'] = datetime.now().strftime('_%d_%m_%Y_%H-%Mpm')
            else:
                self.window_params['name_date_time'] = datetime.now().strftime('_%d_%m_%Y_%H-%Mam')
            self.window_params['include_date_time'] = True
        else:
            self.window_params['name_date_time'] = ''
            self.window_params['include_date_time'] = False
        self._update_example_filename()
            
    def _append_gain_rospeed(self):
        """ Logic function for the append gain/read-out speed checkbox.
        """
        state = self._file_gain_rospeed_checkbox.checkState()
        if state == 2:
            self.window_params['name_gain_rospeed'] = '_'+ self.window_params['gain'] + '_' + self.window_params['rospeed']
            self.window_params['include_gain_rospeed'] = True
        else:
            self.window_params['name_gain_rospeed'] = ''
            self.window_params['include_gain_rospeed'] = False
        self._update_example_filename()
        
    def _append_temp(self):
        """ Logic function for the append temperature checkbox.
        """
        state = self._file_temp_checkbox.checkState()
        if state == 2:
            self.window_params['name_temp'] = '_-55C'
            self.window_params['include_temp'] = True
        else:
            self.window_params['name_temp'] = ''
            self.window_params['include_temp'] = False
        self._update_example_filename()
            
    def _append_exp_time(self):
        """ Logic function for the append exposure time checkbox.
        """
        state = self._file_exp_time_checkbox.checkState()
        if state == 2:
            self.window_params['name_exp_time'] = '_'+str(self.window_params['exp_array'][0])+self.window_params['time_unit']
            self.window_params['include_exp_time'] = True
        else:
            self.window_params['name_exp_time'] = ''
            self.window_params['include_exp_time'] = False
        self._update_example_filename()
        
    def _update_example_filename(self):
        """ Combine all file name elements for the example file name.
        """
        self._example_file_name_value = ('Example: '
                                         + self.window_params['nameprefix']
                                         + self.window_params['name_date_time']
                                         + self.window_params['name_gain_rospeed']
                                         + self.window_params['name_temp']
                                         + self.window_params['name_exp_time']
                                         + '_1.fits')
        self._file_name_label.setText(self._example_file_name_value)
    
    def _open_ROIWindow(self):
        """ Open the ROI settings window, instantiating the ROIWindow class and
            saving all variables once closed to the window params dict.
        """
        dialog = ROIWindow(self)
        dialog.exec_()
        if dialog.data:
            d = dialog.data
            self.window_params['start_x'] = d['start_x']
            self.window_params['start_y'] = d['start_y']
            self.window_params['width'] = d['width']
            self.window_params['height'] = d['height']
            self.window_params['bin_x'] = d['bin_x']
            self.window_params['bin_y'] = d['bin_y']
            self._roi_label.setText(f'ROI: X={int(d["start_x"]/d["bin_x"])}->{int((d["start_x"]+d["width"])/d["bin_x"])},'
                                    f' Y={int(d["start_y"]/d["bin_y"])}->{int((d["start_y"]+d["height"])/d["bin_y"])},'
                                    f'  {int(d["width"]/d["bin_x"])}x{int(d["height"]/d["bin_y"])} (Bin {d["bin_x"]}x{d["bin_y"]})')


    ####################
    # HELPER FUNCTIONS #
    ####################

    @staticmethod
    def _create_row(label_text=None, widget=None, spacing=0):
        """Add widgets side-by-side to a QtWidgets QHBoxLayout.

        Args:
            label_text (str): A row label widget preceeding the 
                              functional widgets.
                              
            widget (QtWidgets.QWidget): The list of widgets to 
                                        add.

            spacing (float): Defines the vertical spacing of the
                             row.

        Returns:
            row (QtWidgets.QWidget): The full QHBoxLayout as a 
                                     placable QWidget.
        """
        row = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(row)
        if label_text:
            layout.addWidget(QtWidgets.QLabel(label_text))
        if widget:
            # Logic for if widget contains 1 or >1 QWidget.
            widgets = widget if isinstance(widget, (list, tuple)) else (widget,)
            for w in widgets:
                layout.addWidget(w)
        layout.addStretch()
        layout.setContentsMargins(2, 2, 2, 2)
        # Tweakable vertical spacing.
        layout.setSpacing(spacing)
        return row

    @staticmethod
    def _create_frame(widgets, spacing=0):
        """ Add widgets vertically & sequentially to a QtWidgets.QFrame
            with a QVBoxLayout.

        Args:
            widgets (list): The list of widgets to add.

            spacing (float): Defines the vertical spacing of the frame.

        Returns:
            frame (QtWidgets.QFrame): The full QVBoxLayout as a placable 
                                      QFrame.
        """
        frame = QtWidgets.QFrame()
        # Define the styling options.
        frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        frame.setLineWidth(1)
        layout = QtWidgets.QVBoxLayout(frame)
        # Tweakable vertical spacing.
        layout.setSpacing(spacing)
        layout.setContentsMargins(2, 2, 2, 2)
        # Loop through the widgets and add.
        for widget in widgets:
            layout.addWidget(widget)
        return frame
       
    @staticmethod
    def _create_separator():
        """ Create a stylised visual separating line from a placable QFrame.

        Returns:
            sep (QtWidgets.QFrame): The stylised separator as a placable
                                    QFrame.
        """
        # Define and stylise a QFrame.
        sep = QtWidgets.QFrame()
        sep.setFrameShape(QtWidgets.QFrame.HLine)
        sep.setFrameShadow(QtWidgets.QFrame.Sunken)
        return sep
    
        
class CollapsibleWidget(QtWidgets.QWidget):
    """ Create the container for a custom collapsible window widget.
        An instance of CollapsibleWidgetInternalGUI is added to the container layout
        as a QWidget.
    """
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        
        # Master layout.
        self._layout = QtWidgets.QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)

        # Header.
        self._header = QtWidgets.QWidget()
        self._header.setObjectName("collapsibleHeader")
        self._header.setStyleSheet(
            "#collapsibleHeader {background-color: lightgrey; border: 1px solid grey; "
            "border-radius: 1px;}"
        )
        header_layout = QtWidgets.QHBoxLayout(self._header)
        header_layout.setContentsMargins(5, 5, 5, 5)

        # Open/Closed icon.
        self._icon = QtWidgets.QLabel("▼")
        self._icon.setFixedWidth(16)
        header_layout.addWidget(self._icon)

        # (Double click) Editable title.
        self._title_label = QtWidgets.QLabel(title)
        self._title_label.setObjectName("titleLabel")
        self._title_label.setStyleSheet("#titleLabel {background-color: transparent;}")
        self._title_label.mouseDoubleClickEvent = lambda e: self._start_edit()

        self._title_edit = QtWidgets.QLineEdit(title)
        self._title_edit.hide()
        self._title_edit.returnPressed.connect(self._finish_edit)
        self._title_edit.editingFinished.connect(self._finish_edit)

        # Add widgets.
        header_layout.addWidget(self._title_label)
        header_layout.addWidget(self._title_edit)
        header_layout.addStretch()

        # Add the toggle on click event.
        self._header.mousePressEvent = lambda e: self.toggle()

        # Define and stylise the content.
        self._content = QtWidgets.QWidget()
        self._content.setObjectName("collapsibleContent")
        self._content.setStyleSheet(
            "#collapsibleContent {background-color: lightgrey; border: 1px solid grey; "
            "margin: 0 0px 0px 0px;}"
        )
        
        # Add widgets.
        self._layout.addWidget(self._header)
        self._layout.addWidget(self._content)

    def _start_edit(self):
        """ Start editing the title.
        """
        self._title_label.hide()
        self._title_edit.setText(self._title_label.text())
        self._title_edit.show()
        self._title_edit.setFocus()
        self._title_edit.selectAll()

    def _finish_edit(self):
        """ Finish editing and update the title.
        """
        self._title_edit.hide()
        self._title_label.setText(self._title_edit.text())
        self._title_label.show()

    def toggle(self):
        """ Toggle content visibility.
        """
        visible = not self._content.isVisible()
        self._content.setVisible(visible)
        self._icon.setText("▼" if visible else "▶")

    @property
    def contentWidget(self):
        """ Content widget for the child layouts.
        """
        return self._content

    @property
    def title(self):
        """ Current title text.
        """
        return self._title_label.text()
    
    @property
    def internalGui(self):
        """ Access the embedded CollapsibleWidgetInternalGUI instance.
        """
        # Get the internal layout and return.
        layout = self._content.layout()
        return layout.itemAt(0).widget() if layout and layout.count() else None


class ROIWidget(QtWidgets.QGraphicsView):
    """ Interactive graphics display for the ROI selection. Added to the
        ROI settings window in the same manner as a QWidget. 
    """
    
    # Define the inter-thread signal to be sent (4 different ints).
    roi_changed = QtCore.pyqtSignal(int, int, int, int)
    
    def __init__(self, width=640, height=512):
        super().__init__()
        self._w, self._h = width, height
        left, top, right, bottom = 30, 5, 35, 25
        
        # Turn off the panning scroll-bars.
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        
        # Initialise the graphics scene.
        self._scene = QtWidgets.QGraphicsScene(-left, -top, width + left + right, height + top + bottom)
        self.setScene(self._scene)
        
        buffer = 4
        self.setFixedSize(width + left + right + buffer, height + top + bottom + buffer)
        self.setRenderHint(QtGui.QPainter.Antialiasing)
        
        self._rect = None
        self._handles = []
        self._state = None
        self._offset = QtCore.QPointF()
        
        # Draw the gridlines.
        self._draw_grid()
        self._create_roi(0, 0, width, height)
        
    def _draw_grid(self):
        """ Create grid-lines for the plot area of the ROI window
            using QtGui elements.
        """
        pen_minor = QtGui.QPen(QtCore.Qt.gray, 0.5)
        pen_minor.setStyle(QtCore.Qt.DashLine)
        pen_major = QtGui.QPen(QtCore.Qt.black, 1)
        pen_major.setStyle(QtCore.Qt.DashLine)
        pen_solid = QtGui.QPen(QtCore.Qt.black, 1)
        
        # Dashed internal grid lines.
        for x in range(0, self._w + 1, 16):
            self._scene.addLine(x, 0, x, self._h, pen_major if x % 32 == 0 else pen_minor)
        for y in range(0, self._h + 1, 16):
            self._scene.addLine(0, y, self._w, y, pen_major if y % 32 == 0 else pen_minor)
            
        # Solid outer boundary.
        self._scene.addRect(0, 0, self._w, self._h, pen_solid)
        
        # Label ticks.
        font = QtGui.QFont('Arial', 8)
        for i in range(0, self._w + 1, 32):
            self._scene.addLine(i, self._h, i, self._h + 3, pen_solid)
            t = self._scene.addText(str(i), font)
            t.setPos(i - 10, self._h + 5)
        for i in range(0, self._h + 1, 32):
            self._scene.addLine(0, i, -3, i, pen_solid)
            t = self._scene.addText(str(i), font)
            t.setPos(-25, i - 5)
            
    def _create_roi(self, x, y, w, h):
        """ Create a new ROI setting (removes and recreates the ROI region indicator on the graphics 
            scene).
        """
        # Clear the graphics scene of the previous region indicator.
        if self._rect:
            self._scene.removeItem(self._rect)
            for handle in self._handles:
                self._scene.removeItem(handle)
        # Add a new region indicator corresponding to the new ROI.
        self._rect = self._scene.addRect(x, y, w, h, QtGui.QPen(QtCore.Qt.green, 2), 
                                        QtGui.QBrush(QtGui.QColor(128, 128, 128, 51)))
        self._rect.setZValue(10)

        # Create the resize handles on the indicator.
        self._handles = []
        pos = [(x, y), (x + w/2, y), (x + w, y), (x, y + h/2), 
               (x + w, y + h/2), (x, y + h), (x + w/2, y + h), (x + w, y + h)]
        for px, py in pos:
            self._handles.append(self._scene.addRect(px - 3, py - 3, 6, 6, 
                                   QtGui.QPen(QtCore.Qt.black), QtGui.QBrush(QtCore.Qt.white)))
            self._handles[-1].setZValue(11)
            
    def mousePressEvent(self, e):
        """ Define the mouse click event for the indicator region & handles.
        """
        p = self.mapToScene(e.pos())
        self._state = next((i for i, h in enumerate(self._handles) if h.rect().contains(p)), None)
        # If the mouse position, on click, is within the indicator region, start the drag logic.
        if self._state is not None and self._rect is not None:
            self._offset = p - self._rect.rect().topLeft()
        elif self._rect.rect().contains(p):
            self._state = 'drag'
            self._offset = p - self._rect.rect().topLeft()
            
    def mouseMoveEvent(self, e):
        """ Define the mouse move event for the indicator region & handles.
        """
        if self._state is None:
            return
        p = self.mapToScene(e.pos())
        r = self._rect.rect()
        # On mouse move, drag the indicator region (if state is drag), or resize the region.
        if self._state == 'drag':
            self.set_roi(int(max(0, min(p.x() - self._offset.x(), self._w - r.width()))),
                        int(max(0, min(p.y() - self._offset.y(), self._h - r.height()))),
                        int(r.width()), int(r.height()))
        else:
            self._resize(p)
            
    def mouseReleaseEvent(self, e):
        """ On mouse release, reset to idle state.
        """
        self._state = None
        
    def _resize(self, p):
        """ Indicator region resizing logic.
        """
        r = self._rect.rect()
        x, y, w, h = r.x(), r.y(), r.width(), r.height()
        a = self._state
        # Camera specific ROI coordinate restrictions applied here.
        if a in [0, 3, 5]:
            nw = max(16, min(self._w - x, w + (x - p.x())))
            x, w = max(0, min(p.x(), self._w - nw)), nw
        if a in [2, 4, 7]:
            w = max(16, min(self._w - x, p.x() - x))
        if a in [0, 1, 2]:
            nh = max(16, min(self._h - y, h + (y - p.y())))
            y, h = max(0, min(p.y(), self._h - nh)), nh
        if a in [5, 6, 7]:
            h = max(16, min(self._h - y, p.y() - y))
        self.set_roi(int(x), int(y), int(w), int(h))
        
    def get_roi(self):
        """ Get the current ROI region.
        """
        r = self._rect.rect()
        return int(r.x()), int(r.y()), int(r.width()), int(r.height())
        
    def set_roi(self, x, y, w, h):
        """ Set the current ROI region and emit an inter-thread, thread-safe signal
            using the predefined roi_changed signal.
        """
        self._create_roi(x, y, w, h)
        # Emit the ROI properties as a thread-safe signal using predefined roi_changed signal.
        self.roi_changed.emit(x, y, w, h)


class ROIWindow(QtWidgets.QDialog):
    """ Create the full ROI settings window. Add the ROI graphics view in
        the same manner as a QWidget to the window.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        # Initialise the window.
        self.setWindowTitle('ROI Settings')
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
        [c.addItems(['1x', '2x', '4x', '8x']) for c in self._bins]
        
        self._edits = [QtWidgets.QLineEdit() for _ in range(4)]
        vals = [(0, 640), (0, 512), (0, 640), (0, 512)]
        [self._edits[i].setValidator(QtGui.QIntValidator(*vals[i])) for i in range(4)]
        
        self._full = QtWidgets.QPushButton('Full Sensor')
        self._centre = QtWidgets.QPushButton('Centre ROI')
        self._ROI_label = QtWidgets.QLabel('640x512')
        self._apply = QtWidgets.QPushButton('Apply Current ROI')
        
        self._last_data = None
        
        self._setup_ui()
        self._connect()
        self._view.set_roi(0, 0, 640, 512)
        
    def _setup_ui(self):
        """ Add elements to the main ROI window layout.
        """
        bin_row = self._row([QtWidgets.QLabel('X:'), self._bins[0], 
                            QtWidgets.QLabel('Y:'), self._bins[1]], 10)
        self._left_layout.addWidget(self._frame([bin_row], 'Binning'))
        
        labels = ['Start X:', 'Start Y:', 'Width:', 'Height:']
        
        # Use _form_row for aligned labels.
        rows = [self._form_row(labels[i], self._edits[i]) for i in range(4)]
        rows.extend([self._row([self._full, self._centre, self._ROI_label], 10), self._apply])
        
        self._left_layout.addWidget(self._frame(rows, 'Region-of-interest (ROI)'))
        self._left_layout.addStretch()
        
        self._view = ROIWidget()
        self._main.addWidget(self._left)
        self._main.addWidget(self._view)
        
    def _connect(self):
        """ Connect logic functions to settings widgets.
        """
        [e.textChanged.connect(self._sync_edits) for e in self._edits]
        self._view.roi_changed.connect(self._sync_view)
        self._full.clicked.connect(lambda: self._view.set_roi(0, 0, 640, 512))
        self._centre.clicked.connect(self._on_centre)
        self._apply.clicked.connect(self._on_apply)
        
    def _sync_edits(self):
        """ Sync all ROI related settings widgets.
        """
        try:
            x, y, w, h = [int(e.text() or 0) for e in self._edits]
            w, h = min(w, 640 - x), min(h, 512 - y)
            self._view.set_roi(x, y, w, h)
        except:
            pass
            
    def _sync_view(self, x, y, w, h):
        """ Sync the ROI related labels.
        """
        [e.setText(str(v)) for e, v in zip(self._edits, [x, y, w, h])]
        self._ROI_label.setText(f'{w}x{h} unbinned')
        
    def _on_centre(self):
        """ Logic function for recentring the ROI.
        """
        x, y, w, h = self._view.get_roi()
        self._view.set_roi((640 - w) // 2, (512 - h) // 2, w, h) 
        
    def _on_apply(self):
        """ Logic function for applying current ROI settings to the ROI.
        """
        # Get ROI coords.
        start_x = int(self._edits[0].text() or 0)
        start_y = int(self._edits[1].text() or 0)
        width = int(self._edits[2].text() or 0)
        height = int(self._edits[3].text() or 0)
        
        # Get binning factors.
        bin_x = int(self._bins[0].currentText()[0])
        bin_y = int(self._bins[1].currentText()[0])
        
        # Apply corrections to make coordinates divisible by binning factors.
        start_x = start_x - (start_x % bin_x)
        width = width - (width % bin_x)
        start_y = start_y - (start_y % bin_y)
        height = height - (height % bin_y)
        
        if bin_x == 1 or bin_x == 2:
            # Apply corrections to force x to be divisible by 4 (hardware requirement).
            width = width - ((start_x + width) % 4)
        
            
        # Bounds checking with re-correction.
        if start_x + width > 640:
            width = 640 - start_x
            width = width - (width % bin_x)
            
        if start_x + width > 640:
            width = 640 - start_x
            width = width - (width % 4)
                
        if start_y + height > 512:
            height = 512 - start_y
            height = height - (height % bin_y)
        
        if width % bin_x != 0 or height % bin_y != 0:
            QtWidgets.QMessageBox.warning(self, "ROI Error", f"Width/Height must be divisible by binning {bin_x}x{bin_y}")
            return

        # Set the ROI.
        self._view.set_roi(start_x, int(self._edits[1].text() or 0), width, int(self._edits[3].text() or 0))
        
        # Send the ROI properties to a params dict.
        self._last_data = {'start_x': start_x,
                           'start_y': int(self._edits[1].text() or 0),
                           'width': width,
                           'height': int(self._edits[3].text() or 0),
                           'bin_x': int(self._bins[0].currentText()[0]),
                           'bin_y': int(self._bins[1].currentText()[0])}
        
    @staticmethod
    def _row(widgets, spacing=0):
        """ Add widgets side-by-side to a QtWidgets QHBoxLayout.

        Args:        
            widgets (QtWidgets.QWidget): The list of widgets to 
                                         add.

            spacing (float): Defines the vertical spacing of the
                             row.

        Returns:
            row (QtWidgets.QWidget): The full QHBoxLayout as a 
                                     placable QWidget.
        """
        row = QtWidgets.QFrame()
        row.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        row.setLineWidth(1)
        row.setStyleSheet('QFrame {background-color: lightgrey}')
        layout = QtWidgets.QHBoxLayout(row)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(spacing)
        [layout.addWidget(w) for w in widgets]
        layout.addStretch()
        return row
        
    @staticmethod
    def _form_row(label_text, widget):
        """ Add widgets side-by-side to a QtWidgets QHBoxLayout.
            This sspecific method is used to ensure fixed label
            sizing.

        Args:
            label_text (str): A row label widget preceeding the 
                              functional widgets.
                              
            widget (QtWidgets.QWidget): The list of widgets to 
                                        add.

        Returns:
            row (QtWidgets.QWidget): The full QHBoxLayout as a 
                                     placable QWidget.
        """
        row = QtWidgets.QFrame()
        row.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        row.setLineWidth(1)
        row.setStyleSheet('QFrame {background-color: lightgrey}')
        layout = QtWidgets.QHBoxLayout(row)
        layout.setContentsMargins(2, 2, 2, 2)
        label = QtWidgets.QLabel(label_text)
        label.setFixedWidth(60)  # Consistent label width
        layout.addWidget(label)
        layout.addWidget(widget)
        layout.addStretch()
        return row
        
    @staticmethod
    def _frame(widgets, title=''):
        """ Add widgets vertically & sequentially to a QtWidgets.QFrame
            with a QVBoxLayout.

        Args:
            widgets (list): The list of widgets to add.

            title (str): The title of the frame to be displayed.

        Returns:
            frame (QtWidgets.QFrame): The full QVBoxLayout as a placable 
                                      QFrame.
        """
        frame = QtWidgets.QFrame()
        frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        frame.setLineWidth(1)
        frame.setStyleSheet('QFrame {background-color: lightgrey}')
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
        """ The ROI data dict.
        """
        return self._last_data
    


class LogWidget(QtWidgets.QWidget):
    """ A message logging window, use as a QWidget.
    """
    def __init__(self):
        super().__init__()
        self.setFixedSize(640, 256)
        self.te = QtWidgets.QTextEdit(readOnly=True, styleSheet="background-color: black;")
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.te)
        layout.setContentsMargins(0, 0, 0, 0)

    def log(self, msg, color):
        """ Log a message with timestamp and color styling.

        Args:
            msg (str): The message to be logged.
                              
            color (str): The colour code string.
                         E.g.  '#1dff35', 'orange' etc.
        """
        ts = datetime.now().strftime("[%d/%m/%y: %H:%M:%S]")
        self.te.append(f'<span style="color:{color}; font-family:Arial; font-size:10pt; '
                      f'font-weight:bold;">{ts} {msg}</span>')



class ImageViewerWidget(QtWidgets.QWidget):
    """ Embeddable camera image viewer widget with histogram and contrast control.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(640, 722)
        self.zoom = 1.0 # Default zoom level
        self.item = self.data = self.rgb = None # Current data
        self.contrast_low = 0 # Minimum contrast cutoff
        self.contrast_high = 65535 # Maximum contrast cutoff
        self.dragging_bar = None # Contrast control bars
        self.auto_contrast_enabled = True  # Track auto state

        # Histogram setup
        self.fig = Figure(figsize=(4.8, 1.3), dpi=100, facecolor='black')
        self.ax = self.fig.add_subplot(111, facecolor='black')
        self.ax.set_xlim(0, 65535)
        self.ax.set_ylabel('')
        self.ax.set_xlabel('Pixel Value (DN)', fontsize=8, fontweight='medium', color='#e0e0e0')
        self.ax.set_yticks([])
        self.ax.tick_params(axis='x', labelsize=8, colors='#e0e0e0')
        for spine in self.ax.spines.values():
            visible = spine.spine_type == 'bottom'
            spine.set_visible(visible)
            if visible:
                spine.set_color('white')
        self.fig.subplots_adjust(left=0.0625, right=0.95, top=1, bottom=0.25)
        self.hist_canvas = FigureCanvas(self.fig)
        self.hist_canvas.setFixedSize(480, 150)
        self.fig.canvas.mpl_connect('button_press_event', self._on_bar_press)
        self.fig.canvas.mpl_connect('motion_notify_event', self._on_bar_motion)
        self.fig.canvas.mpl_connect('button_release_event', self._on_bar_release)
        self.bar_low = self.ax.axvline(0, color='purple', linewidth=3)
        self.bar_high = self.ax.axvline(65535, color='purple', linewidth=3)

        # Stats box setup
        self.stats_box = QtWidgets.QFrame()
        self.stats_box.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        self.stats_box.setLineWidth(1)
        self.stats_box.setStyleSheet('QFrame {background-color: lightgrey}')
        self.stats_box.setFixedSize(160, 150)
        layout = QtWidgets.QVBoxLayout(self.stats_box)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(3)
        self.auto_btn = QtWidgets.QPushButton('Auto Contrast')
        self.auto_btn.clicked.connect(self._auto_contrast)
        layout.addWidget(self.auto_btn)
        self.stats_labels = {}
        for name in ('Mean', 'Median', 'Max', 'Min', 'Std. Dev.', 'No. Pix.'):
            self.stats_labels[name] = QtWidgets.QLabel(f'{name}: ')
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
        self.info_row.setStyleSheet('QFrame {background-color: lightgrey}')
        self.info_row_layout = QtWidgets.QVBoxLayout(self.info_row)
        self.info = QtWidgets.QLabel()
        self.info.setFixedSize(640, 20)
        
        # Button that generates a dummy image for testing the image display
        #self.btn = QtWidgets.QPushButton('Open Image')
        #self.btn.setFixedSize(640, 40)
        #self.btn.clicked.connect(self._open_image)
        
        self.info_row_layout.addWidget(self.info)
        #self.info_row_layout.addWidget(self.btn)

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
        """ Call relevant functions to display array. 
        """
        self.data = arr
        self._update_stats()
        # Always update histogram (xlims) for new data
        if self.dragging_bar is None:
            self._update_histogram()
        self._redraw_image()

    def _update_histogram(self):
        """ Update xlims and redraw histogram plot.
        """
        if self.data is None:
            return
        counts = np.bincount(self.data.ravel(), minlength=65536)
        self.ax.clear()
        self.ax.set_facecolor('black')
        
        mean, std = np.mean(self.data), np.std(self.data)
            
        # Set xlims based on incoming array
        xlim_low = max(0, mean - 5 * std)
        xlim_high = min(65535, mean + 5 * std)
        self.ax.set_xlim(xlim_low, xlim_high)
        
        # Auto-calculate contrast (pm 3 std dev.)
        if self.auto_contrast_enabled:
            self.contrast_low = max(0, mean - 3 * std)
            self.contrast_high = min(65535, mean + 3 * std)
        
        # Plot histogram and contrast bars
        self.ax.set_ylabel('')
        self.ax.set_xlabel('Pixel Value (DN)', fontsize=8, fontweight='medium', color='#e0e0e0')
        self.ax.set_yticks([])
        self.ax.tick_params(axis='x', labelsize=8, colors='#e0e0e0')
        for spine in self.ax.spines.values():
            if spine.spine_type == 'bottom':
                spine.set_visible(True); spine.set_color('#808080'); spine.set_linewidth(1.2)
            else:
                spine.set_visible(False)
        self.ax.plot(counts, color='#FFD700')
        self.bar_low = self.ax.axvline(self.contrast_low, color='purple', linewidth=3)
        self.bar_high = self.ax.axvline(self.contrast_high, color='purple', linewidth=3)
        self.hist_canvas.draw()

    def _auto_contrast(self):
        """ Enable auto contrast flag and redraw.
        """
        if self.data is None:
            return
        self.auto_contrast_enabled = True
        self._update_histogram()
        self._redraw_image()

    def _on_bar_press(self, event):
        """ Logic for clicking on a contrast control bar.
        """
        if event.inaxes != self.ax or self.data is None:
            return
        if abs(event.xdata - self.contrast_low) < 20:
            self.dragging_bar = 'low'
        elif abs(event.xdata - self.contrast_high) < 20:
            self.dragging_bar = 'high'

    def _on_bar_motion(self, event):
        """ Logic for dragging a contrast control bar.
        """
        if self.dragging_bar is None or event.inaxes != self.ax:
            return
        xdata = int(np.clip(event.xdata, 0, 65535))
        (self.bar_low if self.dragging_bar == 'low' else self.bar_high).set_xdata([xdata, xdata])
        self.hist_canvas.draw()

    def _on_bar_release(self, event):
        """ Logic for releasing a dragged contrast control bar.
        """
        if self.dragging_bar is None or event.inaxes != self.ax:
            return
        xdata = int(np.clip(event.xdata, 0, 65535))
        if self.dragging_bar == 'low':
            self.contrast_low = xdata
        else:
            self.contrast_high = xdata
        self.dragging_bar = None
        self.auto_contrast_enabled = False  # Disable auto on manual adjust
        self._redraw_image()

    def _redraw_image(self):
        """ Transform from raw data to rgb and draw as an image.
        """
        if self.data is None: return
        h, w = self.data.shape
        stretched = self._stretch_contrast(self.data)
        self.rgb = np.zeros((h, w, 3), np.uint8)
        # Saturation point
        sat = self.data == 65535
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
        """ Apply contrast from the contrast control bars.
        """
        if self.contrast_high == self.contrast_low:
            return np.zeros(arr.shape, np.uint8)
        scaled = (arr.astype(np.float32) - self.contrast_low) * 255.0 / (self.contrast_high - self.contrast_low)
        return np.clip(scaled, 0, 255).astype(np.uint8)

    def _update_stats(self):
        """ Recalculate all the relevant statistics.
        """  
        if self.data is None: return
        arr = self.data
        self.stats_labels['Mean'].setText(f'Mean: {np.mean(arr):.1f} DN')
        self.stats_labels['Median'].setText(f'Median: {np.median(arr):.1f} DN')
        self.stats_labels['Max'].setText(f'Max: {np.max(arr)} DN')
        self.stats_labels['Min'].setText(f'Min: {np.min(arr)} DN')
        self.stats_labels['Std. Dev.'].setText(f'Std. Dev.: {np.std(arr):.1f} DN')
        self.stats_labels['No. Pix.'].setText(f'No. Pix.: {arr.size}')

    def _open_image(self):
        """ Generate a dummy data array.
        """  
        arr = np.random.randint(0, 65536, (512, 640), np.uint16)
        arr[100:110, 100:110] = 65535
        self.display_array(arr)

    def _wheel_event(self, event):
        """ Determine image zoom based on scrolling the mouse-wheel.
        """  
        delta = 1.2 if event.angleDelta().y() > 0 else 1 / 1.2
        self.zoom = max(1, min(self.zoom * delta, 20))
        self._update_transform()

    def _update_transform(self):
        """ Apply a given zoom level and add/remove scrollbars if necessary.
        """  
        self.view.setTransform(QtGui.QTransform().scale(self.zoom, self.zoom))
        sb = QtCore.Qt.ScrollBarAsNeeded if self.zoom > 1.0 else QtCore.Qt.ScrollBarAlwaysOff
        self.view.setVerticalScrollBarPolicy(sb)
        self.view.setHorizontalScrollBarPolicy(sb)

    def _mouse_move_event(self, event):
        """ Get the current pixel and associated value at the mouse position after moving.
        """  
        if self.data is None:
            QtWidgets.QGraphicsView.mouseMoveEvent(self.view, event); return
        pos = self.view.mapToScene(event.pos())
        x, y = int(pos.x()), int(pos.y())
        if 0 <= x < self.data.shape[1] and 0 <= y < self.data.shape[0]:
            self.info.setText(f'x:{x} y:{y} Signal:{self.data[y, x]} DN zoom:{self.zoom:.1f}x')
        else:
            self.info.clear()
        QtWidgets.QGraphicsView.mouseMoveEvent(self.view, event)
        
        
class PreviewThread(QtCore.QThread):
    """ PyQt thread for previewing images from the camera.
    """  
    # Define the inter-thread signals and states dict
    frame = QtCore.pyqtSignal(np.ndarray)
    status = QtCore.pyqtSignal(float, str)
    error = QtCore.pyqtSignal(str)
    _states = {0: 'Idle', 1: 'Waiting', 2: 'Exposing', 3: 'Reading', 4: 'Download', 5: 'Error'}
    
    def __init__(self, cam, params):
        super().__init__()
        # Get and assign params
        self.cam, self.p, self.run_flag = cam, params, True

    def run(self):
        """ Loop for acquiring images from the camera to preview.
        """  
        #pythoncom.CoInitialize()
        try:
            cam = self.cam
            p = self.p
            
            # ROI setup
            cam.BinX, cam.BinY = p['bin_x'], p['bin_y']
            cam.StartX, cam.StartY = p['start_x'], p['start_y']
            cam.NumX, cam.NumY = p['width'], p['height']
            
            # Pre-calculate exposure time
            unit = p['time_unit']
            mult = {'s': 1.0, 'ms': 1e-3, 'μs': 1e-6}[unit]
            exp_time = p['exp_array'][0] * mult
            
            states = self._states
            while self.run_flag:
                try:
                    self.status.emit(cam.CCDTemperature, 'Exposing')
                    cam.StartExposure(exp_time, True)
                    
                    poll_interval = min(0.01, max(0.0001, exp_time * 0.1))
                    
                    while not cam.ImageReady and self.run_flag:
                        time.sleep(poll_interval)
                    
                    if not self.run_flag:
                        break
                        
                    img = cam.ImageArray
                    arr = np.transpose(np.array(img, dtype=np.int16).astype(np.uint16))
                    self.frame.emit(arr)
                    self.status.emit(cam.CCDTemperature, 
                                   states.get(cam.CameraState, 'Unknown'))
                except Exception as e:
                    self.error.emit(str(e))
                    break
        finally:
            #pythoncom.CoUninitialize()
            pass
                
    def abort(self, timeout=5.0):
        """ Abort preview.
        """  
        self.run_flag = False
        self.cam.AbortExposure()
        start = time.time()
        while self.cam.CameraState not in (0, 5) and time.time() - start < timeout:
            time.sleep(0.1)
        if time.time() - start >= timeout:
            self.error.emit("Camera abort timeout!")


class ExperimentThread(QtCore.QThread):
    """ PyQt thread for acquiring images from the camera.
    """  

    # Define the inter-thread signals and the states dict
    frame = QtCore.pyqtSignal(np.ndarray)
    status = QtCore.pyqtSignal(float, str)
    progress = QtCore.pyqtSignal(int, str, int, dict, int, int, str, int)
    error = QtCore.pyqtSignal(str)
    log_requested = QtCore.pyqtSignal(str, str)
    save_requested = QtCore.pyqtSignal(np.ndarray, int, int, int, object, dict, float, float, str, datetime)
    finished = QtCore.pyqtSignal()
    _states = {0: 'Idle', 1: 'Waiting', 2: 'Exposing', 3: 'Reading', 4: 'Download', 5: 'Error'}
    
    def __init__(self, camera, containers, saver):
        super().__init__()
        # Get the params
        self.cam, self.containers, self.saver = camera, containers, saver
        self.run_flag = True
    
    def run(self):
        """ PyQt thread for acquiring images from the camera.
        """  
        #pythoncom.CoInitialize()
        try:
            cam = self.cam
            states = self._states
            # The individual experiment sub-windows
            containers = self.containers
            
            # Loop over the experiments
            for exp_idx, container in enumerate(containers):
                if not self.run_flag:
                    break
                    
                # params can be modified by GUI during a running experiment
                # (i.e. from set ROI, set temperature etc.).
                # This creates a deepcopy that is independent of params.
                params = copy.deepcopy(container.internalGui.window_params)
                name = container.title
                
                bin_x = params['bin_x']
                bin_y = params['bin_y']
                start_x = params['start_x']
                start_y = params['start_y']
                width = params['width']
                height = params['height']
                
                # Configure ROI
                cam.BinX, cam.BinY = bin_x, bin_y
                cam.StartX, cam.StartY = start_x, start_y
                cam.NumX, cam.NumY = width, height
                
                # Pre-calculate exposure time
                unit = params['time_unit']
                mult = {'s': 1.0, 'ms': 1e-3, 'μs': 1e-6}[unit]
                exp_array = np.asarray(params['exp_array']) * mult
                
                # Get experiment settings
                experiment_delay = params['delay']
                experiment_repeats = params['repeats']

                # Image acquisition loops
                total_exp = len(exp_array)
                for j in range(int(experiment_repeats)):
                    self.log_requested.emit(f'Waiting {experiment_delay}s', 'orange')
                    time.sleep(float(experiment_delay))
                    for i, exp in enumerate(exp_array):
                        if not self.run_flag:
                            break
    
                        try:
                            temp = cam.CCDTemperature
                            target_temp = cam.SetCCDTemperature
                            ts_utc = datetime.now(timezone.utc).isoformat()
                            ts_ltc = datetime.now()
                            self.status.emit(cam.CCDTemperature, 'Exposing')
                            
                            cam.StartExposure(exp, True)
                            
                            poll_interval = min(0.01, max(0.0001, exp * 0.1))
                            while not cam.ImageReady and self.run_flag:
                                time.sleep(poll_interval)
                            
                            if not self.run_flag:
                                break
                                
                            img = cam.ImageArray
                            arr = np.transpose(np.array(img, dtype=np.int16).astype(np.uint16))
                            self.frame.emit(arr)

                            self.save_requested.emit(arr, exp_idx, i, j, exp, params,
                                                   temp, target_temp, ts_utc, ts_ltc)
                            self.progress.emit(exp_idx, name, params['exp_array'][i], params, i + 1, j + 1, str(experiment_repeats), total_exp)
                            self.status.emit(cam.CCDTemperature, 
                                           states.get(cam.CameraState, 'Unknown'))
                        except Exception as e:
                            self.error.emit(str(e))
                            return
                self.finished.emit()
        finally:
            #pythoncom.CoUninitialize()
            pass

    def abort(self, timeout=5.0):
        """ Abort image acquisition.
        """  
        self.run_flag = False
        self.cam.AbortExposure()
        start = time.time()
        while self.cam.CameraState not in (0, 5) and time.time() - start < timeout:
            time.sleep(0.1)
        if time.time() - start >= timeout:
            self.error.emit("Camera abort timeout!")


class SaverThread(QtCore.QThread):
    """ PyQt thread for saving a queue of images.
    """  

    # Define the inter-thread signals
    progress = QtCore.pyqtSignal(str)
    log_error = QtCore.pyqtSignal(str, str)
    
    def __init__(self):
        super().__init__()
        self.queue = Queue(maxsize=100) # Cap queue to avoid potential memory bloat for long runs.
        self.run_flag = True

    def run(self):
        """ File saving loop.
        """  
        #pythoncom.CoInitialize()
        try:
            from queue import Empty
            q = self.queue
            while self.run_flag:
                try:
                    arr, hdr, path = q.get(timeout=1.0)
                    fits.writeto(path, arr, hdr, overwrite=True)
                    self.progress.emit(path)
                except Empty:
                    pass
        finally:
            #pythoncom.CoUninitialize()
            pass

    def save(self, arr, exp_idx, img_idx, exp_rep_idx, exp, p, temp, target_temp, ts_utc, ts_ltc):
        """ Add a non-blocking job to the queue.
        """
        path = self._build_path(exp_idx, img_idx, exp_rep_idx, exp, p, temp, ts_ltc)
        hdr = self._build_header(p, exp, temp, target_temp, ts_utc, ts_ltc)
        try:
            self.queue.put((arr, hdr, path), timeout=2.0)  # Raise if full
        except:
            self.log_error.emit(f'Save queue full, dropping frame: {img_idx+1}, exp: {exp_idx+1}, rep: {exp_rep_idx+1}!', 'red')

    def _build_path(self, exp_idx, img_idx, exp_rep_idx, exp, p, temp, ts_ltc):
        """ Helper function to build the correct file name and directory path.
        """  
        if p['include_date_time']:
            if int(ts_ltc.strftime("%H")) >= 12:
                _date_time_str = ts_ltc.strftime('_%d_%m_%Y_%H-%Mpm')
            else:
                _date_time_str = ts_ltc.strftime('_%d_%m_%Y_%H-%Mam')
        else:
            _date_time_str = ''
        
        if p['include_gain_rospeed']:
            _gain_rospeed_str = '_'+ p['gain'] + '_' + p['rospeed']
        else:
            _gain_rospeed_str = ''
            
        if p['include_temp']:
            _temp_str = '_' + str(temp) + 'C'
        else:
            _temp_str = ''
            
        if p['include_exp_time']:
            _exp_time_str = '_' + str(exp) + p['time_unit']
        else:
            _exp_time_str = ''
            
        base = (p['savedirectory'] + '/' + p['nameprefix'] +
                _date_time_str + _gain_rospeed_str +
                _temp_str + _exp_time_str)
        return f"{base}_exp{exp_idx+1}_rep{exp_rep_idx+1}_img{img_idx+1}.fits"
    
    def _build_header(self, p, exp, temp, target_temp, ts_utc, ts_ltc):
        """ Helper function to build the fits header.
        """  
        hdr = fits.Header()
        hdr['SIMPLE'] = (True, 'File conforms to FITS standard')
        hdr['COMMENT'] = ('FITS (Flexible Image Transport System) format')
        hdr['BITPIX'] = (16, 'Number of bits per data pixel')
        hdr['NAXIS'] = (2, 'Number of data axes')
        hdr['NAXIS1'] = (int(p['width']/p['bin_x']), '')
        hdr['NAXIS2'] = (int(p['height']/p['bin_y']), '')
        hdr['EXTEND'] = (True, 'FITS dataset may contain extensions')    
        hdr['OBSERVER'] = ('obs', 'Name of observer')
        hdr['BZERO'] = (32768.0, 'Offset data range')
        hdr['BSCALE'] = (1.0, 'Scale factor for the data')
        hdr['XORGSUBF'] = (int(p['start_x']/p['bin_x']), 'Subframe upper-left X in pixel coordinates')
        hdr['YORGSUBF'] = (int(p['start_y']/p['bin_y']), 'Subframe upper-left Y in pixel coordinates')
        hdr['XBINNING'] = (p['bin_x'], 'Binning factor in width')
        hdr['YBINNING'] = (p['bin_y'], 'Binning factor in height')
        hdr['XPIXSZ'] = (20*p['bin_x'], 'Pixel width in microns, after binning')
        hdr['YPIXSZ'] = (20*p['bin_y'], 'Pixel height in microns, after binning')
        hdr['INTRUME '] = ('NIRvana HS', 'The model of the camera')
        hdr['HISTORY'] = ('Created using Python ASCOM client')
        hdr['DATE-OBS'] = (ts_utc, 'UTC at exposure start')
        hdr['LOCALTIM'] = (ts_ltc.isoformat(), 'Local time at exposure start')
        hdr['SEN_TEMP'] = (temp, 'Temperature of the sensor')
        hdr['SET_TEMP'] = (target_temp, 'Cooler setpoint in degrees C')

        return hdr
    
    def stop(self):
        """ Stop the save queue.
        """  
        self.run_flag = False
        # Drain remaining queue items
        while not self.queue.empty():
            time.sleep(0.05)
        self.wait()
        
        
class MainWindow(QtWidgets.QWidget):
    """ The main GUI window defined and usable as a QWidget.
    """  
    # GREEN/RED button styles
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
        
    # Status update from image capture threads to update UI
    status_update = QtCore.pyqtSignal(float, str)
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ASCOM Camera Controller")
        
        # Timer for idle polling (main thread only)
        self._poll_timer = QtCore.QTimer()
        self._poll_timer.setInterval(500)  # 500 ms
        self._poll_timer.timeout.connect(self._idle_poll)
        
        # Thread placeholders
        self._preview_thread = None
        self._exp_thread = None
        self._saver_thread = None
        
        # Wire signal to UI slot
        self.status_update.connect(self._update_status)
        
        # Master window layout
        self._master_layout = QtWidgets.QHBoxLayout(self)
        self.setLayout(self._master_layout)
        
        self.camera = None
        self.is_connected = False
        
        # Build UI elements
        self._build_left_side()
        self._build_right_side()
        
    def _update_status(self, temp, state):
        """ Update the current camera temperature label.
        """
        if temp == self._temp_target:
            self._current_temp_label.setText(f'Temperature: {temp:.1f}°C     State: {state}')
        else:
            self._current_temp_label.setText(f'Temperature: ⚠{temp:.1f}°C ({self._temp_target}°C target)     State: {state}')

    def _idle_poll(self):
        """ Main thread polling when camera is idle.
        """
        if self.camera and self.camera.Connected:
            try:
                _state = {
                    0: 'Idle', 1: 'Waiting', 2: 'Exposing',
                    3: 'Reading', 4: 'Download', 5: 'Error'
                    }.get(self.camera.CameraState, 'Unknown')
                _temp = self.camera.CCDTemperature
                
                self._update_status(_temp, 
                                   _state)
            except Exception as e:
                self.log.log(f'Error reading temperature/state: {e}', 'red')
        else:
            self.log.log('Cannot read temperature/state, no camera connected', 'red')
        
    def _update_containers(self):
        """ Update the number of experiment containers.
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
            container = CollapsibleWidget(f'Experiment {i + 1}')
            self._container_layout.insertWidget(i, container)
            self._containers.append(container)

            contents_layout = QtWidgets.QVBoxLayout(container.contentWidget)
            contents_layout.addWidget(CollapsibleWidgetInternalGUI())

    def _get_experiments(self):
        """ Extract a tuple of (index, title, params) from the experiment containers.
        """
        return [(i, c.title, c.internalGui.window_params) for i, c in enumerate(self._containers)]
    
    # Dummy function for testing retrieval of experiment settings:
    # Dumps all experiment settings to the log widget.
    def _grab_contents_fun(self):
        """ Dummy function for testing retrieval of experiment settings:
            Dumps all experiment settings to the log widget.
        """       
        for idx, title, params in self._get_experiments():
            self.log.log(f'Running experiment {idx}: {title}', '#1dff35')
            self.log.log(f'Params: {params}', '#1dff35')
    
    
    def _connect_pressed(self):
        """ Toggle between connect/disconnect
        """   
        if not self.is_connected:
            # Camera connection logic
            try:
                self.camera_running = True
                
                # Connect to NIRvana HS
                #self.camera = win32com.client.Dispatch('ASCOM.PI.Camera.1')

                # Connect to ASCOM SIM
                self.camera = Camera('localhost:32323', 0)

                self.log.log('Connecting Camera', 'orange')
                self.camera.Connected = True
                self.log.log(f'Connected Camera: {self.camera.Name}', '#1dff35')
                
                # Initialize ROI settings
                self.camera.BinX = 1
                self.camera.BinY = 1
                self.camera.StartX = 0
                self.camera.StartY = 0
                self.camera.NumX = self.camera.CameraXSize // self.camera.BinX
                self.camera.NumY = self.camera.CameraYSize // self.camera.BinY
                self.log.log('Initialised ROI settings', '#1dff35')
                
                # Set target temperature
                self.camera.CoolerOn = True
                self._temp_target = -40.0
                self.camera.SetCCDTemperature = self._temp_target
                self.log.log(f'Set target temp. {self._temp_target}°C', '#1dff35')
                
                # Start temperature monitoring poll
                self._poll_timer.start()
                self.log.log('Started temperature monitor', '#1dff35')
                
                # Switch from connect -> disconnect
                self._connect_button.setStyleSheet(self.RED_STYLE)
                self._connect_button.setText('Disconnect')
                self._preview_run_buttons_row.setVisible(True) 
                
                self.is_connected = True
                
            except Exception as e:
                self.camera_running = False
                self.log.log(f'Error connecting to camera: {e}', 'red')

        else:
            try:
                # Camera disconnection logic
                if self.camera:
                    try:
                        if self._exp_thread is not None or self._preview_thread is not None:
                            msg = ('<b>Camera previewing/running experiments.</b><br><br>'
                                   '(Consider halting preview/aborting experiments before disconnecting camera)')
                            btn = QtWidgets.QMessageBox.critical(
                                self, 'Preview/Experiments running', msg,
                                QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return
                            
                        temp = self.camera.CCDTemperature
                        if temp < -15.0:
                            msg = (f'<b>Camera at {temp:.1f}°C. Disconnecting now may cause thermal shock.</b><br><br>'
                                   '(Consider warming the sensor above -15.0°C before disconnecting camera)')
                            btn = QtWidgets.QMessageBox.warning(
                                self, 'Thermal Warning', msg,
                                QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                                QtWidgets.QMessageBox.Cancel
                            )
                            if btn == QtWidgets.QMessageBox.Ignore:
                                pass
                            else:
                                return
                    except:
                        pass
                    
                    if self._preview_thread:
                        self.log.log('Halting preview', 'orange')
                        self._preview_thread.abort()
                        self._preview_thread.quit()
                        if not self._preview_thread.wait(2000):  # 2 second timeout
                            self.log.log(f"Warning: Thread {self._preview_thread} didn't terminate gracefully", 'orange')
                        self._preview_thread.deleteLater()
                        self._preview_thread = None
                        
                    if self._exp_thread:
                        self.log.log('Aborting experiments', 'orange')
                        self._exp_thread.abort()
                        self._exp_thread.quit()
                        if not self._exp_thread.wait(2000):  # 2 second timeout
                            self.log.log(f"Warning: Thread {self._exp_thread} didn't terminate gracefully", 'orange')
                        self._exp_thread.deleteLater()
                        self._exp_thread = None
                            
                    if self._saver_thread:
                        self.log.log('Stopping file-saver thread', 'orange')
                        self._saver_thread.stop()
                        self._saver_thread.quit()
                        if not self._saver_thread.wait(2000):  # 2 second timeout
                            self.log.log(f"Warning: Thread {self._exp_thread} didn't terminate gracefully", 'orange')
                        self._saver_thread.deleteLater()
                        self._saver_thread = None
                    
                    self._poll_timer.stop()
                    self.log.log('Stopped temperature monitor', '#1dff35')
                    
                    self.log.log(f'Disconnecting Camera: {self.camera.Name}', 'orange')
                    self.camera.Connected = False
                    del self.camera
                    self.camera = None
                    self.log.log('Camera disconnected', '#1dff35')
                    
                    self._current_temp_label.setText('Current Temperature: XX.x°C     State: X')

                
                    # Switch from disconnect -> connect
                    self._connect_button.setStyleSheet(self.GREEN_STYLE)
                    self._connect_button.setText('Connect')
                    self._preview_run_buttons_row.setVisible(False)
            
                    self.is_connected = False
                    
            except Exception as e:
                    self.log.log(f'Error disconnecting camera: {e}', 'red')
                
    def _run_experiments_pressed(self):
        """ Start the experiment and saving threads with the current experiments.
        """
        btn = self._run_experiments_button
        if not self._exp_thread:  
            # Start experiments
            if not (self.camera and self.camera.Connected):
                self.log.log("Run failed: camera not connected", "red")
                return
            if not self._containers:
                self.log.log("Run failed: no experiments", "red")
                return
            

            msg = ('<b>Have you set FITS header options as approriate?</b> <br><br>'
                   '(Note: setting gain/read-speed etc. via camera settings does not update the FITS headers, '
                   'these need to be updated by the user within each experiment settings block as appropriate)')
            btn_msg_box = QtWidgets.QMessageBox.warning(
                self, 'FITS Header reminder', msg,
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.Yes
            )
            if btn_msg_box == QtWidgets.QMessageBox.No:
                return
            else:
                pass

            # Start the threads and connect the inter-thread signals
            self._saver_thread = SaverThread()
            self._saver_thread.progress.connect(self._log_saved_progress)
            self._saver_thread.log_error.connect(self.log.log)
            self._saver_thread.start()
            self._exp_thread = ExperimentThread(self.camera, self._containers, self._saver_thread)
            self._exp_thread.frame.connect(self.imageviewer.display_array)
            self._exp_thread.status.connect(self._update_status)
            self._exp_thread.progress.connect(self._log_progress)
            self._exp_thread.log_requested.connect(self.log.log)
            self._exp_thread.save_requested.connect(self._saver_thread.save)
            self._exp_thread.error.connect(lambda e: self.log.log(f'Experiment error: {e}', 'red'))
            self._exp_thread.finished.connect(self._experiments_finished)
            self._poll_timer.stop()
            self._exp_thread.start()
            
            # Change the button style
            btn.setStyleSheet(self.RED_STYLE)
            btn.setText('Abort Experiments')
            self.log.log("Experiments started", "#1dff35")
        else:  # Abort
            if self._exp_thread:
                self.log.log('Aborting experiments', 'orange')
                self._exp_thread.abort()
                self._exp_thread.quit()
                if not self._exp_thread.wait(2000):  # 2 second timeout
                    self.log.log(f"Warning: Thread {self._exp_thread} didn't terminate gracefully", 'orange')
                self._exp_thread.deleteLater()
                self._exp_thread = None
            if self._saver_thread:
                self.log.log('Stopping save queue', 'orange')
                self._saver_thread.stop()
                self._saver_thread.quit()
                if not self._saver_thread.wait(2000):  # 2 second timeout
                    self.log.log(f"Warning: Thread {self._saver_thread} didn't terminate gracefully", 'orange')
                self._saver_thread.deleteLater()
                self._saver_thread = None
            
            # Restart the idle polling
            self._poll_timer.start()
            btn.setStyleSheet(self.GREEN_STYLE)
            btn.setText('Run Experiments')
            self.log.log('Experiments aborted', '#1dff35')

    def _log_progress(self, exp_idx, exp_name, exp, params, exp_num, exp_rep_num, exp_reps, total_exp):
        """ Log the current image acquisition.
        """
        self.log.log(f'[Exp {exp_idx+1}/{len(self._containers)}][{exp_rep_num}/{exp_reps}] {exp_name}: '
                    f'Image {exp_num}/{total_exp}: {exp}{params["time_unit"]}', '#1dff35')
    
    def _log_saved_progress(self, path):
        """ Log the latest saved file.
        """
        self.log.log(f'Image saved to → {os.path.basename(path)}', '#1dff35')
        
    def _experiments_finished(self):
        """ Stop the experiment and saving threads.
        """
        self._run_experiments_button.setStyleSheet(self.GREEN_STYLE)
        self._run_experiments_button.setText('Run Experiments')
        self._exp_thread = None 
        self._poll_timer.start()
        if self._saver_thread:
            self._saver_thread.stop()
            self._saver_thread = None 
        self.log.log("All experiments complete", "#1dff35")
            
    def _preview_pressed(self):
        """ Start the image preview thread.
        """
        if self._preview_button.styleSheet() == self.GREEN_STYLE:
            if not self.camera or not self.camera.Connected:
                self.log.log("Preview failed: camera not connected", "red")
                return
            if not self._containers:
                self.log.log("Preview failed: no experiments", "red")
                return
            p = self._containers[0].internalGui.window_params
            if not list(p.get('exp_array')):
                self.log.log("Preview failed: no exposures", "red")
                return
            self._preview_thread = PreviewThread(self.camera, p)
            self._preview_thread.frame.connect(self.imageviewer.display_array)
            self._preview_thread.status.connect(self._update_status)
            self._preview_thread.error.connect(lambda e: self.log.log(f'Preview error: {e}', 'red'))
            self._poll_timer.stop()
            self._preview_thread.start()
            self._preview_button.setStyleSheet(self.RED_STYLE)
            self._preview_button.setText('Halt Preview')
            self.log.log("Preview started", "#1dff35")
        else:
            if self._preview_thread:
                self._preview_thread.abort()
                self._preview_thread = None
            if self.camera and self.camera.Connected:
                self._poll_timer.start()
            self._preview_button.setStyleSheet(self.GREEN_STYLE)
            self._preview_button.setText('Preview')
            self.log.log("Preview halted", "#1dff35")
        
    def _set_temp_target(self):
        """ Set and log the cooler target temperature.
        """
        self._temp_target = float(self._temp_target_line_edit.text())
    
        try:
            self.camera.SetCCDTemperature = self._temp_target
            self.log.log(f'Setting target temp. to {self._temp_target}°C', '#1dff35')
        except Exception as e:
            self.log.log(f'Error setting target temp.: {e}', 'red')
        
    def _open_camera_settings(self):
        """ Routine to open the camera's ASCOM settings (requires disconnecting the camera).
            The Alpyca simulator does not support this and so this will crash the GUI if using
            the simulator.
        """
        if not self.camera:
                self.log.log("Cannot open camera settings, connect camera first", "red")
        else:
            try:
                self.log.log("Disconnecting camera while modifying settings", "orange")
                self.camera.Connected = False
                self.camera.SetupDialog()
                self.log.log("Reconnecting camera after modifying settings", "orange")
                self.camera.Connected = True
            except Exception as e:
                self.log.log(f"Error opening camera setup: {e}", "red")
                    
    def _build_left_side(self):
        """ Build the entire left half of the GUI (experiment settings windows, camera settings etc.).
        """
        # Master left half layout
        self._left_frame = QtWidgets.QFrame()
        self._left_layout = QtWidgets.QVBoxLayout(self._left_frame)
        self._left_layout.setSpacing(0)
        
        self._master_layout.addWidget(self._left_frame)
        
        
        # BUILD THE COLLAPSIBLE WINDOW SECTION
        
        
        # Experiment window number control panel.
        _control = QtWidgets.QFrame()
        _control.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        _control.setLineWidth(1)
        _control.setStyleSheet('QFrame {background-color: lightgrey}')
        _control.setToolTip('Create the specified number of experiments' 
                           '\nand their associated settings windows below.'
                           '\nWindows can be renamed by double clicking the title.')
        _control_layout = QtWidgets.QHBoxLayout(_control)
        _control_layout.setContentsMargins(5, 5, 5, 5)
        _control_layout.addWidget(QtWidgets.QLabel("Experiments:"))
        
        self._spinbox = QtWidgets.QSpinBox()
        self._spinbox.setMinimum(1)
        self._spinbox.valueChanged.connect(self._update_containers)
        _control_layout.addWidget(self._spinbox)
        
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
        #_grab_contents = QtWidgets.QPushButton('Grab Contents')
        #_grab_contents.clicked.connect(self._grab_contents_fun)
        #self._left_layout.addWidget(_grab_contents)
        
        self._containers = []
        self._update_containers()


        # BUILD TEMPERATURE SECTION
        self._current_temp_label = QtWidgets.QLabel('Current Temperature: XX.x°C     State: X')
        self._temp_target_line_edit = QtWidgets.QLineEdit('-55')
        self._temp_target_line_edit.setValidator(QtGui.QDoubleValidator(bottom=-40.0, top=20.0, decimals=1))
        self._set_temp_target_button = QtWidgets.QPushButton('Set Target')
        self._set_temp_target_button.clicked.connect(self._set_temp_target)
        self._camera_settings_button = QtWidgets.QPushButton('Camera Settings')
        self._camera_settings_button.clicked.connect(self._open_camera_settings)
        self._cam_settings_row = CollapsibleWidgetInternalGUI._create_row(widget=[self._temp_target_line_edit, self._set_temp_target_button, self._camera_settings_button])
        self._cam_settings_frame = CollapsibleWidgetInternalGUI._create_frame(widgets=[self._current_temp_label, self._cam_settings_row])
        self._cam_settings_frame.setFrameStyle(QtWidgets.QFrame.Box | QtWidgets.QFrame.Raised)
        self._cam_settings_frame.setLineWidth(1)
        self._cam_settings_frame.setStyleSheet('QFrame {background-color: lightgrey}')
        
        self._left_layout.addWidget(self._cam_settings_frame)
        
        self._connect_button = QtWidgets.QPushButton('Connect')
        self._connect_button.setStyleSheet(self.GREEN_STYLE)
        self._connect_button.clicked.connect(self._connect_pressed)
        self._left_layout.addWidget(self._connect_button)
        
        self._run_experiments_button =QtWidgets.QPushButton('Run Experiments')
        self._run_experiments_button.setStyleSheet(self.GREEN_STYLE)
        self._run_experiments_button.clicked.connect(self._run_experiments_pressed)
        

        self._preview_button =QtWidgets.QPushButton('Preview')
        self._preview_button.setStyleSheet(self.GREEN_STYLE)
        self._preview_button.clicked.connect(self._preview_pressed)
        
        self._preview_run_buttons_row = CollapsibleWidgetInternalGUI._create_row(widget=[self._run_experiments_button, self._preview_button], spacing=5)
        self._left_layout.addWidget(self._preview_run_buttons_row)
        self._preview_run_buttons_row.setVisible(False)
        
        self._left_layout.addStretch()
        
    def _build_right_side(self):
        """ Build the entire right half of the GUI (image display/statistics and log).
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
            
    def closeEvent(self, event):
        """ Logic for closing the GUI, block unsafe shutdown when camera is cooled.
        """
        cam = self.camera
        if cam and cam.Connected:
            try:
                temp = cam.CCDTemperature
                
                if self._exp_thread is not None or self._preview_thread is not None:
                    msg = ('<b>Camera previewing/running experiments.</b><br><br>'
                           'Consider halting preview/aborting experiments before closing.')
                    btn = QtWidgets.QMessageBox.critical(
                        self, 'Preview/Experiments running', msg,
                        QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                        QtWidgets.QMessageBox.Cancel
                    )
                    if btn == QtWidgets.QMessageBox.Ignore:
                        # Force-stop all threads
                        if self._preview_thread:
                            self._preview_thread.abort()
                            self._preview_thread.quit()
                            if not self._preview_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._preview_thread} didn't terminate gracefully", 'orange')
                            self._preview_thread.deleteLater()
                            self._preview_thread = None
                        if self._exp_thread:
                            self._exp_thread.abort()
                            self._exp_thread.quit()
                            if not self._exp_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._exp_thread} didn't terminate gracefully", 'orange')
                            self._exp_thread.deleteLater()
                            self._exp_thread = None
                        if self._saver_thread:
                            self._saver_thread.stop()
                            self._saver_thread.quit()
                            if not self._saver_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._saver_thread} didn't terminate gracefully", 'orange')
                            self._saver_thread.deleteLater()
                            self._saver_thread = None
                        self._poll_timer.stop()
                        event.accept()
                    else:
                        event.ignore()
                        return
                
                if temp < -15.0:
                    msg = (f'<b>Camera at {temp:.1f}°C. Closing now may cause thermal shock.</b><br><br>'
                           '(Consider warming the sensor above -15.0°C before closing)')
                    btn = QtWidgets.QMessageBox.warning(
                        self, 'Thermal Warning', msg,
                        QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.Ignore,
                        QtWidgets.QMessageBox.Cancel
                    )
                    if btn == QtWidgets.QMessageBox.Ignore:
                        # Force-stop all threads
                        if self._preview_thread:
                            self._preview_thread.abort()
                            self._preview_thread.quit()
                            if not self._preview_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._preview_thread} didn't terminate gracefully", 'orange')
                            self._preview_thread.deleteLater()
                            self._preview_thread = None
                        if self._exp_thread:
                            self._exp_thread.abort()
                            self._exp_thread.quit()
                            if not self._exp_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._exp_thread} didn't terminate gracefully", 'orange')
                            self._exp_thread.deleteLater()
                            self._exp_thread = None
                        if self._saver_thread:
                            self._saver_thread.stop()
                            self._saver_thread.quit()
                            if not self._saver_thread.wait(2000):  # 2 second timeout
                                self.log.log(f"Warning: Thread {self._saver_thread} didn't terminate gracefully", 'orange')
                            self._saver_thread.deleteLater()
                            self._saver_thread = None
                        self._poll_timer.stop()
                        event.accept()
                    else:
                        event.ignore()
                        return

            except Exception as e:
                self.log.log(f'Error checking temperature on close: {e}', 'red')
        event.accept()
    
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())