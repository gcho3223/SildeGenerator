#!/usr/bin/env python3
"""
Keynote Slide Generator GUI - Main Application
A GUI application to generate Keynote presentations with plots
Modularized version with separate CPV and DRC components
"""

import sys
import os
import traceback
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QRadioButton, QGroupBox, QButtonGroup, QCheckBox,
    QFileDialog, QMessageBox, QPushButton, QInputDialog, QComboBox
)
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QFont

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import configurations
try:
    from Config.config_cpv import cpv_config
    from Config.config_drc import drc_config
except ImportError:
    from ..Config.config_cpv import cpv_config
    from ..Config.config_drc import drc_config

# Import modularized components
from GUI.common import create_directory_ui, create_status_ui, create_action_buttons
from GUI.gui_cpv import create_cpv_options_ui, create_cpv_objects_ui, create_cpv_steps_ui, run_cpv_generation
from GUI.gui_drc import (
    create_drc_options_ui, create_drc_channels_ui, 
    create_drc_preview_ui, create_energy_checkboxes, update_energy_preview, 
    run_drc_generation
)
from GUI.gui_loopDefined import (
    create_loopDefined_options_ui, create_loopDefined_objects_ui, 
    create_loopDefined_variables_ui, create_loopDefined_preview_ui,
    update_loopDefined_preview
)
from GUI.gui_dragdrop import (
    create_dragdrop_preview_ui, update_dragdrop_preview
)

# Import settings and manual dialogs
from GUI.Setting import SettingsDialog, ManualDialog


class GeneratorThread(QThread):
    """Thread for running slide generation without blocking UI"""
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(self, mode, systematic, input_dir, output_file, selected_objects, selected_steps, 
                 observable_text, jobversion, particle=None, case=None, c_checked=None, 
                 s_selected=None, drcor_selected=None, selected_energies=None,
                 resol_with_noise=False, resol_without_noise=False, linearity=False, file_format="pdf",
                 user_settings=None, loopDefined_path_template="", loopDefined_filename_template="", loopDefined_variables=None,
                 loopDefined_row_values=None, loopDefined_column_values=None, loopDefined_scan_mode=False,
                 dragdrop_file_paths=None, dragdrop_arrangement=None, dragdrop_slide_count=1,
                 sc_overlay_checked=False):
        super().__init__()
        self.mode = mode
        self.systematic = systematic
        self.input_dir = input_dir
        self.output_file = output_file
        self.selected_objects = selected_objects
        self.selected_steps = selected_steps
        self.observable_text = observable_text
        self.jobversion = jobversion
        self.particle = particle
        self.case = case
        self.c_checked = c_checked
        self.s_selected = s_selected if s_selected is not None else ""
        self.drcor_selected = drcor_selected if drcor_selected is not None else ""
        self.sc_overlay_checked = sc_overlay_checked
        self.selected_energies = selected_energies if selected_energies is not None else []
        self.resol_with_noise = resol_with_noise
        self.resol_without_noise = resol_without_noise
        self.linearity = linearity
        self.file_format = file_format
        self.user_settings = user_settings  # Store user settings
        self.loopDefined_path_template = loopDefined_path_template
        self.loopDefined_filename_template = loopDefined_filename_template
        self.loopDefined_variables = loopDefined_variables if loopDefined_variables is not None else {}
        self.loopDefined_row_values = loopDefined_row_values if loopDefined_row_values is not None else []
        self.loopDefined_column_values = loopDefined_column_values if loopDefined_column_values is not None else []
        self.loopDefined_scan_mode = loopDefined_scan_mode
        self.dragdrop_file_paths = dragdrop_file_paths if dragdrop_file_paths is not None else {}
        self.dragdrop_arrangement = dragdrop_arrangement if dragdrop_arrangement else "2*3"
        self.dragdrop_slide_count = dragdrop_slide_count if dragdrop_slide_count else 1
    
    def run(self):
        """Run the generation process"""
        try:
            self.log_signal.emit("=" * 60)
            self.log_signal.emit(f"Mode: {self.mode.upper()}")
            if self.systematic and self.mode == "cpv":
                self.log_signal.emit("Systematic mode: ENABLED")
            
            # Check for interruption before starting
            if self.isInterruptionRequested():
                self.log_signal.emit("Generation cancelled by user")
                self.finished_signal.emit(False, "Generation cancelled")
                return
            
            if self.mode == "cpv":
                success, message = run_cpv_generation(self)
                if not self.isInterruptionRequested():
                    self.finished_signal.emit(success, message)
                else:
                    self.log_signal.emit("Generation cancelled by user")
                    self.finished_signal.emit(False, "Generation cancelled")
            elif self.mode == "drc":
                success, message = run_drc_generation(self)
                if not self.isInterruptionRequested():
                    self.finished_signal.emit(success, message)
                else:
                    self.log_signal.emit("Generation cancelled by user")
                    self.finished_signal.emit(False, "Generation cancelled")
            elif self.mode == "loopDefined":
                from GUI.gui_loopDefined import run_loopDefined_generation
                success, message = run_loopDefined_generation(self)
                if not self.isInterruptionRequested():
                    self.finished_signal.emit(success, message)
                else:
                    self.log_signal.emit("Generation cancelled by user")
                    self.finished_signal.emit(False, "Generation cancelled")
            elif self.mode == "dragdrop":
                from GUI.gui_dragdrop import run_dragdrop_generation
                success, message = run_dragdrop_generation(self)
                if not self.isInterruptionRequested():
                    self.finished_signal.emit(success, message)
                else:
                    self.log_signal.emit("Generation cancelled by user")
                    self.finished_signal.emit(False, "Generation cancelled")
                
        except Exception as e:
            self.log_signal.emit("=" * 60)
            self.log_signal.emit("❌ An error occurred during slide generation!")
            self.log_signal.emit("=" * 60)
            self.log_signal.emit("Error Details:")
            self.log_signal.emit(f"Error Type: {type(e).__name__}")
            self.log_signal.emit(f"Error Message: {str(e)}")
            self.log_signal.emit("-" * 60)
            self.log_signal.emit("Stack Trace:")
            tb_str = traceback.format_exc()
            self.log_signal.emit(tb_str)
            self.log_signal.emit("=" * 60)
            
            self.finished_signal.emit(False, f"An error occurred: {type(e).__name__}\n{str(e)}")


class KeynoteSlideGeneratorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        # Window settings
        self.setWindowTitle("Silde Maker for Keynote v2.2")
        self.setGeometry(100, 100, 750, 900)
        self.setMinimumSize(750, 900)
        
        # Variables
        self.mode = "cpv"
        self.systematic = False
        self.input_dir = cpv_config["input_dir"]
        self.output_file = cpv_config["output_file"]
        self.year = cpv_config["runPeriod"]
        self.channel = cpv_config["channel"]
        
        # Dictionary for storing checkboxes
        self.object_checks = {}
        self.energy_checks = {}
        self.step_checks = {}
        
        # Settings storage
        self.user_settings = None  # Will store settings from SettingsDialog
        
        # Thread
        self.generator_thread = None
        
        # High/Low energy values for DRC mode
        self.high_energy_values = [
            ["10GeV", "20GeV", "30GeV", "40GeV", "50GeV"],
            ["60GeV", "70GeV", "80GeV", "90GeV", "100GeV"],
            ["110GeV", "120GeV"],
        ]
        
        self.low_energy_values = [
            ["0.5GeV", "1.0GeV", "1.5GeV", "2.0GeV", "2.5GeV"],
            ["3.0GeV", "3.5GeV", "4.0GeV", "4.5GeV", "5.0GeV"],
        ]
        
        # Create UI
        self.init_ui()
    
    def init_ui(self):
        """Initialize UI components"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # ============ Top Bar with Settings and Manual buttons ============
        top_bar_layout = QHBoxLayout()
        top_bar_layout.addStretch()
        
        # Settings button (gear icon)
        self.settings_btn = QPushButton("⚙")
        self.settings_btn.setFixedSize(20, 20)
        self.settings_btn.setToolTip("Settings")
        self.settings_btn.setFont(QFont("Arial", 16))
        self.settings_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """)
        self.settings_btn.clicked.connect(self.show_settings)
        
        # Manual button (book icon)
        self.manual_btn = QPushButton("📖")
        self.manual_btn.setFixedSize(20, 20)
        self.manual_btn.setToolTip("Manual")
        self.manual_btn.setFont(QFont("Arial", 16))
        self.manual_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """)
        self.manual_btn.clicked.connect(self.show_manual)
        
        top_bar_layout.addWidget(self.settings_btn)
        top_bar_layout.addWidget(self.manual_btn)
        
        main_layout.addLayout(top_bar_layout)
        
        # ============ Mode Selection ============
        mode_group = QGroupBox("Mode")
        mode_layout = QHBoxLayout()
        
        self.mode_group = QButtonGroup()
        self.cpv_radio = QRadioButton("CPV")
        self.drc_radio = QRadioButton("DRC")
        self.loopDefined_radio = QRadioButton("Loop-defined")
        self.dragdrop_radio = QRadioButton("Drag and Drop")
        self.cpv_radio.setChecked(True)
        
        self.mode_group.addButton(self.cpv_radio)
        self.mode_group.addButton(self.drc_radio)
        self.mode_group.addButton(self.loopDefined_radio)
        self.mode_group.addButton(self.dragdrop_radio)
        
        self.cpv_radio.toggled.connect(self.on_mode_change)
        self.drc_radio.toggled.connect(self.on_mode_change)
        self.loopDefined_radio.toggled.connect(self.on_mode_change)
        self.dragdrop_radio.toggled.connect(self.on_mode_change)
        
        self.systematic_check = QCheckBox("Systematic")
        self.systematic_check.toggled.connect(self.on_systematic_change)
        
        # Preset dropdown and folder button for Loop-defined mode
        self.preset_dropdown = QComboBox()
        self.preset_dropdown.setMinimumWidth(150)
        self.preset_dropdown.addItem("None")
        self.preset_dropdown.currentTextChanged.connect(self.on_preset_selected)
        # Folder button (📁 icon)
        self.preset_folder_btn = QPushButton("📁")
        self.preset_folder_btn.setToolTip("Select preset folder")
        self.preset_folder_btn.setFixedSize(30, 30)
        self.preset_folder_btn.clicked.connect(self.on_preset_folder_clicked)
        # Initially hidden (will be shown in Loop-defined mode)
        self.preset_dropdown.setVisible(False)
        self.preset_folder_btn.setVisible(False)
        
        mode_layout.addWidget(self.cpv_radio)
        mode_layout.addWidget(self.drc_radio)
        mode_layout.addWidget(self.loopDefined_radio)
        mode_layout.addWidget(self.dragdrop_radio)
        mode_layout.addStretch()
        # Preset UI (for Loop-defined mode)
        mode_layout.addWidget(self.preset_dropdown)  # Dropdown
        mode_layout.addWidget(self.preset_folder_btn)  # Folder button (📁)
        # Systematic checkbox (for CPV mode)
        mode_layout.addWidget(self.systematic_check)
        
        mode_group.setLayout(mode_layout)
        main_layout.addWidget(mode_group)
        
        # ============ Directory Settings ============
        (self.dir_group, self.input_dir_edit, self.output_file_edit) = create_directory_ui(self)
        main_layout.addWidget(self.dir_group)
        
        # ============ CPV Options ============
        (self.year_channel_group, self.year_combo, self.channel_combo, 
         self.cpv_file_format_combo) = create_cpv_options_ui(self)
        main_layout.addWidget(self.year_channel_group)
        
        # ============ DRC Options ============
        (self.drc_options_group, self.particle_combo, self.case_combo, 
         self.file_format_combo) = create_drc_options_ui(self)
        main_layout.addWidget(self.drc_options_group)
        self.drc_options_group.setVisible(False)
        
        # ============ Loop-defined Options ============
        (self.loopDefined_options_group, self.loopDefined_file_format_combo, 
         self.loopDefined_path_template_input, self.loopDefined_filename_template_input,
         self.loopDefined_scan_mode_check) = create_loopDefined_options_ui(self)
        # Connect template input changes to update variable inputs
        self.loopDefined_path_template_input.textChanged.connect(self.update_loopDefined_template_info)
        # Also update axis dropdowns when path template changes
        self.loopDefined_path_template_input.textChanged.connect(self._update_loopDefined_axis_dropdowns)
        self.loopDefined_filename_template_input.textChanged.connect(self.update_loopDefined_template_info)
        # Also update axis dropdowns when filename template changes
        self.loopDefined_filename_template_input.textChanged.connect(self._update_loopDefined_axis_dropdowns)
        # Connect scan mode checkbox to update SCAN button visibility and path template input state
        self.loopDefined_scan_mode_check.stateChanged.connect(self._update_scan_button_visibility)
        self.loopDefined_scan_mode_check.stateChanged.connect(self._update_path_template_input_state)
        # Initialize path template input state
        self._update_path_template_input_state()
        main_layout.addWidget(self.loopDefined_options_group)
        self.loopDefined_options_group.setVisible(False)
        
        # ============ Object and Step Selection ============
        selection_layout = QHBoxLayout()
        
        # RightLayout (Objects in CPV / Channel & Energy in DRC)
        self.right_layout = QGroupBox("Objects")
        object_layout = QVBoxLayout()
        
        # CPV Objects
        (self.cpv_objects_widget, cpv_object_checks, 
         self.after_top_reco_objects) = create_cpv_objects_ui(self)
        self.object_checks.update(cpv_object_checks)
        object_layout.addWidget(self.cpv_objects_widget)
        
        # DRC Channels & Energy
        (self.drc_channels_widget, self.c_check, self.s_combo, self.sc_overlay_check,
         self.drcor_combo, drc_energy_checks, self.low_energy_check,
         self.energy_checkboxes_layout, self.resol_with_noise_check,
         self.resol_without_noise_check, self.lin_check) = create_drc_channels_ui(self)
        self.energy_checks.update(drc_energy_checks)
        object_layout.addWidget(self.drc_channels_widget)
        self.drc_channels_widget.setVisible(False)
        
        # Loop-defined Objects & Variables
        (self.loopDefined_objects_widget, loopDefined_object_inputs) = create_loopDefined_objects_ui(self)
        # Store separately for Loop-defined mode (text input instead of checkboxes)
        self.object_inputs = loopDefined_object_inputs
        # Connect object input to sync with variable input
        if 'object' in self.object_inputs:
            self.object_inputs['object'].textChanged.connect(self._on_object_input_changed)
        object_layout.addWidget(self.loopDefined_objects_widget)
        self.loopDefined_objects_widget.setVisible(False)
        
        (self.loopDefined_variables_widget, self.loopDefined_variable_inputs) = create_loopDefined_variables_ui(self)
        object_layout.addWidget(self.loopDefined_variables_widget)
        self.loopDefined_variables_widget.setVisible(False)
        
        # Create initial high energy checkboxes for DRC
        create_energy_checkboxes(self, self.high_energy_values)
        
        self.right_layout.setLayout(object_layout)
        selection_layout.addWidget(self.right_layout)
        
        # LeftLayout container (Steps in CPV / PID+Preview in DRC)
        left_side_layout = QVBoxLayout()
        
        # Header with "Event selection" label and "Select all" checkbox (outside QGroupBox)
        # Only visible in CPV mode
        header_layout = QHBoxLayout()
        self.header_label = QLabel("Event selection")
        header_layout.addWidget(self.header_label)
        header_layout.addStretch()
        # Select all checkbox for steps
        self.select_all_steps_check = QCheckBox("Select all")
        self.select_all_steps_check.setChecked(False)
        header_layout.addWidget(self.select_all_steps_check)
        left_side_layout.addLayout(header_layout)
        
        # LeftLayout main section (Step Selection / PID)
        self.left_layout = QGroupBox()
        step_main_layout = QVBoxLayout()
        
        # CPV Steps
        self.cpv_steps_widget, cpv_step_checks = create_cpv_steps_ui(self)
        self.step_checks.update(cpv_step_checks)
        step_main_layout.addWidget(self.cpv_steps_widget)
        
        # DRC PID - removed (no longer needed)
        
        # Loop-defined Steps (no longer used - steps are handled in Variables section)
        # Removed: Steps are now input directly in Variables section
        
        self.left_layout.setLayout(step_main_layout)
        left_side_layout.addWidget(self.left_layout)
        
        # Preview Section (DRC and Loop-defined)
        self.preview_group, self.energy_preview_cells, self.drc_page_tabs, self.drc_tab_grids, self.drc_tab_cells = create_drc_preview_ui(self)
        left_side_layout.addWidget(self.preview_group)
        self.preview_group.setVisible(False)
        
        # Loop-defined Preview Section
        (self.loopDefined_preview_group, self.loopDefined_preview_cells, 
         self.loopDefined_arrangement_combo, self.loopDefined_row_inputs_container,
         self.loopDefined_row_inputs_layout, self.loopDefined_column_input,
         self.loopDefined_preview_grid_layout, self.loopDefined_drc_plot_check,
         self.loopDefined_drc_energy_container, self.loopDefined_energy_checkboxes_layout,
         self.loopDefined_drc_controls_container, self.loopDefined_slide_axis_combo,
         self.loopDefined_grid_axis_combo, self.loopDefined_page_tabs) = create_loopDefined_preview_ui(self)
        left_side_layout.addWidget(self.loopDefined_preview_group)
        self.loopDefined_preview_group.setVisible(False)
        
        # Drag and Drop Preview Section
        (self.dragdrop_preview_group, self.dragdrop_preview_cells,
         self.dragdrop_arrangement_combo, self.dragdrop_slide_count_combo,
         self.dragdrop_slide_count_input, self.dragdrop_page_tabs,
         self.dragdrop_tab_grids, self.dragdrop_tab_cells) = create_dragdrop_preview_ui(self)
        left_side_layout.addWidget(self.dragdrop_preview_group)
        self.dragdrop_preview_group.setVisible(False)
        
        # Store row inputs dictionary
        self.loopDefined_row_inputs = {}
        
        # Store energy checks dictionary for DRC-style preview
        self.loopDefined_energy_checks = {}
        
        # Connect arrangement change to update row inputs
        self.loopDefined_arrangement_combo.currentTextChanged.connect(self.on_loopDefined_arrangement_changed)
        # Connect column input changes to update preview
        self.loopDefined_column_input.textChanged.connect(self.update_loopDefined_preview)
        # Connect DRC plot checkbox to toggle preview style
        self.loopDefined_drc_plot_check.stateChanged.connect(self.on_loopDefined_drc_plot_changed)
        # Connect axis dropdowns to update preview
        self.loopDefined_slide_axis_combo.currentTextChanged.connect(self.on_loopDefined_axis_changed)
        self.loopDefined_grid_axis_combo.currentTextChanged.connect(self.on_loopDefined_axis_changed)
        self.loopDefined_page_tabs.currentChanged.connect(self.on_loopDefined_page_changed)
        
        # Store scan results for DRC plot mode
        self.loopDefined_scan_results = None
        self.loopDefined_confirmed_scan_results = None
        
        # Initialize row inputs based on default arrangement
        self.on_loopDefined_arrangement_changed("2*3")
        
        # Connect dragdrop arrangement change
        self.dragdrop_arrangement_combo.currentTextChanged.connect(self.on_dragdrop_arrangement_changed)
        # Connect dragdrop slide count change
        self.dragdrop_slide_count_combo.currentTextChanged.connect(self.on_dragdrop_slide_count_changed)
        self.dragdrop_slide_count_input.valueChanged.connect(self.on_dragdrop_slide_count_changed)
        # Initialize dragdrop preview
        self.on_dragdrop_arrangement_changed("2*3")
        
        selection_layout.addLayout(left_side_layout)
        main_layout.addLayout(selection_layout)
        
        # ============ Action Buttons ============
        button_layout = create_action_buttons(self)
        main_layout.addLayout(button_layout)
        
        # ============ Status/Log Area ============
        status_group, self.log_text = create_status_ui()
        main_layout.addWidget(status_group)
        
        # ============ Connect Select All Checkboxes ============
        self.connect_select_all_handlers()
        
        # Connect Loop-defined mode select all checkboxes
        if hasattr(self, 'select_all_loopDefined_check'):
            self.select_all_loopDefined_check.toggled.connect(
                lambda checked: self._toggle_all_loopDefined_objects(checked)
            )
        # Loop-defined steps select all checkbox removed (steps are now in Variables section)
        
        # Initialize visibility: CPV mode is default, so show header and select all checkbox
        # (DRC mode will hide them in on_mode_change)
        if hasattr(self, 'header_label'):
            self.header_label.setVisible(True)
        if hasattr(self, 'select_all_steps_check'):
            self.select_all_steps_check.setVisible(True)
    
    def on_mode_change(self):
        """Handle mode change between CPV, DRC, Loop-defined, and Drag and Drop"""
        if self.cpv_radio.isChecked():
            self.mode = "cpv"
            self.input_dir_edit.setText(cpv_config["input_dir"])
            self.output_file_edit.setText(cpv_config["output_file"])
            self.systematic_check.setEnabled(True)
            self.systematic_check.setVisible(True)  # Show systematic checkbox in CPV mode
            # Show directory settings in CPV mode
            self.dir_group.setVisible(True)
            # Hide preset UI in CPV mode
            self.preset_dropdown.setVisible(False)
            self.preset_folder_btn.setVisible(False)
            self.year_channel_group.setVisible(True)
            self.drc_options_group.setVisible(False)
            self.loopDefined_options_group.setVisible(False)
            # Hide dragdrop preview in CPV mode
            self.dragdrop_preview_group.setVisible(False)
            # RightLayout: Objects
            self.right_layout.setTitle("Objects")
            self.right_layout.setVisible(True)  # Show right_layout in CPV mode
            self.cpv_objects_widget.setVisible(True)
            self.drc_channels_widget.setVisible(False)
            self.loopDefined_objects_widget.setVisible(False)
            self.loopDefined_variables_widget.setVisible(False)
            # LeftLayout: No title (header label is outside)
            self.left_layout.setTitle("")
            self.left_layout.setVisible(True)  # Show left_layout in CPV mode
            # Show header label and select all checkbox for CPV mode
            self.header_label.setVisible(True)
            self.select_all_steps_check.setVisible(True)
            self.cpv_steps_widget.setVisible(True)
            self.preview_group.setVisible(False)
            self.loopDefined_preview_group.setVisible(False)  # Hide Loop-defined preview in CPV mode
            # Hide SCAN button in CPV mode
            if hasattr(self, 'scan_button'):
                self.scan_button.setVisible(False)
        elif self.drc_radio.isChecked():
            self.mode = "drc"
            self.systematic_check.setChecked(False)
            self.systematic_check.setEnabled(False)
            self.systematic_check.setVisible(False)  # Hide systematic checkbox in DRC mode
            # Show directory settings in DRC mode
            self.dir_group.setVisible(True)
            # Hide preset UI in DRC mode
            self.preset_dropdown.setVisible(False)
            self.preset_folder_btn.setVisible(False)
            self.input_dir_edit.setText(drc_config["base_dir"])
            self.output_file_edit.setText(drc_config["output_file"])
            self.year_channel_group.setVisible(False)
            self.drc_options_group.setVisible(True)
            self.loopDefined_options_group.setVisible(False)
            # Hide dragdrop preview in DRC mode
            self.dragdrop_preview_group.setVisible(False)
            # Hide SCAN button in DRC mode
            if hasattr(self, 'scan_button'):
                self.scan_button.setVisible(False)
            # RightLayout: Channel && Energy points
            self.right_layout.setTitle("Channel && Energy points")
            self.right_layout.setVisible(True)  # Show right_layout in DRC mode
            self.cpv_objects_widget.setVisible(False)
            self.drc_channels_widget.setVisible(True)
            self.loopDefined_objects_widget.setVisible(False)
            self.loopDefined_variables_widget.setVisible(False)
            # LeftLayout: Preview (no PID)
            self.left_layout.setTitle("")
            self.left_layout.setVisible(False)  # Hide left_layout in DRC mode (no PID)
            self.cpv_steps_widget.setVisible(False)
            self.preview_group.setVisible(True)
            self.loopDefined_preview_group.setVisible(False)  # Hide Loop-defined preview in DRC mode
            # Hide header label and select all checkbox for DRC mode
            self.header_label.setVisible(False)
            self.select_all_steps_check.setVisible(False)
            update_energy_preview(self)
        elif self.loopDefined_radio.isChecked():  # Loop-defined mode
            self.mode = "loopDefined"
            self.systematic_check.setChecked(False)
            self.systematic_check.setEnabled(False)
            self.systematic_check.setVisible(False)  # Hide systematic checkbox in Loop-defined mode
            # Show directory settings in Loop-defined mode
            self.dir_group.setVisible(True)
            # Show preset UI in Loop-defined mode
            self.preset_dropdown.setVisible(True)
            self.preset_folder_btn.setVisible(True)
            # Load and populate presets
            self._update_preset_dropdown()
            # Use default paths (user will set via directory UI)
            self.year_channel_group.setVisible(False)
            self.drc_options_group.setVisible(False)
            self.loopDefined_options_group.setVisible(True)
            # Hide dragdrop preview in Loop-defined mode
            self.dragdrop_preview_group.setVisible(False)
            # Update SCAN button visibility and path template input state
            self._update_scan_button_visibility()
            self._update_path_template_input_state()
            # RightLayout: Variables only
            self.right_layout.setTitle("Variables")
            self.right_layout.setVisible(True)  # Show right_layout in Loop-defined mode
            self.cpv_objects_widget.setVisible(False)
            self.drc_channels_widget.setVisible(False)
            self.loopDefined_objects_widget.setVisible(False)  # Hide Objects section
            self.loopDefined_variables_widget.setVisible(True)
            # Update template info from settings (this also updates variable inputs)
            # Do this after making widget visible to ensure proper layout update
            self.update_loopDefined_template_info()
            # LeftLayout: Preview (Loop-defined mode)
            self.left_layout.setTitle("")
            self.left_layout.setVisible(False)  # Hide empty left_layout in Loop-defined mode
            self.cpv_steps_widget.setVisible(False)
            self.preview_group.setVisible(False)
            self.loopDefined_preview_group.setVisible(True)
            # Hide header label and select all checkbox for Loop-defined mode
            self.header_label.setVisible(False)
            self.select_all_steps_check.setVisible(False)
            # Update preview when switching to Loop-defined mode
            update_loopDefined_preview(self)
            
            # Sync object and step values to disabled input fields (after variable inputs are created)
            # Use QTimer to ensure this runs after UI is updated
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(100, self._sync_object_step_to_variables)
            
            # Update SCAN button visibility and path template input state based on scan mode checkbox
            self._update_scan_button_visibility()
            self._update_path_template_input_state()
        elif self.dragdrop_radio.isChecked():  # Drag and Drop mode
            self.mode = "dragdrop"
            self.systematic_check.setChecked(False)
            self.systematic_check.setEnabled(False)
            self.systematic_check.setVisible(False)  # Hide systematic checkbox in Drag and Drop mode
            # Hide preset UI in Drag and Drop mode
            self.preset_dropdown.setVisible(False)
            self.preset_folder_btn.setVisible(False)
            # Hide directory settings in Drag and Drop mode
            self.dir_group.setVisible(False)
            # Hide all options groups
            self.year_channel_group.setVisible(False)
            self.drc_options_group.setVisible(False)
            self.loopDefined_options_group.setVisible(False)
            # Hide SCAN button in Drag and Drop mode
            if hasattr(self, 'scan_button'):
                self.scan_button.setVisible(False)
            # Hide right layout (Objects/Variables) in Drag and Drop mode
            self.right_layout.setVisible(False)
            # Hide left layout in Drag and Drop mode
            self.left_layout.setVisible(False)
            # Hide header label and select all checkbox
            self.header_label.setVisible(False)
            self.select_all_steps_check.setVisible(False)
            # Hide other preview groups
            self.preview_group.setVisible(False)
            self.loopDefined_preview_group.setVisible(False)
            # Show dragdrop preview
            self.dragdrop_preview_group.setVisible(True)
            # Update dragdrop preview
            from GUI.gui_dragdrop import update_dragdrop_preview
            update_dragdrop_preview(self)
    
    def on_dragdrop_arrangement_changed(self, arrangement):
        """Handle arrangement change in Drag and Drop mode"""
        from GUI.gui_dragdrop import update_dragdrop_preview
        update_dragdrop_preview(self)
    
    def on_dragdrop_slide_count_changed(self):
        """Handle slide count change in Drag and Drop mode"""
        # Show/hide custom input based on combo selection
        if hasattr(self, 'dragdrop_slide_count_combo'):
            if self.dragdrop_slide_count_combo.currentText() == "Custom":
                if hasattr(self, 'dragdrop_slide_count_input'):
                    self.dragdrop_slide_count_input.setVisible(True)
            else:
                if hasattr(self, 'dragdrop_slide_count_input'):
                    self.dragdrop_slide_count_input.setVisible(False)
        
        # Update preview to reflect new slide count
        from GUI.gui_dragdrop import update_dragdrop_preview
        update_dragdrop_preview(self)
    
    def update_loopDefined_template_info(self):
        """Update user-defined variable input fields based on templates"""
        # Get templates from input fields
        path_template = self.loopDefined_path_template_input.text().strip()
        filename_template = self.loopDefined_filename_template_input.text().strip()
        
        # Use defaults if empty
        default_path_template = "{sample}/{channel}/{object}/{step}"
        default_filename_template = "h_{object}_{kinematic}_{step}.{format}"
        
        if not path_template:
            path_template = default_path_template
            self.loopDefined_path_template_input.setText(path_template)
        if not filename_template:
            filename_template = default_filename_template
            self.loopDefined_filename_template_input.setText(filename_template)
        
        # Extract variables from templates and update variable inputs dynamically
        self._update_loopDefined_variable_inputs(path_template, filename_template)
    
    def _update_loopDefined_variable_inputs(self, path_template, filename_template):
        """Dynamically create variable input fields based on templates"""
        try:
            from Config.Rules.loopDefined_template_engine import create_template_parser
            from PyQt5.QtWidgets import QFormLayout, QLabel, QLineEdit, QFrame
            from PyQt5.QtCore import Qt
            
            parser = create_template_parser(path_template, filename_template)
            
            # Get variables from path and filename templates separately
            path_variables = parser.path_variables
            filename_variables = parser.filename_variables
            
            # Store path_variables for later use (to disable them in scan mode)
            self.loopDefined_path_variables = path_variables
            
            # Remove format from filename variables (handled automatically)
            filename_variables = [v for v in filename_variables if v != 'format']
            
            # Clear existing layout
            if hasattr(self, 'loopDefined_variables_widget'):
                layout = self.loopDefined_variables_widget.layout()
                if layout:
                    # Clear all items
                    while layout.count():
                        item = layout.takeAt(0)
                        if item.widget():
                            item.widget().deleteLater()
                        elif item.layout():
                            # Clear sub-layout
                            sub_layout = item.layout()
                            while sub_layout.count():
                                sub_item = sub_layout.takeAt(0)
                                if sub_item.widget():
                                    sub_item.widget().deleteLater()
                                elif sub_item.layout():
                                    # Clear nested layout
                                    nested_layout = sub_item.layout()
                                    while nested_layout.count():
                                        nested_item = nested_layout.takeAt(0)
                                        if nested_item.widget():
                                            nested_item.widget().deleteLater()
                
                # Path Template Variables Section
                if path_variables:
                    path_label = QLabel("<b>Path Template Variables:</b>")
                    layout.addWidget(path_label)
                    
                    path_form = QFormLayout()
                    self.loopDefined_variable_inputs = {}
                    
                    for var_name in path_variables:
                        var_input = QLineEdit()
                        var_input.setPlaceholderText(f"e.g., value1,value2,value3")
                        var_input.setVisible(True)  # Ensure visibility
                        var_input.setMinimumWidth(200)  # Ensure minimum width
                        # Create label explicitly
                        var_label = QLabel(f"{var_name}:")
                        path_form.addRow(var_label, var_input)
                        self.loopDefined_variable_inputs[var_name] = var_input
                    
                    layout.addLayout(path_form)
                    
                    # Force layout update
                    self.loopDefined_variables_widget.update()
                    self.loopDefined_variables_widget.repaint()
                
                # Separator line
                if path_variables and filename_variables:
                    separator = QFrame()
                    separator.setFrameShape(QFrame.HLine)
                    separator.setFrameShadow(QFrame.Sunken)
                    layout.addWidget(separator)
                
                # Filename Template Variables Section
                if filename_variables:
                    filename_label = QLabel("<b>Filename Template Variables:</b>")
                    layout.addWidget(filename_label)
                    
                    filename_form = QFormLayout()
                    
                    for var_name in filename_variables:
                        # If already added from path template, create a read-only display
                        if var_name in self.loopDefined_variable_inputs:
                            # Get existing input from path template
                            var_input = self.loopDefined_variable_inputs[var_name]
                            # Create a read-only display that syncs with the path template input
                            var_display = QLineEdit()
                            var_display.setEnabled(False)
                            var_display.setStyleSheet("background-color: #f0f0f0; color: #666;")
                            var_display.setToolTip(f"'{var_name}' is already defined in Path Template above. Edit it there.")
                            # Sync display with the actual input
                            def update_display(text, display=var_display):
                                display.setText(text)
                            var_input.textChanged.connect(update_display)
                            # Set initial value
                            var_display.setText(var_input.text())
                            # Add asterisk to label
                            filename_form.addRow(f"{var_name}*:", var_display)
                        else:
                            # New variable, create new input
                            var_input = QLineEdit()
                            var_input.setPlaceholderText(f"e.g., value1,value2,value3")
                            filename_form.addRow(f"{var_name}:", var_input)
                            self.loopDefined_variable_inputs[var_name] = var_input
                    
                    layout.addLayout(filename_form)
                
                layout.addStretch()
                
                # Force widget update to ensure all fields are visible
                self.loopDefined_variables_widget.update()
                self.loopDefined_variables_widget.repaint()
                if hasattr(self, 'loopDefined_variables_widget'):
                    self.loopDefined_variables_widget.show()
                
                # Update path template variables state based on scan mode checkbox
                self._update_path_template_input_state()
                
        except Exception as e:
            # If template parsing fails, keep default variables
            import traceback
            print(f"Error updating variable inputs: {e}")
            traceback.print_exc()
    
    def _sync_object_step_to_variables(self):
        """Sync object value from Objects section to disabled variable input field (no longer needed)"""
        # Objects section is removed, so this function is no longer needed
        pass
    
    def _on_object_input_changed(self):
        """Handle object input text change - sync to variable input (no longer needed)"""
        # Objects section is removed, so this function is no longer needed
        pass
    
    def toggle_energy_range(self):
        """Toggle between high energy (10-120 GeV) and low energy (0.5-5 GeV)"""
        if self.low_energy_check.isChecked():
            create_energy_checkboxes(self, self.low_energy_values)
        else:
            create_energy_checkboxes(self, self.high_energy_values)
        update_energy_preview(self)
    
    def update_energy_preview(self):
        """Update energy preview - delegates to gui_drc module"""
        update_energy_preview(self)
    
    def on_loopDefined_arrangement_changed(self, arrangement):
        """Handle arrangement change in Loop-defined mode"""
        from GUI.gui_loopDefined import update_loopDefined_row_inputs
        update_loopDefined_row_inputs(self, arrangement)
        self.update_loopDefined_preview()
    
    def update_loopDefined_preview(self):
        """Update Loop-defined preview - delegates to gui_loopDefined module"""
        from GUI.gui_loopDefined import update_loopDefined_preview
        update_loopDefined_preview(self)
    
    def on_loopDefined_drc_plot_changed(self, state):
        """Handle DRC plot checkbox change in Loop-defined mode"""
        drc_plot_enabled = (state == 2)  # Qt.Checked = 2
        
        if drc_plot_enabled:
            # Show DRC controls (axis dropdowns and tabs)
            if hasattr(self, 'loopDefined_drc_controls_container'):
                self.loopDefined_drc_controls_container.setVisible(True)
            
            # Hide row/column inputs container
            if hasattr(self, 'loopDefined_row_inputs_container'):
                self.loopDefined_row_inputs_container.setVisible(False)
            
            # Show info label
            if not hasattr(self, 'loopDefined_drc_info_label'):
                from PyQt5.QtWidgets import QLabel
                from PyQt5.QtCore import Qt
                info_label = QLabel("DRC plot mode uses Slide axis + Grid axis.")
                info_label.setStyleSheet("color: #666; padding: 5px; font-size: 10pt;")
                info_label.setAlignment(Qt.AlignCenter)
                # Insert in preview layout before preview grid
                if hasattr(self, 'loopDefined_preview_group'):
                    preview_layout = self.loopDefined_preview_group.layout()
                    if preview_layout:
                        # Find the index of preview grid widget
                        grid_widget = None
                        for i in range(preview_layout.count()):
                            item = preview_layout.itemAt(i)
                            if item and item.widget():
                                widget = item.widget()
                                if hasattr(widget, 'layout') and widget.layout() == self.loopDefined_preview_grid_layout:
                                    grid_widget = widget
                                    break
                        if grid_widget:
                            # Insert before grid widget
                            preview_layout.insertWidget(preview_layout.indexOf(grid_widget), info_label)
                        else:
                            # If grid widget not found, add at the end
                            preview_layout.insertWidget(preview_layout.count() - 1, info_label)
                self.loopDefined_drc_info_label = info_label
            else:
                self.loopDefined_drc_info_label.setVisible(True)
            
            # Update axis dropdowns with available variables
            self._update_loopDefined_axis_dropdowns()
            
            # Update preview (DRC mode)
            if self.loopDefined_confirmed_scan_results:
                self._update_loopDefined_drc_preview()
        else:
            # DRC plot disabled - hide DRC controls and show row/column inputs
            if hasattr(self, 'loopDefined_drc_controls_container'):
                self.loopDefined_drc_controls_container.setVisible(False)
            
            # Show row/column inputs container
            if hasattr(self, 'loopDefined_row_inputs_container'):
                self.loopDefined_row_inputs_container.setVisible(True)
            
            # Hide info label
            if hasattr(self, 'loopDefined_drc_info_label'):
                self.loopDefined_drc_info_label.setVisible(False)
            
            # Clear and hide tabs
            if hasattr(self, 'loopDefined_page_tabs'):
                # Block signals to prevent recursion
                self.loopDefined_page_tabs.blockSignals(True)
                # Clear all tabs
                while self.loopDefined_page_tabs.count():
                    self.loopDefined_page_tabs.removeTab(0)
                # Hide tabs
                self.loopDefined_page_tabs.setVisible(False)
                self.loopDefined_page_tabs.blockSignals(False)
            
            # Clear tab data
            if hasattr(self, 'loopDefined_tab_grids'):
                for tab_cells in self.loopDefined_tab_cells:
                    for cell in tab_cells:
                        if cell:
                            cell.deleteLater()
                self.loopDefined_tab_grids.clear()
                self.loopDefined_tab_cells.clear()
            
            # Use normal preview (row/col based)
            self.update_loopDefined_preview()
    
    def on_loopDefined_axis_changed(self):
        """Handle axis dropdown change"""
        # Only handle axis changes when DRC plot is enabled
        if hasattr(self, 'loopDefined_drc_plot_check') and self.loopDefined_drc_plot_check.isChecked():
            if self.loopDefined_confirmed_scan_results:
                self._update_loopDefined_drc_preview()
        # If DRC plot is not enabled, axis changes are ignored (use normal preview)
    
    def on_loopDefined_page_changed(self, index):
        """Handle page tab change in DRC plot mode"""
        # Tab change doesn't need to update preview - each tab already has its own grid
        # The preview is already created when tabs are generated
        # No action needed - just switching between already-created tabs
        pass
    
    def _update_loopDefined_drc_preview(self):
        """Update preview for DRC plot mode based on slide axis and grid axis"""
        if not hasattr(self, 'loopDefined_confirmed_scan_results') or not self.loopDefined_confirmed_scan_results:
            return
        
        if not hasattr(self, 'loopDefined_slide_axis_combo') or not hasattr(self, 'loopDefined_grid_axis_combo'):
            return
        
        slide_axis = self.loopDefined_slide_axis_combo.currentText()
        grid_axis = self.loopDefined_grid_axis_combo.currentText()
        
        if not slide_axis or not grid_axis or slide_axis == grid_axis:
            # Invalid selection - clear preview and hide main grid
            if hasattr(self, 'loopDefined_preview_cells'):
                for cell in self.loopDefined_preview_cells:
                    if cell:
                        cell.deleteLater()
                self.loopDefined_preview_cells.clear()
            if hasattr(self, 'loopDefined_preview_grid_layout'):
                while self.loopDefined_preview_grid_layout.count():
                    item = self.loopDefined_preview_grid_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()
            
            # Hide tabs if they exist
            if hasattr(self, 'loopDefined_page_tabs'):
                self.loopDefined_page_tabs.setVisible(False)
            
            # Hide the main preview grid widget
            if hasattr(self, 'loopDefined_preview_group'):
                preview_layout = self.loopDefined_preview_group.layout()
                if preview_layout:
                    for i in range(preview_layout.count()):
                        item = preview_layout.itemAt(i)
                        if item and item.widget():
                            widget = item.widget()
                            # Check if this widget contains the preview_grid_layout
                            if hasattr(widget, 'layout') and widget.layout() == self.loopDefined_preview_grid_layout:
                                widget.setVisible(False)
                                break
            return
        
        # Get arrangement
        if hasattr(self, 'loopDefined_arrangement_combo'):
            arrangement = self.loopDefined_arrangement_combo.currentText()
            try:
                num_rows, num_cols = map(int, arrangement.split('*'))
            except:
                num_rows, num_cols = 2, 3
        else:
            num_rows, num_cols = 2, 3
        
        # Get scan results
        scan_results = self.loopDefined_confirmed_scan_results
        all_results = scan_results['all_results']
        variable_combinations = scan_results['variable_combinations']
        
        # Get slide axis values (from variable_combinations or from scan results)
        slide_axis_values = []
        if slide_axis in variable_combinations:
            slide_axis_values = variable_combinations[slide_axis]
        else:
            # Extract unique values from scan results
            slide_values_set = set()
            for file_path, parsed_vars, status in all_results:
                if slide_axis in parsed_vars:
                    slide_values_set.add(parsed_vars[slide_axis])
            slide_axis_values = sorted(list(slide_values_set))
        
        # Get grid axis values
        grid_axis_values = []
        if grid_axis in variable_combinations:
            grid_axis_values = variable_combinations[grid_axis]
        else:
            # Extract unique values from scan results
            grid_values_set = set()
            for file_path, parsed_vars, status in all_results:
                if grid_axis in parsed_vars:
                    grid_values_set.add(parsed_vars[grid_axis])
            grid_axis_values = sorted(list(grid_values_set))
        
        # Create matching dictionary: (slide_axis_value, grid_axis_value) -> (file_path, status)
        matching_dict = {}
        for file_path, parsed_vars, status in all_results:
            if slide_axis in parsed_vars and grid_axis in parsed_vars:
                key = (parsed_vars[slide_axis], parsed_vars[grid_axis])
                if key not in matching_dict:
                    matching_dict[key] = []
                matching_dict[key].append((file_path, parsed_vars, status))
        
        # Create page tabs with grids
        if hasattr(self, 'loopDefined_page_tabs'):
            from PyQt5.QtWidgets import QWidget, QGridLayout
            from GUI.common import DraggableLabel
            from PyQt5.QtCore import Qt
            
            # Block signals to prevent recursion during tab creation
            self.loopDefined_page_tabs.blockSignals(True)
            
            # Clear existing tabs
            while self.loopDefined_page_tabs.count():
                self.loopDefined_page_tabs.removeTab(0)
            
            # Store tab grids and cells for each page
            if not hasattr(self, 'loopDefined_tab_grids'):
                self.loopDefined_tab_grids = []
            if not hasattr(self, 'loopDefined_tab_cells'):
                self.loopDefined_tab_cells = []
            
            # Clear previous tab data
            for tab_cells in self.loopDefined_tab_cells:
                for cell in tab_cells:
                    if cell:
                        cell.deleteLater()
            self.loopDefined_tab_grids.clear()
            self.loopDefined_tab_cells.clear()
            
            # Create tabs for each slide axis value
            for slide_idx, slide_value in enumerate(slide_axis_values):
                # Create tab widget
                tab_widget = QWidget()
                tab_layout = QVBoxLayout()
                tab_layout.setContentsMargins(0, 0, 0, 0)
                
                # Create grid layout for this tab
                tab_grid_layout = QGridLayout()
                tab_grid_layout.setSpacing(5)
                tab_grid_layout.setContentsMargins(5, 5, 5, 5)
                
                # Create grid widget
                grid_widget = QWidget()
                grid_widget.setLayout(tab_grid_layout)
                tab_layout.addWidget(grid_widget)
                
                tab_widget.setLayout(tab_layout)
                self.loopDefined_page_tabs.addTab(tab_widget, f"{slide_axis}={slide_value}")
                
                # Store grid layout
                self.loopDefined_tab_grids.append(tab_grid_layout)
                self.loopDefined_tab_cells.append([])
                
                # Create grid cells for this tab
                grid_size = num_rows * num_cols
                cell_idx = 0
                
                for grid_value in grid_axis_values:
                    if cell_idx >= grid_size:
                        break  # Grid is full
                    
                    row = cell_idx // num_cols
                    col = cell_idx % num_cols
                    
                    # Find matching file
                    key = (slide_value, grid_value)
                    matches = matching_dict.get(key, [])
                    
                    # Determine cell status and text
                    if len(matches) == 0:
                        cell_text = f"Not found\n{grid_axis}={grid_value}"
                        cell_status = "not_found"
                    elif len(matches) > 1:
                        cell_text = f"Ambiguous\n{grid_axis}={grid_value}"
                        cell_status = "ambiguous"
                    else:
                        file_path, parsed_vars, status = matches[0]
                        cell_text = f"{grid_axis}={grid_value}"
                        cell_status = "matched"
                    
                    # Create cell
                    cell = DraggableLabel(cell_text, self, cell_list=self.loopDefined_tab_cells[slide_idx])
                    cell.setAlignment(Qt.AlignCenter)
                    cell.setMinimumSize(100, 60)
                    
                    # Set style based on status
                    if cell_status == "not_found":
                        cell.setStyleSheet(
                            "border: 2px dashed #ff6b6b; "
                            "background-color: #ffe0e0; "
                            "color: #cc0000; "
                            "border-radius: 5px; "
                            "padding: 5px;"
                        )
                    elif cell_status == "ambiguous":
                        cell.setStyleSheet(
                            "border: 2px solid #ffa500; "
                            "background-color: #fff4e0; "
                            "color: #cc6600; "
                            "font-weight: bold; "
                            "border-radius: 5px; "
                            "padding: 5px;"
                        )
                    else:  # matched
                        cell.setStyleSheet(
                            "border: 2px solid #0066cc; "
                            "background-color: #e6f2ff; "
                            "color: #0066cc; "
                            "font-weight: bold; "
                            "border-radius: 5px; "
                            "padding: 5px;"
                        )
                    
                    tab_grid_layout.addWidget(cell, row, col)
                    self.loopDefined_tab_cells[slide_idx].append(cell)
                    
                    cell_idx += 1
                
                # Fill remaining cells with "Empty"
                while cell_idx < grid_size:
                    row = cell_idx // num_cols
                    col = cell_idx % num_cols
                    
                    cell = DraggableLabel("Empty", self, cell_list=self.loopDefined_tab_cells[slide_idx])
                    cell.setAlignment(Qt.AlignCenter)
                    cell.setMinimumSize(100, 60)
                    cell.setStyleSheet(
                        "border: 2px dashed #ccc; "
                        "background-color: #f5f5f5; "
                        "color: #999; "
                        "border-radius: 5px; "
                        "padding: 5px;"
                    )
                    
                    tab_grid_layout.addWidget(cell, row, col)
                    self.loopDefined_tab_cells[slide_idx].append(cell)
                    
                    cell_idx += 1
            
            self.loopDefined_page_tabs.setVisible(len(slide_axis_values) > 0)
            
            # Unblock signals after tab creation is complete
            self.loopDefined_page_tabs.blockSignals(False)
        
        # Hide the main preview grid widget when using tabs (DRC plot mode)
        if hasattr(self, 'loopDefined_preview_group'):
            preview_layout = self.loopDefined_preview_group.layout()
            if preview_layout:
                for i in range(preview_layout.count()):
                    item = preview_layout.itemAt(i)
                    if item and item.widget():
                        widget = item.widget()
                        # Check if this widget contains the preview_grid_layout
                        if hasattr(widget, 'layout') and widget.layout() == self.loopDefined_preview_grid_layout:
                            widget.setVisible(False)
                            break
    
    def _update_loopDefined_preview_with_tabs(self):
        """Update preview with tabs based on row/column values and axis dropdowns (when DRC plot is not checked)"""
        if not hasattr(self, 'loopDefined_row_inputs') or not hasattr(self, 'loopDefined_column_input'):
            return
        
        if not hasattr(self, 'loopDefined_slide_axis_combo') or not hasattr(self, 'loopDefined_grid_axis_combo'):
            return
        
        slide_axis = self.loopDefined_slide_axis_combo.currentText()
        grid_axis = self.loopDefined_grid_axis_combo.currentText()
        
        if not slide_axis or not grid_axis or slide_axis == grid_axis:
            # Invalid selection - use normal preview
            self.update_loopDefined_preview()
            return
        
        # Get arrangement
        if hasattr(self, 'loopDefined_arrangement_combo'):
            arrangement = self.loopDefined_arrangement_combo.currentText()
            try:
                num_rows, num_cols = map(int, arrangement.split('*'))
            except:
                num_rows, num_cols = 2, 3
        else:
            num_rows, num_cols = 2, 3
        
        # Get row and column values
        column_text = self.loopDefined_column_input.text().strip()
        column_values = [v.strip() for v in column_text.split(',') if v.strip()] if column_text else []
        
        row_values_list = []
        for i in range(num_rows):
            if i in self.loopDefined_row_inputs:
                row_text = self.loopDefined_row_inputs[i].text().strip()
                row_values = [v.strip() for v in row_text.split(',') if v.strip()] if row_text else []
                row_values_list.append(row_values)
            else:
                row_values_list.append([])
        
        # Map row/column values to slide axis and grid axis
        # For now, we'll use row values as slide axis and column values as grid axis
        # This is a simple mapping - can be enhanced later
        slide_axis_values = []
        grid_axis_values = column_values
        
        # Get slide axis values from row inputs
        # Combine all row values as slide axis values
        for row_values in row_values_list:
            for val in row_values:
                if val and val not in slide_axis_values:
                    slide_axis_values.append(val)
        
        # If no slide axis values from rows, use column values as slide axis
        if not slide_axis_values:
            slide_axis_values = column_values[:num_rows] if len(column_values) >= num_rows else column_values
            grid_axis_values = column_values[num_rows:] if len(column_values) > num_rows else []
        
        # Create page tabs with grids
        if hasattr(self, 'loopDefined_page_tabs'):
            from PyQt5.QtWidgets import QWidget, QGridLayout, QVBoxLayout
            from GUI.common import DraggableLabel
            from PyQt5.QtCore import Qt
            
            # Block signals to prevent recursion during tab creation
            self.loopDefined_page_tabs.blockSignals(True)
            
            # Clear existing tabs
            while self.loopDefined_page_tabs.count():
                self.loopDefined_page_tabs.removeTab(0)
            
            # Store tab grids and cells for each page
            if not hasattr(self, 'loopDefined_tab_grids'):
                self.loopDefined_tab_grids = []
            if not hasattr(self, 'loopDefined_tab_cells'):
                self.loopDefined_tab_cells = []
            
            # Clear previous tab data
            for tab_cells in self.loopDefined_tab_cells:
                for cell in tab_cells:
                    if cell:
                        cell.deleteLater()
            self.loopDefined_tab_grids.clear()
            self.loopDefined_tab_cells.clear()
            
            # Create tabs for each slide axis value
            for slide_idx, slide_value in enumerate(slide_axis_values):
                # Create tab widget
                tab_widget = QWidget()
                tab_layout = QVBoxLayout()
                tab_layout.setContentsMargins(0, 0, 0, 0)
                
                # Create grid layout for this tab
                tab_grid_layout = QGridLayout()
                tab_grid_layout.setSpacing(5)
                tab_grid_layout.setContentsMargins(5, 5, 5, 5)
                
                # Create grid widget
                grid_widget = QWidget()
                grid_widget.setLayout(tab_grid_layout)
                tab_layout.addWidget(grid_widget)
                
                tab_widget.setLayout(tab_layout)
                self.loopDefined_page_tabs.addTab(tab_widget, f"{slide_axis}={slide_value}")
                
                # Store grid layout
                self.loopDefined_tab_grids.append(tab_grid_layout)
                self.loopDefined_tab_cells.append([])
                
                # Create grid cells for this tab
                grid_size = num_rows * num_cols
                cell_idx = 0
                
                for grid_value in grid_axis_values:
                    if cell_idx >= grid_size:
                        break  # Grid is full
                    
                    row = cell_idx // num_cols
                    col = cell_idx % num_cols
                    
                    # Create cell text
                    cell_text = f"{grid_axis}={grid_value}"
                    cell_status = "matched"
                    
                    # Create cell
                    cell = DraggableLabel(cell_text, self, cell_list=self.loopDefined_tab_cells[slide_idx])
                    cell.setAlignment(Qt.AlignCenter)
                    cell.setMinimumSize(100, 60)
                    
                    # Set style
                    cell.setStyleSheet(
                        "border: 2px solid #0066cc; "
                        "background-color: #e6f2ff; "
                        "color: #0066cc; "
                        "font-weight: bold; "
                        "border-radius: 5px; "
                        "padding: 5px;"
                    )
                    
                    tab_grid_layout.addWidget(cell, row, col)
                    self.loopDefined_tab_cells[slide_idx].append(cell)
                    
                    cell_idx += 1
                
                # Fill remaining cells with "Empty"
                while cell_idx < grid_size:
                    row = cell_idx // num_cols
                    col = cell_idx % num_cols
                    
                    cell = DraggableLabel("Empty", self, cell_list=self.loopDefined_tab_cells[slide_idx])
                    cell.setAlignment(Qt.AlignCenter)
                    cell.setMinimumSize(100, 60)
                    cell.setStyleSheet(
                        "border: 2px dashed #ccc; "
                        "background-color: #f5f5f5; "
                        "color: #999; "
                        "border-radius: 5px; "
                        "padding: 5px;"
                    )
                    
                    tab_grid_layout.addWidget(cell, row, col)
                    self.loopDefined_tab_cells[slide_idx].append(cell)
                    
                    cell_idx += 1
            
            self.loopDefined_page_tabs.setVisible(len(slide_axis_values) > 0)
            
            # Unblock signals after tab creation is complete
            self.loopDefined_page_tabs.blockSignals(False)
        
        # Hide the main preview grid widget when using tabs
        if hasattr(self, 'loopDefined_preview_group'):
            preview_layout = self.loopDefined_preview_group.layout()
            if preview_layout:
                for i in range(preview_layout.count()):
                    item = preview_layout.itemAt(i)
                    if item and item.widget():
                        widget = item.widget()
                        # Check if this widget contains the preview_grid_layout
                        if hasattr(widget, 'layout') and widget.layout() == self.loopDefined_preview_grid_layout:
                            widget.setVisible(False)
                            break
    
    def _update_loopDefined_axis_dropdowns(self):
        """Update axis dropdowns with available variables from path and filename templates (for DRC plot mode)"""
        if not hasattr(self, 'loopDefined_filename_template_input') or not hasattr(self, 'loopDefined_path_template_input'):
            return
        
        path_template = self.loopDefined_path_template_input.text().strip()
        filename_template = self.loopDefined_filename_template_input.text().strip()
        
        if not path_template and not filename_template:
            return
        
        # Extract variables from both path and filename templates
        from Config.Rules.loopDefined_template_engine import create_template_parser
        all_variables = set()
        
        try:
            parser = create_template_parser(path_template, filename_template)
            # Get variables from path template
            if hasattr(parser, 'path_variables'):
                all_variables.update(parser.path_variables)
            # Get variables from filename template (excluding 'format')
            if hasattr(parser, 'filename_variables'):
                filename_variables = [v for v in parser.filename_variables if v != 'format']
                all_variables.update(filename_variables)
        except:
            # Fallback: try to extract manually using regex
            import re
            if path_template:
                path_vars = re.findall(r'\{(\w+)\}', path_template)
                all_variables.update(path_vars)
            if filename_template:
                filename_vars = re.findall(r'\{(\w+)\}', filename_template)
                filename_vars = [v for v in filename_vars if v != 'format']
                all_variables.update(filename_vars)
        
        all_variables = sorted(list(all_variables))
        
        # Update dropdowns (for DRC plot mode)
        if hasattr(self, 'loopDefined_slide_axis_combo'):
            current_slide = self.loopDefined_slide_axis_combo.currentText()
            self.loopDefined_slide_axis_combo.clear()
            self.loopDefined_slide_axis_combo.addItems([""] + all_variables)
            if current_slide in all_variables:
                self.loopDefined_slide_axis_combo.setCurrentText(current_slide)
        
        if hasattr(self, 'loopDefined_grid_axis_combo'):
            current_grid = self.loopDefined_grid_axis_combo.currentText()
            self.loopDefined_grid_axis_combo.clear()
            self.loopDefined_grid_axis_combo.addItems([""] + all_variables)
            if current_grid in all_variables:
                self.loopDefined_grid_axis_combo.setCurrentText(current_grid)
    
    def _update_loopDefined_slide_axis_dropdown_normal(self):
        """Update Slide axis dropdown with available variables from filename template (for normal mode, DRC plot OFF)"""
        if not hasattr(self, 'loopDefined_filename_template_input'):
            return
        
        filename_template = self.loopDefined_filename_template_input.text().strip()
        if not filename_template:
            return
        
        # Extract variables from filename template
        from Config.Rules.loopDefined_template_engine import create_template_parser
        try:
            parser = create_template_parser("", filename_template)  # Only filename template
            filename_variables = [v for v in parser.filename_variables if v != 'format']
        except:
            filename_variables = []
        
        # Update Slide axis dropdown (for normal mode)
        if hasattr(self, 'loopDefined_slide_axis_combo_normal'):
            current_slide = self.loopDefined_slide_axis_combo_normal.currentText()
            self.loopDefined_slide_axis_combo_normal.clear()
            self.loopDefined_slide_axis_combo_normal.addItem("")  # None option
            self.loopDefined_slide_axis_combo_normal.addItems(filename_variables)
            if current_slide in filename_variables:
                self.loopDefined_slide_axis_combo_normal.setCurrentText(current_slide)
    
    def on_loopDefined_slide_axis_normal_changed(self, text):
        """Handle Slide axis dropdown change in normal mode (DRC plot OFF)"""
        # Update preview when slide axis changes
        self.update_loopDefined_preview()
    
    def on_systematic_change(self):
        """Handle systematic checkbox change"""
        pass
    
    def _update_preset_dropdown(self):
        """Update preset dropdown with available presets"""
        from GUI.preset_manager import list_presets
        
        # Block signals to prevent triggering during update
        self.preset_dropdown.blockSignals(True)
        
        # Save current selection
        current_text = self.preset_dropdown.currentText()
        
        # Clear and rebuild
        self.preset_dropdown.clear()
        self.preset_dropdown.addItem("None")
        
        # Add presets
        presets = list_presets(self.user_settings)
        preset_names = [p['name'] for p in presets]
        if preset_names:
            self.preset_dropdown.addItems(preset_names)
        
        # Add separator and Save as...
        if preset_names:
            self.preset_dropdown.insertSeparator(self.preset_dropdown.count())
        self.preset_dropdown.addItem("Save as…")
        
        # Restore selection if still valid
        if current_text and current_text in [self.preset_dropdown.itemText(i) for i in range(self.preset_dropdown.count())]:
            index = self.preset_dropdown.findText(current_text)
            if index >= 0:
                self.preset_dropdown.setCurrentIndex(index)
        else:
            self.preset_dropdown.setCurrentIndex(0)  # None
        
        self.preset_dropdown.blockSignals(False)
    
    def on_preset_selected(self, text):
        """Handle preset dropdown selection"""
        if text == "None":
            self.log("Preset cleared (None)")
            return
        
        if text == "Save as…":
            # Save current selection to restore after save
            current_index = self.preset_dropdown.currentIndex()
            self._save_preset_as()
            # Note: _save_preset_as will update dropdown and select saved preset
            # If save was cancelled, restore to previous selection
            # (This is handled by _save_preset_as not changing selection if cancelled)
            return
        
        # Load and apply preset
        from GUI.preset_manager import load_preset, apply_preset
        
        preset_data = load_preset(text, self.user_settings)
        if preset_data:
            apply_preset(preset_data, self)
            self.log(f"Preset applied: {text}")
            # Update preview after applying preset
            self.update_loopDefined_preview()
        else:
            self.log(f"Failed to load preset: {text}")
            # Reset to None if load failed
            self.preset_dropdown.blockSignals(True)
            self.preset_dropdown.setCurrentIndex(0)
            self.preset_dropdown.blockSignals(False)
    
    def _save_preset_as(self):
        """Save current settings as a new preset"""
        from GUI.preset_manager import create_preset_data, save_preset, list_presets
        
        # Save current selection to restore if cancelled
        previous_text = None
        for i in range(self.preset_dropdown.count()):
            if self.preset_dropdown.itemText(i) != "Save as…" and self.preset_dropdown.itemText(i) != "None":
                if not self.preset_dropdown.itemText(i).startswith("---"):  # Not separator
                    previous_text = self.preset_dropdown.itemText(i)
                    break
        
        # Get preset name from user
        name, ok = QInputDialog.getText(self, "Save Preset", "Enter preset name:")
        if not ok or not name.strip():
            # Restore previous selection if cancelled
            if previous_text:
                index = self.preset_dropdown.findText(previous_text)
                if index >= 0:
                    self.preset_dropdown.blockSignals(True)
                    self.preset_dropdown.setCurrentIndex(index)
                    self.preset_dropdown.blockSignals(False)
            else:
                self.preset_dropdown.blockSignals(True)
                self.preset_dropdown.setCurrentIndex(0)  # None
                self.preset_dropdown.blockSignals(False)
            return
        
        name = name.strip()
        
        # Check for duplicate
        existing_presets = list_presets(self.user_settings)
        existing_names = [p['name'] for p in existing_presets]
        
        if name in existing_names:
            reply = QMessageBox.question(
                self, "Overwrite Preset",
                f"Preset '{name}' already exists. Overwrite?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                # Restore previous selection if cancelled
                if previous_text:
                    index = self.preset_dropdown.findText(previous_text)
                    if index >= 0:
                        self.preset_dropdown.blockSignals(True)
                        self.preset_dropdown.setCurrentIndex(index)
                        self.preset_dropdown.blockSignals(False)
                else:
                    self.preset_dropdown.blockSignals(True)
                    self.preset_dropdown.setCurrentIndex(0)  # None
                    self.preset_dropdown.blockSignals(False)
                return
        
        # Create preset data
        preset_data = create_preset_data(self)
        preset_data['name'] = name
        
        # Save preset
        if save_preset(preset_data, self.user_settings):
            self.log(f"Preset saved: {name}")
            # Update dropdown
            self._update_preset_dropdown()
            # Select and apply the saved preset
            index = self.preset_dropdown.findText(name)
            if index >= 0:
                self.preset_dropdown.blockSignals(True)
                self.preset_dropdown.setCurrentIndex(index)
                self.preset_dropdown.blockSignals(False)
                # Apply the preset
                from GUI.preset_manager import apply_preset
                apply_preset(preset_data, self)
                self.update_loopDefined_preview()
            self.log(f"Preset applied: {name}")
        else:
            QMessageBox.warning(self, "Save Error", f"Failed to save preset: {name}")
            # Restore previous selection if save failed
            if previous_text:
                index = self.preset_dropdown.findText(previous_text)
                if index >= 0:
                    self.preset_dropdown.blockSignals(True)
                    self.preset_dropdown.setCurrentIndex(index)
                    self.preset_dropdown.blockSignals(False)
            else:
                self.preset_dropdown.blockSignals(True)
                self.preset_dropdown.setCurrentIndex(0)  # None
                self.preset_dropdown.blockSignals(False)
    
    def on_preset_folder_clicked(self):
        """Handle preset folder button click"""
        folder = QFileDialog.getExistingDirectory(
            self, "Select Preset Folder", 
            str(self._get_preset_folder_path())
        )
        
        if folder:
            # Save folder path to settings
            if not self.user_settings:
                self.user_settings = {}
            if 'preset' not in self.user_settings:
                self.user_settings['preset'] = {}
            
            self.user_settings['preset']['folder_path'] = folder
            
            # Save settings to file (if there's a settings file mechanism)
            # For now, just log it
            self.log(f"Preset folder set: {folder}")
            
            # Update preset dropdown to reflect new folder
            self._update_preset_dropdown()
    
    def _get_preset_folder_path(self):
        """Get current preset folder path"""
        from GUI.preset_manager import get_preset_folder_path
        return get_preset_folder_path(self.user_settings)
    
    def show_settings(self):
        """Show settings dialog"""
        dialog = SettingsDialog(self)
        if dialog.exec_():
            # Settings were saved
            settings = dialog.get_settings()
            # Store settings in instance variable
            self.user_settings = settings
            self.log_text.append("Settings saved successfully.")
            
            # Note: update_loopDefined_template_info() is not called here to preserve
            # user input in path/filename template fields and preview row/column inputs.
            # Template info is updated automatically when templates are changed via
            # textChanged signals connected to update_loopDefined_template_info()
    
    def show_manual(self):
        """Show manual dialog"""
        dialog = ManualDialog(self)
        dialog.exec_()
    
    def on_observable_changed(self):
        """Auto-select/deselect Observable step based on Observable object"""
        observable_checked = False
        for obj_name in self.object_checks:
            if "Observable" in obj_name and self.object_checks[obj_name].isChecked():
                observable_checked = True
                break
        
        if "Observable" in self.step_checks:
            self.step_checks["Observable"].setChecked(observable_checked)
    
    def on_after_top_reco_object_changed(self):
        """Auto-select/deselect afterTopReco based on objects below separator"""
        any_selected = False
        for obj_name in self.after_top_reco_objects:
            if obj_name in self.object_checks and self.object_checks[obj_name].isChecked():
                any_selected = True
                break
        
        if "afterTopReco" in self.step_checks:
            self.step_checks["afterTopReco"].setChecked(any_selected)
    
    def browse_input_dir(self):
        """Browse for input directory"""
        directory = QFileDialog.getExistingDirectory(
            self, "Select Input Directory", self.input_dir_edit.text()
        )
        if directory:
            self.input_dir_edit.setText(directory)
    
    def browse_output_file(self):
        """Browse for output file"""
        filename, _ = QFileDialog.getSaveFileName(
            self, "Select Output File",
            self.output_file_edit.text(),
            "Keynote files (*.key);;All files (*.*)"
        )
        if filename:
            self.output_file_edit.setText(filename)
    
    def log(self, message):
        """Add message to log"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def generate_slides(self):
        """Generate Keynote slides based on selections"""
        try:
            self.log_text.clear()
            
            mode = self.mode
            systematic = self.systematic_check.isChecked()
            input_dir = self.input_dir_edit.text()
            output_file = self.output_file_edit.text()
            # Get jobversion from config
            from Config.config_cpv import cpv_config
            jobversion = cpv_config.get("jobversion", "")
            
            # Get file format based on mode
            if mode == "cpv":
                file_format = self.cpv_file_format_combo.currentText()
            elif mode == "drc":
                file_format = self.file_format_combo.currentText()
            elif mode == "dragdrop":
                file_format = "pdf"  # Default for dragdrop mode
            else:  # loopDefined
                file_format = self.loopDefined_file_format_combo.currentText()
            
            # DRC options
            particle = self.particle_combo.currentText()
            case = self.case_combo.currentText()
            c_checked = self.c_check.isChecked()
            s_selected = self.s_combo.currentText()
            drcor_selected = self.drcor_combo.currentText()
            sc_overlay_checked = self.sc_overlay_check.isChecked() if hasattr(self, 'sc_overlay_check') else False
            
            # Get selected energy points from preview grid (DRC) or checkboxes (CPV/Loop-defined)
            if mode == "drc":
                # Read from each tab (C, S, DRcor) - each tab has its own energy order
                if hasattr(self, 'drc_page_tabs') and hasattr(self, 'drc_tab_cells') and hasattr(self, 'drc_tab_grids'):
                    # Map tab index to channel name: 0=C, 1=S, 2=DRcor
                    channel_names = ["C", "S", "DRcor"]
                    selected_energies = {}  # Dictionary: channel_name -> list of energies
                    
                    for tab_idx, tab_cells in enumerate(self.drc_tab_cells):
                        if tab_idx < len(channel_names):
                            channel_name = channel_names[tab_idx]
                            # Read energy order from grid layout (preserves drag-and-drop order)
                            tab_grid_layout = self.drc_tab_grids[tab_idx] if tab_idx < len(self.drc_tab_grids) else None
                            
                            if tab_grid_layout:
                                # Read cells in grid order (row by row, left to right)
                                energies_list = []
                                for row in range(2):  # 2 rows
                                    for col in range(3):  # 3 columns
                                        item = tab_grid_layout.itemAtPosition(row, col)
                                        if item and item.widget():
                                            cell = item.widget()
                                            cell_text = cell.text()
                                            if cell_text and cell_text != "Empty":
                                                energies_list.append(cell_text)
                                            else:
                                                energies_list.append("Empty")
                                selected_energies[channel_name] = energies_list
                            else:
                                # Fallback: read from cell list in order
                                energies_list = [cell.text() if cell.text() != "" and cell.text() != "Empty" else "Empty" 
                                               for cell in tab_cells]
                                selected_energies[channel_name] = energies_list
                else:
                    # Fallback to old behavior (single list)
                    selected_energies = [cell.text() if cell.text() != "" else "Empty" 
                                       for cell in self.energy_preview_cells]
            else:
                selected_energies = []
            
            if mode == "cpv":
                # Collect selected objects for CPV mode
                selected_objects = []
                lep_group = []
                jet_group = []
                bjet_group = []
                
                if self.object_checks["Lep1"].isChecked():
                    lep_group.append("Lep1")
                if self.object_checks["Lep2"].isChecked():
                    lep_group.append("Lep2")
                    
                if self.object_checks["Jet1"].isChecked():
                    jet_group.append("Jet1")
                if self.object_checks["Jet2"].isChecked():
                    jet_group.append("Jet2")
                    
                if self.object_checks["bJet1"].isChecked():
                    bjet_group.append("bJet1")
                if self.object_checks["bJet2"].isChecked():
                    bjet_group.append("bJet2")
                
                if lep_group:
                    selected_objects.append(lep_group)
                if jet_group:
                    selected_objects.append(jet_group)
                if bjet_group:
                    selected_objects.append(bjet_group)
                
                # After Top Reco object groups
                top_group = []
                nu_antinu_group = []
                bjet_antibjet_group = []
                
                if self.object_checks.get("Top", None) and self.object_checks["Top"].isChecked():
                    top_group.append("Top")
                if self.object_checks.get("AnTop", None) and self.object_checks["AnTop"].isChecked():
                    top_group.append("AnTop")
                
                if self.object_checks.get("Nu", None) and self.object_checks["Nu"].isChecked():
                    nu_antinu_group.append("Nu")
                if self.object_checks.get("AnNu", None) and self.object_checks["AnNu"].isChecked():
                    nu_antinu_group.append("AnNu")
                
                if self.object_checks.get("bJet", None) and self.object_checks["bJet"].isChecked():
                    bjet_antibjet_group.append("bJet")
                if self.object_checks.get("AnbJet", None) and self.object_checks["AnbJet"].isChecked():
                    bjet_antibjet_group.append("AnbJet")
                
                if top_group:
                    selected_objects.append(top_group)
                if nu_antinu_group:
                    selected_objects.append(nu_antinu_group)
                if bjet_antibjet_group:
                    selected_objects.append(bjet_antibjet_group)
                
                # Single objects (not grouped)
                single_objects = ["Num_PV", "Mass", "Num_Jets", "MET", "Jets", "bJets"]
                for obj in single_objects:
                    if obj in self.object_checks and self.object_checks[obj].isChecked():
                        selected_objects.append([obj])
                
                # Observable
                for obj_name in self.object_checks:
                    if "Observable" in obj_name and self.object_checks[obj_name].isChecked():
                        if "(" in obj_name and ")" in obj_name:
                            obs_text = obj_name.split("(")[1].split(")")[0]
                            obs_list = [o.strip() for o in obs_text.split(",")]
                            selected_objects.append(obs_list)
                        break
                
                if not selected_objects:
                    QMessageBox.warning(self, "No Selection", "Please select at least one object.")
                    return
                
                # Get selected steps
                selected_steps = [step for step, check in self.step_checks.items() if check.isChecked()]
                
                # Auto-add afterTopReco step if any afterTopReco objects are selected
                any_after_top_reco_selected = False
                for obj_name in self.after_top_reco_objects:
                    if obj_name in self.object_checks and self.object_checks[obj_name].isChecked():
                        any_after_top_reco_selected = True
                        break
                
                if any_after_top_reco_selected and "afterTopReco" not in selected_steps:
                    selected_steps.append("afterTopReco")
                    # Also update the hidden checkbox state
                    if "afterTopReco" in self.step_checks:
                        self.step_checks["afterTopReco"].setChecked(True)
                
                # Auto-add Observable step if Observable object is selected
                observable_selected = False
                for obj_name in self.object_checks:
                    if "Observable" in obj_name and self.object_checks[obj_name].isChecked():
                        observable_selected = True
                        break
                
                if observable_selected and "Observable" not in selected_steps:
                    selected_steps.append("Observable")
                    # Also update the hidden checkbox state
                    if "Observable" in self.step_checks:
                        self.step_checks["Observable"].setChecked(True)
                
                if not selected_steps:
                    QMessageBox.warning(self, "No Selection", "Please select at least one step.")
                    return
                
                observable_text = ""
                resol_with_noise = False
                resol_without_noise = False
                linearity = False
            elif mode == "drc":
                # DRC mode
                selected_objects = []
                selected_steps = []
                observable_text = ""
                
                resol_with_noise = self.resol_with_noise_check.isChecked()
                resol_without_noise = self.resol_without_noise_check.isChecked()
                linearity = self.lin_check.isChecked()
                
                # Check if at least one option is selected
                has_energy = any(e != "Empty" for e in selected_energies)
                if not has_energy and not resol_with_noise and not resol_without_noise and not linearity:
                    QMessageBox.warning(self, "No Selection", 
                                      "Please select at least one option:\n- Energy point\n- Resolution (with/without Noise term)\n- Linearity")
                    return
            else:
                # Loop-defined mode
                # Objects and steps are now handled in Variables section
                # Check if object and step variables are in templates
                from Config.Rules.loopDefined_template_engine import create_template_parser
                path_template = self.loopDefined_path_template if hasattr(self, 'loopDefined_path_template') else ""
                filename_template = self.loopDefined_filename_template if hasattr(self, 'loopDefined_filename_template') else ""
                
                has_object = False
                has_step = False
                if path_template or filename_template:
                    parser = create_template_parser(path_template, filename_template)
                    has_object = "object" in parser.all_variables
                    has_step = "step" in parser.all_variables
                
                # Collect objects from Variables section (only if object is in template)
                selected_objects = []
                if has_object:
                    if hasattr(self, 'loopDefined_variable_inputs') and 'object' in self.loopDefined_variable_inputs:
                        object_input = self.loopDefined_variable_inputs['object']
                        object_text = object_input.text().strip()
                        if object_text:
                            # Parse comma-separated object names
                            object_names = [name.strip() for name in object_text.split(',') if name.strip()]
                            # Each object becomes its own group
                            selected_objects = [[obj_name] for obj_name in object_names]
                    
                    if not selected_objects:
                        QMessageBox.warning(self, "No Selection", "Please enter at least one object name in Variables section.")
                        return
                # If object is not in template, selected_objects will be empty list
                # This will be handled in keynoteCtrl_loopDefined.py
                
                # Get selected steps from Variables section (only if step is in template)
                selected_steps = []
                if has_step:
                    if hasattr(self, 'loopDefined_variable_inputs') and 'step' in self.loopDefined_variable_inputs:
                        step_input = self.loopDefined_variable_inputs['step']
                        step_text = step_input.text().strip()
                        if step_text:
                            # Parse comma-separated step names
                            selected_steps = [step.strip() for step in step_text.split(',') if step.strip()]
                    
                    if not selected_steps:
                        QMessageBox.warning(self, "No Selection", "Please enter at least one step name in Variables section.")
                        return
                # If step is not in template, selected_steps will be empty list
                # This will be handled in keynoteCtrl_loopDefined.py
                
                observable_text = ""
                resol_with_noise = False
                resol_without_noise = False
                linearity = False
                
                # Get user-defined templates and variables from mode screen
                loopDefined_path_template = self.loopDefined_path_template_input.text().strip()
                loopDefined_filename_template = self.loopDefined_filename_template_input.text().strip()
                loopDefined_variables = {}
                
                # Use default templates if empty
                default_path_template = "{sample}/{channel}/{object}/{step}"
                default_filename_template = "h_{object}_{kinematic}_{step}.{format}"
                
                if not loopDefined_path_template:
                    loopDefined_path_template = default_path_template
                if not loopDefined_filename_template:
                    loopDefined_filename_template = default_filename_template
                
                # Get variable values from UI
                for var_name, var_input in self.loopDefined_variable_inputs.items():
                    # Get from input field (all variables are now in Variables section)
                    var_text = var_input.text().strip()
                    if var_text:
                        # Parse comma-separated values
                        loopDefined_variables[var_name] = [v.strip() for v in var_text.split(',') if v.strip()]
                
                # Get row and column values from preview inputs
                # Get column values (single input)
                column_text = self.loopDefined_column_input.text().strip()
                loopDefined_column_values = [v.strip() for v in column_text.split(',') if v.strip()] if column_text else []
                
                # Get row values from each row input field
                loopDefined_row_values = []
                if hasattr(self, 'loopDefined_row_inputs'):
                    for i in sorted(self.loopDefined_row_inputs.keys()):
                        row_input = self.loopDefined_row_inputs[i]
                        row_text = row_input.text().strip()
                        row_values = [v.strip() for v in row_text.split(',') if v.strip()] if row_text else []
                        loopDefined_row_values.append(row_values)
            
            # Create and start thread
            if mode == "dragdrop":
                # Dragdrop mode: collect file paths from tab cells
                dragdrop_file_paths_dict = {}  # Format: {slide_idx: [file_paths]}
                arrangement = "2*3"  # Default
                slide_count = 1  # Default
                
                # Get arrangement
                if hasattr(self, 'dragdrop_arrangement_combo'):
                    arrangement = self.dragdrop_arrangement_combo.currentText()
                
                # Get slide count
                if hasattr(self, 'dragdrop_slide_count_combo'):
                    slide_count_text = self.dragdrop_slide_count_combo.currentText()
                    if slide_count_text == "Custom":
                        if hasattr(self, 'dragdrop_slide_count_input'):
                            slide_count = self.dragdrop_slide_count_input.value()
                    else:
                        try:
                            slide_count = int(slide_count_text)
                        except:
                            slide_count = 1
                
                try:
                    num_rows, num_cols = map(int, arrangement.split('*'))
                except:
                    num_rows, num_cols = 2, 3
                    arrangement = "2*3"
                
                # Read file paths from each tab
                if hasattr(self, 'dragdrop_tab_grids') and hasattr(self, 'dragdrop_tab_cells'):
                    for slide_idx in range(slide_count):
                        if slide_idx < len(self.dragdrop_tab_grids):
                            tab_grid_layout = self.dragdrop_tab_grids[slide_idx]
                            file_paths = []
                            
                            # Read cells in grid order (row by row, left to right)
                            for row in range(num_rows):
                                for col in range(num_cols):
                                    item = tab_grid_layout.itemAtPosition(row, col)
                                    if item and item.widget():
                                        cell = item.widget()
                                        if hasattr(cell, 'file_path') and cell.file_path:
                                            file_paths.append(cell.file_path)
                                        else:
                                            file_paths.append(None)
                                    else:
                                        file_paths.append(None)
                            
                            dragdrop_file_paths_dict[slide_idx] = file_paths
                
                # Check if at least one file is dropped in any slide
                has_files = False
                for slide_idx, file_paths in dragdrop_file_paths_dict.items():
                    if any(path for path in file_paths if path):
                        has_files = True
                        break
                
                if not has_files:
                    QMessageBox.warning(self, "No Files", 
                                      "Please drag and drop at least one plot file into the preview grid.")
                    return
                
                self.generator_thread = GeneratorThread(
                    mode, False, input_dir, output_file,
                    [], [], "", "",  # No objects, steps, observable, jobversion for dragdrop
                    None, None, None, None, None, [],  # DRC options
                    False, False, False, file_format,  # Resolution/linearity options
                    self.user_settings,  # User settings
                    dragdrop_file_paths=dragdrop_file_paths_dict,  # File paths dict
                    dragdrop_arrangement=arrangement,  # Arrangement
                    dragdrop_slide_count=slide_count  # Slide count
                )
            elif mode == "loopDefined":
                # Get scan mode checkbox state
                loopDefined_scan_mode = False
                if hasattr(self, 'loopDefined_scan_mode_check'):
                    loopDefined_scan_mode = self.loopDefined_scan_mode_check.isChecked()
                
                # Get DRC plot mode information
                loopDefined_drc_plot = False
                loopDefined_slide_axis = None
                loopDefined_grid_axis = None
                loopDefined_arrangement = None
                loopDefined_confirmed_scan_results = None
                loopDefined_preview_cell_order = None  # Store preview cell order for each slide
                
                if hasattr(self, 'loopDefined_drc_plot_check'):
                    loopDefined_drc_plot = self.loopDefined_drc_plot_check.isChecked()
                    if loopDefined_drc_plot:
                        if hasattr(self, 'loopDefined_slide_axis_combo'):
                            loopDefined_slide_axis = self.loopDefined_slide_axis_combo.currentText()
                        if hasattr(self, 'loopDefined_grid_axis_combo'):
                            loopDefined_grid_axis = self.loopDefined_grid_axis_combo.currentText()
                        if hasattr(self, 'loopDefined_arrangement_combo'):
                            loopDefined_arrangement = self.loopDefined_arrangement_combo.currentText()
                        if hasattr(self, 'loopDefined_confirmed_scan_results'):
                            loopDefined_confirmed_scan_results = self.loopDefined_confirmed_scan_results
                        
                        # Read preview cell order from tabs for DRC plot mode (preserves drag-and-drop order)
                        if (hasattr(self, 'loopDefined_tab_grids') and 
                            hasattr(self, 'loopDefined_tab_cells') and 
                            hasattr(self, 'loopDefined_page_tabs') and
                            loopDefined_grid_axis and
                            loopDefined_arrangement):
                            loopDefined_preview_cell_order = {}
                            try:
                                num_rows, num_cols = map(int, loopDefined_arrangement.split('*'))
                            except:
                                num_rows, num_cols = 2, 3
                            
                            for slide_idx, tab_grid_layout in enumerate(self.loopDefined_tab_grids):
                                if tab_grid_layout:
                                    # 그리드 레이아웃에서 각 위치의 셀을 읽어옴 (DRC mode 방식)
                                    # 위치 순서대로 정렬: row * num_cols + col
                                    # Format: List of (position_idx, grid_value) tuples
                                    cell_positions = []
                                    
                                    for row in range(num_rows):
                                        for col in range(num_cols):
                                            # Get cell at this position
                                            item = tab_grid_layout.itemAtPosition(row, col)
                                            if item and item.widget():
                                                cell = item.widget()
                                                cell_text = cell.text()
                                                
                                                # Extract grid axis value from cell text
                                                grid_value = None
                                                if cell_text and cell_text != "Empty":
                                                    if f"{loopDefined_grid_axis}=" in cell_text:
                                                        # Format: "grid_axis=value" or "grid_axis=value\n..."
                                                        value_part = cell_text.split(f"{loopDefined_grid_axis}=")[1]
                                                        # Take first line if multi-line
                                                        grid_value = value_part.split('\n')[0].strip()
                                                
                                                # Calculate position index (same as DRC mode: row * num_cols + col)
                                                position_idx = row * num_cols + col
                                                cell_positions.append((position_idx, grid_value))
                                    
                                    # Sort by position index to preserve grid layout order
                                    cell_positions.sort(key=lambda x: x[0])
                                    # Store as list of (position_idx, grid_value) tuples
                                    # This preserves both the position and the value
                                    loopDefined_preview_cell_order[slide_idx] = cell_positions
                    else:
                        # Normal mode (DRC plot OFF): Get slide axis from normal dropdown
                        if hasattr(self, 'loopDefined_slide_axis_combo_normal'):
                            slide_axis_text = self.loopDefined_slide_axis_combo_normal.currentText()
                            if slide_axis_text:  # Not empty
                                loopDefined_slide_axis = slide_axis_text
                
                self.generator_thread = GeneratorThread(
                    mode, systematic, input_dir, output_file,
                    selected_objects, selected_steps, observable_text, jobversion,
                    particle, case, c_checked, s_selected, drcor_selected, selected_energies,
                    resol_with_noise, resol_without_noise, linearity, file_format,
                    self.user_settings, loopDefined_path_template, loopDefined_filename_template, loopDefined_variables,
                    loopDefined_row_values, loopDefined_column_values, loopDefined_scan_mode
                )
                
                # Store DRC plot mode information in thread instance
                self.generator_thread.loopDefined_drc_plot = loopDefined_drc_plot
                self.generator_thread.loopDefined_slide_axis = loopDefined_slide_axis
                self.generator_thread.loopDefined_grid_axis = loopDefined_grid_axis
                self.generator_thread.loopDefined_arrangement = loopDefined_arrangement
                self.generator_thread.loopDefined_confirmed_scan_results = loopDefined_confirmed_scan_results
                self.generator_thread.loopDefined_preview_cell_order = loopDefined_preview_cell_order
            else:
                self.generator_thread = GeneratorThread(
                    mode, systematic, input_dir, output_file,
                    selected_objects, selected_steps, observable_text, jobversion,
                    particle, case, c_checked, s_selected, drcor_selected, selected_energies,
                    resol_with_noise, resol_without_noise, linearity, file_format,
                    self.user_settings,  # Pass user settings
                    sc_overlay_checked=sc_overlay_checked
                )
            self.generator_thread.log_signal.connect(self.log)
            self.generator_thread.finished_signal.connect(self.on_generation_finished)
            
            # Enable Stop button and disable Generate button
            if hasattr(self, 'stop_button'):
                self.stop_button.setEnabled(True)
            # Disable Generate button during generation
            for btn in self.findChildren(QPushButton):
                if btn.text() == "GENERATE":
                    btn.setEnabled(False)
            
            self.generator_thread.start()
        
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            self.log("-" * 60)
            self.log("Error Details:")
            self.log(f"Error Type: {type(e).__name__}")
            self.log(f"Error Message: {str(e)}")
            self.log("-" * 60)
            self.log("Stack Trace:")
            tb_str = traceback.format_exc()
            self.log(tb_str)
            self.log("=" * 60)
            
            QMessageBox.critical(self, "Error", 
                               f"An error occurred: {type(e).__name__}\n{str(e)}\n\nCheck Status log for details.")
    
    def on_generation_finished(self, success, message):
        """Handle generation completion"""
        # Re-enable Generate button and disable Stop button
        if hasattr(self, 'stop_button'):
            self.stop_button.setEnabled(False)
        # Re-enable Generate button
        for btn in self.findChildren(QPushButton):
            if btn.text() == "GENERATE":
                btn.setEnabled(True)
        
        # Clean up thread
        if self.generator_thread:
            self.generator_thread.quit()
            self.generator_thread.wait()
            self.generator_thread = None
    
    def _toggle_all_loopDefined_objects(self, checked):
        """Toggle all user-defined objects based on select all checkbox"""
        for check in self.object_checks.values():
            if hasattr(check, 'setChecked'):
                check.setChecked(checked)
    
    def _toggle_all_loopDefined_steps(self, checked):
        """Toggle all user-defined steps (no longer used - steps are in Variables section)"""
        # This function is no longer needed since steps are handled in Variables section
        pass
    
    def _update_scan_button_visibility(self):
        """Update SCAN button visibility based on mode and scan mode checkbox"""
        if hasattr(self, 'scan_button'):
            if self.mode == "loopDefined" and hasattr(self, 'loopDefined_scan_mode_check'):
                # Show SCAN button only if scan mode is enabled
                self.scan_button.setVisible(self.loopDefined_scan_mode_check.isChecked())
            else:
                self.scan_button.setVisible(False)
    
    def _update_path_template_input_state(self):
        """Update path template input and path template variables enabled/disabled state based on scan mode checkbox"""
        if hasattr(self, 'loopDefined_path_template_input') and hasattr(self, 'loopDefined_scan_mode_check'):
            # Disable path template input when scan mode is enabled
            scan_mode_enabled = self.loopDefined_scan_mode_check.isChecked()
            self.loopDefined_path_template_input.setEnabled(not scan_mode_enabled)
            
            # Update tooltip to explain why it's disabled
            if scan_mode_enabled:
                self.loopDefined_path_template_input.setToolTip(
                    "Path template is disabled in scan mode.\n"
                    "In scan mode, files are found by matching filename_template only."
                )
            else:
                self.loopDefined_path_template_input.setToolTip(
                    "Enter path template using variables in braces.\n"
                    "Example: {sample}/{channel}/{object}/{step}\n"
                    "You can use ANY variable names. Variables will be replaced with actual values during plot insertion."
                )
        
        # Disable path template variables in Variables panel when scan mode is enabled
        if hasattr(self, 'loopDefined_scan_mode_check') and hasattr(self, 'loopDefined_path_variables'):
            scan_mode_enabled = self.loopDefined_scan_mode_check.isChecked()
            path_variables = getattr(self, 'loopDefined_path_variables', [])
            
            if hasattr(self, 'loopDefined_variable_inputs'):
                for var_name in path_variables:
                    if var_name in self.loopDefined_variable_inputs:
                        var_input = self.loopDefined_variable_inputs[var_name]
                        var_input.setEnabled(not scan_mode_enabled)
                        
                        # Update tooltip
                        if scan_mode_enabled:
                            var_input.setToolTip(
                                f"'{var_name}' is from Path Template and is disabled in scan mode.\n"
                                "In scan mode, only Filename Template variables are used."
                            )
                        else:
                            var_input.setToolTip(f"Enter values for {var_name} variable (comma-separated)")
    
    def scan_files(self):
        """Scan files and show results in popup dialog"""
        try:
            # Get input directory
            input_dir = self.input_dir_edit.text().strip()
            if not input_dir or not os.path.exists(input_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                  "Please select a valid input directory.")
                return
            
            # Get filename template
            filename_template = self.loopDefined_filename_template_input.text().strip()
            if not filename_template:
                QMessageBox.warning(self, "Missing Template", 
                                  "Please enter a filename template.")
                return
            
            # Get file format
            file_format = self.loopDefined_file_format_combo.currentText()
            
            # Get variable combinations from variable inputs
            variable_combinations = {}
            if hasattr(self, 'loopDefined_variable_inputs'):
                for var_name, var_input in self.loopDefined_variable_inputs.items():
                    var_text = var_input.text().strip()
                    if var_text:
                        # Parse comma-separated values
                        variable_combinations[var_name] = [v.strip() for v in var_text.split(',') if v.strip()]
            
            if not variable_combinations:
                QMessageBox.warning(self, "No Variables", 
                                  "Please enter variable values in the Variables panel.")
                return
            
            # Perform scan
            from Config.Rules.rules_loopDefined import find_loopDefined_files_scan
            
            # Create a simple log callback that collects messages
            log_messages = []
            def log_callback(msg):
                log_messages.append(msg)
            
            # Get path template
            path_template = self.loopDefined_path_template_input.text().strip()
            
            matched_files, stats, all_results = find_loopDefined_files_scan(
                base_dir=input_dir,
                filename_template=filename_template,
                variable_combinations=variable_combinations,
                file_format=file_format,
                log_callback=log_callback,
                path_template=path_template
            )
            
            # Show results in popup dialog
            from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTextEdit, QPushButton, QLabel
            from PyQt5.QtCore import Qt
            
            dialog = QDialog(self)
            dialog.setWindowTitle("Scan Results")
            dialog.setMinimumSize(800, 600)
            
            layout = QVBoxLayout()
            
            # Title
            title = QLabel("File Scan Results")
            title_font = QFont()
            title_font.setPointSize(16)
            title_font.setBold(True)
            title.setFont(title_font)
            layout.addWidget(title)
            
            # Results text area
            results_text = QTextEdit()
            results_text.setReadOnly(True)
            results_text.setFont(QFont("Courier", 10))
            
            # Format results
            output = []
            output.append("=" * 60)
            output.append("SCAN RESULTS")
            output.append("=" * 60)
            output.append(f"Total files scanned: {stats['total_scanned']}")
            output.append(f"Matched files: {stats['matched_count']}")
            output.append(f"Not found: {stats['not_found_count']}")
            output.append(f"Ambiguous (multiple matches): {stats['ambiguous_count']}")
            output.append("-" * 60)
            
            # Show matched files
            if stats['matched_count'] > 0:
                output.append("\nMatched files:")
                for file_path, parsed_vars, status in all_results:
                    if status == "Matched" and file_path:
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        output.append(f"  [{var_str}] {file_path}")
                output.append("-" * 60)
            
            # Show not found combinations
            if stats['not_found_count'] > 0:
                output.append("\nNot found (missing variable combinations):")
                for file_path, parsed_vars, status in all_results:
                    if status == "Not found":
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        output.append(f"  [{var_str}]")
                output.append("-" * 60)
            
            # Show ambiguous files
            if stats['ambiguous_count'] > 0:
                output.append("\nAmbiguous files (multiple matches for same variable combination):")
                for file_path, parsed_vars, status in all_results:
                    if status == "Ambiguous" and file_path:
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        output.append(f"  [{var_str}] {file_path}")
                output.append("-" * 60)
            
            results_text.setPlainText("\n".join(output))
            layout.addWidget(results_text)
            
            # Buttons layout
            buttons_layout = QHBoxLayout()
            buttons_layout.addStretch()
            
            # Correct button
            correct_btn = QPushButton("Correct")
            correct_btn.setMinimumSize(120, 40)
            correct_font = QFont()
            correct_font.setPointSize(12)
            correct_font.setBold(True)
            correct_btn.setFont(correct_font)
            correct_btn.setStyleSheet("color: green;")
            
            def on_correct_clicked():
                # Store scan results
                self.loopDefined_scan_results = {
                    'matched_files': matched_files,
                    'stats': stats,
                    'all_results': all_results,
                    'variable_combinations': variable_combinations
                }
                # Confirm scan results (for DRC plot mode)
                self.loopDefined_confirmed_scan_results = self.loopDefined_scan_results
                # Update axis dropdowns if DRC plot is enabled
                if hasattr(self, 'loopDefined_drc_plot_check') and self.loopDefined_drc_plot_check.isChecked():
                    self._update_loopDefined_axis_dropdowns()
                # Always update preview when Correct is clicked (includes Ambiguous files)
                if hasattr(self, '_update_loopDefined_drc_preview'):
                    self._update_loopDefined_drc_preview()
                dialog.accept()
            
            correct_btn.clicked.connect(on_correct_clicked)
            buttons_layout.addWidget(correct_btn)
            
            buttons_layout.addSpacing(20)
            
            # No button
            no_btn = QPushButton("Nooooo.....")
            no_btn.setMinimumSize(120, 40)
            no_font = QFont()
            no_font.setPointSize(12)
            no_font.setBold(True)
            no_btn.setFont(no_font)
            no_btn.setStyleSheet("color: red;")
            no_btn.clicked.connect(dialog.reject)
            buttons_layout.addWidget(no_btn)
            
            buttons_layout.addStretch()
            layout.addLayout(buttons_layout)
            
            dialog.setLayout(layout)
            dialog.exec_()
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Scan Error", 
                               f"An error occurred during file scanning:\n\n{str(e)}\n\n{traceback.format_exc()}")
    
    def stop_generation(self):
        """Stop the current slide generation process"""
        if self.generator_thread and self.generator_thread.isRunning():
            self.log("=" * 60)
            self.log("⚠️ Stopping slide generation...")
            self.log("=" * 60)
            # Request interruption
            self.generator_thread.requestInterruption()
            # Wait for thread to finish (with timeout)
            if not self.generator_thread.wait(3000):  # Wait up to 3 seconds
                # Force terminate if it doesn't stop gracefully
                self.generator_thread.terminate()
                self.generator_thread.wait()
            self.log("Generation stopped by user")
            # Re-enable Generate button and disable Stop button
            if hasattr(self, 'stop_button'):
                self.stop_button.setEnabled(False)
            for btn in self.findChildren(QPushButton):
                if btn.text() == "GENERATE":
                    btn.setEnabled(True)
            self.generator_thread = None
    
    def connect_select_all_handlers(self):
        """Connect Select All checkboxes to their handlers"""
        if hasattr(self, 'select_all_above_check'):
            self.select_all_above_check.toggled.connect(self.on_select_all_above)
        if hasattr(self, 'select_all_below_check'):
            self.select_all_below_check.toggled.connect(self.on_select_all_below)
        if hasattr(self, 'select_all_steps_check'):
            self.select_all_steps_check.toggled.connect(self.on_select_all_steps)
    
    def on_select_all_above(self, checked):
        """Handle Select All for Event selection stage objects"""
        # Objects above separator (Event selection stage)
        object_names_above = [
            "Num_PV", "Mass", "Num_Jets", "Lep1", "Lep2",
            "Jet1", "Jet2", "bJet1", "bJet2", "MET", "Observable (O1, O3)"
        ]
        
        # Block signals to prevent recursive updates
        if hasattr(self, 'select_all_above_check'):
            self.select_all_above_check.blockSignals(True)
        
        for obj_name in object_names_above:
            if obj_name in self.object_checks:
                # Block signals temporarily to avoid triggering other handlers
                self.object_checks[obj_name].blockSignals(True)
                self.object_checks[obj_name].setChecked(checked)
                self.object_checks[obj_name].blockSignals(False)
        
        if hasattr(self, 'select_all_above_check'):
            self.select_all_above_check.blockSignals(False)
    
    def on_select_all_below(self, checked):
        """Handle Select All for After Top quark reconstruction stage objects"""
        # Objects below separator (After Top quark reconstruction stage)
        object_names_below = ["Nu", "AnNu", "bJet", "AnbJet", "Top", "AnTop"]
        
        # Block signals to prevent recursive updates
        if hasattr(self, 'select_all_below_check'):
            self.select_all_below_check.blockSignals(True)
        
        for obj_name in object_names_below:
            if obj_name in self.object_checks:
                # Block signals temporarily to avoid triggering other handlers
                self.object_checks[obj_name].blockSignals(True)
                self.object_checks[obj_name].setChecked(checked)
                self.object_checks[obj_name].blockSignals(False)
        
        if hasattr(self, 'select_all_below_check'):
            self.select_all_below_check.blockSignals(False)
        
        # After unblocking signals, manually trigger afterTopReco update
        # to ensure afterTopReco step is selected when objects are selected
        self.on_after_top_reco_object_changed()
    
    def on_select_all_steps(self, checked):
        """Handle Select All for steps"""
        # All step names (excluding afterTopReco and Observable - they are auto-selected)
        step_names = ["initial", "step1", "step2", "step3", "step4", "step5", "step6"]
        
        # Block signals to prevent recursive updates
        if hasattr(self, 'select_all_steps_check'):
            self.select_all_steps_check.blockSignals(True)
        
        for step_name in step_names:
            if step_name in self.step_checks:
                # Block signals temporarily to avoid triggering other handlers
                self.step_checks[step_name].blockSignals(True)
                self.step_checks[step_name].setChecked(checked)
                self.step_checks[step_name].blockSignals(False)
        
        if hasattr(self, 'select_all_steps_check'):
            self.select_all_steps_check.blockSignals(False)
    
    def update_select_all_above_state(self):
        """Update Select All checkbox state based on individual object checkboxes"""
        if not hasattr(self, 'select_all_above_check'):
            return
        
        object_names_above = [
            "Num_PV", "Mass", "Num_Jets", "Lep1", "Lep2",
            "Jet1", "Jet2", "bJet1", "bJet2", "MET", "Observable (O1, O3)"
        ]
        
        # Check if all objects are checked
        all_checked = True
        for obj_name in object_names_above:
            if obj_name in self.object_checks:
                if not self.object_checks[obj_name].isChecked():
                    all_checked = False
                    break
        
        # Update select_all_above_check without triggering signal
        self.select_all_above_check.blockSignals(True)
        self.select_all_above_check.setChecked(all_checked)
        self.select_all_above_check.blockSignals(False)
    
    def update_select_all_below_state(self):
        """Update Select All checkbox state based on individual object checkboxes"""
        if not hasattr(self, 'select_all_below_check'):
            return
        
        object_names_below = ["Nu", "AnNu", "bJet", "AnbJet", "Top", "AnTop"]
        
        # Check if all objects are checked
        all_checked = True
        for obj_name in object_names_below:
            if obj_name in self.object_checks:
                if not self.object_checks[obj_name].isChecked():
                    all_checked = False
                    break
        
        # Update select_all_below_check without triggering signal
        self.select_all_below_check.blockSignals(True)
        self.select_all_below_check.setChecked(all_checked)
        self.select_all_below_check.blockSignals(False)
    
    def update_select_all_steps_state(self):
        """Update Select All checkbox state based on individual step checkboxes"""
        if not hasattr(self, 'select_all_steps_check'):
            return
        
        step_names = ["initial", "step1", "step2", "step3", "step4", "step5", "step6"]
        
        # Check if all steps are checked
        all_checked = True
        for step_name in step_names:
            if step_name in self.step_checks:
                if not self.step_checks[step_name].isChecked():
                    all_checked = False
                    break
        
        # Update select_all_steps_check without triggering signal
        self.select_all_steps_check.blockSignals(True)
        self.select_all_steps_check.setChecked(all_checked)
        self.select_all_steps_check.blockSignals(False)


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    
    # Set application style for modern look
    app.setStyle('Fusion')
    
    window = KeynoteSlideGeneratorGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
