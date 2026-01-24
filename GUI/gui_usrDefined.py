#!/usr/bin/env python3
"""
User-defined Mode UI components for Keynote Slide Generator
User-defined template-based plot insertion
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox,
    QGroupBox, QComboBox, QLineEdit, QTextEdit, QPushButton,
    QFormLayout, QScrollArea, QFrame, QGridLayout
)
from PyQt5.QtCore import Qt
from GUI.common import DraggableLabel


def create_usrDefined_options_ui(parent):
    """
    Create User-defined mode options UI
    Returns: (usrDefined_options_group, file_format_combo, path_template_input, filename_template_input)
    """
    usrDefined_options_group = QGroupBox("User-defined Options")
    usrDefined_options_layout = QVBoxLayout()
    
    # File Format selection
    format_layout = QHBoxLayout()
    format_layout.addWidget(QLabel("File Format:"))
    file_format_combo = QComboBox()
    file_format_combo.addItems(["pdf", "png"])
    file_format_combo.setCurrentText("pdf")
    format_layout.addWidget(file_format_combo)
    format_layout.addStretch()
    # Add scan mode checkbox
    scan_mode_check = QCheckBox("Use SCAN mode")
    scan_mode_check.setToolTip(
        "When enabled, scans Input Directory recursively to find files matching filename_template.\n"
        "Variables are extracted from filenames and matched against allowed values in Variables panel.\n"
        "When disabled, uses template-based path generation (default mode)."
    )
    format_layout.addWidget(scan_mode_check)
    usrDefined_options_layout.addLayout(format_layout)
    
    # Path Template input
    path_template_layout = QFormLayout()
    path_template_input = QLineEdit()
    path_template_input.setText("{sample}/{channel}/{object}/{step}")  # Default value
    path_template_input.setToolTip(
        "Enter path template using variables in braces.\n"
        "Example: {sample}/{channel}/{object}/{step}\n"
        "You can use ANY variable names. Variables will be replaced with actual values during plot insertion."
    )
    path_template_layout.addRow("Path Template:", path_template_input)
    usrDefined_options_layout.addLayout(path_template_layout)
    
    # Filename Template input
    filename_template_layout = QFormLayout()
    filename_template_input = QLineEdit()
    filename_template_input.setText("h_{object}_{kinematic}_{step}.{format}")  # Default value
    filename_template_input.setToolTip(
        "Enter filename template using variables in braces.\n"
        "Example: h_{object}_{kinematic}_{step}.{format}\n"
        "You can use ANY variable names. The {format} variable will be automatically replaced with the selected file format."
    )
    filename_template_layout.addRow("Filename Template:", filename_template_input)
    usrDefined_options_layout.addLayout(filename_template_layout)
    
    usrDefined_options_group.setLayout(usrDefined_options_layout)
    
    return usrDefined_options_group, file_format_combo, path_template_input, filename_template_input, scan_mode_check


def create_usrDefined_objects_ui(parent):
    """
    Create User-defined mode object input UI
    Users can directly input object names in text fields
    Returns: (usrDefined_objects_widget, object_inputs)
    """
    usrDefined_objects_widget = QWidget()
    usrDefined_objects_layout = QVBoxLayout()
    usrDefined_objects_layout.setContentsMargins(0, 0, 0, 0)
    
    # Object input field (no label or instructions)
    object_input = QLineEdit()
    object_input.setPlaceholderText("e.g., Lep1,Lep2,Jet1")
    usrDefined_objects_layout.addWidget(object_input)
    
    # Store as dict for consistency with other modes
    object_inputs = {"object": object_input}
    
    usrDefined_objects_widget.setLayout(usrDefined_objects_layout)
    
    return usrDefined_objects_widget, object_inputs


def create_usrDefined_variables_ui(parent):
    """
    Create User-defined mode variable values input UI
    Users input values for template variables defined in Settings
    Returns: (usrDefined_variables_widget, variable_inputs)
    """
    usrDefined_variables_widget = QWidget()
    usrDefined_variables_layout = QVBoxLayout()
    usrDefined_variables_layout.setContentsMargins(0, 0, 0, 0)
    
    # Instructions
    info_label = QLabel(
        "Enter values for template variables defined in Settings → User-defined tab.\n"
        "These values will be used to replace variables in your path/filename templates.\n"
        "Example: If template is {sample}/{channel}, enter sample=dy,ttbar and channel=MuMu,ee\n"
        "Note: 'object' and 'step' are handled separately in Objects and Steps sections."
    )
    info_label.setWordWrap(True)
    info_label.setStyleSheet("color: #666; padding: 5px;")
    usrDefined_variables_layout.addWidget(info_label)
    
    # Variable inputs (excluding object and step, which are handled separately)
    variable_inputs = {}
    
    # Common variables (excluding object and step)
    common_vars = ["sample", "channel", "kinematic"]
    
    form_layout = QFormLayout()
    for var_name in common_vars:
        var_input = QLineEdit()
        var_input.setPlaceholderText(f"e.g., value1,value2,value3")
        variable_inputs[var_name] = var_input
        form_layout.addRow(f"{var_name}:", var_input)
    
    usrDefined_variables_layout.addLayout(form_layout)
    usrDefined_variables_layout.addStretch()
    
    usrDefined_variables_widget.setLayout(usrDefined_variables_layout)
    
    return usrDefined_variables_widget, variable_inputs


def create_usrDefined_preview_ui(parent):
    """
    Create User-defined mode preview UI with Arrangement selection
    Returns: (preview_group, preview_cells, arrangement_combo, row_inputs, column_input, preview_grid_layout)
    """
    from PyQt5.QtWidgets import QTabWidget
    
    preview_group = QGroupBox("Preview")
    preview_layout = QVBoxLayout()
    
    # Arrangement selection with DRC plot checkbox
    arrangement_layout = QHBoxLayout()
    arrangement_layout.addWidget(QLabel("Arrangement"))
    arrangement_combo = QComboBox()
    arrangement_combo.addItems(["1*1", "1*2", "1*3", "1*4", "2*1", "2*2", "2*3", "2*4"])
    arrangement_combo.setCurrentText("2*3")
    arrangement_combo.setMinimumWidth(100)
    arrangement_layout.addWidget(arrangement_combo)
    arrangement_layout.addStretch()
    # DRC plot checkbox (flag for generation logic, no UI changes)
    drc_plot_check = QCheckBox("DRC plot")
    drc_plot_check.setToolTip("Use DRC-style generation logic")
    arrangement_layout.addWidget(drc_plot_check)
    preview_layout.addLayout(arrangement_layout)
    
    # DRC plot mode controls (shown only when DRC plot is checked)
    drc_controls_container = QWidget()
    drc_controls_layout = QVBoxLayout()
    drc_controls_layout.setContentsMargins(0, 0, 0, 0)
    
    # Slide axis and Grid axis dropdowns
    axis_layout = QHBoxLayout()
    axis_layout.addWidget(QLabel("Slide axis:"))
    slide_axis_combo = QComboBox()
    slide_axis_combo.setMinimumWidth(120)
    axis_layout.addWidget(slide_axis_combo)
    axis_layout.addSpacing(20)
    axis_layout.addWidget(QLabel("Grid axis:"))
    grid_axis_combo = QComboBox()
    grid_axis_combo.setMinimumWidth(120)
    axis_layout.addWidget(grid_axis_combo)
    axis_layout.addStretch()
    drc_controls_layout.addLayout(axis_layout)
    
    # Page selector (Tabs)
    page_tabs = QTabWidget()
    page_tabs.setVisible(False)  # Initially hidden, shown when pages are created
    drc_controls_layout.addWidget(page_tabs)
    
    drc_controls_container.setLayout(drc_controls_layout)
    drc_controls_container.setVisible(False)  # Initially hidden
    preview_layout.addWidget(drc_controls_container)
    
    # Row and Column input fields container (always shown)
    row_column_container = QWidget()
    row_column_layout = QVBoxLayout()
    row_column_layout.setContentsMargins(0, 0, 0, 0)
    
    # Column input (will be placed next to row 1)
    column_input = QLineEdit()
    column_input.setPlaceholderText("e.g., pt,eta,phi")
    column_input.setMinimumWidth(145)
    column_input.setFixedWidth(145)
    
    # Row inputs container (will be dynamically created based on arrangement)
    # First row will have row 1 and column side by side
    row_inputs_container = QWidget()
    row_inputs_layout = QVBoxLayout()
    row_inputs_layout.setContentsMargins(0, 0, 0, 0)
    row_inputs_container.setLayout(row_inputs_layout)
    row_column_layout.addWidget(row_inputs_container)
    
    row_column_container.setLayout(row_column_layout)
    preview_layout.addWidget(row_column_container)
    
    # DRC-style energy checkboxes container (always hidden - not used)
    drc_energy_container = QWidget()
    drc_energy_layout = QVBoxLayout()
    drc_energy_layout.setContentsMargins(0, 0, 0, 0)
    
    # Energy checkboxes layout (similar to DRC mode)
    energy_checkboxes_layout = QVBoxLayout()
    energy_checkboxes_layout.setContentsMargins(0, 0, 0, 0)
    drc_energy_container.setLayout(drc_energy_layout)
    drc_energy_layout.addLayout(energy_checkboxes_layout)
    drc_energy_container.setVisible(False)  # Always hidden
    preview_layout.addWidget(drc_energy_container)
    
    # Create grid layout for visual preview
    preview_widget = QWidget()
    preview_grid_layout = QGridLayout()
    preview_grid_layout.setSpacing(5)
    preview_grid_layout.setContentsMargins(5, 5, 5, 5)
    
    # Initialize preview cells (will be dynamically resized based on row/column input)
    preview_cells = []
    
    preview_widget.setLayout(preview_grid_layout)
    preview_layout.addWidget(preview_widget)
    
    # Add info label below preview grid
    info_label = QLabel("You can adjust the size and postions of plots in Settings.")
    info_label.setWordWrap(True)
    info_label.setStyleSheet("color: #666; padding: 5px; font-size: 10pt;")
    info_label.setAlignment(Qt.AlignCenter)
    preview_layout.addWidget(info_label)
    
    preview_group.setLayout(preview_layout)
    
    return (preview_group, preview_cells, arrangement_combo, row_inputs_container, 
            row_inputs_layout, column_input, preview_grid_layout, drc_plot_check, 
            drc_energy_container, energy_checkboxes_layout, drc_controls_container,
            slide_axis_combo, grid_axis_combo, page_tabs)


def update_usrDefined_row_inputs(parent, arrangement):
    """Update row input fields based on arrangement selection"""
    if not hasattr(parent, 'usrDefined_row_inputs_layout'):
        return
    
    # Parse arrangement to get number of rows
    try:
        rows, cols = map(int, arrangement.split('*'))
    except:
        rows, cols = 2, 3  # Default
    
    # Clear existing row inputs
    # Note: column_input is stored in parent, so it won't be deleted
    while parent.usrDefined_row_inputs_layout.count():
        item = parent.usrDefined_row_inputs_layout.takeAt(0)
        if item.layout():
            # Remove layout items, but don't delete column_input
            layout = item.layout()
            while layout.count():
                sub_item = layout.takeAt(0)
                if sub_item.widget():
                    widget = sub_item.widget()
                    # Don't delete column_input, it's reused
                    if widget != parent.usrDefined_column_input:
                        widget.deleteLater()
        elif item.widget() and item.widget() != parent.usrDefined_column_input:
            item.widget().deleteLater()
    
    # Clear row inputs dictionary
    if hasattr(parent, 'usrDefined_row_inputs'):
        parent.usrDefined_row_inputs.clear()
    else:
        parent.usrDefined_row_inputs = {}
    
    # Create row input fields
    for i in range(rows):
        row_layout = QHBoxLayout()
        row_label = QLabel(f"row {i+1}")
        row_input = QLineEdit()
        row_input.setPlaceholderText(f"e.g., value1,value2,value3")
        row_input.setMinimumWidth(150)
        row_input.setFixedWidth(150)  # Match column input width
        row_layout.addWidget(row_label)
        row_layout.addWidget(row_input)
        
        # For row 1, add column input next to it
        if i == 0:
            row_layout.addSpacing(20)
            column_label = QLabel("column")
            row_layout.addWidget(column_label)
            # Column input should already exist, just add it
            if hasattr(parent, 'usrDefined_column_input'):
                row_layout.addWidget(parent.usrDefined_column_input)
        
        row_layout.addStretch()
        parent.usrDefined_row_inputs_layout.addLayout(row_layout)
        parent.usrDefined_row_inputs[i] = row_input
        # Connect to update preview
        row_input.textChanged.connect(parent.update_usrDefined_preview)
    
    # Add Slide axis dropdown below column (after last row)
    if not hasattr(parent, 'usrDefined_slide_axis_combo_normal') or parent.usrDefined_slide_axis_combo_normal is None:
        slide_axis_layout = QHBoxLayout()
        slide_axis_label = QLabel("Slide axis:")
        slide_axis_combo_normal = QComboBox()
        slide_axis_combo_normal.setMinimumWidth(120)
        slide_axis_combo_normal.addItem("")  # None option
        slide_axis_layout.addWidget(slide_axis_label)
        slide_axis_layout.addWidget(slide_axis_combo_normal)
        slide_axis_layout.addStretch()
        parent.usrDefined_row_inputs_layout.addLayout(slide_axis_layout)
        parent.usrDefined_slide_axis_combo_normal = slide_axis_combo_normal
        # Connect to update preview and axis dropdowns
        slide_axis_combo_normal.currentTextChanged.connect(parent.on_usrDefined_slide_axis_normal_changed)
        # Update dropdown items when template changes
        if hasattr(parent, 'usrDefined_filename_template_input'):
            parent.usrDefined_filename_template_input.textChanged.connect(
                lambda: parent._update_usrDefined_slide_axis_dropdown_normal()
            )
    else:
        # Slide axis combo already exists, just update it
        parent.usrDefined_slide_axis_combo_normal.setParent(None)
        slide_axis_layout = QHBoxLayout()
        slide_axis_label = QLabel("Slide axis:")
        slide_axis_layout.addWidget(slide_axis_label)
        slide_axis_layout.addWidget(parent.usrDefined_slide_axis_combo_normal)
        slide_axis_layout.addStretch()
        parent.usrDefined_row_inputs_layout.addLayout(slide_axis_layout)
    
    # Update dropdown items
    if hasattr(parent, '_update_usrDefined_slide_axis_dropdown_normal'):
        parent._update_usrDefined_slide_axis_dropdown_normal()


def update_usrDefined_preview(parent):
    """Update the preview layout in User-defined mode based on row and column inputs"""
    # Check if DRC plot is checked (if checked, don't use normal preview)
    drc_plot_checked = False
    if hasattr(parent, 'usrDefined_drc_plot_check'):
        drc_plot_checked = parent.usrDefined_drc_plot_check.isChecked()
    
    # If DRC plot is checked and has scan results, use DRC preview instead
    if drc_plot_checked:
        if hasattr(parent, 'usrDefined_confirmed_scan_results') and parent.usrDefined_confirmed_scan_results:
            # Use DRC preview - don't show normal preview
            if hasattr(parent, '_update_usrDefined_drc_preview'):
                parent._update_usrDefined_drc_preview()
            # Hide main grid
            if hasattr(parent, 'usrDefined_preview_group'):
                preview_layout = parent.usrDefined_preview_group.layout()
                if preview_layout:
                    for i in range(preview_layout.count()):
                        item = preview_layout.itemAt(i)
                        if item and item.widget():
                            widget = item.widget()
                            if hasattr(widget, 'layout') and widget.layout() == parent.usrDefined_preview_grid_layout:
                                widget.setVisible(False)
                                break
            return
    
    # DRC plot not checked - always use normal preview (row/col based)
    # Hide tabs if they exist
    if hasattr(parent, 'usrDefined_page_tabs'):
        parent.usrDefined_page_tabs.setVisible(False)
    
    # Use normal preview with row/column inputs
    if not hasattr(parent, 'usrDefined_row_inputs') or not hasattr(parent, 'usrDefined_column_input'):
        return
    
    if not hasattr(parent, 'usrDefined_preview_grid_layout'):
        return
    
    # Get arrangement
    if hasattr(parent, 'usrDefined_arrangement_combo'):
        arrangement = parent.usrDefined_arrangement_combo.currentText()
        try:
            num_rows, num_cols = map(int, arrangement.split('*'))
        except:
            num_rows, num_cols = 2, 3
    else:
        num_rows, num_cols = 2, 3
    
    # Get column values
    column_text = parent.usrDefined_column_input.text().strip()
    column_values = [v.strip() for v in column_text.split(',') if v.strip()] if column_text else []
    
    # Get row values from each row input
    row_values_list = []
    for i in range(num_rows):
        if i in parent.usrDefined_row_inputs:
            row_text = parent.usrDefined_row_inputs[i].text().strip()
            # Parse comma-separated values, handling spaces
            row_values = [v.strip() for v in row_text.split(',') if v.strip()] if row_text else []
            row_values_list.append(row_values)
        else:
            row_values_list.append([])
    
    # Clear existing cells
    if hasattr(parent, 'usrDefined_preview_cells'):
        for cell in parent.usrDefined_preview_cells:
            if cell:
                cell.deleteLater()
        parent.usrDefined_preview_cells.clear()
    else:
        parent.usrDefined_preview_cells = []
    
    # Clear grid layout
    while parent.usrDefined_preview_grid_layout.count():
        item = parent.usrDefined_preview_grid_layout.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
    
    # Create new cells based on arrangement
    # Each row's all values are shown in first line, column value in second line
    # row 1: [Lep1, Jet1, bJet1] -> col 1: "Lep1, Jet1, bJet1\npt"
    for row in range(num_rows):
        for col in range(num_cols):
            # Get row values for this specific row
            row_values = row_values_list[row] if row < len(row_values_list) else []
            
            # First line: all row values joined with comma (e.g., "Lep1, Jet1, bJet1")
            if row_values:
                row_value = ", ".join(row_values)  # All row values in first line
            else:
                row_value = ""
            
            # Get column value for this column position
            if col < len(column_values):
                col_value = column_values[col]  # Second line: value from column input at column index
            else:
                col_value = ""
            
            # Create cell text: first line is all row values, second line is column value
            if row_value and col_value:
                cell_text = f"{row_value}\n{col_value}"
            elif row_value:
                cell_text = row_value
            elif col_value:
                cell_text = col_value
            else:
                cell_text = "Empty"
            
            cell = DraggableLabel(cell_text, parent, cell_list=parent.usrDefined_preview_cells)
            cell.setAlignment(Qt.AlignCenter)
            cell.setMinimumSize(100, 60)
            
            if cell_text == "Empty":
                cell.setStyleSheet(
                    "border: 2px dashed #ccc; "
                    "background-color: #f5f5f5; "
                    "color: #999; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )
            else:
                cell.setStyleSheet(
                    "border: 2px solid #0066cc; "
                    "background-color: #e6f2ff; "
                    "color: #0066cc; "
                    "font-weight: bold; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )
            
            parent.usrDefined_preview_grid_layout.addWidget(cell, row, col)
            parent.usrDefined_preview_cells.append(cell)
    
    # Show the main preview grid widget when not using tabs
    if hasattr(parent, 'usrDefined_preview_group'):
        preview_layout = parent.usrDefined_preview_group.layout()
        if preview_layout:
            for i in range(preview_layout.count()):
                item = preview_layout.itemAt(i)
                if item and item.widget():
                    widget = item.widget()
                    # Check if this widget contains the preview_grid_layout
                    if hasattr(widget, 'layout') and widget.layout() == parent.usrDefined_preview_grid_layout:
                        widget.setVisible(True)
                        break


# User-defined Steps UI functions removed
# Steps are now handled directly in Variables section
# No separate Steps UI needed


def create_usrDefined_energy_checkboxes(parent, energy_values):
    """
    Create energy checkboxes for DRC-style preview in User-defined mode
    Similar to create_energy_checkboxes in gui_drc.py
    """
    if not hasattr(parent, 'usrDefined_energy_checkboxes_layout'):
        return
    
    # Clear existing checkboxes
    while parent.usrDefined_energy_checkboxes_layout.count():
        child = parent.usrDefined_energy_checkboxes_layout.takeAt(0)
        if child.widget():
            child.widget().deleteLater()
        elif child.layout():
            while child.layout().count():
                subchild = child.layout().takeAt(0)
                if subchild.widget():
                    subchild.widget().deleteLater()
    
    # Clear energy_checks dictionary
    if not hasattr(parent, 'usrDefined_energy_checks'):
        parent.usrDefined_energy_checks = {}
    else:
        parent.usrDefined_energy_checks.clear()
    
    # Create new checkboxes
    for row in energy_values:
        row_layout = QHBoxLayout()
        for energy in row:
            check = QCheckBox(str(energy))
            check.setChecked(False)
            parent.usrDefined_energy_checks[str(energy)] = check
            # Connect to update preview
            check.stateChanged.connect(lambda state, p=parent: update_usrDefined_drc_preview(p))
            row_layout.addWidget(check)
        row_layout.addStretch()
        parent.usrDefined_energy_checkboxes_layout.addLayout(row_layout)


def update_usrDefined_drc_preview(parent):
    """Update the DRC-style energy preview layout in User-defined mode"""
    if not hasattr(parent, 'usrDefined_energy_checks'):
        return
    
    if not hasattr(parent, 'usrDefined_preview_grid_layout'):
        return
    
    selected_energies = [energy for energy, check in parent.usrDefined_energy_checks.items() if check.isChecked()]
    
    # Sort energies by numeric value (handle string values)
    def sort_key(x):
        try:
            return float(str(x).replace("GeV", "").strip())
        except:
            return float('inf')
    
    sorted_energies = sorted(selected_energies, key=sort_key)
    
    # Clear existing cells
    if hasattr(parent, 'usrDefined_preview_cells'):
        for cell in parent.usrDefined_preview_cells:
            if cell:
                cell.deleteLater()
        parent.usrDefined_preview_cells.clear()
    else:
        parent.usrDefined_preview_cells = []
    
    # Clear grid layout
    while parent.usrDefined_preview_grid_layout.count():
        item = parent.usrDefined_preview_grid_layout.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
    
    # Create 2x3 grid for DRC-style preview
    for row in range(2):
        for col in range(3):
            idx = row * 3 + col
            if idx < len(sorted_energies):
                energy = sorted_energies[idx]
                cell_text = str(energy)
                cell_style = (
                    "border: 2px solid #0066cc; "
                    "background-color: #e6f2ff; "
                    "color: #0066cc; "
                    "font-weight: bold; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )
            else:
                cell_text = "Empty"
                cell_style = (
                    "border: 2px dashed #ccc; "
                    "background-color: #f5f5f5; "
                    "color: #999; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )
            
            cell = DraggableLabel(cell_text, parent, cell_list=parent.usrDefined_preview_cells)
            cell.setAlignment(Qt.AlignCenter)
            cell.setMinimumSize(100, 60)
            cell.setStyleSheet(cell_style)
            
            parent.usrDefined_preview_grid_layout.addWidget(cell, row, col)
            parent.usrDefined_preview_cells.append(cell)


def run_usrDefined_generation(thread_instance):
    """
    Run User-defined mode slide generation
    Called from GeneratorThread.run()
    """
    # Check for interruption at the start
    if thread_instance.isInterruptionRequested():
        return False, "Generation cancelled"
    
    try:
        from Config.KeynoteCtrl import prepare_keynote
        from Config.KeynoteCtrl.keynoteCtrl_usrDefined import insert_pdfs_into_slide_usrDefined
        from Config.Rules.usrDefined_template_engine import create_template_parser, validate_template
        from Config.Rules.rules_usrDefined import find_usrDefined_files, find_usrDefined_files_scan
        
        # Check scan mode
        scan_mode = getattr(thread_instance, 'usrDefined_scan_mode', False)
        
        thread_instance.log_signal.emit(f"Selected objects: {thread_instance.selected_objects}")
        thread_instance.log_signal.emit(f"Selected steps: {thread_instance.selected_steps}")
        
        output_file = thread_instance.output_file
        base_dir = thread_instance.input_dir
        
        thread_instance.log_signal.emit(f"Output file: {output_file}")
        thread_instance.log_signal.emit(f"Base directory: {base_dir}")
        thread_instance.log_signal.emit(f"Scan mode: {'ENABLED' if scan_mode else 'DISABLED'}")
        thread_instance.log_signal.emit("-" * 60)
        
        # Get templates from thread_instance
        path_template = thread_instance.usrDefined_path_template
        filename_template = thread_instance.usrDefined_filename_template
        variable_combinations = thread_instance.usrDefined_variables
        
        if not filename_template:
            return False, "Filename template not configured. Please set filename template in Settings → User-defined tab."
        
        # In scan mode, path_template is not used for path generation
        # but filename_template is required
        if not scan_mode and not path_template:
            return False, "Path template not configured. Please set path template in Settings → User-defined tab."
        
        # Validate templates
        if not scan_mode:
            path_valid, path_error = validate_template(path_template)
            if not path_valid:
                return False, f"Invalid path template: {path_error}"
        
        filename_valid, filename_error = validate_template(filename_template)
        if not filename_valid:
            return False, f"Invalid filename template: {filename_error}"
        
        if scan_mode:
            thread_instance.log_signal.emit(f"Filename template: {filename_template}")
        else:
            thread_instance.log_signal.emit(f"Path template: {path_template}")
            thread_instance.log_signal.emit(f"Filename template: {filename_template}")
        thread_instance.log_signal.emit(f"Variable combinations: {variable_combinations}")
        
        # In scan mode, perform scan and show status
        if scan_mode:
            thread_instance.log_signal.emit("=" * 60)
            thread_instance.log_signal.emit("SCAN MODE: Scanning directory for matching files...")
            thread_instance.log_signal.emit("=" * 60)
            
            matched_files, stats, all_results = find_usrDefined_files_scan(
                base_dir=base_dir,
                filename_template=filename_template,
                variable_combinations=variable_combinations,
                file_format=thread_instance.file_format,
                log_callback=thread_instance.log_signal.emit
            )
            
            # Display scan results
            thread_instance.log_signal.emit("=" * 60)
            thread_instance.log_signal.emit("SCAN RESULTS:")
            thread_instance.log_signal.emit("=" * 60)
            thread_instance.log_signal.emit(f"Total files scanned: {stats['total_scanned']}")
            thread_instance.log_signal.emit(f"Matched files: {stats['matched_count']}")
            thread_instance.log_signal.emit(f"Not found: {stats['not_found_count']}")
            thread_instance.log_signal.emit(f"Ambiguous (multiple matches): {stats['ambiguous_count']}")
            thread_instance.log_signal.emit("-" * 60)
            
            # Show matched files
            if stats['matched_count'] > 0:
                thread_instance.log_signal.emit("Matched files:")
                for file_path, parsed_vars, status in all_results:
                    if status == "Matched" and file_path:
                        # Format variables for display
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        thread_instance.log_signal.emit(f"  [{var_str}] {file_path}")
                thread_instance.log_signal.emit("-" * 60)
            
            # Show not found combinations
            if stats['not_found_count'] > 0:
                thread_instance.log_signal.emit("Not found (missing variable combinations):")
                for file_path, parsed_vars, status in all_results:
                    if status == "Not found":
                        # Format variables for display
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        thread_instance.log_signal.emit(f"  [{var_str}]")
                thread_instance.log_signal.emit("-" * 60)
            
            # Show ambiguous files
            if stats['ambiguous_count'] > 0:
                thread_instance.log_signal.emit("Ambiguous files (multiple matches for same variable combination):")
                for file_path, parsed_vars, status in all_results:
                    if status == "Ambiguous" and file_path:
                        # Format variables for display
                        var_str = ", ".join([f"{k}={v}" for k, v in sorted(parsed_vars.items()) if k != "format"])
                        thread_instance.log_signal.emit(f"  [{var_str}] {file_path}")
                thread_instance.log_signal.emit("-" * 60)
            
            if stats['matched_count'] == 0:
                return False, "No files matched the template. Please check your filename template and variable combinations."
            
            thread_instance.log_signal.emit("=" * 60)
            
            # Store scan results for use in insert_pdfs_into_slide_usrDefined
            thread_instance.scan_matched_files = matched_files
            thread_instance.scan_stats = stats
        else:
            # Normal mode: no scan results
            thread_instance.scan_matched_files = None
            thread_instance.scan_stats = None
        
        # Create Keynote file
        thread_instance.log_signal.emit("Creating Keynote file...")
        prepare_keynote(output_file, theme="White")
        thread_instance.log_signal.emit("Keynote file created successfully!")
        
        # Check if kinematic variable is needed (only in non-scan mode)
        if not scan_mode:
            from Config.Rules.usrDefined_template_engine import create_template_parser
            parser = create_template_parser(path_template, filename_template)
            has_kinematic = "kinematic" in parser.all_variables
            
            # If kinematic is in template but not in variable_combinations, try to get from object
            if has_kinematic and "kinematic" not in variable_combinations:
                from Config.Rules.rules import get_kinematics
                # Get kinematics for first selected object
                if thread_instance.selected_objects:
                    first_obj_group = thread_instance.selected_objects[0]
                    if first_obj_group:
                        first_obj = first_obj_group[0] if isinstance(first_obj_group, list) else first_obj_group
                        kinematic_list = get_kinematics(first_obj)
                        variable_combinations["kinematic"] = kinematic_list
                        thread_instance.log_signal.emit(f"Auto-detected kinematics for {first_obj}: {kinematic_list}")
        
        # Insert plots into slides
        thread_instance.log_signal.emit("Inserting plots into slides...")
        
        # Get DRC plot mode information
        drc_plot = getattr(thread_instance, 'usrDefined_drc_plot', False)
        slide_axis = getattr(thread_instance, 'usrDefined_slide_axis', None)
        grid_axis = getattr(thread_instance, 'usrDefined_grid_axis', None)
        arrangement = getattr(thread_instance, 'usrDefined_arrangement', None)
        confirmed_scan_results = getattr(thread_instance, 'usrDefined_confirmed_scan_results', None)
        preview_cell_order = getattr(thread_instance, 'usrDefined_preview_cell_order', None)
        
        insert_pdfs_into_slide_usrDefined(
            output_file=output_file,
            base_dir=base_dir,
            path_template=path_template if not scan_mode else None,  # path_template not used in scan mode
            filename_template=filename_template,
            selected_objects=thread_instance.selected_objects,
            selected_steps=thread_instance.selected_steps,
            variable_combinations=variable_combinations,
            file_format=thread_instance.file_format,
            log_callback=thread_instance.log_signal.emit,
            thread_instance=thread_instance,
            row_values=thread_instance.usrDefined_row_values,
            column_values=thread_instance.usrDefined_column_values,
            scan_mode=scan_mode,  # Pass scan mode flag
            drc_plot=drc_plot,  # Pass DRC plot mode flag
            slide_axis=slide_axis,  # Pass slide axis variable name
            grid_axis=grid_axis,  # Pass grid axis variable name
            arrangement=arrangement,  # Pass arrangement (e.g., "2*3")
            confirmed_scan_results=confirmed_scan_results,  # Pass confirmed scan results
            preview_cell_order=preview_cell_order  # Pass preview cell order (preserves drag-and-drop order)
        )
        
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("✅ Slide generation completed successfully!")
        thread_instance.log_signal.emit("=" * 60)
        
        return True, "Slide generation completed successfully!"
        
    except Exception as e:
        import traceback
        error_msg = f"Error during generation: {str(e)}"
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("❌ An error occurred during slide generation!")
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("Error Details:")
        thread_instance.log_signal.emit(f"Error Type: {type(e).__name__}")
        thread_instance.log_signal.emit(f"Error Message: {str(e)}")
        thread_instance.log_signal.emit("-" * 60)
        thread_instance.log_signal.emit("Stack Trace:")
        tb_str = traceback.format_exc()
        thread_instance.log_signal.emit(tb_str)
        thread_instance.log_signal.emit("=" * 60)
        return False, error_msg
