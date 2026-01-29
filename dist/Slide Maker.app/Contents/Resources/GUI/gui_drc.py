#!/usr/bin/env python3
"""
DRC Mode UI components and logic for Keynote Slide Generator
"""

import os
import sys
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox,
    QGroupBox, QComboBox, QFrame, QGridLayout
)
from PyQt5.QtCore import Qt
from GUI.common import DraggableLabel

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def create_drc_options_ui(parent):
    """
    Create DRC-specific options UI (Particle, Case, File Format)
    Returns: (drc_options_group, particle_combo, case_combo, file_format_combo)
    """
    drc_options_group = QGroupBox("DRC Options")
    drc_options_layout = QVBoxLayout()
    
    # First row: Particle and Case
    particle_case_layout = QHBoxLayout()
    
    # Particle selection
    particle_case_layout.addWidget(QLabel("Particle:"))
    particle_combo = QComboBox()
    particle_combo.addItems(["EM", "Pion", "Proton", "Kaon"])
    particle_combo.setCurrentText("Proton")
    particle_case_layout.addWidget(particle_combo)
    
    particle_case_layout.addSpacing(20)
    
    # Case selection
    particle_case_layout.addWidget(QLabel("Case:"))
    case_combo = QComboBox()
    case_combo.addItems(["Normal", "Rotation & Tilting", "Interaction Target"])
    case_combo.setCurrentText("Normal")
    case_combo.setFixedWidth(100)
    particle_case_layout.addWidget(case_combo)
    
    particle_case_layout.addSpacing(20)
    
    # File Format selection
    particle_case_layout.addWidget(QLabel("File Format:"))
    file_format_combo = QComboBox()
    file_format_combo.addItems(["pdf", "png"])
    file_format_combo.setCurrentText("pdf")
    particle_case_layout.addWidget(file_format_combo)
    
    particle_case_layout.addStretch()
    drc_options_layout.addLayout(particle_case_layout)
    
    drc_options_group.setLayout(drc_options_layout)
    
    return drc_options_group, particle_combo, case_combo, file_format_combo


def create_drc_channels_ui(parent):
    """
    Create DRC channel and energy selection UI
    Returns: (drc_channels_widget, c_check, s_combo, drcor_combo, energy_checks, 
              low_energy_check, energy_checkboxes_layout, resol_with_noise_check, 
              resol_without_noise_check, lin_check)
    """
    drc_channels_widget = QWidget()
    drc_channels_layout = QVBoxLayout()
    drc_channels_layout.setContentsMargins(0, 0, 0, 0)
    
    # Channel label
    channel_label_layout = QHBoxLayout()
    channel_label_layout.addWidget(QLabel("Channel"))
    channel_label_layout.addStretch()
    drc_channels_layout.addLayout(channel_label_layout)
    
    # Channel selection (C, S, DRcor)
    channel_options_layout = QHBoxLayout()
    
    # C checkbox
    c_check = QCheckBox("C")
    c_check.setChecked(True)
    channel_options_layout.addWidget(c_check)
    channel_options_layout.addSpacing(20)
    
    # S dropdown
    channel_options_layout.addWidget(QLabel("S:"))
    s_combo = QComboBox()
    s_combo.addItems(["None", "S", "ATTcor", "LCcor", "LC+ATTcor"])
    s_combo.setCurrentText("LC+ATTcor")
    s_combo.setMinimumWidth(100)
    channel_options_layout.addWidget(s_combo)
    channel_options_layout.addSpacing(20)
    
    # DRcor dropdown
    channel_options_layout.addWidget(QLabel("DRcor:"))
    drcor_combo = QComboBox()
    drcor_combo.addItems(["None", "DRcor", "ATTcor", "LCcor", "LC+ATTcor"])
    drcor_combo.setCurrentText("LC+ATTcor")
    drcor_combo.setMinimumWidth(100)
    channel_options_layout.addWidget(drcor_combo)
    
    channel_options_layout.addStretch()
    drc_channels_layout.addLayout(channel_options_layout)

    # Add separator line
    separator = QFrame()
    separator.setFrameShape(QFrame.HLine)
    separator.setFrameShadow(QFrame.Sunken)
    separator.setLineWidth(2)
    drc_channels_layout.addWidget(separator)

    # Energy points label with low energy toggle
    energy_label_layout = QHBoxLayout()
    energy_label_layout.addWidget(QLabel("Energy"))
    
    low_energy_check = QCheckBox("Low Energy")
    low_energy_check.stateChanged.connect(parent.toggle_energy_range)
    energy_label_layout.addWidget(low_energy_check)
    
    energy_label_layout.addStretch()
    drc_channels_layout.addLayout(energy_label_layout)
    
    # Energy points checkboxes container
    energy_checkboxes_layout = QVBoxLayout()
    energy_checks = {}
    
    # Note: Energy checkboxes will be created by create_energy_checkboxes() in parent
    drc_channels_layout.addLayout(energy_checkboxes_layout)
    
    # Add separator line
    separator2 = QFrame()
    separator2.setFrameShape(QFrame.HLine)
    separator2.setFrameShadow(QFrame.Sunken)
    separator2.setLineWidth(2)
    drc_channels_layout.addWidget(separator2)

    # Resolution and linearity label
    resol_label_layout = QHBoxLayout()
    resol_label_layout.addWidget(QLabel("Resolution & Linearity"))
    resol_label_layout.addStretch()
    drc_channels_layout.addLayout(resol_label_layout)

    # Energy resolution
    resol_layout = QHBoxLayout()
    resol_layout.addWidget(QLabel("Resolution:"))
    
    resol_with_noise_check = QCheckBox("with Noise term")
    resol_with_noise_check.setChecked(True)
    resol_layout.addWidget(resol_with_noise_check)
    
    resol_without_noise_check = QCheckBox("without Noise term")
    resol_without_noise_check.setChecked(True)
    resol_layout.addWidget(resol_without_noise_check)
    
    resol_layout.addStretch()
    drc_channels_layout.addLayout(resol_layout)
    
    # Linearity
    lin_layout = QHBoxLayout()
    lin_check = QCheckBox("Linearity")
    lin_check.setChecked(True)
    lin_layout.addWidget(lin_check)
    lin_layout.addStretch()
    drc_channels_layout.addLayout(lin_layout)

    drc_channels_widget.setLayout(drc_channels_layout)
    
    return (drc_channels_widget, c_check, s_combo, drcor_combo, energy_checks, 
            low_energy_check, energy_checkboxes_layout, resol_with_noise_check, 
            resol_without_noise_check, lin_check)


def create_drc_pid_ui():
    """
    Create DRC PID info widget (text only, no actual steps)
    Returns: drc_pid_widget
    """
    drc_pid_widget = QWidget()
    drc_pid_layout = QVBoxLayout()
    drc_pid_layout.setContentsMargins(0, 0, 0, 0)
    
    drc_pid_info = QLabel(
        "PID steps are not supported yet...\n"
        "You can only generate slides for the last selection."
    )
    drc_pid_info.setWordWrap(True)
    drc_pid_info.setStyleSheet("color: #555; padding: 5px;")
    drc_pid_layout.addWidget(drc_pid_info)
    
    drc_pid_widget.setLayout(drc_pid_layout)
    
    return drc_pid_widget


def create_drc_preview_ui(parent):
    """
    Create DRC energy preview UI with tabs (C, S, DRcor)
    Returns: (preview_group, energy_preview_cells, page_tabs, tab_grids, tab_cells)
    """
    from PyQt5.QtWidgets import QTabWidget, QWidget
    
    preview_group = QGroupBox("Preview")
    preview_layout = QVBoxLayout()
    
    # Create tab widget for C, S, DRcor
    page_tabs = QTabWidget()
    page_tabs.setVisible(True)
    
    # Store tab grids and cells
    tab_grids = []
    tab_cells = []
    energy_preview_cells = []  # Keep for backward compatibility
    
    # Create tabs for C, S, DRcor
    channel_names = ["C", "S", "DRcor"]
    for channel_idx, channel_name in enumerate(channel_names):
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
        page_tabs.addTab(tab_widget, channel_name)
        
        # Store grid layout and cells
        tab_grids.append(tab_grid_layout)
        tab_cells.append([])
        
        # Initialize preview cells (2x3 grid for energy points) - Use DraggableLabel
        for row in range(2):
            for col in range(3):
                cell = DraggableLabel("Empty", parent, cell_list=tab_cells[channel_idx])
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
                tab_cells[channel_idx].append(cell)
                energy_preview_cells.append(cell)  # Keep for backward compatibility
    
    preview_layout.addWidget(page_tabs)
    preview_group.setLayout(preview_layout)
    
    return preview_group, energy_preview_cells, page_tabs, tab_grids, tab_cells


def create_energy_checkboxes(parent, energy_values):
    """
    Create energy checkboxes from given values
    Called when switching between high/low energy modes
    """
    # Clear existing checkboxes
    while parent.energy_checkboxes_layout.count():
        child = parent.energy_checkboxes_layout.takeAt(0)
        if child.widget():
            child.widget().deleteLater()
        elif child.layout():
            while child.layout().count():
                subchild = child.layout().takeAt(0)
                if subchild.widget():
                    subchild.widget().deleteLater()
    
    # Clear energy_checks dictionary
    parent.energy_checks.clear()
    
    # Create new checkboxes
    for row in energy_values:
        row_layout = QHBoxLayout()
        for energy in row:
            check = QCheckBox(energy)
            check.setChecked(False)
            parent.energy_checks[energy] = check
            # Connect to update preview
            check.stateChanged.connect(parent.update_energy_preview)
            row_layout.addWidget(check)
        row_layout.addStretch()
        parent.energy_checkboxes_layout.addLayout(row_layout)


def update_energy_preview(parent):
    """Update the energy preview layout in DRC mode (with tabs for C, S, DRcor)"""
    selected_energies = [energy for energy, check in parent.energy_checks.items() if check.isChecked()]
    
    # Sort energies by numeric value
    sorted_energies = sorted(selected_energies, key=lambda x: float(x.replace("GeV", "")))
    
    # Update each tab (C, S, DRcor)
    if hasattr(parent, 'drc_tab_cells') and parent.drc_tab_cells:
        for channel_idx, tab_cells in enumerate(parent.drc_tab_cells):
            # Reset all cells to empty for this tab
            for cell in tab_cells:
                cell.setText("Empty")
                cell.setStyleSheet(
                    "border: 2px dashed #ccc; "
                    "background-color: #f5f5f5; "
                    "color: #999; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )
            
            # Fill in selected energies (2x3 layout) for this tab
            for idx, energy in enumerate(sorted_energies):
                if idx < len(tab_cells):
                    # Style for selected energy
                    tab_cells[idx].setText(energy)
                    tab_cells[idx].setStyleSheet(
                        "border: 2px solid #0066cc; "
                        "background-color: #e6f2ff; "
                        "color: #0066cc; "
                        "font-weight: bold; "
                        "border-radius: 5px; "
                        "padding: 5px;"
                    )
    else:
        # Fallback to old behavior if tabs not available
        # Reset all cells to empty
        for cell in parent.energy_preview_cells:
            cell.setText("Empty")
            cell.setStyleSheet(
                "border: 2px dashed #ccc; "
                "background-color: #f5f5f5; "
                "color: #999; "
                "border-radius: 5px; "
                "padding: 5px;"
            )
        
        # Fill in selected energies (2x3 layout)
        for idx, energy in enumerate(sorted_energies):
            if idx < len(parent.energy_preview_cells):
                # Style for selected energy
                parent.energy_preview_cells[idx].setText(energy)
                parent.energy_preview_cells[idx].setStyleSheet(
                    "border: 2px solid #0066cc; "
                    "background-color: #e6f2ff; "
                    "color: #0066cc; "
                    "font-weight: bold; "
                    "border-radius: 5px; "
                    "padding: 5px;"
                )


def run_drc_generation(thread_instance):
    """
    Run DRC mode slide generation
    Called from GeneratorThread.run()
    """
    # Check for interruption at the start
    if thread_instance.isInterruptionRequested():
        return False, "Generation cancelled"
    
    try:
        from Config.config_drc import drc_config
        from Config.KeynoteCtrl import prepare_keynote, insert_pdfs_into_slide_drc, insert_pdfs_into_slide_drc_resol_linearity
        
        # Map particle to internal name
        particle_map = {
            "EM": "em",
            "Pion": "pi",
            "Proton": "proton",
            "Kaon": "kaon"
        }
        particle_name = particle_map.get(thread_instance.particle, "proton")
        
        # Map case to suffix
        if thread_instance.case == "Normal":
            case_suffix = f"_{particle_name}"
        elif thread_instance.case == "Rotation & Tilting":
            case_suffix = "_Rot"
        elif thread_instance.case == "Interaction Target":
            case_suffix = "_IT"
        else:
            case_suffix = f"_{particle_name}"
        
        # Construct program name
        program = f"{particle_name}{case_suffix}"
        programs = [program]
        
        # Construct channels list
        channels = []
        if thread_instance.c_checked:
            channels.append("C")
        
        # Add S channel (convert LC+ATTcor to LCATTcor) - skip if "none"
        if thread_instance.s_selected != "None":
            s_channel_name = thread_instance.s_selected.replace("LC+ATTcor", "LCATTcor")
            channels.append(f"S_{s_channel_name}")
        
        # Add DRcor channel (convert LC+ATTcor to LCATTcor) - skip if "none"
        if thread_instance.drcor_selected != "None":
            drcor_channel_name = thread_instance.drcor_selected.replace("LC+ATTcor", "LCATTcor")
            channels.append(f"DRcor_{drcor_channel_name}")
        
        # Convert selected energies from GUI format (e.g., "20GeV") to internal format (e.g., "energy20_run20")
        # Keep "Empty" to preserve positions
        # The order is already set by the preview grid (after drag and drop)
        # selected_energies can be either a dict (channel_name -> energies) or a list (legacy)
        energies_dict = {}  # Dictionary: channel_name -> list of energies (internal format)
        
        if thread_instance.selected_energies:
            if isinstance(thread_instance.selected_energies, dict):
                # New format: dictionary with channel names as keys
                for channel_name, energies_gui in thread_instance.selected_energies.items():
                    energies_internal = []
                    for idx, energy_gui in enumerate(energies_gui):
                        if energy_gui == "Empty":
                            energies_internal.append("Empty")
                            thread_instance.log_signal.emit(f"Channel {channel_name}, Position {idx}: Empty (No plot)")
                        else:
                            # Extract number from "XXGeV" format
                            energy_num = energy_gui.replace("GeV", "")
                            energy_internal = f"energy{energy_num}_run{energy_num}"
                            energies_internal.append(energy_internal)
                            thread_instance.log_signal.emit(f"Channel {channel_name}, Position {idx}: {energy_gui} -> {energy_internal}")
                    energies_dict[channel_name] = energies_internal
            else:
                # Legacy format: single list (convert to dict for all channels)
                energies = []
                for idx, energy_gui in enumerate(thread_instance.selected_energies):
                    if energy_gui == "Empty":
                        energies.append("Empty")
                        thread_instance.log_signal.emit(f"Position {idx}: Empty (No plot)")
                    else:
                        # Extract number from "XXGeV" format
                        energy_num = energy_gui.replace("GeV", "")
                        energy_internal = f"energy{energy_num}_run{energy_num}"
                        energies.append(energy_internal)
                        thread_instance.log_signal.emit(f"Position {idx}: {energy_gui} -> {energy_internal}")
                # Use same energies for all channels (legacy behavior)
                for channel in channels:
                    energies_dict[channel] = energies
        
        base_dir = thread_instance.input_dir
        output_file = thread_instance.output_file
        
        thread_instance.log_signal.emit(f"Particle: {thread_instance.particle} -> {particle_name}")
        thread_instance.log_signal.emit(f"Case: {thread_instance.case} -> {case_suffix}")
        thread_instance.log_signal.emit(f"Programs: {programs}")
        thread_instance.log_signal.emit(f"C checked: {thread_instance.c_checked}")
        thread_instance.log_signal.emit(f"S selected: {thread_instance.s_selected}")
        thread_instance.log_signal.emit(f"DRcor selected: {thread_instance.drcor_selected}")
        thread_instance.log_signal.emit(f"Channels: {channels}")
        if energies_dict:
            for channel_name, energies_list in energies_dict.items():
                thread_instance.log_signal.emit(f"Energies for {channel_name}: {energies_list}")
        else:
            thread_instance.log_signal.emit(f"Energies: None (Resolution/Linearity only)")
        thread_instance.log_signal.emit(f"Resolution with Noise: {thread_instance.resol_with_noise}")
        thread_instance.log_signal.emit(f"Resolution without Noise: {thread_instance.resol_without_noise}")
        thread_instance.log_signal.emit(f"Linearity: {thread_instance.linearity}")
        thread_instance.log_signal.emit(f"Base directory: {base_dir}")
        thread_instance.log_signal.emit(f"Output file: {output_file}")
        thread_instance.log_signal.emit("-" * 60)
        
        # Create Keynote file
        thread_instance.log_signal.emit("Creating Keynote file...")
        prepare_keynote(output_file, theme="White")
        thread_instance.log_signal.emit("Keynote file created successfully!")
        
        # Insert energy plots if selected
        if energies_dict:
            thread_instance.log_signal.emit("Inserting DRC energy plots...")
            
            # Get size and positions from config_drc.py (default)
            size = drc_config["size"]
            positions = drc_config["positions"]
            thread_instance.log_signal.emit(f"Using DRC config positions: {positions}")
            custom_positions = None
            plot_width = None
            arrangement_sizes = None
            
            if thread_instance.user_settings:
                drc_settings = thread_instance.user_settings.get('drc', {})
                plot_settings = thread_instance.user_settings.get('plot', {})
                
                if drc_settings.get('use_custom', False):
                    # Use custom positions and sizes from Plot tab
                    thread_instance.log_signal.emit("Using customized size and positions for DRC")
                    custom_positions = plot_settings.get('custom_positions', {})
                    arrangement_sizes = plot_settings.get('arrangement_sizes', {})
                    thread_instance.log_signal.emit(f"[DEBUG] custom_positions keys: {list(custom_positions.keys())}")
                    thread_instance.log_signal.emit(f"[DEBUG] arrangement_sizes: {arrangement_sizes}")
                    
                    # Get arrangement-specific width for 2*3 layout if available
                    if arrangement_sizes and "2*3" in arrangement_sizes:
                        plot_width = arrangement_sizes["2*3"]
                        thread_instance.log_signal.emit(f"[DEBUG] Using arrangement-specific width for 2*3: {plot_width}")
                    
                    # Check for 2*3 arrangement (DRC uses 2*3 layout)
                    # Try different position options for 2*3
                    for pos_opt in ["Center", "Upper", "Bottom"]:
                        arrangement_key = f"2*3_{pos_opt}"
                        if custom_positions and arrangement_key in custom_positions:
                            positions = custom_positions[arrangement_key]
                            thread_instance.log_signal.emit(f"Using custom positions for {arrangement_key}")
                            break
                    else:
                        # If no custom position found, use default DRC positions from config_drc.py
                        positions = drc_config["positions"]
                        thread_instance.log_signal.emit("Custom positions not found, using default DRC positions from config_drc.py")
                else:
                    thread_instance.log_signal.emit(f"Using default DRC size and positions from config_drc.py: size={size}, positions={positions}")
            
            insert_pdfs_into_slide_drc(
                output_file=output_file,
                base_dir=base_dir,
                programs=programs,
                channels=channels,
                energies=energies_dict,  # Pass dictionary instead of list
                size=size,
                positions=positions,
                log_callback=thread_instance.log_signal.emit,
                plot_width=plot_width,
                thread_instance=thread_instance
            )
        
        # Insert Resolution/Linearity plots if selected
        if thread_instance.resol_with_noise or thread_instance.resol_without_noise or thread_instance.linearity:
            thread_instance.log_signal.emit("Inserting Resolution/Linearity plots...")
            
            # Get plot width and position option from settings if available
            plot_width = None
            position_opt = "Center"  # Default
            if thread_instance.user_settings:
                drc_settings = thread_instance.user_settings.get('drc', {})
                plot_settings = thread_instance.user_settings.get('plot', {})
                
                # Get default position option for Resolution/Linearity plots
                position_opt = drc_settings.get('position_opt', 'Center')
                
                if drc_settings.get('use_custom', False):
                    # Get arrangement-specific sizes for Resolution/Linearity plots
                    arrangement_sizes = plot_settings.get('arrangement_sizes', {})
                    # Resolution/Linearity plots use different layouts (1*1, 1*2, 1*3)
                    # Determine which layout based on number of plots
                    num_plots = sum([
                        thread_instance.resol_with_noise,
                        thread_instance.resol_without_noise,
                        thread_instance.linearity
                    ])
                    
                    if num_plots == 3:
                        # 1*3 layout
                        if arrangement_sizes and "1*3" in arrangement_sizes:
                            plot_width = arrangement_sizes["1*3"]
                    elif num_plots == 2:
                        # 1*2 layout
                        if arrangement_sizes and "1*2" in arrangement_sizes:
                            plot_width = arrangement_sizes["1*2"]
                    elif num_plots == 1:
                        # 1*1 layout
                        if arrangement_sizes and "1*1" in arrangement_sizes:
                            plot_width = arrangement_sizes["1*1"]
                    
                    if plot_width:
                        thread_instance.log_signal.emit(f"[DEBUG] Using arrangement-specific width for Resolution/Linearity: {plot_width}")
                elif 'width' in plot_settings:
                    # Legacy: single width for all arrangements
                    plot_width = plot_settings.get('width')
            
            insert_pdfs_into_slide_drc_resol_linearity(
                output_file=output_file,
                base_dir=base_dir,
                programs=programs,
                particle_name=particle_name,
                resol_with_noise=thread_instance.resol_with_noise,
                resol_without_noise=thread_instance.resol_without_noise,
                linearity=thread_instance.linearity,
                file_format=thread_instance.file_format,
                log_callback=thread_instance.log_signal.emit,
                plot_width=plot_width,
                positionOpt=position_opt,
                thread_instance=thread_instance
            )
        
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("✅ Slides generated successfully!")
        thread_instance.log_signal.emit("=" * 60)
        
        return True, "Keynote slides generated successfully!"
        
    except Exception as e:
        raise e
