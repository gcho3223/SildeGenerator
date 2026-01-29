#!/usr/bin/env python3
"""
CPV Mode UI components and logic for Keynote Slide Generator
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox,
    QGroupBox, QComboBox, QFrame
)


def create_cpv_options_ui(parent):
    """
    Create CPV-specific options UI (Year, Channel, File Format)
    Returns: (year_channel_group, year_combo, channel_combo, cpv_file_format_combo)
    """
    year_channel_group = QGroupBox("CPV Options")
    year_channel_layout = QHBoxLayout()
    
    # Year selection
    year_channel_layout.addWidget(QLabel("Year:"))
    year_combo = QComboBox()
    year_combo.addItems(["UL2016PreVFP", "UL2016PostVFP", "UL2017", "UL2018"])
    year_combo.setCurrentText(parent.year)
    year_channel_layout.addWidget(year_combo)
    
    year_channel_layout.addSpacing(20)
    
    # Channel selection
    year_channel_layout.addWidget(QLabel("Channel:"))
    channel_combo = QComboBox()
    channel_combo.addItems(["Dimuon", "Dielectron", "EMu"])
    channel_combo.setCurrentText(parent.channel)
    year_channel_layout.addWidget(channel_combo)
    
    year_channel_layout.addSpacing(20)
    
    # File Format selection (CPV)
    year_channel_layout.addWidget(QLabel("File Format:"))
    cpv_file_format_combo = QComboBox()
    cpv_file_format_combo.addItems(["pdf", "png"])
    cpv_file_format_combo.setCurrentText("pdf")
    year_channel_layout.addWidget(cpv_file_format_combo)
    
    year_channel_layout.addStretch()
    
    year_channel_group.setLayout(year_channel_layout)
    
    return year_channel_group, year_combo, channel_combo, cpv_file_format_combo


def create_cpv_objects_ui(parent):
    """
    Create CPV object selection UI
    Returns: (cpv_objects_widget, object_checks, after_top_reco_objects)
    """
    cpv_objects_widget = QWidget()
    cpv_objects_layout = QVBoxLayout()
    cpv_objects_layout.setContentsMargins(0, 0, 0, 0)

    cpv_channel_label_layout = QHBoxLayout()
    cpv_channel_label_layout.addWidget(QLabel("Event selection stage"))
    cpv_channel_label_layout.addStretch()
    # Select all checkbox for event selection stage objects
    select_all_above_check = QCheckBox("Select all")
    select_all_above_check.setChecked(False)
    cpv_channel_label_layout.addWidget(select_all_above_check)
    cpv_objects_layout.addLayout(cpv_channel_label_layout)
    
    # Objects above separator (regular objects)
    object_names_above = [
        ["Num_PV"],
        ["Mass"],
        ["Num_Jets"],
        ["Lep1", "Lep2"],
        ["Jet1", "Jet2"],
        ["bJet1", "bJet2"],
        ["MET"],
        ["Observable (O1, O3)"],
    ]
    
    object_checks = {}
    for row in object_names_above:
        row_layout = QHBoxLayout()
        for obj_name in row:
            check = QCheckBox(obj_name)
            object_checks[obj_name] = check
            # Connect Observable checkbox to auto-select Observable step
            if "Observable" in obj_name:
                check.toggled.connect(parent.on_observable_changed)
            # Connect to update select_all_above_check state
            check.toggled.connect(lambda checked, name=obj_name: parent.update_select_all_above_state())
            row_layout.addWidget(check)
        if len(row) == 1:
            row_layout.addStretch()
        cpv_objects_layout.addLayout(row_layout)
    
    # Add separator line
    separator = QFrame()
    separator.setFrameShape(QFrame.HLine)
    separator.setFrameShadow(QFrame.Sunken)
    separator.setLineWidth(2)
    cpv_objects_layout.addWidget(separator)

    cpv_top_reco_label_layout = QHBoxLayout()
    cpv_top_reco_label_layout.addWidget(QLabel("After Top quark reconstruction stage"))
    cpv_top_reco_label_layout.addStretch()
    # Select all checkbox for after top quark reco objects
    select_all_below_check = QCheckBox("Select all")
    select_all_below_check.setChecked(False)
    cpv_top_reco_label_layout.addWidget(select_all_below_check)
    cpv_objects_layout.addLayout(cpv_top_reco_label_layout)
    
    # Objects below separator (afterTopReco objects)
    object_names_below = [
        ["Nu", "AnNu"],
        ["bJet", "AnbJet"],
        ["Top", "AnTop"],
    ]
    
    # Track objects below separator
    after_top_reco_objects = []
    
    for row in object_names_below:
        row_layout = QHBoxLayout()
        for obj_name in row:
            check = QCheckBox(obj_name)
            object_checks[obj_name] = check
            after_top_reco_objects.append(obj_name)
            # Connect signal to auto-select afterTopReco
            check.toggled.connect(parent.on_after_top_reco_object_changed)
            # Connect to update select_all_below_check state
            check.toggled.connect(lambda checked, name=obj_name: parent.update_select_all_below_state())
            row_layout.addWidget(check)
        if len(row) == 1:
            row_layout.addStretch()
        cpv_objects_layout.addLayout(row_layout)
    
    cpv_objects_widget.setLayout(cpv_objects_layout)
    
    # Store select all checkboxes in parent for later connection
    parent.select_all_above_check = select_all_above_check
    parent.select_all_below_check = select_all_below_check
    
    return cpv_objects_widget, object_checks, after_top_reco_objects


def create_cpv_steps_ui(parent):
    """
    Create CPV step selection UI
    Returns: (cpv_steps_widget, step_checks)
    """
    from PyQt5.QtWidgets import QGridLayout, QCheckBox
    
    cpv_steps_widget = QWidget()
    cpv_steps_main_layout = QVBoxLayout()
    cpv_steps_main_layout.setContentsMargins(0, 0, 0, 0)
    
    # Step selection using grid layout for alignment
    cpv_steps_layout = QGridLayout()
    cpv_steps_layout.setContentsMargins(0, 0, 0, 0)
    
    # Format: (row, col, display_label, internal_value, colspan)
    # Note: "afterTopReco" and "Observable" are not shown as checkboxes
    # but are auto-selected based on object selection (logic preserved)
    step_data = [
        (0, 0, "Initial stage", "initial", 2),
        (1, 0, "Dilepton mass cut (step1)", "step1", 1),
        (1, 1, "Z mass veto (step2)", "step2", 1),
        (2, 0, "# of Jets ≥ 2 (step3)", "step3", 1),
        (2, 1, "MET cut (step4)", "step4", 1),
        (3, 0, "b-tagging jet ≥ 1 (step5)", "step5", 1),
        (3, 1, "Top Reco (step6)", "step6", 1),
        # afterTopReco and Observable are not shown but logic is preserved
    ]
    
    step_checks = {}
    
    for row, col, display_label, internal_value, colspan in step_data:
        check = QCheckBox(display_label)
        # Default: no selection (all unchecked)
        check.setChecked(False)
        step_checks[internal_value] = check  # Use internal value as key
        # Connect to update select_all_steps_check state
        check.toggled.connect(lambda checked, name=internal_value: parent.update_select_all_steps_state())
        cpv_steps_layout.addWidget(check, row, col, 1, colspan)
    
    # Add hidden checkboxes for afterTopReco and Observable (for auto-selection logic)
    # These are not displayed but needed for the auto-selection logic to work
    hidden_afterTopReco = QCheckBox()
    hidden_afterTopReco.setVisible(False)
    step_checks["afterTopReco"] = hidden_afterTopReco
    
    hidden_observable = QCheckBox()
    hidden_observable.setVisible(False)
    step_checks["Observable"] = hidden_observable
    
    cpv_steps_main_layout.addLayout(cpv_steps_layout)
    cpv_steps_widget.setLayout(cpv_steps_main_layout)
    
    return cpv_steps_widget, step_checks


def run_cpv_generation(thread_instance):
    """
    Run CPV mode slide generation
    Called from GeneratorThread.run()
    """
    # Check for interruption at the start
    if thread_instance.isInterruptionRequested():
        return False, "Generation cancelled"
    
    try:
        from Config.config_cpv import cpv_config
        from Config.KeynoteCtrl import prepare_keynote, insert_pdfs_into_slide, insert_pdfs_into_slide_systematic
        
        thread_instance.log_signal.emit(f"Selected objects: {thread_instance.selected_objects}")
        thread_instance.log_signal.emit(f"Selected steps: {thread_instance.selected_steps}")
        
        output_file = thread_instance.output_file
        
        # Add _sys suffix if systematic
        if thread_instance.systematic:
            base, ext = os.path.splitext(output_file)
            output_file = f"{base}_sys{ext}"
        
        thread_instance.log_signal.emit(f"Output file: {output_file}")
        thread_instance.log_signal.emit("-" * 60)
        
        # Create Keynote file
        thread_instance.log_signal.emit("Creating Keynote file...")
        prepare_keynote(output_file, theme="White")
        thread_instance.log_signal.emit("Keynote file created successfully!")
        
        if thread_instance.systematic:
            # Systematic mode
            sample_dirs = cpv_config["systematic_samples"]
            sample_names = cpv_config["categorizedMC"]
            
            thread_instance.log_signal.emit("Inserting systematic comparison plots...")
            
            # Get plot width from settings if available (height will be proportional)
            plot_width = None
            if thread_instance.user_settings:
                plot_settings = thread_instance.user_settings.get('plot', {})
                if 'width' in plot_settings:
                    plot_width = plot_settings.get('width')
            
            insert_pdfs_into_slide_systematic(
                output_file=output_file,
                input_dir=thread_instance.input_dir,
                selected_objects=thread_instance.selected_objects,
                selected_steps=thread_instance.selected_steps,
                sample_dirs=sample_dirs,
                sample_names=sample_names,
                file_format=thread_instance.file_format,
                log_callback=thread_instance.log_signal.emit,
                plot_width=plot_width,
                thread_instance=thread_instance
            )
        else:
            # Normal CPV mode
            # Build plot_dir: input_dir + jobversion (same as v1.0: base_dir + "/" + version)
            # v1.0: plot_dir = base_dir + "/" + version
            plot_dir = os.path.join(thread_instance.input_dir, thread_instance.jobversion) if thread_instance.jobversion else thread_instance.input_dir
            thread_instance.log_signal.emit(f"Plot directory: {plot_dir}")
            thread_instance.log_signal.emit("Inserting plots...")
            
            # Determine position option from settings
            position_opt = "Center"  # Default
            custom_positions = None
            plot_width = None
            arrangement_positions = None  # Initialize
            arrangement_sizes = None  # Initialize
            
            if thread_instance.user_settings:
                cpv_settings = thread_instance.user_settings.get('cpv', {})
                plot_settings = thread_instance.user_settings.get('plot', {})
                
                if cpv_settings.get('use_custom', False):
                    # Use custom positions and sizes from Plot tab
                    thread_instance.log_signal.emit("Using customized size and positions")
                    custom_positions = plot_settings.get('custom_positions', {})
                    arrangement_positions = plot_settings.get('arrangement_positions', {})
                    arrangement_sizes = plot_settings.get('arrangement_sizes', {})
                    thread_instance.log_signal.emit(f"[DEBUG] custom_positions keys: {list(custom_positions.keys())}")
                    thread_instance.log_signal.emit(f"[DEBUG] arrangement_positions: {arrangement_positions}")
                    thread_instance.log_signal.emit(f"[DEBUG] arrangement_sizes: {arrangement_sizes}")
                else:
                    # Use default position
                    position_opt = cpv_settings.get('position_opt', 'Center')
                    thread_instance.log_signal.emit(f"Using default position: {position_opt}")
            
            insert_pdfs_into_slide(
                output_file=output_file,
                plot_dir=plot_dir,
                selected_objects=thread_instance.selected_objects,
                selected_steps=thread_instance.selected_steps,
                positionOpt=position_opt,
                file_format=thread_instance.file_format,
                log_callback=thread_instance.log_signal.emit,
                custom_positions=custom_positions,
                plot_width=plot_width,
                arrangement_positions=arrangement_positions,
                arrangement_sizes=arrangement_sizes,
                thread_instance=thread_instance
            )
        
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("✅ Slides generated successfully!")
        thread_instance.log_signal.emit("=" * 60)
        
        return True, "Keynote slides generated successfully!"
        
    except Exception as e:
        raise e
