#!/usr/bin/env python3
"""
Drag and Drop Mode UI components for Keynote Slide Generator
Simplified mode with only preview (no preset, directory, options, or variables)
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QGroupBox, QComboBox, QGridLayout, QTabWidget, QLineEdit, QSpinBox
)
from PyQt5.QtCore import Qt
from GUI.common import DraggableLabel
import os


def create_dragdrop_preview_ui(parent):
    """
    Create Drag and Drop mode preview UI with Arrangement and Slide count selection
    Returns: (preview_group, preview_cells, arrangement_combo, slide_count_combo, 
              slide_count_input, page_tabs, tab_grids, tab_cells)
    """
    preview_group = QGroupBox("Preview")
    preview_layout = QVBoxLayout()
    
    # Arrangement and Slide count selection
    arrangement_layout = QHBoxLayout()
    arrangement_layout.addWidget(QLabel("Arrangement"))
    arrangement_combo = QComboBox()
    arrangement_combo.addItems(["1*1", "1*2", "1*3", "1*4", "2*1", "2*2", "2*3", "2*4"])
    arrangement_combo.setCurrentText("2*3")
    arrangement_combo.setMinimumWidth(100)
    arrangement_layout.addWidget(arrangement_combo)
    
    arrangement_layout.addSpacing(20)
    arrangement_layout.addWidget(QLabel("Slides:"))
    slide_count_combo = QComboBox()
    slide_count_combo.addItems([str(i) for i in range(1, 11)] + ["Custom"])
    slide_count_combo.setCurrentText("1")
    slide_count_combo.setMinimumWidth(80)
    arrangement_layout.addWidget(slide_count_combo)
    
    # Custom slide count input (initially hidden)
    slide_count_input = QSpinBox()
    slide_count_input.setMinimum(1)
    slide_count_input.setMaximum(100)
    slide_count_input.setValue(1)
    slide_count_input.setMinimumWidth(80)
    slide_count_input.setVisible(False)
    arrangement_layout.addWidget(slide_count_input)
    
    arrangement_layout.addStretch()
    preview_layout.addLayout(arrangement_layout)
    
    # Tab widget for multiple slides
    page_tabs = QTabWidget()
    page_tabs.setVisible(True)
    preview_layout.addWidget(page_tabs)
    
    # Store tab grids and cells
    tab_grids = []
    tab_cells = []
    preview_cells = []  # Keep for compatibility
    
    # Initialize file paths dictionary if not exists
    if not hasattr(parent, 'dragdrop_file_paths'):
        parent.dragdrop_file_paths = {}  # Format: {slide_idx: {cell_index: file_path}}
    
    # Create initial tab (1 slide)
    _create_dragdrop_tab(page_tabs, tab_grids, tab_cells, parent, 0)
    
    # Add info label below tabs
    info_label = QLabel("You can adjust the size and positions of plots in Settings.")
    info_label.setWordWrap(True)
    info_label.setStyleSheet("color: #666; padding: 5px; font-size: 10pt;")
    info_label.setAlignment(Qt.AlignCenter)
    preview_layout.addWidget(info_label)
    
    preview_group.setLayout(preview_layout)
    
    return (preview_group, preview_cells, arrangement_combo, slide_count_combo, 
            slide_count_input, page_tabs, tab_grids, tab_cells)


def _create_dragdrop_tab(page_tabs, tab_grids, tab_cells, parent, slide_idx):
    """Create a single tab for dragdrop preview"""
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
    page_tabs.addTab(tab_widget, f"Slide {slide_idx + 1}")
    
    # Store grid layout
    tab_grids.append(tab_grid_layout)
    tab_cells.append([])
    
    # Initialize file paths for this slide if not exists
    if not hasattr(parent, 'dragdrop_file_paths'):
        parent.dragdrop_file_paths = {}
    if slide_idx not in parent.dragdrop_file_paths:
        parent.dragdrop_file_paths[slide_idx] = {}
    
    return tab_grid_layout, tab_cells[slide_idx]


def update_dragdrop_preview(parent):
    """Update the preview layout in Drag and Drop mode based on arrangement and slide count"""
    if not hasattr(parent, 'dragdrop_page_tabs') or not hasattr(parent, 'dragdrop_tab_grids'):
        return
    
    # Get arrangement
    if hasattr(parent, 'dragdrop_arrangement_combo'):
        arrangement = parent.dragdrop_arrangement_combo.currentText()
        try:
            num_rows, num_cols = map(int, arrangement.split('*'))
        except:
            num_rows, num_cols = 2, 3
    else:
        num_rows, num_cols = 2, 3
    
    # Get slide count
    slide_count = _get_slide_count(parent)
    
    # Initialize file paths dictionary if not exists
    if not hasattr(parent, 'dragdrop_file_paths'):
        parent.dragdrop_file_paths = {}
    
    # Update tabs to match slide count
    page_tabs = parent.dragdrop_page_tabs
    tab_grids = parent.dragdrop_tab_grids
    tab_cells = parent.dragdrop_tab_cells
    
    # CRITICAL: Save all existing file paths BEFORE making any changes
    # This ensures file paths are preserved even when tabs are deleted or grids are updated
    current_tab_count = page_tabs.count()
    for slide_idx in range(current_tab_count):
        if slide_idx < len(tab_grids):
            if slide_idx not in parent.dragdrop_file_paths:
                parent.dragdrop_file_paths[slide_idx] = {}
            
            # Get current arrangement to calculate cell_index correctly
            grid_layout = tab_grids[slide_idx]
            
            # Iterate through grid positions to find cells and their file paths
            for row in range(num_rows):
                for col in range(num_cols):
                    cell_index = row * num_cols + col
                    item = grid_layout.itemAtPosition(row, col)
                    if item and item.widget():
                        cell = item.widget()
                        # Try to get file_path from cell
                        file_path = None
                        if hasattr(cell, 'file_path') and cell.file_path:
                            file_path = cell.file_path
                        elif cell.text() and cell.text() != "Empty":
                            # If cell has non-empty text, it might be a file path
                            # Try to get full path from stored attribute
                            if hasattr(cell, 'file_path'):
                                file_path = cell.file_path
                        
                        # Save file path if we found one
                        if file_path:
                            parent.dragdrop_file_paths[slide_idx][cell_index] = file_path
    
    # Add or remove tabs as needed
    current_tab_count = page_tabs.count()
    
    if slide_count > current_tab_count:
        # Add new tabs - preserve existing file paths
        for slide_idx in range(current_tab_count, slide_count):
            _create_dragdrop_tab(page_tabs, tab_grids, tab_cells, parent, slide_idx)
    elif slide_count < current_tab_count:
        # Remove excess tabs - only remove file paths for deleted slides
        while page_tabs.count() > slide_count:
            # Remove last tab
            last_idx = page_tabs.count() - 1
            
            # Clear cells
            if last_idx < len(tab_cells):
                for cell in tab_cells[last_idx]:
                    if cell:
                        cell.deleteLater()
                tab_cells[last_idx].clear()
            # Remove from lists
            if last_idx < len(tab_grids):
                tab_grids.pop(last_idx)
            if last_idx < len(tab_cells):
                tab_cells.pop(last_idx)
            # Remove file paths for this slide (only the deleted slide)
            if last_idx in parent.dragdrop_file_paths:
                del parent.dragdrop_file_paths[last_idx]
            # Remove tab
            page_tabs.removeTab(last_idx)
    
    # Update each tab's grid - this will restore file paths from dragdrop_file_paths
    for slide_idx in range(slide_count):
        if slide_idx < len(tab_grids):
            _update_dragdrop_tab_grid(parent, slide_idx, tab_grids[slide_idx], tab_cells[slide_idx], num_rows, num_cols)


def _get_slide_count(parent):
    """Get slide count from combo box or input"""
    if not hasattr(parent, 'dragdrop_slide_count_combo'):
        return 1
    
    slide_count_text = parent.dragdrop_slide_count_combo.currentText()
    
    if slide_count_text == "Custom":
        if hasattr(parent, 'dragdrop_slide_count_input'):
            return parent.dragdrop_slide_count_input.value()
        return 1
    else:
        try:
            return int(slide_count_text)
        except:
            return 1


def _update_dragdrop_tab_grid(parent, slide_idx, tab_grid_layout, tab_cells_list, num_rows, num_cols):
    """Update a single tab's grid layout"""
    # Clear existing cells
    for cell in tab_cells_list:
        if cell:
            cell.deleteLater()
    tab_cells_list.clear()
    
    # Clear grid layout
    while tab_grid_layout.count():
        item = tab_grid_layout.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
    
    # Initialize file paths for this slide if not exists
    if not hasattr(parent, 'dragdrop_file_paths'):
        parent.dragdrop_file_paths = {}
    if slide_idx not in parent.dragdrop_file_paths:
        parent.dragdrop_file_paths[slide_idx] = {}
    
    # Create new cells based on arrangement
    for row in range(num_rows):
        for col in range(num_cols):
            cell_index = row * num_cols + col
            cell_text = "Empty"
            
            # Check if there's a saved file path for this cell in this slide
            file_path = None
            if (hasattr(parent, 'dragdrop_file_paths') and 
                slide_idx in parent.dragdrop_file_paths and 
                cell_index in parent.dragdrop_file_paths[slide_idx]):
                file_path = parent.dragdrop_file_paths[slide_idx][cell_index]
                if file_path and os.path.exists(file_path):
                    cell_text = os.path.basename(file_path) if len(file_path) < 60 else file_path[:57] + "..."
                elif file_path:
                    # File path exists but file doesn't exist - still show it
                    cell_text = os.path.basename(file_path) if len(file_path) < 60 else file_path[:57] + "..."
            
            cell = DraggableLabel(cell_text, parent, cell_list=tab_cells_list)
            cell.setAlignment(Qt.AlignCenter)
            cell.setMinimumSize(100, 60)
            
            # Store file path if exists
            if file_path:
                cell.file_path = file_path
            cell.slide_idx = slide_idx
            cell.cell_index = cell_index
            
            # Update style based on content
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
            
            tab_grid_layout.addWidget(cell, row, col)
            tab_cells_list.append(cell)


def run_dragdrop_generation(thread_instance):
    """
    Run Drag and Drop mode slide generation
    Called from GeneratorThread.run()
    """
    # Check for interruption at the start
    if thread_instance.isInterruptionRequested():
        return False, "Generation cancelled"
    
    try:
        from Config.KeynoteCtrl import prepare_keynote
        from Config.KeynoteCtrl.keynoteCtrl_dragdrop import insert_pdfs_into_slide_dragdrop
        
        thread_instance.log_signal.emit("=" * 60)
        thread_instance.log_signal.emit("Mode: DRAG AND DROP")
        thread_instance.log_signal.emit("=" * 60)
        
        output_file = thread_instance.output_file
        file_paths_dict = thread_instance.dragdrop_file_paths  # Dict: {slide_idx: [file_paths]}
        arrangement = thread_instance.dragdrop_arrangement
        slide_count = thread_instance.dragdrop_slide_count
        
        thread_instance.log_signal.emit(f"Output file: {output_file}")
        thread_instance.log_signal.emit(f"Arrangement: {arrangement}")
        thread_instance.log_signal.emit(f"Number of slides: {slide_count}")
        for slide_idx in range(slide_count):
            slide_files = file_paths_dict.get(slide_idx, [])
            file_count = len([p for p in slide_files if p])
            thread_instance.log_signal.emit(f"  Slide {slide_idx + 1}: {file_count} file(s)")
        thread_instance.log_signal.emit("-" * 60)
        
        # Create Keynote file
        thread_instance.log_signal.emit("Creating Keynote file...")
        prepare_keynote(output_file, theme="White")
        thread_instance.log_signal.emit("Keynote file created successfully!")
        
        # Insert plots into slides
        thread_instance.log_signal.emit("Inserting plots into slides...")
        
        insert_pdfs_into_slide_dragdrop(
            output_file=output_file,
            file_paths_dict=file_paths_dict,
            arrangement=arrangement,
            slide_count=slide_count,
            log_callback=thread_instance.log_signal.emit,
            thread_instance=thread_instance,
            user_settings=thread_instance.user_settings
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
