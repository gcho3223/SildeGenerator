#!/usr/bin/env python3
"""
Common UI components and utilities for Keynote Slide Generator GUI
Shared between CPV and DRC modes
"""

from PyQt5.QtWidgets import QLabel
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDrag


class DraggableLabel(QLabel):
    """Custom QLabel that supports drag and drop for preview cells"""
    def __init__(self, text, parent=None, cell_list=None):
        super().__init__(text, parent)
        self.setAcceptDrops(True)
        self.parent_widget = parent
        # Store reference to the cell list this cell belongs to
        # This allows the drop event to find and swap with the correct source cell
        self.cell_list = cell_list
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.text() not in ["Empty", ""]:
            drag = QDrag(self)
            mime_data = QMimeData()
            mime_data.setText(self.text())
            drag.setMimeData(mime_data)
            drag.exec_(Qt.MoveAction)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.accept()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        if event.mimeData().hasText():
            # Swap the contents
            source_text = event.mimeData().text()
            target_text = self.text()
            
            # Update this cell
            self.setText(source_text)
            self._update_style(source_text)
            
            # Find and update the source cell
            # Try multiple cell lists to support different modes
            cell_lists_to_check = []
            
            # If cell_list is explicitly set, use it
            if self.cell_list:
                cell_lists_to_check.append(self.cell_list)
            
            # Also check common cell list attributes in parent_widget
            if self.parent_widget:
                # DRC mode
                if hasattr(self.parent_widget, 'energy_preview_cells'):
                    cell_lists_to_check.append(self.parent_widget.energy_preview_cells)
                
                # User-defined mode - main preview grid
                if hasattr(self.parent_widget, 'usrDefined_preview_cells'):
                    cell_lists_to_check.append(self.parent_widget.usrDefined_preview_cells)
                
                # User-defined mode - tab-based preview (list of lists)
                if hasattr(self.parent_widget, 'usrDefined_tab_cells'):
                    for tab_cells in self.parent_widget.usrDefined_tab_cells:
                        cell_lists_to_check.append(tab_cells)
            
            # Search for source cell in all candidate lists
            source_cell = None
            for cell_list in cell_lists_to_check:
                if cell_list:
                    for cell in cell_list:
                        if cell != self and cell.text() == source_text:
                            source_cell = cell
                            break
                    if source_cell:
                        break
            
            # Update source cell if found
            if source_cell:
                source_cell.setText(target_text)
                source_cell._update_style(target_text)
            
            event.accept()
    
    def _update_style(self, text):
        """Update cell style based on text content"""
        if text == "Empty":
            self.setStyleSheet(
                "border: 2px dashed #ccc; "
                "background-color: #f5f5f5; "
                "color: #999; "
                "border-radius: 5px; "
                "padding: 5px;"
            )
        elif "Not found" in text or "not_found" in text:
            self.setStyleSheet(
                "border: 2px dashed #ff6b6b; "
                "background-color: #ffe0e0; "
                "color: #cc0000; "
                "border-radius: 5px; "
                "padding: 5px;"
            )
        elif "Ambiguous" in text or "ambiguous" in text:
            self.setStyleSheet(
                "border: 2px solid #ffa500; "
                "background-color: #fff4e0; "
                "color: #cc6600; "
                "font-weight: bold; "
                "border-radius: 5px; "
                "padding: 5px;"
            )
        else:
            self.setStyleSheet(
                "border: 2px solid #0066cc; "
                "background-color: #e6f2ff; "
                "color: #0066cc; "
                "font-weight: bold; "
                "border-radius: 5px; "
                "padding: 5px;"
            )


def create_directory_ui(parent):
    """
    Create directory selection UI (used by both CPV and DRC modes)
    Returns: (dir_group, input_dir_edit, output_file_edit)
    """
    from PyQt5.QtWidgets import QGroupBox, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QWidget
    
    dir_group = QGroupBox("Directories")
    dir_layout = QVBoxLayout()
    
    # Input directory
    input_layout = QHBoxLayout()
    input_label = QLabel("Input Directory")
    input_label.setFixedWidth(90)
    input_layout.addWidget(input_label)
    input_dir_edit = QLineEdit(parent.input_dir)
    input_layout.addWidget(input_dir_edit, 1)
    input_browse = QPushButton("Browse")
    input_browse.clicked.connect(parent.browse_input_dir)
    input_layout.addWidget(input_browse)
    dir_layout.addLayout(input_layout)
    
    # Output file
    output_layout = QHBoxLayout()
    output_label = QLabel("Output File")
    output_label.setFixedWidth(90)
    output_layout.addWidget(output_label)
    output_file_edit = QLineEdit(parent.output_file)
    output_layout.addWidget(output_file_edit, 1)
    output_browse = QPushButton("Browse")
    output_browse.clicked.connect(parent.browse_output_file)
    output_layout.addWidget(output_browse)
    dir_layout.addLayout(output_layout)
    
    dir_group.setLayout(dir_layout)
    
    return dir_group, input_dir_edit, output_file_edit


def create_status_ui():
    """Create status log UI (used by both CPV and DRC modes)"""
    from PyQt5.QtWidgets import QGroupBox, QVBoxLayout, QTextEdit
    
    status_group = QGroupBox("Status")
    status_layout = QVBoxLayout()
    status_log = QTextEdit()
    status_log.setReadOnly(True)
    status_layout.addWidget(status_log)
    status_group.setLayout(status_layout)
    
    return status_group, status_log


def create_action_buttons(parent):
    """Create Generate, Stop, Scan (for User-defined mode), and Exit buttons"""
    from PyQt5.QtWidgets import QHBoxLayout, QPushButton
    from PyQt5.QtGui import QFont
    
    button_layout = QHBoxLayout()
    
    # Add stretch before buttons to center them
    button_layout.addStretch()
    
    # Scan button (for User-defined mode with scan mode enabled)
    scan_btn = QPushButton("SCAN")
    scan_btn.setMinimumSize(120, 40)
    scan_font = QFont()
    scan_font.setPointSize(14)
    scan_font.setBold(True)
    scan_btn.setFont(scan_font)
    scan_btn.setStyleSheet("color: blue;")
    scan_btn.setVisible(False)  # Hidden by default
    if hasattr(parent, 'scan_files'):
        scan_btn.clicked.connect(parent.scan_files)
    button_layout.addWidget(scan_btn)
    
    # Store scan button reference in parent
    parent.scan_button = scan_btn
    
    button_layout.addSpacing(20)
    
    # Stop button with custom style (initially disabled)
    stop_btn = QPushButton("STOP")
    stop_btn.setMinimumSize(120, 40)
    stop_font = QFont()
    stop_font.setPointSize(14)
    stop_font.setBold(True)
    stop_btn.setFont(stop_font)
    stop_btn.setStyleSheet("color: orange;")
    stop_btn.setEnabled(False)  # Initially disabled
    stop_btn.clicked.connect(parent.stop_generation)
    button_layout.addWidget(stop_btn)
    
    # Store stop button reference in parent
    parent.stop_button = stop_btn
    
    button_layout.addSpacing(20)
    
    # Generate button with custom style
    generate_btn = QPushButton("GENERATE")
    generate_btn.setMinimumSize(120, 40)
    font = QFont()
    font.setPointSize(14)
    font.setBold(True)
    generate_btn.setFont(font)
    generate_btn.setStyleSheet("color: red;")
    generate_btn.clicked.connect(parent.generate_slides)
    button_layout.addWidget(generate_btn)
    
    button_layout.addSpacing(20)
    
    # Exit button with custom style
    exit_btn = QPushButton("Exit")
    exit_btn.setMinimumSize(120, 40)
    exit_font = QFont()
    exit_font.setPointSize(14)
    exit_font.setBold(True)
    exit_btn.setFont(exit_font)
    exit_btn.clicked.connect(parent.close)
    button_layout.addWidget(exit_btn)
    
    # Add stretch after buttons to center them
    button_layout.addStretch()
    
    return button_layout
