"""
Settings Dialog Module
Provides configuration settings for the application
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QLineEdit, QGroupBox, QFormLayout,
                            QSpinBox, QCheckBox, QTabWidget, QWidget, QComboBox,
                            QGridLayout, QFrame, QMessageBox, QInputDialog)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
import os
import json
from pathlib import Path


class PositionEditDialog(QDialog):
    """Dialog for editing X, Y coordinates"""
    
    def __init__(self, plot_num, current_x, current_y, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Edit Plot {plot_num} Position")
        self.setModal(True)
        
        layout = QVBoxLayout()
        
        # Info label
        info_label = QLabel(f"Enter new coordinates for Plot {plot_num}:")
        info_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info_label)
        
        # Form layout for X and Y inputs
        form_layout = QFormLayout()
        
        # X coordinate input
        self.x_input = QSpinBox()
        self.x_input.setRange(0, 1024)
        self.x_input.setValue(current_x)
        self.x_input.setSuffix(" px")
        form_layout.addRow("X coordinate (0-1024):", self.x_input)
        
        # Y coordinate input
        self.y_input = QSpinBox()
        self.y_input.setRange(0, 768)
        self.y_input.setValue(current_y)
        self.y_input.setSuffix(" px")
        form_layout.addRow("Y coordinate (0-768):", self.y_input)
        
        layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setDefault(True)
        button_layout.addWidget(ok_button)
        
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def get_coordinates(self):
        """Return the entered coordinates"""
        return self.x_input.value(), self.y_input.value()


class ClickablePositionLabel(QLabel):
    """Clickable label for plot position preview"""
    clicked = pyqtSignal(int, int, int)  # plot_num, original_x, original_y
    
    def __init__(self, text, plot_num, original_x, original_y, parent=None):
        super().__init__(text, parent)
        self.plot_num = plot_num
        self.original_x = original_x
        self.original_y = original_y
        self.setCursor(Qt.PointingHandCursor)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.plot_num, self.original_x, self.original_y)
        super().mousePressEvent(event)


class SettingsDialog(QDialog):
    """Settings Dialog for application configuration"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(550)  # 600 → 550
        self.setMinimumHeight(550)  # 700 → 550
        
        # Initialize settings from parent if available
        self.parent_gui = parent
        
        # Determine current mode (CPV or DRC) from parent
        self.current_mode = None
        if parent and hasattr(parent, 'mode'):
            self.current_mode = parent.mode
        
        # Store custom positions (key: arrangement_position, value: list of (x, y) tuples)
        self.custom_positions = {}
        
        # Import config to get default sizes and positions
        try:
            from Config.config import size_, positions_
        except ImportError:
            # Try alternative import path
            import sys
            import os
            current_dir = os.path.dirname(os.path.abspath(__file__))
            parent_dir = os.path.dirname(os.path.dirname(current_dir))
            if parent_dir not in sys.path:
                sys.path.insert(0, parent_dir)
            from Config.config import size_, positions_
        
        # Store position option for each arrangement (key: arrangement, value: position)
        # Default: Center for all arrangements
        self.arrangement_positions = {
            "1*1": "Center",
            "1*2": "Center",
            "1*3": "Center",
            "1*4": "Center",
        }
        
        # Store size (width) for each arrangement (key: arrangement, value: width)
        # Get default values from config.py size_
        # size_[0] = (314, 305) # 1*3, 2*1
        # size_[1] = (500, 485) # 1*1, 1*2
        # size_[2] = (300, 291) # 2*2
        # size_[3] = (323, 313) # 2*3
        # size_[4] = (250, 243) # 1*4, 2*4
        self.arrangement_sizes = {
            "1*1": size_[1][0],  # 500
            "1*2": size_[1][0],  # 500
            "1*3": size_[0][0],  # 314
            "1*4": size_[4][0],  # 250
            "2*1": size_[2][0],  # 300
            "2*2": size_[2][0],  # 300
            "2*3": 300,  # DRC default: 300 (from config_drc.py size: (300, 291))
            "2*4": size_[4][0],  # 250
        }
        
        # Store default values for resetting to "None" preset
        self.default_arrangement_sizes = self.arrangement_sizes.copy()
        self.default_arrangement_positions = self.arrangement_positions.copy()
        self.default_custom_positions = {}
        
        self.setup_ui()
        self.load_current_settings()
        
        # After UI is set up, if DRC custom mode is enabled, fix arrangement to 2*3
        if hasattr(self, 'use_custom_drc') and self.use_custom_drc.isChecked():
            if hasattr(self, 'arrangement_combo'):
                self.arrangement_combo.setCurrentText("2*3")
                self.arrangement_combo.setEnabled(False)
    
    def setup_ui(self):
        """Setup the settings dialog UI"""
        layout = QVBoxLayout()
        
        # Create tab widget for different setting categories
        tab_widget = QTabWidget()
        
        # Plot Settings Tab
        plot_tab = self.create_plot_tab()
        tab_widget.addTab(plot_tab, "Plot")

        # CPV Settings Tab
        cpv_tab = self.create_cpv_tab()
        tab_widget.addTab(cpv_tab, "CPV")
        
        # DRC Settings Tab
        drc_tab = self.create_drc_tab()
        tab_widget.addTab(drc_tab, "DRC")
        
        layout.addWidget(tab_widget)
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_settings)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def create_cpv_tab(self):
        """Create CPV settings tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # CPV Systematic Settings
        sys_group = QGroupBox("Systematic Settings")
        sys_layout = QFormLayout()
        
        self.sys_suffix = QLineEdit("_Central_vs_Up_vs_Down")
        sys_layout.addRow("Systematic Suffix:", self.sys_suffix)
        
        self.auto_detect_samples = QCheckBox("Auto-detect sample directories")
        self.auto_detect_samples.setChecked(True)
        sys_layout.addRow("", self.auto_detect_samples)
        
        sys_group.setLayout(sys_layout)
        layout.addWidget(sys_group)
        
        # Plot Position Settings
        position_group = QGroupBox("Plot Settings")
        position_layout = QVBoxLayout()
        
        # Use customized size and positions checkbox
        self.use_custom_cpv = QCheckBox("Use customized size and positions")
        self.use_custom_cpv.setChecked(False)
        self.use_custom_cpv.stateChanged.connect(self.on_custom_cpv_changed)
        position_layout.addWidget(self.use_custom_cpv)
        
        # Default Position (combobox)
        default_pos_layout = QHBoxLayout()
        default_pos_label = QLabel("Default Position:")
        self.position_opt = QComboBox()
        self.position_opt.addItems(["Upper", "Center", "Bottom"])
        self.position_opt.setCurrentText("Center")
        default_pos_layout.addWidget(default_pos_label)
        default_pos_layout.addWidget(self.position_opt)
        default_pos_layout.addStretch()
        position_layout.addLayout(default_pos_layout)
        
        # Store label for enabling/disabling
        self.cpv_default_pos_label = default_pos_label
        
        position_group.setLayout(position_layout)
        layout.addWidget(position_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def create_drc_tab(self):
        """Create DRC settings tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Plot Position Settings
        position_group = QGroupBox("Plot Settings")
        position_layout = QVBoxLayout()
        
        # Use customized size and positions checkbox
        self.use_custom_drc = QCheckBox("Use customized size and positions")
        self.use_custom_drc.setChecked(False)
        self.use_custom_drc.stateChanged.connect(self.on_custom_drc_changed)
        position_layout.addWidget(self.use_custom_drc)
        
        # Default Position (combobox) - for Resolution/Linearity plots
        default_pos_layout = QHBoxLayout()
        default_pos_label = QLabel("Default Position:")
        self.position_opt_drc = QComboBox()
        self.position_opt_drc.addItems(["Upper", "Center", "Bottom"])
        self.position_opt_drc.setCurrentText("Center")
        default_pos_layout.addWidget(default_pos_label)
        default_pos_layout.addWidget(self.position_opt_drc)
        default_pos_layout.addStretch()
        position_layout.addLayout(default_pos_layout)
        
        # Store label for enabling/disabling
        self.drc_default_pos_label = default_pos_label
        
        position_group.setLayout(position_layout)
        layout.addWidget(position_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        return widget
    
    def create_plot_tab(self):
        """Create plot settings tab"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # Customized Plot Settings
        customized_group = QGroupBox("User-defined Plot Settings")
        customized_layout = QVBoxLayout()
        
        # Arrangement, Position, and Size in one row
        arrangement_position_size_row = QHBoxLayout()
        
        # Arrangement selection
        arrangement_label = QLabel("Arrangement:")
        self.arrangement_combo = QComboBox()
        # Allow all arrangements initially
        self.arrangement_combo.addItems(["1*1", "1*2", "1*3", "1*4", "2*1", "2*2", "2*3", "2*4"])
        self.arrangement_combo.setCurrentText("2*3")
        self.arrangement_combo.currentTextChanged.connect(self.on_arrangement_changed)
        
        arrangement_position_size_row.addWidget(arrangement_label)
        arrangement_position_size_row.addWidget(self.arrangement_combo)
        
        # Position option (only for 1*1, 1*2, 1*3, and 1*4) - 같은 줄에 배치
        self.position_option_label = QLabel("Position:")
        self.position_option_combo = QComboBox()
        self.position_option_combo.addItems(["Upper", "Center", "Bottom"])
        self.position_option_combo.setCurrentText("Center")
        # Connect to save position when changed
        self.position_option_combo.currentTextChanged.connect(self.on_position_option_changed)
        self.position_option_combo.currentTextChanged.connect(self.update_arrangement_preview)
        
        arrangement_position_size_row.addWidget(self.position_option_label)
        arrangement_position_size_row.addWidget(self.position_option_combo)
        
        # Size (width) input - 같은 줄에 배치
        size_label = QLabel("Size(width):")
        self.set_size_width = QSpinBox()
        self.set_size_width.setRange(100, 1000)
        # Default: 2*3 layout width (323 for CPV, 300 for DRC)
        default_width = 300 if self.current_mode == "drc" else 323
        self.set_size_width.setValue(default_width)
        self.set_size_width.valueChanged.connect(self.on_size_width_changed)
        
        arrangement_position_size_row.addWidget(size_label)
        arrangement_position_size_row.addWidget(self.set_size_width)
        arrangement_position_size_row.addStretch()
        
        customized_layout.addLayout(arrangement_position_size_row)
        
        # Note about height (proportional scaling)
        height_note = QLabel("Note: Height will be automatically set proportionally based on width.")
        height_note.setStyleSheet("color: #666; font-style: italic; padding-top: 5px;")
        height_note.setWordWrap(True)
        customized_layout.addWidget(height_note)
        
        # Initially hide position option
        self.position_option_label.setVisible(False)
        self.position_option_combo.setVisible(False)
        
        # Preview area
        preview_label = QLabel("Preview:")
        preview_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        customized_layout.addWidget(preview_label)
        
        # Preview frame - 검정 테두리가 슬라이드 역할
        self.preview_frame = QFrame()
        self.preview_frame.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.preview_frame.setLineWidth(2)
        self.preview_frame.setFixedSize(400, 300)  # 4:3 비율 (1024:768 = 400:300)
        # 슬라이드처럼 보이도록 스타일 설정
        self.preview_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #333;
                border-radius: 5px;
            }
        """)
        
        customized_layout.addWidget(self.preview_frame)
        
        customized_group.setLayout(customized_layout)
        layout.addWidget(customized_group)
        
        # Preset section for Plot settings
        preset_group = QGroupBox("Preset")
        preset_layout = QVBoxLayout()
        
        # Preset dropdown and buttons row
        preset_controls_row = QHBoxLayout()
        
        preset_label = QLabel("Preset:")
        self.plot_preset_combo = QComboBox()
        self.plot_preset_combo.setMinimumWidth(200)
        self.plot_preset_combo.currentTextChanged.connect(self.on_plot_preset_selected)
        
        # Load preset button
        load_preset_btn = QPushButton("Load")
        load_preset_btn.clicked.connect(self.load_plot_preset)
        
        # Save preset button
        save_preset_btn = QPushButton("Save")
        save_preset_btn.clicked.connect(self.save_plot_preset)
        
        # Delete preset button
        delete_preset_btn = QPushButton("Delete")
        delete_preset_btn.clicked.connect(self.delete_plot_preset)
        
        preset_controls_row.addWidget(preset_label)
        preset_controls_row.addWidget(self.plot_preset_combo)
        preset_controls_row.addWidget(load_preset_btn)
        preset_controls_row.addWidget(save_preset_btn)
        preset_controls_row.addWidget(delete_preset_btn)
        preset_controls_row.addStretch()
        
        preset_layout.addLayout(preset_controls_row)
        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)
        
        layout.addStretch()
        widget.setLayout(layout)
        
        # Initialize preview
        self.update_arrangement_preview()
        
        # Load preset list
        self._update_plot_preset_dropdown()
        
        return widget
    
    def on_custom_cpv_changed(self):
        """Handle CPV custom checkbox state change"""
        is_custom = self.use_custom_cpv.isChecked()
        
        # Disable/enable default position when custom is checked/unchecked
        self.position_opt.setEnabled(not is_custom)
        self.cpv_default_pos_label.setEnabled(not is_custom)
    
    def on_custom_drc_changed(self):
        """Handle DRC custom checkbox state change"""
        is_custom = self.use_custom_drc.isChecked()
        # Note: Default position is always enabled because it's used for Resolution/Linearity plots
        
        # If custom mode is enabled, fix arrangement to 2*3 (DRC uses fixed 2*3 layout)
        if hasattr(self, 'arrangement_combo'):
            if is_custom:
                # Save current arrangement before fixing
                if self.arrangement_combo.currentText() != "2*3":
                    # Temporarily store non-2*3 arrangement if needed
                    pass
                self.arrangement_combo.setCurrentText("2*3")
                self.arrangement_combo.setEnabled(False)
            else:
                # Re-enable arrangement selection when custom mode is disabled
                self.arrangement_combo.setEnabled(True)
    
    def on_arrangement_changed(self):
        """Handle arrangement change - show/hide position option and restore saved position/size"""
        arrangement = self.arrangement_combo.currentText()
        
        # Show position option for 1*1, 1*2, 1*3, and 1*4
        if arrangement in ["1*1", "1*2", "1*3", "1*4"]:
            self.position_option_label.setVisible(True)
            self.position_option_combo.setVisible(True)
            # Restore saved position for this arrangement
            saved_position = self.arrangement_positions.get(arrangement, "Center")
            # Temporarily disconnect to avoid triggering save
            try:
                self.position_option_combo.currentTextChanged.disconnect(self.on_position_option_changed)
            except:
                pass
            try:
                self.position_option_combo.currentTextChanged.disconnect(self.update_arrangement_preview)
            except:
                pass
            self.position_option_combo.setCurrentText(saved_position)
            self.position_option_combo.currentTextChanged.connect(self.on_position_option_changed)
            self.position_option_combo.currentTextChanged.connect(self.update_arrangement_preview)
        else:
            self.position_option_label.setVisible(False)
            self.position_option_combo.setVisible(False)
        
        # Restore saved size for this arrangement
        saved_size = self.arrangement_sizes.get(arrangement, 323)
        # Temporarily disconnect to avoid triggering save
        try:
            self.set_size_width.valueChanged.disconnect(self.on_size_width_changed)
        except:
            pass
        self.set_size_width.setValue(saved_size)
        self.set_size_width.valueChanged.connect(self.on_size_width_changed)
        
        # Update preview
        self.update_arrangement_preview()
    
    def on_size_width_changed(self, value):
        """Handle size width change - save size for current arrangement"""
        arrangement = self.arrangement_combo.currentText()
        self.arrangement_sizes[arrangement] = value
    
    def on_position_option_changed(self, position):
        """Handle position option change - save position for current arrangement"""
        arrangement = self.arrangement_combo.currentText()
        if arrangement in ["1*1", "1*2", "1*3", "1*4"]:
            # Save position for this arrangement
            self.arrangement_positions[arrangement] = position
    
    def update_arrangement_preview(self):
        """Update the arrangement preview based on selected arrangement"""
        # Clear previous preview labels - 더 확실하게 제거
        children = self.preview_frame.findChildren(QLabel)
        for child in children:
            child.setParent(None)  # 부모 관계 해제
            child.hide()  # 숨김
            child.deleteLater()  # 삭제 예약
        
        # 이벤트 처리를 통해 즉시 삭제 반영
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        
        arrangement = self.arrangement_combo.currentText()
        
        # Get position option for 1*1, 1*2, 1*3, and 1*4
        position = "Center"
        if arrangement in ["1*1", "1*2", "1*3", "1*4"]:
            position = self.position_option_combo.currentText()
        
        # Get frame dimensions
        frame_width = self.preview_frame.width()
        frame_height = self.preview_frame.height()
        
        if frame_width < 100:
            frame_width = 400  # 4:3 비율 유지
        if frame_height < 100:
            frame_height = 300  # 4:3 비율 유지
        
        # Calculate plot positions based on config.py positions_
        from Config.config import positions_
        
        # Map arrangement to positions_ index (매칭: config.py)
        # Note: positions_ array order is Upper, Center, Bottom
        position_map = {
            "1*1": {"Upper": 0, "Center": 1, "Bottom": 2},
            "1*2": {"Upper": 3, "Center": 4, "Bottom": 5},
            "1*3": {"Upper": 6, "Center": 7, "Bottom": 8},
            "1*4": {"Upper": 9, "Center": 10, "Bottom": 11},
            "2*1": 12,  # 2*1 : 12th index
            "2*2": 13,  # 2*2 : 13th index
            "2*3": 14,  # 2*3 : 14th index (CPV)
            "2*4": 16,  # 2*4 : 16th index
        }
        
        # Get positions for this arrangement
        custom_key = f"{arrangement}_{position}"
        if custom_key in self.custom_positions:
            # Use custom positions
            plot_positions = self.custom_positions[custom_key]
        elif arrangement == "1*1":
            pos_idx = position_map["1*1"][position]
            plot_positions = positions_[pos_idx]
        elif arrangement == "1*2":
            pos_idx = position_map["1*2"][position]
            plot_positions = positions_[pos_idx]
        elif arrangement == "1*3":
            pos_idx = position_map["1*3"][position]
            plot_positions = positions_[pos_idx]
        elif arrangement == "1*4":
            pos_idx = position_map["1*4"][position]
            plot_positions = positions_[pos_idx]
        elif arrangement in position_map and position_map[arrangement] is not None:
            plot_positions = positions_[position_map[arrangement]]
        else:
            # For 2*1, create vertical positions
            plot_positions = [[(frame_width // 2 - 35, 100)], [(frame_width // 2 - 35, 225)]]
        
        # Scale factor: 실제 Keynote 슬라이드(1024×768) → 프레임 크기
        scale_x = frame_width / 1024
        scale_y = frame_height / 768
        
        # Create plot position indicators
        plot_num = 1
        for row_positions in plot_positions:
            for x, y in row_positions:
                # Scale positions
                scaled_x = int(x * scale_x)
                scaled_y = int(y * scale_y)
                
                # Create clickable label with plot number and coordinates
                label_text = f"{plot_num}\n({x}, {y})"
                label = ClickablePositionLabel(label_text, plot_num, x, y, self.preview_frame)
                label.setAlignment(Qt.AlignCenter)
                label.setStyleSheet("""
                    QLabel {
                        background-color: rgba(74, 144, 226, 0.7);
                        border: 2px solid #4a90e2;
                        border-radius: 4px;
                        font-size: 10px;
                        font-weight: bold;
                        color: white;
                    }
                    QLabel:hover {
                        background-color: rgba(74, 144, 226, 0.9);
                        border: 2px solid #2a70c2;
                    }
                """)
                
                # Connect click signal
                label.clicked.connect(self.on_plot_position_clicked)
                
                # Plot indicator size (scaled) - 더 크게 표시
                label_width = int(180 * scale_x)  # 140 → 180
                label_height = int(110 * scale_y)  # 90 → 110
                label.setGeometry(scaled_x, scaled_y, label_width, label_height)
                label.show()
                
                plot_num += 1
    
    def on_plot_position_clicked(self, plot_num, current_x, current_y):
        """Handle plot position label click - allow editing coordinates"""
        # Get current arrangement and position
        arrangement = self.arrangement_combo.currentText()
        position = "Center"
        if arrangement in ["1*1", "1*2", "1*3", "1*4"]:
            position = self.position_option_combo.currentText()
        
        # Show position edit dialog
        dialog = PositionEditDialog(plot_num, current_x, current_y, self)
        
        if dialog.exec_() != QDialog.Accepted:
            return
        
        # Get new coordinates
        new_x, new_y = dialog.get_coordinates()
        
        # Store custom position
        custom_key = f"{arrangement}_{position}"
        
        # Get current positions
        from Config.config import positions_
        # Note: positions_ array order is Upper, Center, Bottom
        position_map = {
            "1*1": {"Upper": 0, "Center": 1, "Bottom": 2},
            "1*2": {"Upper": 3, "Center": 4, "Bottom": 5},
            "1*3": {"Upper": 6, "Center": 7, "Bottom": 8},
            "1*4": {"Upper": 9, "Center": 10, "Bottom": 11},
            "2*1": 12,  # 2*1 : 12th index
            "2*2": 13,  # 2*2 : 13th index
            "2*3": 14,  # 2*3 : 14th index (CPV)
            "2*4": 16,  # 2*4 : 16th index
        }
        
        # Initialize custom positions if not exists
        if custom_key not in self.custom_positions:
            # Copy from config
            if arrangement == "1*1":
                pos_idx = position_map["1*1"][position]
                self.custom_positions[custom_key] = [list(row) for row in positions_[pos_idx]]
            elif arrangement == "1*2":
                pos_idx = position_map["1*2"][position]
                self.custom_positions[custom_key] = [list(row) for row in positions_[pos_idx]]
            elif arrangement == "1*3":
                pos_idx = position_map["1*3"][position]
                self.custom_positions[custom_key] = [list(row) for row in positions_[pos_idx]]
            elif arrangement == "1*4":
                pos_idx = position_map["1*4"][position]
                self.custom_positions[custom_key] = [list(row) for row in positions_[pos_idx]]
            elif arrangement in position_map and position_map[arrangement] is not None:
                self.custom_positions[custom_key] = [list(row) for row in positions_[position_map[arrangement]]]
            else:
                self.custom_positions[custom_key] = [[(512, 100)], [(512, 300)]]
        
        # Update the specific position
        current_plot = 1
        for row_idx, row_positions in enumerate(self.custom_positions[custom_key]):
            for col_idx, (x, y) in enumerate(row_positions):
                if current_plot == plot_num:
                    self.custom_positions[custom_key][row_idx][col_idx] = (new_x, new_y)
                    # Refresh preview
                    self.update_arrangement_preview()
                    return
                current_plot += 1
    
    def load_current_settings(self):
        """Load current settings from parent GUI"""
        if not self.parent_gui:
            return
        
        # Check if parent has saved settings
        if not hasattr(self.parent_gui, 'user_settings') or not self.parent_gui.user_settings:
            return
        
        settings = self.parent_gui.user_settings
        
        # Load General settings
        # (No general settings to load currently)
        
        # Load CPV settings
        if 'cpv' in settings:
            cpv = settings['cpv']
            if 'sys_suffix' in cpv:
                self.sys_suffix.setText(cpv['sys_suffix'])
            if 'auto_detect_samples' in cpv:
                self.auto_detect_samples.setChecked(cpv['auto_detect_samples'])
            if 'use_custom' in cpv:
                self.use_custom_cpv.setChecked(cpv['use_custom'])
            if 'position_opt' in cpv:
                self.position_opt.setCurrentText(cpv['position_opt'])
        
        # Load DRC settings
        if 'drc' in settings:
            drc = settings['drc']
            if 'use_custom' in drc:
                self.use_custom_drc.setChecked(drc['use_custom'])
            if 'position_opt' in drc:
                self.position_opt_drc.setCurrentText(drc['position_opt'])
        
        # Load Plot settings
        if 'plot' in settings:
            plot = settings['plot']
            if 'arrangement' in plot:
                self.arrangement_combo.setCurrentText(plot['arrangement'])
            if 'custom_positions' in plot:
                self.custom_positions = plot['custom_positions'].copy()
            # Load arrangement-specific positions
            if 'arrangement_positions' in plot:
                self.arrangement_positions.update(plot['arrangement_positions'])
                # Set current arrangement's position
                current_arrangement = self.arrangement_combo.currentText()
                if current_arrangement in self.arrangement_positions:
                    saved_position = self.arrangement_positions[current_arrangement]
                    self.position_option_combo.setCurrentText(saved_position)
            elif 'position' in plot:
                # Legacy: single position for all arrangements
                # Apply to all 1*1, 1*2, 1*3, 1*4
                for arr in ["1*1", "1*2", "1*3", "1*4"]:
                    self.arrangement_positions[arr] = plot['position']
                self.position_option_combo.setCurrentText(plot['position'])
            # Load arrangement-specific sizes
            if 'arrangement_sizes' in plot:
                self.arrangement_sizes.update(plot['arrangement_sizes'])
                # Set current arrangement's size
                current_arrangement = self.arrangement_combo.currentText()
                if current_arrangement in self.arrangement_sizes:
                    saved_size = self.arrangement_sizes[current_arrangement]
                    self.set_size_width.setValue(saved_size)
            elif 'width' in plot:
                # Legacy: single width for all arrangements
                # Apply to all arrangements
                for arr in ["1*1", "1*2", "1*3", "1*4", "2*1", "2*2", "2*3", "2*4"]:
                    self.arrangement_sizes[arr] = plot['width']
                self.set_size_width.setValue(plot['width'])
            
            # Load selected preset and apply it
            if 'selected_preset' in plot and hasattr(self, 'plot_preset_combo'):
                saved_preset = plot['selected_preset']
                # Temporarily disconnect signal to avoid triggering during preset selection
                try:
                    self.plot_preset_combo.currentTextChanged.disconnect(self.on_plot_preset_selected)
                except:
                    pass
                
                # Set the preset in dropdown
                index = self.plot_preset_combo.findText(saved_preset)
                if index >= 0:
                    self.plot_preset_combo.setCurrentIndex(index)
                    # Apply the preset settings
                    if saved_preset == "Default":
                        self._restore_default_plot_settings()
                    else:
                        self._load_plot_preset_internal(saved_preset, show_message=False)
                
                # Reconnect signal
                self.plot_preset_combo.currentTextChanged.connect(self.on_plot_preset_selected)
        
        # Load User-defined settings
        if 'loopDefined' in settings:
            loopDefined = settings['loopDefined']
            # Templates are now in mode screen, not here
        
        # Update preview with loaded settings (only if preset wasn't loaded)
        if 'plot' not in settings or 'selected_preset' not in settings.get('plot', {}):
            self.update_arrangement_preview()
    
    def save_settings(self):
        """Save settings and apply them"""
        # Here you would save the settings to a config file or apply them
        # For now, we'll just accept the dialog
        
        # You can add logic here to:
        # 1. Save to a config file (JSON, INI, etc.)
        # 2. Apply settings to the parent GUI
        # 3. Update global configuration
        
        self.accept()
    
    def get_settings(self):
        """Return dictionary of all settings"""
        settings = {
            'general': {
                # General settings will be added here in the future
            },
            'plot': {
                'arrangement': self.arrangement_combo.currentText(),
                'custom_positions': self.custom_positions,  # Add custom positions
            },
            'cpv': {
                'sys_suffix': self.sys_suffix.text(),
                'auto_detect_samples': self.auto_detect_samples.isChecked(),
                'use_custom': self.use_custom_cpv.isChecked(),
                'position_opt': self.position_opt.currentText(),
            },
            'drc': {
                'use_custom': self.use_custom_drc.isChecked(),
                'position_opt': self.position_opt_drc.currentText(),
            },
            'loopDefined': {
                # Templates are now in mode screen, not saved here
            }
        }
        
        # Save arrangement-specific positions and sizes
        settings['plot']['arrangement_positions'] = self.arrangement_positions.copy()
        settings['plot']['arrangement_sizes'] = self.arrangement_sizes.copy()
        
        # Save selected preset name
        if hasattr(self, 'plot_preset_combo'):
            settings['plot']['selected_preset'] = self.plot_preset_combo.currentText()
        
        return settings
    
    def _get_plot_preset_dir(self):
        """Get directory for plot presets"""
        from GUI.preset_manager import get_app_support_path
        preset_dir = get_app_support_path()
        plot_preset_dir = preset_dir / "plot_presets"
        plot_preset_dir.mkdir(parents=True, exist_ok=True)
        return plot_preset_dir
    
    def _update_plot_preset_dropdown(self):
        """Update the plot preset dropdown with available presets"""
        if not hasattr(self, 'plot_preset_combo'):
            return
        
        # Save current selection
        current_text = self.plot_preset_combo.currentText()
        
        # Temporarily disconnect signal to avoid triggering during update
        try:
            self.plot_preset_combo.currentTextChanged.disconnect(self.on_plot_preset_selected)
        except:
            pass
        
        # Clear and add "Default"
        self.plot_preset_combo.clear()
        self.plot_preset_combo.addItem("Default")
        
        # Load presets
        preset_dir = self._get_plot_preset_dir()
        presets = []
        preset_names_set = set()  # Track unique preset names
        
        if preset_dir.exists():
            for file_path in preset_dir.glob("*.json"):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        if 'name' in data:
                            preset_name = data['name']
                            # Only add if not already in set (avoid duplicates)
                            if preset_name not in preset_names_set:
                                presets.append(preset_name)
                                preset_names_set.add(preset_name)
                except Exception:
                    continue
        
        # Sort and add to combo
        presets.sort(key=str.lower)
        for preset_name in presets:
            self.plot_preset_combo.addItem(preset_name)
        
        # Reconnect signal
        self.plot_preset_combo.currentTextChanged.connect(self.on_plot_preset_selected)
        
        # Restore selection if still available
        index = self.plot_preset_combo.findText(current_text)
        if index >= 0:
            self.plot_preset_combo.setCurrentIndex(index)
        else:
            self.plot_preset_combo.setCurrentIndex(0)  # "Default"
    
    def on_plot_preset_selected(self, preset_name):
        """Handle plot preset selection change - auto-load preset when selected"""
        # If "Default" is selected, restore default values
        if preset_name == "Default":
            self._restore_default_plot_settings()
        else:
            # Auto-load the selected preset
            self._load_plot_preset_internal(preset_name)
    
    def _restore_default_plot_settings(self):
        """Restore default plot settings (when preset is set to None)"""
        # Restore default arrangement sizes
        self.arrangement_sizes = self.default_arrangement_sizes.copy()
        
        # Restore default arrangement positions
        self.arrangement_positions = self.default_arrangement_positions.copy()
        
        # Clear custom positions
        self.custom_positions = {}
        
        # Update current arrangement's size
        current_arrangement = self.arrangement_combo.currentText()
        if current_arrangement in self.arrangement_sizes:
            saved_size = self.arrangement_sizes[current_arrangement]
            try:
                self.set_size_width.valueChanged.disconnect(self.on_size_width_changed)
            except:
                pass
            self.set_size_width.setValue(saved_size)
            self.set_size_width.valueChanged.connect(self.on_size_width_changed)
        
        # Update current arrangement's position
        if current_arrangement in self.arrangement_positions:
            saved_position = self.arrangement_positions[current_arrangement]
            try:
                self.position_option_combo.currentTextChanged.disconnect(self.on_position_option_changed)
                self.position_option_combo.currentTextChanged.disconnect(self.update_arrangement_preview)
            except:
                pass
            self.position_option_combo.setCurrentText(saved_position)
            self.position_option_combo.currentTextChanged.connect(self.on_position_option_changed)
            self.position_option_combo.currentTextChanged.connect(self.update_arrangement_preview)
        
        # Update preview to show default positions
        self.update_arrangement_preview()
    
    def _load_plot_preset_internal(self, preset_name, show_message=False):
        """Internal method to load a plot preset (used by both dropdown and Load button)"""
        if preset_name == "Default":
            if show_message:
                QMessageBox.information(self, "No Preset Selected", 
                                      "Please select a preset from the dropdown.")
            return
        
        preset_dir = self._get_plot_preset_dir()
        preset_file = preset_dir / f"{preset_name}.json"
        
        if not preset_file.exists():
            if show_message:
                QMessageBox.warning(self, "Preset Not Found", 
                                  f"Preset '{preset_name}' not found.")
            self._update_plot_preset_dropdown()
            return
        
        try:
            with open(preset_file, 'r', encoding='utf-8') as f:
                preset_data = json.load(f)
            
            # Load arrangement sizes
            if 'arrangement_sizes' in preset_data:
                self.arrangement_sizes.update(preset_data['arrangement_sizes'])
                # Update current arrangement size
                current_arrangement = self.arrangement_combo.currentText()
                if current_arrangement in self.arrangement_sizes:
                    saved_size = self.arrangement_sizes[current_arrangement]
                    try:
                        self.set_size_width.valueChanged.disconnect(self.on_size_width_changed)
                    except:
                        pass
                    self.set_size_width.setValue(saved_size)
                    self.set_size_width.valueChanged.connect(self.on_size_width_changed)
            
            # Load arrangement positions
            if 'arrangement_positions' in preset_data:
                self.arrangement_positions.update(preset_data['arrangement_positions'])
                # Update current arrangement position
                current_arrangement = self.arrangement_combo.currentText()
                if current_arrangement in self.arrangement_positions:
                    saved_position = self.arrangement_positions[current_arrangement]
                    try:
                        self.position_option_combo.currentTextChanged.disconnect(self.on_position_option_changed)
                        self.position_option_combo.currentTextChanged.disconnect(self.update_arrangement_preview)
                    except:
                        pass
                    self.position_option_combo.setCurrentText(saved_position)
                    self.position_option_combo.currentTextChanged.connect(self.on_position_option_changed)
                    self.position_option_combo.currentTextChanged.connect(self.update_arrangement_preview)
            
            # Load custom positions
            if 'custom_positions' in preset_data:
                self.custom_positions = preset_data['custom_positions'].copy()
            
            # Update preview
            self.update_arrangement_preview()
            
            if show_message:
                QMessageBox.information(self, "Preset Loaded", 
                                      f"Preset '{preset_name}' loaded successfully.")
            
        except Exception as e:
            if show_message:
                QMessageBox.critical(self, "Error Loading Preset", 
                                   f"Failed to load preset:\n{str(e)}")
            else:
                # Silent error for auto-load, just log
                print(f"Error loading preset '{preset_name}': {str(e)}")
    
    def load_plot_preset(self):
        """Load a plot preset (called by Load button)"""
        preset_name = self.plot_preset_combo.currentText()
        self._load_plot_preset_internal(preset_name, show_message=True)
    
    def save_plot_preset(self):
        """Save current plot settings as a preset"""
        # Get preset name from user
        preset_name, ok = QInputDialog.getText(
            self, "Save Plot Preset", "Enter preset name:",
            text=""
        )
        
        if not ok or not preset_name.strip():
            return
        
        preset_name = preset_name.strip()
        
        # Create preset data
        preset_data = {
            'schema_version': '1.0',
            'name': preset_name,
            'arrangement_sizes': self.arrangement_sizes.copy(),
            'arrangement_positions': self.arrangement_positions.copy(),
            'custom_positions': self.custom_positions.copy()
        }
        
        # Save to file
        preset_dir = self._get_plot_preset_dir()
        preset_file = preset_dir / f"{preset_name}.json"
        
        try:
            with open(preset_file, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=2, ensure_ascii=False)
            
            # Update dropdown
            self._update_plot_preset_dropdown()
            
            # Select the newly saved preset
            index = self.plot_preset_combo.findText(preset_name)
            if index >= 0:
                self.plot_preset_combo.setCurrentIndex(index)
            
            QMessageBox.information(self, "Preset Saved", 
                                  f"Preset '{preset_name}' saved successfully.")
            
        except Exception as e:
            QMessageBox.critical(self, "Error Saving Preset", 
                               f"Failed to save preset:\n{str(e)}")
    
    def delete_plot_preset(self):
        """Delete a plot preset"""
        preset_name = self.plot_preset_combo.currentText()
        
        if preset_name == "Default":
            QMessageBox.information(self, "No Preset Selected", 
                                  "Please select a preset to delete.")
            return
        
        # Confirm deletion
        reply = QMessageBox.question(
            self, "Delete Preset", 
            f"Are you sure you want to delete preset '{preset_name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Delete file
        preset_dir = self._get_plot_preset_dir()
        preset_file = preset_dir / f"{preset_name}.json"
        
        try:
            if preset_file.exists():
                preset_file.unlink()
            
            # Update dropdown
            self._update_plot_preset_dropdown()
            
            QMessageBox.information(self, "Preset Deleted", 
                                  f"Preset '{preset_name}' deleted successfully.")
            
        except Exception as e:
            QMessageBox.critical(self, "Error Deleting Preset", 
                               f"Failed to delete preset:\n{str(e)}")