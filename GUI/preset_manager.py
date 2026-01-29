"""
Preset Manager for User-defined mode
Handles saving and loading of preset configurations
"""
import os
import json
import platform
from pathlib import Path
from PyQt5.QtWidgets import QFileDialog, QInputDialog, QMessageBox


def get_app_support_path():
    """Get macOS Application Support path or equivalent for other OS"""
    if platform.system() == "Darwin":  # macOS
        home = Path.home()
        app_support = home / "Library" / "Application Support"
        app_name = "SildeMaker"
        preset_dir = app_support / app_name / "presets"
    else:
        # For other OS, use user's home directory
        home = Path.home()
        app_name = "SildeMaker"
        preset_dir = home / f".{app_name}" / "presets"
    
    # Create directory if it doesn't exist
    preset_dir.mkdir(parents=True, exist_ok=True)
    return preset_dir


def get_preset_folder_path(user_settings=None):
    """Get preset folder path, prioritizing user-specified folder"""
    if user_settings and 'preset' in user_settings:
        custom_folder = user_settings['preset'].get('folder_path')
        if custom_folder and os.path.exists(custom_folder):
            return Path(custom_folder)
    
    # Default to Application Support
    return get_app_support_path()


def list_presets(user_settings=None):
    """List all available preset files"""
    preset_dir = get_preset_folder_path(user_settings)
    presets = []
    
    if not preset_dir.exists():
        return presets
    
    for file_path in preset_dir.glob("*.json"):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if 'name' in data:
                    presets.append({
                        'name': data['name'],
                        'file': file_path.name,
                        'path': file_path
                    })
        except Exception:
            continue
    
    # Sort by name
    presets.sort(key=lambda x: x['name'].lower())
    return presets


def load_preset(preset_name, user_settings=None):
    """Load a preset by name"""
    preset_dir = get_preset_folder_path(user_settings)
    preset_file = preset_dir / f"{preset_name}.json"
    
    if not preset_file.exists():
        return None
    
    try:
        with open(preset_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading preset: {e}")
        return None


def save_preset(preset_data, user_settings=None):
    """Save a preset"""
    preset_dir = get_preset_folder_path(user_settings)
    preset_name = preset_data.get('name', 'Untitled')
    preset_file = preset_dir / f"{preset_name}.json"
    
    # Ensure directory exists
    preset_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        with open(preset_file, 'w', encoding='utf-8') as f:
            json.dump(preset_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving preset: {e}")
        return False


def create_preset_data(gui_instance):
    """Create preset data from current GUI state"""
    preset = {
        'schema_version': '1.0',
        'name': '',  # Will be set by caller
        'file_format': '',
        'scan_mode': False,
        'path_template': '',
        'filename_template': '',
        'variables': {},
        'preview': {
            'arrangement': '2*3',
            'row_values': [],
            'column_values': [],
            'slide_axis': ''
        },
        'drc_plot': {
            'enabled': False,
            'slide_axis': '',
            'grid_axis': '',
            'arrangement': ''
        }
    }
    
    # File format
    if hasattr(gui_instance, 'loopDefined_file_format_combo'):
        preset['file_format'] = gui_instance.loopDefined_file_format_combo.currentText()
    
    # Scan mode
    if hasattr(gui_instance, 'loopDefined_scan_mode_check'):
        preset['scan_mode'] = gui_instance.loopDefined_scan_mode_check.isChecked()
    
    # Templates
    if hasattr(gui_instance, 'loopDefined_path_template_input'):
        preset['path_template'] = gui_instance.loopDefined_path_template_input.text().strip()
    if hasattr(gui_instance, 'loopDefined_filename_template_input'):
        preset['filename_template'] = gui_instance.loopDefined_filename_template_input.text().strip()
    
    # Variables
    if hasattr(gui_instance, 'loopDefined_variable_inputs'):
        for var_name, var_input in gui_instance.loopDefined_variable_inputs.items():
            var_text = var_input.text().strip()
            if var_text:
                preset['variables'][var_name] = var_text
    
    # Preview settings
    if hasattr(gui_instance, 'loopDefined_arrangement_combo'):
        preset['preview']['arrangement'] = gui_instance.loopDefined_arrangement_combo.currentText()
    
    if hasattr(gui_instance, 'loopDefined_column_input'):
        column_text = gui_instance.loopDefined_column_input.text().strip()
        if column_text:
            preset['preview']['column_values'] = [v.strip() for v in column_text.split(',') if v.strip()]
    
    if hasattr(gui_instance, 'loopDefined_row_inputs'):
        row_values = []
        for i in sorted(gui_instance.loopDefined_row_inputs.keys()):
            row_input = gui_instance.loopDefined_row_inputs[i]
            row_text = row_input.text().strip()
            if row_text:
                row_values.append([v.strip() for v in row_text.split(',') if v.strip()])
        preset['preview']['row_values'] = row_values
    
    # Slide axis (normal mode)
    if hasattr(gui_instance, 'loopDefined_slide_axis_combo_normal'):
        slide_axis = gui_instance.loopDefined_slide_axis_combo_normal.currentText()
        if slide_axis:
            preset['preview']['slide_axis'] = slide_axis
    
    # DRC plot settings
    if hasattr(gui_instance, 'loopDefined_drc_plot_check'):
        preset['drc_plot']['enabled'] = gui_instance.loopDefined_drc_plot_check.isChecked()
        
        if preset['drc_plot']['enabled']:
            if hasattr(gui_instance, 'loopDefined_slide_axis_combo'):
                preset['drc_plot']['slide_axis'] = gui_instance.loopDefined_slide_axis_combo.currentText()
            if hasattr(gui_instance, 'loopDefined_grid_axis_combo'):
                preset['drc_plot']['grid_axis'] = gui_instance.loopDefined_grid_axis_combo.currentText()
            if hasattr(gui_instance, 'loopDefined_arrangement_combo'):
                preset['drc_plot']['arrangement'] = gui_instance.loopDefined_arrangement_combo.currentText()
    
    return preset


def apply_preset(preset_data, gui_instance):
    """Apply preset data to GUI"""
    # File format
    if 'file_format' in preset_data and hasattr(gui_instance, 'loopDefined_file_format_combo'):
        index = gui_instance.loopDefined_file_format_combo.findText(preset_data['file_format'])
        if index >= 0:
            gui_instance.loopDefined_file_format_combo.setCurrentIndex(index)
    
    # Scan mode
    if 'scan_mode' in preset_data and hasattr(gui_instance, 'loopDefined_scan_mode_check'):
        gui_instance.loopDefined_scan_mode_check.setChecked(preset_data['scan_mode'])
    
    # Templates
    if 'path_template' in preset_data and hasattr(gui_instance, 'loopDefined_path_template_input'):
        gui_instance.loopDefined_path_template_input.setText(preset_data['path_template'])
    
    if 'filename_template' in preset_data and hasattr(gui_instance, 'loopDefined_filename_template_input'):
        gui_instance.loopDefined_filename_template_input.setText(preset_data['filename_template'])
    
    # Wait for template update to complete before setting variable values
    # Use QTimer to ensure variable inputs are created
    from PyQt5.QtCore import QTimer
    def apply_variables():
        if 'variables' in preset_data and hasattr(gui_instance, 'loopDefined_variable_inputs'):
            for var_name, var_value in preset_data['variables'].items():
                if var_name in gui_instance.loopDefined_variable_inputs:
                    gui_instance.loopDefined_variable_inputs[var_name].setText(var_value)
    QTimer.singleShot(100, apply_variables)  # Small delay to ensure template update completes
    
    # Preview settings
    if 'preview' in preset_data:
        preview = preset_data['preview']
        
        # Arrangement
        if 'arrangement' in preview and hasattr(gui_instance, 'loopDefined_arrangement_combo'):
            index = gui_instance.loopDefined_arrangement_combo.findText(preview['arrangement'])
            if index >= 0:
                gui_instance.loopDefined_arrangement_combo.setCurrentIndex(index)
        
        # Column values
        if 'column_values' in preview and hasattr(gui_instance, 'loopDefined_column_input'):
            column_text = ','.join(preview['column_values'])
            gui_instance.loopDefined_column_input.setText(column_text)
        
        # Row values
        if 'row_values' in preview and hasattr(gui_instance, 'loopDefined_row_inputs'):
            for i, row_vals in enumerate(preview['row_values']):
                if i in gui_instance.loopDefined_row_inputs:
                    row_text = ','.join(row_vals)
                    gui_instance.loopDefined_row_inputs[i].setText(row_text)
        
        # Slide axis (normal mode)
        if 'slide_axis' in preview and hasattr(gui_instance, 'loopDefined_slide_axis_combo_normal'):
            slide_axis = preview['slide_axis']
            if slide_axis:
                index = gui_instance.loopDefined_slide_axis_combo_normal.findText(slide_axis)
                if index >= 0:
                    gui_instance.loopDefined_slide_axis_combo_normal.setCurrentIndex(index)
    
    # DRC plot settings
    if 'drc_plot' in preset_data:
        drc_plot = preset_data['drc_plot']
        
        if 'enabled' in drc_plot and hasattr(gui_instance, 'loopDefined_drc_plot_check'):
            gui_instance.loopDefined_drc_plot_check.setChecked(drc_plot['enabled'])
        
        if drc_plot.get('enabled', False):
            if 'slide_axis' in drc_plot and hasattr(gui_instance, 'loopDefined_slide_axis_combo'):
                slide_axis = drc_plot['slide_axis']
                index = gui_instance.loopDefined_slide_axis_combo.findText(slide_axis)
                if index >= 0:
                    gui_instance.loopDefined_slide_axis_combo.setCurrentIndex(index)
            
            if 'grid_axis' in drc_plot and hasattr(gui_instance, 'loopDefined_grid_axis_combo'):
                grid_axis = drc_plot['grid_axis']
                index = gui_instance.loopDefined_grid_axis_combo.findText(grid_axis)
                if index >= 0:
                    gui_instance.loopDefined_grid_axis_combo.setCurrentIndex(index)
            
            if 'arrangement' in drc_plot and hasattr(gui_instance, 'loopDefined_arrangement_combo'):
                arrangement = drc_plot['arrangement']
                index = gui_instance.loopDefined_arrangement_combo.findText(arrangement)
                if index >= 0:
                    gui_instance.loopDefined_arrangement_combo.setCurrentIndex(index)
