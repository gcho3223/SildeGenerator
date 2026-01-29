#!/usr/bin/env python3
"""
Loop-defined Mode Keynote Control Functions
Template-based plot insertion for user-defined file structures
"""

import os
import subprocess
from Config.Rules.loopDefined_template_engine import create_template_parser
from Config.config import size_, positions_


def _find_file_from_scan(variables, thread_instance, log):
    """
    Find file path from scan results based on variable combination
    
    Args:
        variables: Dictionary of variable values
        thread_instance: Thread instance containing scan results
        log: Logging function
    
    Returns:
        File path if found, None otherwise
    """
    if not thread_instance or not hasattr(thread_instance, 'scan_matched_files'):
        return None
    
    scan_matched_files = thread_instance.scan_matched_files
    if not scan_matched_files:
        return None
    
    # Match variables to scan results
    # Compare variable values (excluding format)
    search_vars = {k: str(v) for k, v in variables.items() if k != "format"}
    
    for file_path, parsed_vars in scan_matched_files:
        # Compare parsed variables (excluding format)
        parsed_search = {k: str(v) for k, v in parsed_vars.items() if k != "format"}
        
        # Check if all search variables match
        match = True
        for key, value in search_vars.items():
            if key not in parsed_search or str(parsed_search[key]) != str(value):
                match = False
                break
        
        if match:
            return file_path
    
    return None


def insert_pdfs_into_slide_loopDefined(output_file, base_dir, path_template, filename_template,
                                     selected_objects, selected_steps, variable_combinations,
                                     file_format="pdf", log_callback=None, thread_instance=None,
                                     row_values=None, column_values=None, scan_mode=False,
                                     drc_plot=False, slide_axis=None, grid_axis=None, 
                                     arrangement=None, confirmed_scan_results=None, preview_cell_order=None):
    """
    Insert PDF plots into Keynote slides using user-defined templates
    
    Args:
        output_file (str): Path to output Keynote file
        base_dir (str): Base directory where plots are located
        path_template (str): Path template (e.g., "{sample}/{channel}/{object}/{step}")
        filename_template (str): Filename template (e.g., "h_{object}_{kinematic}_{step}.{format}")
        selected_objects (list): List of object groups (e.g., [["Object1"], ["Object2", "Object3"]])
        selected_steps (list): List of steps (e.g., ["step1", "step2"])
        variable_combinations (dict): Dict mapping variable names to lists of possible values
        file_format (str): File format ("pdf" or "png")
        log_callback (callable): Optional callback function for logging messages
        thread_instance: Thread instance for interruption checking
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    
    abs_output = os.path.abspath(output_file)
    
    # In scan mode, use filename_template only (path_template is not used)
    if scan_mode:
        # For scan mode, we need to extract variables from filename_template only
        from Config.Rules.loopDefined_template_engine import template_to_regex
        # template_to_regex returns (regex_pattern, template_variables, var_instance_to_group)
        # We only need the first two values, so use * to ignore the rest
        regex_pattern, template_variables, *_ = template_to_regex(filename_template, file_format)
        # Create a minimal parser for variable extraction
        # Use empty path_template since it's not used in scan mode
        parser = create_template_parser("", filename_template)
    else:
        # Create template parser (normal mode)
        parser = create_template_parser(path_template, filename_template)
    
    # Check which variables are in template
    has_object = "object" in parser.all_variables
    has_kinematic = "kinematic" in parser.all_variables
    has_step = "step" in parser.all_variables
    
    # Get kinematic values if needed
    kinematic_values = variable_combinations.get("kinematic", [])
    if not kinematic_values and has_kinematic:
        # If kinematic is in template but not provided, try to get default kinematics
        from Config.Rules.rules import get_kinematics
        # Try to get kinematics for first object (if available)
        if selected_objects and selected_objects[0]:
            first_obj = selected_objects[0][0] if isinstance(selected_objects[0], list) else selected_objects[0]
            kinematic_values = get_kinematics(first_obj)
            log(f"Using default kinematics for {first_obj}: {kinematic_values}")
    
    # Determine which variables should be mapped to row_values and column_values
    # If row_values and column_values are provided, try to map them to template variables
    row_var_name = None
    col_var_name = None
    if row_values and column_values:
        # Collect all unique values from row_values and column_values
        all_row_vals = []
        for row_vals in row_values:
            if row_vals:
                all_row_vals.extend([str(v).strip() for v in row_vals])
        all_col_vals = [str(v).strip() for v in column_values if v]
        
        # Remove duplicates while preserving order
        seen = set()
        all_row_vals = [v for v in all_row_vals if not (v in seen or seen.add(v))]
        seen = set()
        all_col_vals = [v for v in all_col_vals if not (v in seen or seen.add(v))]
        
        log(f"Row values from UI: {all_row_vals}")
        log(f"Column values from UI: {all_col_vals}")
        
        # Find variables that have multiple values in variable_combinations
        # Exclude object, step, kinematic, format
        multi_value_vars = []
        for var_name in parser.all_variables:
            if var_name not in ["object", "step", "kinematic", "format"]:
                var_vals = variable_combinations.get(var_name, [])
                if len(var_vals) > 1:
                    multi_value_vars.append(var_name)
        
        # Try to match row_values and column_values with variable_combinations by comparing values
        # First, try exact match by comparing value sets
        row_var_candidates = []
        col_var_candidates = []
        
        for var_name in multi_value_vars:
            var_vals = [str(v).strip() for v in variable_combinations.get(var_name, [])]
            # Check if row values match this variable
            if set(all_row_vals) == set(var_vals) or (len(all_row_vals) > 0 and all(v in var_vals for v in all_row_vals)):
                row_var_candidates.append(var_name)
            # Check if column values match this variable
            if set(all_col_vals) == set(var_vals) or (len(all_col_vals) > 0 and all(v in var_vals for v in all_col_vals)):
                col_var_candidates.append(var_name)
        
        # If we found matches, use them
        if row_var_candidates:
            row_var_name = row_var_candidates[0]  # Use first match
            log(f"Matched row_values to variable '{row_var_name}' by value comparison")
        
        if col_var_candidates:
            # If row_var_name is already set and is in col_var_candidates, skip it
            if row_var_name and row_var_name in col_var_candidates:
                col_var_candidates = [v for v in col_var_candidates if v != row_var_name]
            if col_var_candidates:
                col_var_name = col_var_candidates[0]  # Use first match
                log(f"Matched column_values to variable '{col_var_name}' by value comparison")
        
        # If object/kinematic are in template, prioritize them
        if has_object and has_kinematic:
            # Traditional case: object -> row, kinematic -> column
            # But only if not already matched
            if not row_var_name:
                row_var_name = "object"
            if not col_var_name:
                col_var_name = "kinematic"
        elif not row_var_name or not col_var_name:
            # If we didn't find matches, use multi-value variables in order
            if len(multi_value_vars) >= 2:
                if not row_var_name:
                    row_var_name = multi_value_vars[0]
                if not col_var_name:
                    # Skip row_var_name if it's already set
                    for var_name in multi_value_vars:
                        if var_name != row_var_name:
                            col_var_name = var_name
                            break
            elif len(multi_value_vars) == 1:
                if not row_var_name:
                    row_var_name = multi_value_vars[0]
                # Try to find another variable that could be column
                if not col_var_name:
                    for var_name in parser.all_variables:
                        if var_name not in ["object", "step", "kinematic", "format", row_var_name]:
                            col_var_name = var_name
                            break
            else:
                # No multi-value variables found, but row/column values provided
                if has_object and not row_var_name:
                    row_var_name = "object"
                if has_kinematic and not col_var_name:
                    col_var_name = "kinematic"
                if not row_var_name or not col_var_name:
                    # Use first two variables from template (excluding format)
                    template_vars = [v for v in parser.all_variables if v not in ["format"]]
                    if len(template_vars) >= 2:
                        if not row_var_name:
                            row_var_name = template_vars[0]
                        if not col_var_name:
                            for var_name in template_vars:
                                if var_name != row_var_name:
                                    col_var_name = var_name
                                    break
                    elif len(template_vars) == 1 and not row_var_name:
                        row_var_name = template_vars[0]
        
        if row_var_name and col_var_name:
            log(f"Final mapping: row_values -> '{row_var_name}', column_values -> '{col_var_name}'")
        elif row_var_name:
            log(f"Final mapping: row_values -> '{row_var_name}', column_values -> (not mapped)")
        elif col_var_name:
            log(f"Final mapping: row_values -> (not mapped), column_values -> '{col_var_name}'")
    
    # Open Keynote file once at the beginning
    open_script = f'''
    tell application "Keynote"
    activate
    set theDoc to open (POSIX file "{abs_output}")
    end tell
    '''
    try:
        subprocess.run(["osascript", "-e", open_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        log(f"Error opening Keynote file: {e.stderr.decode() if e.stderr else 'Unknown error'}")
        return
    
    # DRC plot mode: Use slide axis and grid axis to generate slides
    if drc_plot and slide_axis and grid_axis and confirmed_scan_results:
        log("=" * 60)
        log("DRC PLOT MODE: Generating slides based on slide axis and grid axis")
        log("=" * 60)
        log(f"Slide axis: {slide_axis}")
        log(f"Grid axis: {grid_axis}")
        log(f"Arrangement: {arrangement}")
        
        # Get scan results
        all_results = confirmed_scan_results.get('all_results', [])
        variable_combinations = confirmed_scan_results.get('variable_combinations', {})
        
        # Get slide axis values
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
        
        log(f"Slide axis values: {slide_axis_values}")
        log(f"Grid axis values: {grid_axis_values}")
        
        # Parse arrangement
        num_rows, num_cols = 2, 3  # Default
        if arrangement:
            try:
                num_rows, num_cols = map(int, arrangement.split('*'))
            except:
                log(f"Warning: Invalid arrangement '{arrangement}', using default 2*3")
        
        # Create matching dictionary: (slide_axis_value, grid_axis_value) -> (file_path, status)
        matching_dict = {}
        for file_path, parsed_vars, status in all_results:
            if slide_axis in parsed_vars and grid_axis in parsed_vars:
                key = (parsed_vars[slide_axis], parsed_vars[grid_axis])
                if key not in matching_dict:
                    matching_dict[key] = []
                matching_dict[key].append((file_path, parsed_vars, status))
        
        # Get custom settings from user_settings if available
        arrangement_sizes = None
        custom_positions = None
        arrangement_positions = None
        if thread_instance and hasattr(thread_instance, 'user_settings') and thread_instance.user_settings:
            plot_settings = thread_instance.user_settings.get('plot', {})
            arrangement_sizes = plot_settings.get('arrangement_sizes', {})
            custom_positions = plot_settings.get('custom_positions', {})
            arrangement_positions = plot_settings.get('arrangement_positions', {})
            if arrangement_sizes or custom_positions or arrangement_positions:
                log("Using customized size and positions from Settings - Plot tab")
        
        # Determine arrangement and positions
        slide_arrangement = arrangement if arrangement else f"{num_rows}*{num_cols}"
        
        # Get plot size from settings or use default
        if arrangement_sizes and slide_arrangement in arrangement_sizes:
            plot_width = arrangement_sizes[slide_arrangement]
            aspect_ratio = {
                "2*4": 243 / 250, "2*3": 313 / 323, "2*2": 291 / 300,
                "1*4": 243 / 250, "1*3": 305 / 314, "1*2": 485 / 500, "1*1": 485 / 500
            }.get(slide_arrangement, 313 / 323)
            plot_height = int(plot_width * aspect_ratio)
            plot_size = (plot_width, plot_height)
            log(f"Using custom size for {slide_arrangement}: {plot_size}")
        else:
            size_map = {
                "2*4": size_[4], "2*3": size_[3], "2*2": size_[2],
                "1*4": size_[4], "1*3": size_[0], "1*2": size_[1], "1*1": size_[1]
            }
            plot_size = size_map.get(slide_arrangement, size_[3])
        
        # Get plot positions from settings or use default
        position_found = False
        slide_plot_positions = None
        
        if slide_arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
            position_opt = arrangement_positions.get(slide_arrangement, "Center")
            if custom_positions:
                arrangement_key = f"{slide_arrangement}_{position_opt}"
                if arrangement_key in custom_positions:
                    slide_plot_positions = custom_positions[arrangement_key]
                    log(f"Using custom positions for {arrangement_key}")
                    position_found = True
        
        if not position_found and custom_positions:
            for pos_opt in ["Center", "Upper", "Bottom"]:
                arrangement_key = f"{slide_arrangement}_{pos_opt}"
                if arrangement_key in custom_positions:
                    slide_plot_positions = custom_positions[arrangement_key]
                    log(f"Using custom positions for {arrangement_key}")
                    position_found = True
                    break
        
        if not position_found:
            # Use default positions based on arrangement
            pos_idx_map = {
                "2*4": 16, "2*3": 14, "2*2": 13,
                "1*4": 10, "1*3": 7, "1*2": 4, "1*1": 1
            }
            pos_idx = pos_idx_map.get(slide_arrangement, 14)
            slide_plot_positions = positions_[pos_idx]
            log(f"Using default positions for {slide_arrangement}")
        
        # Create slides for each slide axis value
        for slide_idx, slide_value in enumerate(slide_axis_values):
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            log(f"Creating slide {slide_idx + 1}/{len(slide_axis_values)}: {slide_axis}={slide_value}")
            
            # Create title text for this slide
            title_text = f"{slide_axis}={slide_value}"
            
            # Create new slide
            slide_apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items
                    try
                        repeat while (count of text items) > 0
                            delete text item 1
                        end repeat
                    end try
                    
                    -- Add title text box
                    set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
                    set object text of titleBox to "{title_text}"
                    tell object text of titleBox
                        set font of characters 1 thru -1 to "Helvetica Neue"
                        set size of characters 1 thru -1 to 24
                    end tell
            '''
            
            # Process grid axis values in the same order as Preview
            # Use preview cell order if available (preserves drag-and-drop order)
            # DRC mode처럼 프리뷰의 실제 위치 순서대로 처리
            if preview_cell_order and slide_idx in preview_cell_order:
                # Use preview cell order for this slide
                # Format: List of (position_idx, grid_value) tuples
                preview_order = preview_cell_order[slide_idx]
                log(f"Using preview cell order for slide {slide_idx + 1}: {preview_order}")
                
                # Process each cell in preview order
                for position_idx, grid_value in preview_order:
                    if grid_value is None:
                        continue  # Skip Empty/Not found/Ambiguous cells
                    
                    # Calculate row and col from position index (same as DRC mode)
                    row = position_idx // num_cols
                    col = position_idx % num_cols
                    
                    # Find matching file
                    key = (slide_value, grid_value)
                    matches = matching_dict.get(key, [])
                    
                    # Get file path if matched
                    file_path = None
                    if len(matches) == 1:
                        file_path, parsed_vars, status = matches[0]
                        if status == "Matched" and file_path and os.path.exists(file_path):
                            # Insert plot at the position from preview
                            x, y = _get_position_for_plot(position_idx, slide_arrangement, slide_plot_positions, log)
                            if x is not None and y is not None:
                                log(f"Inserting {grid_axis}={grid_value} at grid position ({row}, {col}) -> ({x}, {y})")
                                slide_apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                            else:
                                log(f"Warning: Could not get position for grid cell ({row}, {col})")
                        else:
                            log(f"File not found or not matched: {grid_axis}={grid_value}")
                    elif len(matches) > 1:
                        log(f"Warning: Ambiguous match for {grid_axis}={grid_value}, skipping")
                    else:
                        log(f"Warning: No match found for {grid_axis}={grid_value}, skipping")
            else:
                # Use default grid axis values order (fallback)
                log(f"Using default grid axis values order: {grid_axis_values}")
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
                    
                    # Get file path if matched
                    file_path = None
                    if len(matches) == 1:
                        file_path, parsed_vars, status = matches[0]
                        if status == "Matched" and file_path and os.path.exists(file_path):
                            # Insert plot
                            x, y = _get_position_for_plot(cell_idx, slide_arrangement, slide_plot_positions, log)
                            if x is not None and y is not None:
                                log(f"Inserting {grid_axis}={grid_value} at grid position ({row}, {col}) -> ({x}, {y})")
                                slide_apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                            else:
                                log(f"Warning: Could not get position for grid cell ({row}, {col})")
                        else:
                            log(f"File not found or not matched: {grid_axis}={grid_value}")
                    elif len(matches) > 1:
                        log(f"Warning: Ambiguous match for {grid_axis}={grid_value}, skipping")
                    else:
                        log(f"Warning: No match found for {grid_axis}={grid_value}, skipping")
                    
                    cell_idx += 1
            
            # Close slide and execute
            slide_apple_script += '''
                end tell
            end tell
            end tell
            '''
            
            # Execute this slide's script
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            try:
                subprocess.run(["osascript", "-e", slide_apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
            except subprocess.CalledProcessError as e:
                log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                return
        
        log("=" * 60)
        log("✅ Inserted all plots into Keynote slides (DRC plot mode)")
        log("=" * 60)
        return
    
    # Process plots based on template variables
    # If row_values and column_values are provided, use them regardless of has_object
    if row_values and column_values and row_var_name and col_var_name:
        # Get step values from variable_combinations if selected_steps is empty
        if not selected_steps and has_step:
            step_values = variable_combinations.get("step", [])
            if step_values:
                selected_steps = step_values
                log(f"Using step values from variable_combinations: {selected_steps}")
            else:
                log("Warning: Step is in template but no step values provided")
                return
        elif not selected_steps and not has_step:
            # Neither step nor object in template - create a single dummy step for processing
            selected_steps = ["dummy"]
        
        # Get custom settings from user_settings if available
        arrangement_sizes = None
        custom_positions = None
        arrangement_positions = None
        if thread_instance and hasattr(thread_instance, 'user_settings') and thread_instance.user_settings:
            plot_settings = thread_instance.user_settings.get('plot', {})
            arrangement_sizes = plot_settings.get('arrangement_sizes', {})
            custom_positions = plot_settings.get('custom_positions', {})
            arrangement_positions = plot_settings.get('arrangement_positions', {})
            if arrangement_sizes or custom_positions or arrangement_positions:
                log("Using customized size and positions from Settings - Plot tab")
        
        # Determine number of slides: based on max row length
        max_row_length = max([len(row_vals) for row_vals in row_values if row_vals]) if row_values else 0
        if max_row_length == 0:
            log("Warning: No row values provided, skipping")
            return
        
        # Collect all unique row variable values from row_values and variable_combinations
        all_available_row_values = []
        
        # First, collect from row_values
        for row_vals in row_values:
            if row_vals:
                all_available_row_values.extend(row_vals)
        
        # Also collect from variable_combinations
        # If row_var_name is object, also check obj_group
        if row_var_name == "object" and has_object:
            # obj_group will be handled later in the object loop
            pass
        if row_var_name in variable_combinations:
            for val in variable_combinations[row_var_name]:
                if val not in all_available_row_values:
                    all_available_row_values.append(val)
        
        # Remove duplicates while preserving order
        seen = set()
        all_available_row_values = [x for x in all_available_row_values if not (x in seen or seen.add(x))]
        
        if not all_available_row_values:
            log(f"Warning: No values found for row variable '{row_var_name}' in row_values or variable_combinations")
            return
        
        log(f"Available values for row variable '{row_var_name}': {all_available_row_values}")
        
        # Check if slide_axis is provided (for normal mode, DRC plot OFF)
        # slide_axis is used to create multiple slides (one per slide_axis value)
        slide_axis_values = []
        if slide_axis and not drc_plot:
            # Get slide axis values from variable_combinations first
            if slide_axis in variable_combinations:
                slide_axis_values = variable_combinations[slide_axis]
                log(f"Using slide axis '{slide_axis}' values from variable_combinations: {slide_axis_values}")
            else:
                # If not in variable_combinations, try to extract from scan results
                if scan_mode and confirmed_scan_results:
                    all_results = confirmed_scan_results.get('all_results', [])
                    slide_values_set = set()
                    for file_path, parsed_vars, status in all_results:
                        if slide_axis in parsed_vars:
                            slide_values_set.add(parsed_vars[slide_axis])
                    slide_axis_values = sorted(list(slide_values_set))
                    log(f"Extracted slide axis '{slide_axis}' values from scan results: {slide_axis_values}")
                else:
                    log(f"Warning: slide axis '{slide_axis}' not found in variable_combinations or scan results")
                    slide_axis_values = []
        
        # Process each step first, then for each step create slides
        for step_name in selected_steps:
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Determine slide generation strategy
            if slide_axis and slide_axis_values and not drc_plot:
                # When using slide axis, we need to create slides for each combination of:
                # - slide axis value (e.g., stepnum=3,4,5)
                # - row value pairs (e.g., (Lep1, Lep2), (Jet1, Jet2))
                
                # Get all row value pairs from row_values
                # row_values is like [['Lep1', 'Jet1'], ['Lep2', 'Jet2']]
                # We need to create pairs: (Lep1, Lep2), (Jet1, Jet2)
                row_value_pairs = []
                if row_values and len(row_values) >= 2:
                    # Get the maximum length of row values
                    max_row_vals_length = max([len(row_vals) for row_vals in row_values if row_vals]) if row_values else 0
                    
                    # For each index in row values, create a pair
                    for pair_idx in range(max_row_vals_length):
                        pair = []
                        for row_vals in row_values:
                            if row_vals and pair_idx < len(row_vals):
                                pair.append(row_vals[pair_idx])
                            else:
                                pair.append(None)
                        if any(pair):  # Only add if at least one value exists
                            row_value_pairs.append(pair)
                
                # Create slides: slide axis value × row value pairs
                slide_combinations = []
                for slide_axis_val in slide_axis_values:
                    for row_pair in row_value_pairs:
                        slide_combinations.append((slide_axis_val, row_pair))
                
                log(f"Creating {len(slide_combinations)} slides based on slide axis '{slide_axis}' and row value pairs")
                log(f"Slide combinations: {slide_combinations}")
            else:
                # Use row index to create slides (original behavior)
                slide_combinations = [(None, list(range(max_row_length)))]
                log(f"Creating {len(slide_combinations)} slides based on row index")
            
            # Process each slide
            for slide_idx, (slide_axis_val, row_pair_or_idx) in enumerate(slide_combinations):
                # Check for interruption
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    return
                
                # Determine row values based on slide generation strategy
                if slide_axis and slide_axis_values and not drc_plot:
                    # When using slide axis with row pairs, use the specific row pair
                    # row_pair is like ['Lep1', 'Lep2'] or ['Jet1', 'Jet2']
                    row_vals_at_idx = row_pair_or_idx
                    slide_axis_val_actual = slide_axis_val
                    log(f"Creating slide for {slide_axis}={slide_axis_val_actual} with row values: {row_vals_at_idx}")
                else:
                    # When using row index, use the slide_value as row index
                    row_val_idx = row_pair_or_idx[0] if isinstance(row_pair_or_idx, list) else row_pair_or_idx
                    slide_axis_val_actual = None
                    
                    # Get row values for this row index
                    row_vals_at_idx = []
                    for row_vals in row_values:
                        if row_vals and row_val_idx < len(row_vals):
                            row_vals_at_idx.append(row_vals[row_val_idx])
                        else:
                            row_vals_at_idx.append(None)
                
                if not any(row_vals_at_idx):
                    slide_num = slide_idx + 1
                    log(f"Warning: No row values, skipping slide {slide_num}")
                    continue
                
                # Determine arrangement: num_rows * num_columns (e.g., 2*3)
                # Use the number of non-None row values
                num_rows_in_slide = len([r for r in row_vals_at_idx if r])
                num_cols_in_slide = len(column_values)
                
                # Determine arrangement and positions (same logic as before)
                if num_rows_in_slide == 2 and num_cols_in_slide == 4:
                    slide_arrangement = "2*4"
                    slide_pos_idx = 16
                elif num_rows_in_slide == 2 and num_cols_in_slide == 3:
                    slide_arrangement = "2*3"
                    slide_pos_idx = 14
                elif num_rows_in_slide == 2 and num_cols_in_slide == 2:
                    slide_arrangement = "2*2"
                    slide_pos_idx = 13
                elif num_rows_in_slide == 1:
                    if num_cols_in_slide == 4:
                        slide_arrangement = "1*4"
                        slide_pos_idx = 10
                    elif num_cols_in_slide == 3:
                        slide_arrangement = "1*3"
                        slide_pos_idx = 7
                    elif num_cols_in_slide == 2:
                        slide_arrangement = "1*2"
                        slide_pos_idx = 4
                    else:
                        slide_arrangement = "1*1"
                        slide_pos_idx = 1
                else:
                    slide_arrangement = f"{num_rows_in_slide}*{num_cols_in_slide}"
                    slide_pos_idx = 14
                
                # Get plot size from settings or use default (same logic as before)
                if arrangement_sizes and slide_arrangement in arrangement_sizes:
                    plot_width = arrangement_sizes[slide_arrangement]
                    aspect_ratio = {
                        "2*4": 243 / 250, "2*3": 313 / 323, "2*2": 291 / 300,
                        "1*4": 243 / 250, "1*3": 305 / 314, "1*2": 485 / 500, "1*1": 485 / 500
                    }.get(slide_arrangement, 291 / 300)
                    plot_height = int(plot_width * aspect_ratio)
                    plot_size = (plot_width, plot_height)
                    log(f"Using custom size for {slide_arrangement}: {plot_size}")
                else:
                    size_map = {
                        "2*4": size_[4], "2*3": size_[3], "2*2": size_[2],
                        "1*4": size_[4], "1*3": size_[0], "1*2": size_[1], "1*1": size_[1]
                    }
                    plot_size = size_map.get(slide_arrangement, size_[2])
                
                # Get plot positions from settings or use default (same logic as before)
                position_found = False
                slide_plot_positions = None
                
                if slide_arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
                    position_opt = arrangement_positions.get(slide_arrangement, "Center")
                    if custom_positions:
                        arrangement_key = f"{slide_arrangement}_{position_opt}"
                        if arrangement_key in custom_positions:
                            slide_plot_positions = custom_positions[arrangement_key]
                            log(f"Using custom positions for {arrangement_key}")
                            position_found = True
                
                if not position_found and custom_positions:
                    for pos_opt in ["Center", "Upper", "Bottom"]:
                        arrangement_key = f"{slide_arrangement}_{pos_opt}"
                        if arrangement_key in custom_positions:
                            slide_plot_positions = custom_positions[arrangement_key]
                            log(f"Using custom positions for {arrangement_key}")
                            position_found = True
                            break
                
                if not position_found:
                    if slide_arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
                        position_opt = arrangement_positions.get(slide_arrangement, "Center")
                        position_map = {
                            "1*1": {"Upper": 0, "Center": 1, "Bottom": 2},
                            "1*2": {"Upper": 3, "Center": 4, "Bottom": 5},
                            "1*3": {"Upper": 6, "Center": 7, "Bottom": 8},
                            "1*4": {"Upper": 9, "Center": 10, "Bottom": 11},
                        }
                        if slide_arrangement in position_map and position_opt in position_map[slide_arrangement]:
                            slide_pos_idx = position_map[slide_arrangement][position_opt]
                            slide_plot_positions = positions_[slide_pos_idx]
                            log(f"Using default positions for {slide_arrangement} with {position_opt} option")
                        else:
                            slide_plot_positions = positions_[slide_pos_idx]
                    else:
                        slide_plot_positions = positions_[slide_pos_idx]
                
                # Create title text for this slide
                row_val_str = ", ".join([str(r) for r in row_vals_at_idx if r])
                if slide_axis and slide_axis_val_actual is not None:
                    title_parts = []
                    if has_step:
                        title_parts.append(step_name)
                    if row_val_str:
                        title_parts.append(row_val_str)
                    title_parts.append(f"{slide_axis}={slide_axis_val_actual}")
                    title_text = " - ".join(title_parts)
                elif has_step:
                    title_text = f"{step_name} - {row_val_str}"
                else:
                    title_text = f"{row_val_str}"
                
                # Create new slide for this row index
                slide_apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items
                    try
                        repeat while (count of text items) > 0
                            delete text item 1
                        end repeat
                    end try
                    
                    -- Add title text box
                    set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
                    set object text of titleBox to "{title_text}"
                    tell object text of titleBox
                        set font of characters 1 thru -1 to "Helvetica Neue"
                        set size of characters 1 thru -1 to 24
                    end tell
            '''
                
                # Process each row and column combination
                slide_plot_idx = 0
                
                for slide_row_idx, row_val in enumerate(row_vals_at_idx):
                    if not row_val:
                        continue
                    
                    # Find matching row variable value
                    matching_row_val = None
                    for val in all_available_row_values:
                        if val == row_val or str(val) == str(row_val):
                            matching_row_val = val
                            break
                    
                    if not matching_row_val:
                        row_val_lower = str(row_val).lower().strip()
                        for val in all_available_row_values:
                            val_lower = str(val).lower().strip()
                            if val_lower == row_val_lower or row_val_lower in val_lower or val_lower in row_val_lower:
                                matching_row_val = val
                                break
                    
                    if not matching_row_val:
                        matching_row_val = row_val
                        log(f"Warning: No exact match for row value '{row_val}', using it directly")
                    
                    # Process each column for this row
                    for col_idx, col_val in enumerate(column_values):
                        # Find matching column variable value
                        matching_col_val = None
                        
                        if col_var_name == "kinematic" and kinematic_values:
                            col_val_lower = str(col_val).lower().strip()
                            for kin_name in kinematic_values:
                                kin_name_lower = str(kin_name).lower().strip()
                                if kin_name_lower == col_val_lower:
                                    matching_col_val = kin_name
                                    break
                                elif col_val_lower in kin_name_lower or kin_name_lower in col_val_lower:
                                    matching_col_val = kin_name
                                    break
                            
                            if not matching_col_val and col_idx < len(kinematic_values):
                                matching_col_val = kinematic_values[col_idx]
                        else:
                            if col_var_name in variable_combinations:
                                col_vals = variable_combinations[col_var_name]
                                col_val_lower = str(col_val).lower().strip()
                                for val in col_vals:
                                    val_lower = str(val).lower().strip()
                                    if val_lower == col_val_lower or col_val_lower in val_lower or val_lower in col_val_lower:
                                        matching_col_val = val
                                        break
                                
                                if not matching_col_val and col_idx < len(col_vals):
                                    matching_col_val = col_vals[col_idx]
                        
                        if not matching_col_val:
                            matching_col_val = col_val
                            log(f"No matching value found for column variable '{col_var_name}', using column value '{col_val}' directly")
                        
                        # Build variables dictionary
                        variables = {}
                        
                        # Add variables from variable_combinations (excluding row_var, col_var, slide_axis, object, step, kinematic, format)
                        excluded_vars = [row_var_name, col_var_name, slide_axis, "object", "step", "kinematic", "format"]
                        for var_name, var_values in variable_combinations.items():
                            if var_name not in excluded_vars:
                                if var_values:
                                    variables[var_name] = var_values[0]
                        
                        # Set row variable, column variable, slide axis variable, step, format
                        if row_var_name:
                            variables[row_var_name] = matching_row_val
                        if col_var_name:
                            variables[col_var_name] = matching_col_val
                        if slide_axis and slide_axis_val_actual is not None:
                            variables[slide_axis] = slide_axis_val_actual
                        
                        if has_step:
                            variables["step"] = step_name
                        variables["format"] = file_format
                        
                        # Generate file path
                        if scan_mode:
                            # In scan mode, find file from scan results
                            file_path = _find_file_from_scan(variables, thread_instance, log)
                        else:
                            # Normal mode: build path from template
                            file_path = parser.build_path(base_dir, variables)
                        
                        # Insert plot if exists
                        if file_path and os.path.exists(file_path):
                            grid_plot_idx = slide_row_idx * num_cols_in_slide + col_idx
                            x, y = _get_position_for_plot(grid_plot_idx, slide_arrangement, slide_plot_positions, log)
                            if x is not None and y is not None:
                                row_info = f"{row_var_name}={matching_row_val}" if row_var_name else ""
                                col_info = f"{col_var_name}={matching_col_val}" if col_var_name else ""
                                slide_axis_info = f"{slide_axis}={slide_axis_val_actual}" if slide_axis and slide_axis_val_actual is not None else ""
                                step_info = f"step={step_name}" if has_step else ""
                                info_parts = [p for p in [row_info, col_info, slide_axis_info, step_info] if p]
                                slide_num = slide_idx + 1
                                log(f"Inserting ({', '.join(info_parts)}) in slide {slide_num} at grid position ({slide_row_idx}, {col_idx}) -> ({x}, {y})")
                                slide_apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                                slide_plot_idx += 1
                        else:
                            log(f"File not found: {file_path}")
                
                # Close slide and execute
                slide_apple_script += '''
                end tell
            end tell
            end tell
            '''
                
                # Execute this slide's script
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    return
                
                script_lines = slide_apple_script.strip().split('\n')
                if len(script_lines) >= 3:
                    last_lines = [line.strip() for line in script_lines[-3:]]
                    if all(line == "end tell" for line in last_lines):
                        try:
                            subprocess.run(["osascript", "-e", slide_apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
                        except subprocess.CalledProcessError as e:
                            log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                            return
        
        # Return early since we processed with row/column values
        log("=" * 60)
        log("✅ Inserted all plots into Keynote slides")
        log("=" * 60)
        return
    
    # Process plots based on template variables
    # If object is not in template, process without object loop
    if not has_object:
        # Process without object - use kinematic or other variables directly
        # Get step values from variable_combinations if selected_steps is empty
        if not selected_steps and has_step:
            # Try to get step values from variable_combinations
            step_values = variable_combinations.get("step", [])
            if step_values:
                selected_steps = step_values
                log(f"Using step values from variable_combinations: {selected_steps}")
            else:
                # Step is in template but not provided in either place
                log("Warning: Step is in template but no step values provided")
                return
        elif not selected_steps and not has_step:
            # Neither step nor object in template - create a single dummy step for processing
            selected_steps = ["dummy"]
        
        # Determine layout based on kinematics only
        num_kinematics = len(kinematic_values) if kinematic_values else 1
        total_plots = num_kinematics
        
        if total_plots == 1:
            arrangement = "1*1"
            size_idx = 1
            pos_idx = 1
        elif total_plots == 2:
            arrangement = "1*2"
            size_idx = 1
            pos_idx = 4
        elif total_plots == 3:
            arrangement = "1*3"
            size_idx = 0
            pos_idx = 7
        elif total_plots == 4:
            arrangement = "1*4"
            size_idx = 4
            pos_idx = 10
        else:
            arrangement = "1*1"
            size_idx = 1
            pos_idx = 1
        
        plot_size = size_[size_idx]
        plot_positions = positions_[pos_idx]
        
        log(f"Processing without object variable (arrangement: {arrangement}, total plots: {total_plots})")
        
        # Process each step
        for step_name in selected_steps:
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Create title text for the slide
            if has_step:
                title_text = f"{step_name}"
            else:
                title_text = "Loop-defined plots"
            
            # Create new slide - each slide has complete AppleScript
            apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items
                    try
                        repeat while (count of text items) > 0
                            delete text item 1
                        end repeat
                    end try
                    
                    -- Add title text box
                    set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
                    set object text of titleBox to "{title_text}"
                    tell object text of titleBox
                        set font of characters 1 thru -1 to "Helvetica Neue"
                        set size of characters 1 thru -1 to 24
                    end tell
            '''
            
            # Process kinematic values directly (if kinematic is in template)
            plot_idx = 0
            if has_kinematic and kinematic_values:
                for kin_name in kinematic_values:
                    # Build variables dictionary
                    variables = {}
                    
                    # Add variables from variable_combinations (excluding object, step, kinematic, format)
                    for var_name, var_values in variable_combinations.items():
                        if var_name not in ["object", "step", "kinematic", "format"]:
                            if var_values:
                                variables[var_name] = var_values[0]
                    
                    # Set step, kinematic, format (no object)
                    if has_step:
                        variables["step"] = step_name
                    variables["kinematic"] = kin_name
                    variables["format"] = file_format
                    
                    # Generate file path
                    if scan_mode:
                        # In scan mode, find file from scan results
                        file_path = _find_file_from_scan(variables, thread_instance, log)
                    else:
                        # Normal mode: build path from template
                        file_path = parser.build_path(base_dir, variables)
                    
                    # Insert plot if exists
                    if file_path and os.path.exists(file_path):
                        x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                        if x is not None and y is not None:
                            log(f"Inserting ({step_name if has_step else ''}, {kin_name}) at ({x}, {y})")
                            apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                            plot_idx += 1
                    else:
                        log(f"File not found: {file_path}")
            else:
                # No kinematic, no object - process with just step and other variables
                variables = {}
                
                # Add variables from variable_combinations
                for var_name, var_values in variable_combinations.items():
                    if var_name not in ["object", "step", "format"]:
                        if var_values:
                            variables[var_name] = var_values[0]
                
                # Set step, format (no object, no kinematic)
                if has_step:
                    variables["step"] = step_name
                variables["format"] = file_format
                
                # Generate file path
                file_path = parser.build_path(base_dir, variables)
                
                # Insert plot if exists
                if os.path.exists(file_path):
                    x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                    if x is not None and y is not None:
                        log(f"Inserting ({step_name if has_step else ''}) at ({x}, {y})")
                        apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                        plot_idx += 1
                else:
                    log(f"File not found: {file_path}")
            
            # Close slide and execute immediately
            apple_script += '''
                end tell
            end tell
            end tell
            '''
            
            # Check for interruption before executing
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Only execute if AppleScript is complete (has proper closing)
            script_lines = apple_script.strip().split('\n')
            if len(script_lines) >= 3:
                last_lines = [line.strip() for line in script_lines[-3:]]
                if all(line == "end tell" for line in last_lines):
                    try:
                        subprocess.run(["osascript", "-e", apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
                    except subprocess.CalledProcessError as e:
                        log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                        return
                else:
                    log("Warning: Incomplete AppleScript detected, skipping execution")
                    return
            else:
                log("Warning: Incomplete AppleScript detected, skipping execution")
                return
        
        # Return early since we processed without object
        log("=" * 60)
        log("✅ Inserted all plots into Keynote slides")
        log("=" * 60)
        return
    
    # Process each object group (object is in template)
    # Get object values from variable_combinations if selected_objects is empty
    if not selected_objects and has_object:
        # Try to get object values from variable_combinations
        object_values = variable_combinations.get("object", [])
        if object_values:
            # Convert to selected_objects format: each object becomes its own group
            selected_objects = [[obj_name] for obj_name in object_values]
            log(f"Using object values from variable_combinations: {selected_objects}")
        else:
            # Object is in template but not provided in either place
            log("Warning: Object is in template but no object values provided")
            return
    
    # Get step values from variable_combinations if selected_steps is empty
    if not selected_steps and has_step:
        # Try to get step values from variable_combinations
        step_values = variable_combinations.get("step", [])
        if step_values:
            selected_steps = step_values
            log(f"Using step values from variable_combinations: {selected_steps}")
        elif has_step:
            # Step is in template but not provided in either place
            log("Warning: Step is in template but no step values provided")
            return
    
    for obj_group in selected_objects:
        # Check for interruption
        if thread_instance and thread_instance.isInterruptionRequested():
            log("Generation cancelled by user")
            return
        
        # Determine layout based on number of objects and kinematics
        num_objects = len(obj_group) if has_object else 1
        num_kinematics = len(kinematic_values) if kinematic_values else 1
        
        # Determine arrangement and size/position indices
        total_plots = num_objects * num_kinematics
        
        if total_plots == 1:
            arrangement = "1*1"
            size_idx = 1  # size_[1] = (500, 485)
            pos_idx = 1   # positions_[1] = center
        elif total_plots == 2:
            if num_objects == 1:
                arrangement = "1*2"
                size_idx = 1
                pos_idx = 4   # positions_[4] = center
            else:
                arrangement = "2*1"
                size_idx = 1
                pos_idx = 12  # positions_[12] = 2*1
        elif total_plots == 3:
            arrangement = "1*3"
            size_idx = 0  # size_[0] = (314, 305)
            pos_idx = 7   # positions_[7] = center
        elif total_plots == 4:
            if num_objects == 1:
                arrangement = "1*4"
                size_idx = 4  # size_[4] = (250, 243)
                pos_idx = 10  # positions_[10] = center
            elif num_objects == 2:
                arrangement = "2*2"
                size_idx = 2  # size_[2] = (300, 291)
                pos_idx = 13  # positions_[13] = 2*2
            else:
                arrangement = "1*4"
                size_idx = 4
                pos_idx = 10
        elif total_plots == 6:
            arrangement = "2*3"
            size_idx = 3  # size_[3] = (323, 313)
            pos_idx = 14  # positions_[14] = 2*3
        elif total_plots == 8:
            arrangement = "2*4"
            size_idx = 4  # size_[4] = (250, 243)
            pos_idx = 16  # positions_[16] = 2*4
        else:
            # Default to 1*1 for unknown arrangements
            arrangement = "1*1"
            size_idx = 1
            pos_idx = 1
        
        plot_size = size_[size_idx]
        plot_positions = positions_[pos_idx]
        
        log(f"Processing object group: {obj_group} (arrangement: {arrangement}, total plots: {total_plots})")
        
        # If row_values and column_values are provided, use the general logic above
        # This section is for backward compatibility when has_object is True
        # but row_values/column_values are not provided or not mapped
        # Skip this if we already processed with row/column values above
        if False:  # Disabled - use general logic above instead
            # Collect all unique row variable values from row_values and variable_combinations
            # This ensures we can match any row value to the row variable
            all_available_row_values = []
            
            # First, collect from row_values
            for row_vals in row_values:
                if row_vals:
                    all_available_row_values.extend(row_vals)
            
            # Also collect from variable_combinations if row_var_name is object
            if row_var_name == "object" and obj_group:
                for obj_name in obj_group:
                    if obj_name not in all_available_row_values:
                        all_available_row_values.append(obj_name)
            elif row_var_name and row_var_name in variable_combinations:
                # Add values from variable_combinations
                for val in variable_combinations[row_var_name]:
                    if val not in all_available_row_values:
                        all_available_row_values.append(val)
            
            # Remove duplicates while preserving order
            seen = set()
            all_available_row_values = [x for x in all_available_row_values if not (x in seen or seen.add(x))]
            
            if not all_available_row_values:
                log(f"Warning: No values found for row variable '{row_var_name}' in row_values or variable_combinations")
                return
            
            log(f"Available values for row variable '{row_var_name}': {all_available_row_values}")
            
            # Get custom settings from user_settings if available
            arrangement_sizes = None
            custom_positions = None
            arrangement_positions = None
            if thread_instance and hasattr(thread_instance, 'user_settings') and thread_instance.user_settings:
                plot_settings = thread_instance.user_settings.get('plot', {})
                arrangement_sizes = plot_settings.get('arrangement_sizes', {})
                custom_positions = plot_settings.get('custom_positions', {})
                arrangement_positions = plot_settings.get('arrangement_positions', {})
                if arrangement_sizes or custom_positions or arrangement_positions:
                    log("Using customized size and positions from Settings - Plot tab")
            
            # Determine number of slides: based on max row length
            max_row_length = max([len(row_vals) for row_vals in row_values if row_vals]) if row_values else 0
            if max_row_length == 0:
                log("Warning: No row values provided, skipping")
                return
            
            # Process each step first, then for each step create slides by row index
            for step_name in selected_steps:
                # Check for interruption
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    return
                
                # Process each slide (one per row index)
                for row_val_idx in range(max_row_length):
                    # Check for interruption
                    if thread_instance and thread_instance.isInterruptionRequested():
                        log("Generation cancelled by user")
                        return
                    
                    # Get row values for this row index
                    row_vals_at_idx = []
                    for row_vals in row_values:
                        if row_vals and row_val_idx < len(row_vals):
                            row_vals_at_idx.append(row_vals[row_val_idx])
                        else:
                            row_vals_at_idx.append(None)
                    
                    if not any(row_vals_at_idx):
                        log(f"Warning: No row values for row index {row_val_idx}, skipping slide {row_val_idx+1}")
                        continue
                    
                    # Determine arrangement: num_rows * num_columns (e.g., 2*3)
                    num_rows_in_slide = len([r for r in row_vals_at_idx if r])
                    num_cols_in_slide = len(column_values)
                    
                    # Determine arrangement and positions
                    if num_rows_in_slide == 2 and num_cols_in_slide == 4:
                        slide_arrangement = "2*4"
                        slide_pos_idx = 16  # 2*4
                    elif num_rows_in_slide == 2 and num_cols_in_slide == 3:
                        slide_arrangement = "2*3"
                        slide_pos_idx = 14  # 2*3
                    elif num_rows_in_slide == 2 and num_cols_in_slide == 2:
                        slide_arrangement = "2*2"
                        slide_pos_idx = 13  # 2*2
                    elif num_rows_in_slide == 1:
                        if num_cols_in_slide == 4:
                            slide_arrangement = "1*4"
                            slide_pos_idx = 10  # center
                        elif num_cols_in_slide == 3:
                            slide_arrangement = "1*3"
                            slide_pos_idx = 7  # center
                        elif num_cols_in_slide == 2:
                            slide_arrangement = "1*2"
                            slide_pos_idx = 4  # center
                        else:
                            slide_arrangement = "1*1"
                            slide_pos_idx = 1  # center
                    else:
                        slide_arrangement = f"{num_rows_in_slide}*{num_cols_in_slide}"
                        slide_pos_idx = 14  # default to 2*3
                    
                    # Get plot size from settings or use default
                    if arrangement_sizes and slide_arrangement in arrangement_sizes:
                        plot_width = arrangement_sizes[slide_arrangement]
                        # Calculate height based on aspect ratio (default 2*2 is 300x291)
                        if slide_arrangement == "2*4":
                            aspect_ratio = 243 / 250
                        elif slide_arrangement == "2*3":
                            aspect_ratio = 313 / 323
                        elif slide_arrangement == "2*2":
                            aspect_ratio = 291 / 300
                        elif slide_arrangement == "1*4":
                            aspect_ratio = 243 / 250
                        elif slide_arrangement == "1*3":
                            aspect_ratio = 305 / 314
                        elif slide_arrangement == "1*2":
                            aspect_ratio = 485 / 500
                        elif slide_arrangement == "1*1":
                            aspect_ratio = 485 / 500
                        else:
                            aspect_ratio = 291 / 300  # default
                        plot_height = int(plot_width * aspect_ratio)
                        plot_size = (plot_width, plot_height)
                        log(f"Using custom size for {slide_arrangement}: {plot_size}")
                    else:
                        # Use default size from config.py
                        if slide_arrangement == "2*4":
                            plot_size = size_[4]  # (250, 243)
                        elif slide_arrangement == "2*3":
                            plot_size = size_[3]  # (323, 313)
                        elif slide_arrangement == "2*2":
                            plot_size = size_[2]  # (300, 291)
                        elif slide_arrangement == "1*4":
                            plot_size = size_[4]  # (250, 243)
                        elif slide_arrangement == "1*3":
                            plot_size = size_[0]  # (314, 305)
                        elif slide_arrangement == "1*2":
                            plot_size = size_[1]  # (500, 485)
                        elif slide_arrangement == "1*1":
                            plot_size = size_[1]  # (500, 485)
                        else:
                            plot_size = size_[2]  # default to 2*2 size
                    
                    # Get plot positions from settings or use default
                    position_found = False
                    
                    # For 1*1, 1*2, 1*3, 1*4, check arrangement_positions first
                    if slide_arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
                        position_opt = arrangement_positions.get(slide_arrangement, "Center")
                        if custom_positions:
                            arrangement_key = f"{slide_arrangement}_{position_opt}"
                            if arrangement_key in custom_positions:
                                slide_plot_positions = custom_positions[arrangement_key]
                                log(f"Using custom positions for {arrangement_key} (from arrangement_positions)")
                                position_found = True
                    
                    # If not found yet, try all position options
                    if not position_found and custom_positions:
                        for pos_opt in ["Center", "Upper", "Bottom"]:
                            arrangement_key = f"{slide_arrangement}_{pos_opt}"
                            if arrangement_key in custom_positions:
                                slide_plot_positions = custom_positions[arrangement_key]
                                log(f"Using custom positions for {arrangement_key}")
                                position_found = True
                                break
                    
                    # If still not found, use default positions from config.py
                    if not position_found:
                        # For 1*1, 1*2, 1*3, 1*4, use arrangement_positions to determine default position index
                        if slide_arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
                            position_opt = arrangement_positions.get(slide_arrangement, "Center")
                            position_map = {
                                "1*1": {"Upper": 0, "Center": 1, "Bottom": 2},
                                "1*2": {"Upper": 3, "Center": 4, "Bottom": 5},
                                "1*3": {"Upper": 6, "Center": 7, "Bottom": 8},
                                "1*4": {"Upper": 9, "Center": 10, "Bottom": 11},
                            }
                            if slide_arrangement in position_map and position_opt in position_map[slide_arrangement]:
                                slide_pos_idx = position_map[slide_arrangement][position_opt]
                                slide_plot_positions = positions_[slide_pos_idx]
                                log(f"Using default positions for {slide_arrangement} with {position_opt} option")
                            else:
                                slide_plot_positions = positions_[slide_pos_idx]
                                log(f"Using default positions for {slide_arrangement}")
                        else:
                            slide_plot_positions = positions_[slide_pos_idx]
                            log(f"Using default positions for {slide_arrangement}")
                    
                    # Create title text for this slide
                    row_val_str = ", ".join([r for r in row_vals_at_idx if r])
                    if has_step:
                        title_text = f"{step_name} - {row_val_str}"
                    else:
                        title_text = f"{row_val_str}"
                    
                    # Create new slide for this row index
                    slide_apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items
                    try
                        repeat while (count of text items) > 0
                            delete text item 1
                        end repeat
                    end try
                    
                    -- Add title text box
                    set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
                    set object text of titleBox to "{title_text}"
                    tell object text of titleBox
                        set font of characters 1 thru -1 to "Helvetica Neue"
                        set size of characters 1 thru -1 to 24
                    end tell
            '''
                    
                    # Process each row and column combination for this row index
                    # Grid layout: row by row, column by column
                    slide_plot_idx = 0
                    
                    for slide_row_idx, row_val in enumerate(row_vals_at_idx):
                        if not row_val:
                            continue
                        
                        # Find matching row variable value in all_available_row_values
                        matching_row_val = None
                        for val in all_available_row_values:
                            if val == row_val or str(val) == str(row_val):
                                matching_row_val = val
                                break
                        
                        if not matching_row_val:
                            # Try case-insensitive and substring matching
                            row_val_lower = str(row_val).lower().strip()
                            for val in all_available_row_values:
                                val_lower = str(val).lower().strip()
                                if val_lower == row_val_lower or row_val_lower in val_lower or val_lower in row_val_lower:
                                    matching_row_val = val
                                    break
                        
                        if not matching_row_val:
                            # Use row_val directly if no match found
                            matching_row_val = row_val
                            log(f"Warning: No exact match for row value '{row_val}', using it directly")
                        
                        # Process each column for this row
                        for col_idx, col_val in enumerate(column_values):
                            # Find matching column variable value
                            matching_col_val = None
                            
                            # If column variable is kinematic, use kinematic matching logic
                            if col_var_name == "kinematic" and kinematic_values:
                                col_val_lower = str(col_val).lower().strip()
                                for kin_name in kinematic_values:
                                    kin_name_lower = str(kin_name).lower().strip()
                                    # Try exact match first (case-insensitive)
                                    if kin_name_lower == col_val_lower:
                                        matching_col_val = kin_name
                                        break
                                    # Try substring match (case-insensitive)
                                    elif col_val_lower in kin_name_lower or kin_name_lower in col_val_lower:
                                        matching_col_val = kin_name
                                        break
                                
                                # If still no match, try to use kinematic_values in order
                                if not matching_col_val and col_idx < len(kinematic_values):
                                    matching_col_val = kinematic_values[col_idx]
                                    log(f"Using kinematic at index {col_idx} ({matching_col_val}) for column value '{col_val}'")
                            else:
                                # For other column variables, try to match from variable_combinations
                                if col_var_name in variable_combinations:
                                    col_vals = variable_combinations[col_var_name]
                                    col_val_lower = str(col_val).lower().strip()
                                    for val in col_vals:
                                        val_lower = str(val).lower().strip()
                                        if val_lower == col_val_lower or col_val_lower in val_lower or val_lower in col_val_lower:
                                            matching_col_val = val
                                            break
                                    
                                    # If still no match, try by index
                                    if not matching_col_val and col_idx < len(col_vals):
                                        matching_col_val = col_vals[col_idx]
                                        log(f"Using {col_var_name} at index {col_idx} ({matching_col_val}) for column value '{col_val}'")
                            
                            # If still no match, use column value directly
                            if not matching_col_val:
                                matching_col_val = col_val
                                log(f"No matching value found for column variable '{col_var_name}', using column value '{col_val}' directly")
                            
                            # Build variables dictionary
                            variables = {}
                            
                            # Add variables from variable_combinations (excluding row_var, col_var, object, step, kinematic, format)
                            excluded_vars = [row_var_name, col_var_name, "object", "step", "kinematic", "format"]
                            for var_name, var_values in variable_combinations.items():
                                if var_name not in excluded_vars:
                                    if var_values:
                                        # Use first value for now
                                        variables[var_name] = var_values[0]
                            
                            # Set row variable, column variable, object, step, kinematic, format
                            if row_var_name:
                                variables[row_var_name] = matching_row_val
                            if col_var_name:
                                variables[col_var_name] = matching_col_val
                            
                            # Also set object and kinematic if they exist (for backward compatibility)
                            if has_object and row_var_name != "object":
                                # Object might be in variable_combinations or selected_objects
                                if obj_group:
                                    variables["object"] = obj_group[0] if obj_group else None
                                elif "object" in variable_combinations:
                                    obj_vals = variable_combinations["object"]
                                    if obj_vals:
                                        variables["object"] = obj_vals[0]
                            
                            if has_kinematic and col_var_name != "kinematic":
                                if kinematic_values:
                                    variables["kinematic"] = kinematic_values[0] if kinematic_values else None
                            
                            if has_step:
                                variables["step"] = step_name
                            variables["format"] = file_format
                            
                            # Generate file path
                            file_path = parser.build_path(base_dir, variables)
                            
                            # Insert plot if exists
                            if os.path.exists(file_path):
                                # Calculate grid position: row * num_cols + col
                                grid_plot_idx = slide_row_idx * num_cols_in_slide + col_idx
                                x, y = _get_position_for_plot(grid_plot_idx, slide_arrangement, slide_plot_positions, log)
                                if x is not None and y is not None:
                                    row_info = f"{row_var_name}={matching_row_val}" if row_var_name else ""
                                    col_info = f"{col_var_name}={matching_col_val}" if col_var_name else ""
                                    step_info = f"step={step_name}" if has_step else ""
                                    info_parts = [p for p in [row_info, col_info, step_info] if p]
                                    log(f"Inserting ({', '.join(info_parts)}) in slide {row_val_idx+1} at grid position ({slide_row_idx}, {col_idx}) -> ({x}, {y})")
                                    slide_apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                                    slide_plot_idx += 1
                            else:
                                log(f"File not found: {file_path}")
                    
                    # Close slide and execute
                    slide_apple_script += '''
                end tell
            end tell
            end tell
            '''
                    
                    # Execute this slide's script
                    if thread_instance and thread_instance.isInterruptionRequested():
                        log("Generation cancelled by user")
                        return
                    
                    script_lines = slide_apple_script.strip().split('\n')
                    if len(script_lines) >= 3:
                        last_lines = [line.strip() for line in script_lines[-3:]]
                        if all(line == "end tell" for line in last_lines):
                            try:
                                subprocess.run(["osascript", "-e", slide_apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
                            except subprocess.CalledProcessError as e:
                                log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                                return
            
            # Return early since we processed with row/column values
            log("=" * 60)
            log("✅ Inserted all plots into Keynote slides")
            log("=" * 60)
            return
        
        # Process each step (original logic when row_values/column_values not provided)
        for step_name in selected_steps:
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Create title text for the slide
            if has_object:
                obj_names = " & ".join(obj_group)
                if has_step:
                    title_text = f"{obj_names} - {step_name}"
                else:
                    title_text = f"{obj_names}"
            else:
                if has_step:
                    title_text = f"{step_name}"
                else:
                    title_text = "Loop-defined plots"
            
            # Create new slide - each slide has complete AppleScript
            apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items
                    try
                        repeat while (count of text items) > 0
                            delete text item 1
                        end repeat
                    end try
                    
                    -- Add title text box
                    set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
                    set object text of titleBox to "{title_text}"
                    tell object text of titleBox
                        set font of characters 1 thru -1 to "Helvetica Neue"
                        set size of characters 1 thru -1 to 24
                    end tell
            '''
            
            # Process each object and kinematic combination
            plot_idx = 0
            
            # If object is not in template, process without object loop
            if not has_object:
                # Process kinematic values directly (if kinematic is in template)
                if has_kinematic and kinematic_values:
                    for kin_name in kinematic_values:
                        # Build variables dictionary
                        variables = {}
                        
                        # Add variables from variable_combinations (excluding object, step, kinematic, format)
                        for var_name, var_values in variable_combinations.items():
                            if var_name not in ["object", "step", "kinematic", "format"]:
                                if var_values:
                                    # Use first value for now
                                    variables[var_name] = var_values[0]
                        
                        # Set step, kinematic, format (no object)
                        if has_step:
                            variables["step"] = step_name
                        variables["kinematic"] = kin_name
                        variables["format"] = file_format
                        
                        # Generate file path
                        file_path = parser.build_path(base_dir, variables)
                        
                        # Insert plot if exists
                        if os.path.exists(file_path):
                            x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                            if x is not None and y is not None:
                                log(f"Inserting ({step_name if has_step else ''}, {kin_name}) at ({x}, {y})")
                                apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                                plot_idx += 1
                        else:
                            log(f"File not found: {file_path}")
                else:
                    # No kinematic, no object - process with just step and other variables
                    variables = {}
                    
                    # Add variables from variable_combinations
                    for var_name, var_values in variable_combinations.items():
                        if var_name not in ["object", "step", "format"]:
                            if var_values:
                                variables[var_name] = var_values[0]
                    
                    # Set step, format (no object, no kinematic)
                    if has_step:
                        variables["step"] = step_name
                    variables["format"] = file_format
                    
                    # Generate file path
                    if scan_mode:
                        # In scan mode, find file from scan results
                        file_path = _find_file_from_scan(variables, thread_instance, log)
                    else:
                        # Normal mode: build path from template
                        file_path = parser.build_path(base_dir, variables)
                    
                    # Insert plot if exists
                    if file_path and os.path.exists(file_path):
                        x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                        if x is not None and y is not None:
                            log(f"Inserting ({step_name if has_step else ''}) at ({x}, {y})")
                            apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                            plot_idx += 1
                    else:
                        log(f"File not found: {file_path}")
            else:
                    # Process normally (original logic)
                for obj_name in obj_group:
                    # If kinematic is in template, process each kinematic
                    if has_kinematic and kinematic_values:
                        for kin_name in kinematic_values:
                            # Build variables dictionary
                            variables = {}
                            
                            # Add variables from variable_combinations (excluding object, step, kinematic, format)
                            for var_name, var_values in variable_combinations.items():
                                if var_name not in ["object", "step", "kinematic", "format"]:
                                    if var_values:
                                        # Use first value for now
                                        variables[var_name] = var_values[0]
                            
                            # Set object, step, kinematic, format
                            variables["object"] = obj_name
                            if has_step:
                                variables["step"] = step_name
                            variables["kinematic"] = kin_name
                            variables["format"] = file_format
                            
                            # Generate file path
                            file_path = parser.build_path(base_dir, variables)
                            
                            # Insert plot if exists
                            if os.path.exists(file_path):
                                x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                                if x is not None and y is not None:
                                    log(f"Inserting {obj_name} ({step_name}, {kin_name}) at ({x}, {y})")
                                    apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                                    plot_idx += 1
                            else:
                                log(f"File not found: {file_path}")
                else:
                    # No kinematic variable, process object only
                    variables = {}
                    
                    # Add variables from variable_combinations
                    for var_name, var_values in variable_combinations.items():
                        if var_name not in ["object", "step", "format"]:
                            if var_values:
                                variables[var_name] = var_values[0]
                    
                            # Set object, step, format
                            variables["object"] = obj_name
                            if has_step:
                                variables["step"] = step_name
                            variables["format"] = file_format
                    
                    # Generate file path
                    if scan_mode:
                        # In scan mode, find file from scan results
                        file_path = _find_file_from_scan(variables, thread_instance, log)
                    else:
                        # Normal mode: build path from template
                        file_path = parser.build_path(base_dir, variables)
                    
                    # Insert plot if exists
                    if file_path and os.path.exists(file_path):
                        x, y = _get_position_for_plot(plot_idx, arrangement, plot_positions, log)
                        if x is not None and y is not None:
                            log(f"Inserting {obj_name} ({step_name}) at ({x}, {y})")
                            apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {plot_size[0]}
                    set height of ObjImg to {plot_size[1]}
                    '''
                            plot_idx += 1
                    else:
                        log(f"File not found: {file_path}")
            
            # Close slide and execute immediately
            apple_script += '''
                end tell
            end tell
            end tell
            '''
            
            # Check for interruption before executing
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Only execute if AppleScript is complete (has proper closing)
            script_lines = apple_script.strip().split('\n')
            if len(script_lines) >= 3:
                last_lines = [line.strip() for line in script_lines[-3:]]
                if all(line == "end tell" for line in last_lines):
                    try:
                        subprocess.run(["osascript", "-e", apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
                    except subprocess.CalledProcessError as e:
                        log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                        return
                else:
                    log("Warning: Incomplete AppleScript detected, skipping execution")
                    return
            else:
                log("Warning: Incomplete AppleScript detected, skipping execution")
                return
    
    log("=" * 60)
    log("✅ Inserted all plots into Keynote slides")
    log("=" * 60)


def _get_position_for_plot(plot_idx, arrangement, plot_positions, log):
    """Get position (x, y) for a plot based on index and arrangement"""
    if arrangement in ["2*3", "2*4", "2*2", "2*1"]:
        # 2D layout
        if arrangement == "2*3":
            cols_per_row = 3
        elif arrangement == "2*4":
            cols_per_row = 4
        elif arrangement == "2*2":
            cols_per_row = 2
        else:  # 2*1
            cols_per_row = 1
        
        row = plot_idx // cols_per_row
        col = plot_idx % cols_per_row
        
        if row < len(plot_positions) and col < len(plot_positions[row]):
            return plot_positions[row][col]
        else:
            log(f"Warning: Position out of range for plot {plot_idx} (row {row}, col {col})")
            return None, None
    elif arrangement == "1*3":
        # 1*3 layout: positions structure is [ [(x1,y1)], [(x2,y2)], [(x3,y3)] ]
        # Access: positions[plot_idx][0]
        if plot_idx < len(plot_positions):
            if len(plot_positions[plot_idx]) > 0:
                return plot_positions[plot_idx][0]
            else:
                log(f"Warning: Position sublist is empty for plot {plot_idx}")
                return None, None
        else:
            log(f"Warning: Position out of range for plot {plot_idx} (max index: {len(plot_positions)-1})")
            return None, None
    elif arrangement == "1*4":
        # 1*4 layout: positions structure is [ [(x1,y1)], [(x2,y2)], [(x3,y3)], [(x4,y4)] ]
        # Access: positions[plot_idx][0]
        if plot_idx < len(plot_positions):
            return plot_positions[plot_idx][0]
        else:
            log(f"Warning: Position out of range for plot {plot_idx}")
            return None, None
    else:
        # 1D layout (1*1, 1*2) - structure is [ [(x1,y1), (x2,y2), ...] ]
        # Access: positions[0][plot_idx]
        if plot_idx < len(plot_positions[0]):
            return plot_positions[0][plot_idx]
        else:
            log(f"Warning: Position out of range for plot {plot_idx}")
            return None, None
