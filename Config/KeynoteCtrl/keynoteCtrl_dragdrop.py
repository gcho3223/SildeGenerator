#!/usr/bin/env python3
"""
Drag and Drop Mode Keynote Control Functions
Direct file insertion for drag and drop mode
"""

import os
import subprocess
from Config.config import size_, positions_


def insert_pdfs_into_slide_dragdrop(output_file, file_paths_dict, arrangement, slide_count, 
                                    log_callback=None, thread_instance=None, user_settings=None):
    """
    Insert PDF/PNG files into Keynote slides for drag and drop mode
    
    Args:
        output_file (str): Path to output Keynote file
        file_paths_dict (dict): Dictionary mapping slide_idx to list of file paths (None for empty cells)
                                Format: {slide_idx: [file_paths]}
        arrangement (str): Arrangement string (e.g., "2*3", "2*2")
        slide_count (int): Number of slides to create
        log_callback (callable): Optional callback function for logging messages
        thread_instance: Thread instance for interruption checking
        user_settings (dict): User settings for custom sizes and positions
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    
    abs_output = os.path.abspath(output_file)
    
    # Parse arrangement
    try:
        num_rows, num_cols = map(int, arrangement.split('*'))
    except:
        num_rows, num_cols = 2, 3
        log(f"Invalid arrangement '{arrangement}', using default 2*3")
    
    # Get plot size from settings or use default
    arrangement_sizes = None
    custom_positions = None
    arrangement_positions = None
    
    if user_settings:
        plot_settings = user_settings.get('plot', {})
        arrangement_sizes = plot_settings.get('arrangement_sizes', {})
        custom_positions = plot_settings.get('custom_positions', {})
        arrangement_positions = plot_settings.get('arrangement_positions', {})
        if arrangement_sizes or custom_positions or arrangement_positions:
            log("Using customized size and positions from Settings - Plot tab")
    
    # Get plot size
    if arrangement_sizes and arrangement in arrangement_sizes:
        plot_width = arrangement_sizes[arrangement]
        aspect_ratio = {
            "2*4": 243 / 250, "2*3": 313 / 323, "2*2": 291 / 300,
            "1*4": 243 / 250, "1*3": 305 / 314, "1*2": 485 / 500, "1*1": 485 / 500
        }.get(arrangement, 313 / 323)
        plot_height = int(plot_width * aspect_ratio)
        plot_size = (plot_width, plot_height)
        log(f"Using custom size for {arrangement}: {plot_size}")
    else:
        size_map = {
            "2*4": size_[4], "2*3": size_[3], "2*2": size_[2],
            "1*4": size_[4], "1*3": size_[0], "1*2": size_[1], "1*1": size_[1]
        }
        plot_size = size_map.get(arrangement, size_[3])
        log(f"Using default size for {arrangement}: {plot_size}")
    
    # Get plot positions
    position_found = False
    slide_plot_positions = None
    
    if arrangement in ["1*1", "1*2", "1*3", "1*4"] and arrangement_positions:
        position_opt = arrangement_positions.get(arrangement, "Center")
        if custom_positions:
            arrangement_key = f"{arrangement}_{position_opt}"
            if arrangement_key in custom_positions:
                slide_plot_positions = custom_positions[arrangement_key]
                log(f"Using custom positions for {arrangement_key}")
                position_found = True
    
    if not position_found and custom_positions:
        for pos_opt in ["Center", "Upper", "Bottom"]:
            arrangement_key = f"{arrangement}_{pos_opt}"
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
        pos_idx = pos_idx_map.get(arrangement, 14)
        slide_plot_positions = positions_[pos_idx]
        log(f"Using default positions for {arrangement}")
    
    # Create AppleScript for multiple slides
    apple_script = f'''
    tell application "Keynote"
    activate
    set theDoc to open (POSIX file "{abs_output}")
    tell document 1
    '''
    
    # Create slides and insert files
    total_file_count = 0
    for slide_idx in range(slide_count):
        if thread_instance and thread_instance.isInterruptionRequested():
            log("Generation cancelled by user")
            return
        
        file_paths = file_paths_dict.get(slide_idx, [])
        
        # Create new slide
        apple_script += f'''
        set newSlide{slide_idx} to make new slide
        tell newSlide{slide_idx}
            -- Delete default text items
            try
                repeat while (count of text items) > 0
                    delete text item 1
                end repeat
            end try
            
            -- Add title text box
            set titleBox to make new text item with properties {{position:{{50, 20}}, width:900, height:40}}
            set object text of titleBox to "Drag and Drop Plots ({arrangement}) - Slide {slide_idx + 1}"
            tell object text of titleBox
                set font of characters 1 thru -1 to "Helvetica Neue"
                set size of characters 1 thru -1 to 24
            end tell
        '''
        
        # Insert files in grid order for this slide
        file_count = 0
        for idx, file_path in enumerate(file_paths):
            if file_path is None:
                continue  # Skip empty cells
            
            if not os.path.exists(file_path):
                log(f"Warning: File not found: {file_path}")
                continue
            
            # Calculate row and column from index
            row = idx // num_cols
            col = idx % num_cols
            
            # Get position from slide_plot_positions
            if slide_plot_positions and row < len(slide_plot_positions):
                row_positions = slide_plot_positions[row]
                if col < len(row_positions):
                    x, y = row_positions[col]
                else:
                    log(f"Warning: Column {col} out of range for row {row}, skipping")
                    continue
            else:
                log(f"Warning: Row {row} out of range, skipping")
                continue
            
            file_count += 1
            total_file_count += 1
            file_name = os.path.basename(file_path)
            log(f"Slide {slide_idx + 1}: Inserting {file_name} at position ({row}, {col}) -> ({x}, {y})")
            
            apple_script += f'''
            set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
            set position of ObjImg to {{{x}, {y}}}
            set width of ObjImg to {plot_size[0]}
            set height of ObjImg to {plot_size[1]}
            '''
        
        # Close slide
        apple_script += '''
        end tell
        '''
        
        log(f"Slide {slide_idx + 1}: Inserted {file_count} file(s)")
    
    # Close document and application blocks
    apple_script += '''
    end tell
    end tell
    '''
    
    # Check for interruption before executing
    if thread_instance and thread_instance.isInterruptionRequested():
        log("Generation cancelled by user")
        return
    
    # Execute AppleScript
    try:
        subprocess.run(["osascript", "-e", apple_script], check=True, stderr=subprocess.PIPE, stdout=subprocess.DEVNULL)
        log(f"Successfully inserted {total_file_count} file(s) into {slide_count} Keynote slide(s)")
    except subprocess.CalledProcessError as e:
        log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
        return
    
    log("=" * 60)
    log(f"✅ Inserted all plots into {slide_count} Keynote slide(s) (Drag and Drop mode)")
    log("=" * 60)
