"""
Keynote Control Module - DRC Mode Functions
Detector Response Correction Mode specific functions

================================================================================
DRC MODE FUNCTIONS:
    - insert_pdfs_into_slide_drc()                      : Insert DRC energy plots
    - insert_pdfs_into_slide_drc_resol_linearity()     : Insert DRC resolution/linearity plots
================================================================================
"""
import os
import subprocess
import sys

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# Try relative import first, then absolute import
try:
    from ..config import size_, positions_
    from ..Rules import get_file_path_drc
except ImportError:
    from Config.config import size_, positions_
    from Config.Rules import get_file_path_drc


################################################################################
#                                                                              #
#                          DRC MODE FUNCTIONS                                  #
#                    (Detector Response Correction Mode)                       #
#                                                                              #
################################################################################

def insert_pdfs_into_slide_drc(output_file, base_dir, programs, channels, energies, size, positions, log_callback=None, plot_width=None, thread_instance=None):
    """
    Insert DRC energy deposition plots into Keynote slides
    
    Args:
        output_file (str): Path to output Keynote file
        base_dir (str): Base directory where plots are located
        programs (list): List of programs (e.g., ["proton_Rot"])
        channels (list): List of channels (e.g., ["S_LCATTcor", "C"])
        energies (list or dict): List of energies (e.g., ["energy20_run20", ...]) OR
                                 Dictionary mapping channel names to energy lists (e.g., {"C": [...], "S": [...]})
        size (tuple): Plot size (width, height)
        positions (list): 2D list of positions for 2×3 layout
        log_callback (callable): Optional callback function for logging messages
        plot_width (int): Custom plot width (optional, height will be proportional)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    abs_output = os.path.abspath(output_file)
    
    # Apply custom plot width if provided (height will be proportional)
    if plot_width is not None:
        aspect_ratio = size[1] / size[0] if size[0] > 0 else 1
        new_height = int(plot_width * aspect_ratio)
        size = (plot_width, new_height)
        log(f"Applied custom plot width for DRC: {plot_width}, height: {new_height} (proportional)")
    
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
    
    ################
    # Program loop #
    ################
    for program in programs:
        # Check for interruption
        if thread_instance and thread_instance.isInterruptionRequested():
            log("Generation cancelled by user")
            return
        
        #################
        # Channel loop  #
        #################
        for channel in channels:
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            # Get energies for this channel
            # If energies is a dict, use channel-specific energies; otherwise use the same list for all channels
            if isinstance(energies, dict):
                # Map channel name to tab key: "C_S_overlay_*" -> "C_S_overlay",
                # "S_*" -> "S", "DRcor_*" -> "DRcor", "C" -> "C"
                if channel.startswith("C_S_overlay"):
                    channel_base = "C_S_overlay"
                else:
                    channel_base = channel.split("_")[0]
                channel_energies = energies.get(channel_base, [])
                if not channel_energies:
                    # Fallback: try to find any matching key
                    for key in energies.keys():
                        if key in channel or channel.startswith(key):
                            channel_energies = energies[key]
                            break
                    if not channel_energies:
                        # Last resort: use first available energies
                        channel_energies = list(energies.values())[0] if energies else []
                log(f"Using energies for channel {channel} (base: {channel_base}): {channel_energies}")
            else:
                # Legacy: use same energies for all channels
                channel_energies = energies
            
            # Create title text
            title_text = f"{program} - {channel}"
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
            
            ################
            # Energy loop  #
            ################
            for idx, energy in enumerate(channel_energies):
                # Check for interruption
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    # Close any open tell blocks before returning
                    apple_script += '''
                end tell
            end tell
            end tell
            '''
                    return
                # Skip if energy is "Empty"
                if energy == "Empty":
                    log(f"Skipping position {idx} (Empty)")
                    continue
                
                file_path, file_name = get_file_path_drc(base_dir, program, channel, energy)
                
                if os.path.exists(file_path):
                    # Calculate position in 2×3 array using actual index
                    row = idx // 3  # 0 or 1
                    col = idx % 3   # 0, 1, or 2
                    x, y = positions[row][col]
                    log(f"Inserting {file_name} at position {idx} (row {row}, col {col}) at ({x}, {y})")
                    apple_script += f'''
                    set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                    set position of ObjImg to {{{x}, {y}}}
                    set width of ObjImg to {size[0]}
                    set height of ObjImg to {size[1]}
                    '''
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
                # Don't execute incomplete AppleScript
                return
            
            # Only execute if AppleScript is complete (has proper closing)
            # Check if script ends with proper closing (3 end tell statements)
            script_lines = apple_script.strip().split('\n')
            if len(script_lines) >= 3:
                last_lines = [line.strip() for line in script_lines[-3:]]
                if all(line == "end tell" for line in last_lines):
                    try:
                        subprocess.run(["osascript", "-e", apple_script], check=True, stderr=subprocess.PIPE)
                    except subprocess.CalledProcessError as e:
                        log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                        return
                else:
                    log("Warning: Incomplete AppleScript detected, skipping execution")
                    return
            else:
                log("Warning: Incomplete AppleScript detected, skipping execution")
                return
            apple_script = ""  # Reset for next slide
    
    print("=" * 71)
    print("|| Inserted slides into Keynote. ---> Fill in the title and text!!!! ||")
    print("=" * 71)


##############################################
# DRC: Resolution & Linearity plots          #
##############################################
def insert_pdfs_into_slide_drc_resol_linearity(output_file, base_dir, programs, particle_name, 
                                                resol_with_noise, resol_without_noise, linearity, 
                                                file_format="pdf", log_callback=None, plot_width=None, 
                                                positionOpt="Center", thread_instance=None):
    """
    Insert DRC Resolution and Linearity plots into Keynote slides
    
    Args:
        output_file (str): Path to output Keynote file
        base_dir (str): Base directory where plots are located
        programs (list): List of programs (e.g., ["proton_Rot"])
        particle_name (str): Particle name (e.g., "proton", "pi", "em", "kaon")
        resol_with_noise (bool): Include resolution with noise term
        resol_without_noise (bool): Include resolution without noise term
        linearity (bool): Include linearity
        file_format (str): File format ("pdf" or "png")
        log_callback (callable): Optional callback function for logging messages
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    
    abs_output = os.path.abspath(output_file)
    
    # Build list of plots in order: with noise, without noise, linearity
    plots_to_insert = []
    
    if resol_with_noise:
        plots_to_insert.append({
            "name": "Resolution (with Noise term)",
            "file_name": f"c_{particle_name}Resol_M5-T2.{file_format}"
        })
    
    if resol_without_noise:
        plots_to_insert.append({
            "name": "Resolution (without Noise term)",
            "file_name": f"c_{particle_name}Resol_M5-T2_NNR_From30.{file_format}"
        })
    
    if linearity:
        plots_to_insert.append({
            "name": "Linearity",
            "file_name": f"c_{particle_name}Linearity_M5-T2.{file_format}"
        })
    
    # Determine positions based on number of plots and position option
    num_plots = len(plots_to_insert)
    
    if num_plots == 3:
        # Use 1x3 layout: positions_[6] (upper), [7] (center), [8] (bottom)
        if positionOpt == "Upper":
            positions_list = [positions_[6][0][0], positions_[6][1][0], positions_[6][2][0]]
        elif positionOpt == "Center":
            positions_list = [positions_[7][0][0], positions_[7][1][0], positions_[7][2][0]]
        elif positionOpt == "Bottom":
            positions_list = [positions_[8][0][0], positions_[8][1][0], positions_[8][2][0]]
        else:
            positions_list = [positions_[7][0][0], positions_[7][1][0], positions_[7][2][0]]  # default: center
        size = size_[0]  # 1*3 레이아웃: size_[0] (314, 305)
    elif num_plots == 2:
        # Use 1x2 layout: positions_[3] (upper), [4] (center), [5] (bottom)
        if positionOpt == "Upper":
            positions_list = positions_[3][0]
        elif positionOpt == "Center":
            positions_list = positions_[4][0]
        elif positionOpt == "Bottom":
            positions_list = positions_[5][0]
        else:
            positions_list = positions_[4][0]  # default: center
        size = size_[1]  # 1*2 레이아웃: size_[1] (500, 485)
    else:  # num_plots == 1
        # Use 1x1 layout: positions_[0] (upper), [1] (center), [2] (bottom)
        if positionOpt == "Upper":
            positions_list = [positions_[0][0][0]]
        elif positionOpt == "Center":
            positions_list = [positions_[1][0][0]]
        elif positionOpt == "Bottom":
            positions_list = [positions_[2][0][0]]
        else:
            positions_list = [positions_[1][0][0]]  # default: center
        size = size_[1]  # 1*1 레이아웃: size_[1] (500, 485)
    
    # Apply custom plot width if provided (height will be proportional)
    if plot_width is not None:
        aspect_ratio = size[1] / size[0] if size[0] > 0 else 1
        new_height = int(plot_width * aspect_ratio)
        size = (plot_width, new_height)
        log(f"Applied custom plot width for DRC Resolution/Linearity: {plot_width}, height: {new_height} (proportional)")
    
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
    
    # Initialize AppleScript accumulator (will be executed per slide)
    apple_script = ""
    
    ################
    # Program loop #
    ################
    for program in programs:
        # Create title text for this slide
        title_parts = [plot["name"] for plot in plots_to_insert]
        title_text = f"{program} - " + " + ".join(title_parts)
        
        # Create new slide
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
        
        # Insert each plot
        for idx, plot_info in enumerate(plots_to_insert):
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                # Close any open tell blocks before returning
                apple_script += '''
            end tell
        end tell
        end tell
        '''
                return
            
            file_name = plot_info["file_name"]
            file_path = os.path.join(base_dir, program, file_name)
            
            if os.path.exists(file_path):
                log(f"Inserting {plot_info['name']}: {file_name}")
                x, y = positions_list[idx]
                apple_script += f'''
                -- Add plot {idx + 1}: {plot_info['name']}
                set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                set position of ObjImg to {{{x}, {y}}}
                set width of ObjImg to {size[0]}
                set height of ObjImg to {size[1]}
                '''
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
            # Don't execute incomplete AppleScript
            return
        
        # Only execute if AppleScript is complete (has proper closing)
        # Check if script ends with proper closing (3 end tell statements)
        script_lines = apple_script.strip().split('\n')
        if len(script_lines) >= 3:
            last_lines = [line.strip() for line in script_lines[-3:]]
            if all(line == "end tell" for line in last_lines):
                try:
                    subprocess.run(["osascript", "-e", apple_script], check=True, stderr=subprocess.PIPE)
                except subprocess.CalledProcessError as e:
                    log(f"AppleScript execution error: {e.stderr.decode() if e.stderr else 'Unknown error'}")
                    return
            else:
                log("Warning: Incomplete AppleScript detected, skipping execution")
                return
        else:
            log("Warning: Incomplete AppleScript detected, skipping execution")
            return
        log("Resolution/Linearity plots inserted successfully!")
