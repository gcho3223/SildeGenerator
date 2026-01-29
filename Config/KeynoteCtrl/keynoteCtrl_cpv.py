"""
Keynote Control Module - CPV Mode Functions
CP Violation Analysis Mode specific functions

================================================================================
CPV MODE FUNCTIONS:
    - insert_pdfs_into_slide()              : Insert CPV normal mode plots
    - insert_pdfs_into_slide_systematic()   : Insert CPV systematic comparison plots
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
    from ..Rules import get_file_path, get_file_path_systematic, get_kinematics, get_size_and_positions
except ImportError:
    from Config.config import size_, positions_
    from Config.Rules import get_file_path, get_file_path_systematic, get_kinematics, get_size_and_positions


################################################################################
#                                                                              #
#                          CPV MODE FUNCTIONS                                  #
#                    (CP Violation Analysis Mode)                              #
#                                                                              #
################################################################################

def insert_pdfs_into_slide(output_file, plot_dir, selected_objects, selected_steps, positionOpt="Center", 
                          file_format="pdf", log_callback=None, custom_positions=None, plot_width=None, 
                          arrangement_positions=None, arrangement_sizes=None, thread_instance=None):
    """
    Insert PDF plots into Keynote slides
    Args:
        output_file (str): Path to output Keynote file
        plot_dir (str): Base directory where plots are located
        selected_objects (list): List of object groups to process (e.g., [["Lep1", "Lep2"], ["Jet1"]])
        selected_steps (list): List of steps to process (e.g., ["step1", "step2"])
        positionOpt (str): Position option for single kinematic plots ("Upper", "Center", "Bottom")
        file_format (str): File format ("pdf" or "png")
        log_callback (callable): Optional callback function for logging messages
        custom_positions (dict): Custom positions from settings (optional)
        plot_width (int): Custom plot width (optional, height will be proportional) - deprecated, use arrangement_sizes instead
        arrangement_positions (dict): Arrangement-specific positions from settings (optional)
        arrangement_sizes (dict): Arrangement-specific sizes (width) from settings (optional)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    abs_output = os.path.abspath(output_file)
    
    # Initialize AppleScript
    apple_script = f'''
    tell application "Keynote"
    activate
    set theDoc to open (POSIX file "{abs_output}")
    '''
    
    ########################################
    # Special case: Num_PV (process first) #
    ########################################
    # Check if Num_PV is in selected objects
    num_pv_selected = any("Num_PV" in obj_group for obj_group in selected_objects)
    if num_pv_selected:
        # Filter out afterTopReco and Observable steps (Num_PV not available in these steps)
        valid_steps = [s for s in selected_steps if s != "afterTopReco" and s != "Observable"]
        
        if valid_steps:
            # Num_PV: always use 2*3 layout with fixed size
            pv_positions = positions_[14]  # 2*3 : 14th index (CPV)
            pv_size = size_[3]  # 2*3 레이아웃: size_[3] (323, 313)
            
            # Get arrangement-specific width for 2*3 layout if available
            pv_width = None
            if arrangement_sizes and "2*3" in arrangement_sizes:
                pv_width = arrangement_sizes["2*3"]
            elif plot_width is not None:
                pv_width = plot_width
            
            # Apply custom plot width if provided (height will be proportional)
            if pv_width is not None:
                aspect_ratio = pv_size[1] / pv_size[0] if pv_size[0] > 0 else 1
                new_height = int(pv_width * aspect_ratio)
                pv_size = (pv_width, new_height)
            
            # Sort steps by step number (initial=0, step1=1, step2=2, ..., step6=6)
            def get_step_number(step_name):
                if step_name == "initial":
                    return 0
                elif step_name.startswith("step"):
                    try:
                        return int(step_name.replace("step", ""))
                    except:
                        return 999  # Unknown steps go to end
                else:
                    return 999
            
            sorted_steps = sorted(valid_steps, key=get_step_number)
            
            # Check if initial is selected
            has_initial = "initial" in sorted_steps
            
            # Determine position mapping:
            # - If initial is selected: initial → position 0 (1번 위치), step1 → position 1, step2 → position 2, ...
            # - If initial is NOT selected: step1 → position 0 (1번 위치), step2 → position 1, step3 → position 2, ...
            position_map = {}
            pos_idx = 0
            for step_name in sorted_steps:
                if has_initial:
                    # initial is selected: include all steps in order
                    position_map[step_name] = pos_idx
                    pos_idx += 1
                else:
                    # initial is NOT selected: skip initial, step1 starts at position 0
                    if step_name != "initial":
                        position_map[step_name] = pos_idx
                        pos_idx += 1
            
            # Split steps into slides (max 6 per slide for 2*3 layout)
            slides_steps = []
            current_slide = []
            for step_name in sorted_steps:
                if len(current_slide) < 6:
                    current_slide.append(step_name)
                else:
                    slides_steps.append(current_slide)
                    current_slide = [step_name]
            if current_slide:
                slides_steps.append(current_slide)
            
            # Create slides for Num_PV
            for slide_idx, slide_steps in enumerate(slides_steps):
                # Check for interruption
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    return
                if len(slide_steps) == 1:
                    title_text = f"Num_PV - {slide_steps[0]}"
                else:
                    steps_str = ", ".join(slide_steps)
                    title_text = f"Num_PV - {steps_str}"
                
                apple_script = f'''
                tell application "Keynote"
                tell document 1
                    set newSlide to make new slide
                    tell newSlide
                        -- Delete default text items (placeholder boxes) if they exist
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
                
                # Insert Num_PV plots for steps in this slide
                for step_name in slide_steps:
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
                    # Get position index for this step
                    pos_idx = position_map.get(step_name, 0)
                    # Limit to 0-5 for 2*3 layout
                    pos_idx = min(pos_idx, 5)
                    
                    # Extract step number for file naming
                    if step_name == "initial":
                        step_num = ""
                    else:
                        step_num = step_name.replace("step", "")
                    
                    file_path, file_name = get_file_path(plot_dir, step_name, "Num_PV", "PV", file_format)
                    
                    if os.path.exists(file_path):
                        # Calculate position in 2*3 array
                        row = pos_idx // 3  # 0 or 1
                        col = pos_idx % 3   # 0, 1, or 2
                        x, y = pv_positions[row][col]
                        
                        log(f"Inserting Num_PV {step_name} at position {pos_idx} (row {row}, col {col})")
                        
                        apple_script += f'''
                        set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                        set position of ObjImg to {{{x}, {y}}}
                        set width of ObjImg to {pv_size[0]}
                        set height of ObjImg to {pv_size[1]}
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
    
    ########################################
    # Special case: Mass (process second)  #
    ########################################
    # Check if Mass is in selected objects
    mass_selected = any("Mass" in obj_group for obj_group in selected_objects)
    if mass_selected:
        # Filter out afterTopReco, Observable, and initial steps (Mass not available in these steps)
        valid_steps = [s for s in selected_steps if s != "afterTopReco" and s != "Observable" and s != "initial"]
        
        if valid_steps:
            # Mass: always use 2*3 layout with fixed size
            mass_positions = positions_[14]  # 2*3 : 14th index (CPV)
            mass_size = size_[3]  # 2*3 레이아웃: size_[3] (323, 313)
            
            # Get arrangement-specific width for 2*3 layout if available
            mass_width = None
            if arrangement_sizes and "2*3" in arrangement_sizes:
                mass_width = arrangement_sizes["2*3"]
            elif plot_width is not None:
                mass_width = plot_width
            
            # Apply custom plot width if provided (height will be proportional)
            if mass_width is not None:
                aspect_ratio = mass_size[1] / mass_size[0] if mass_size[0] > 0 else 1
                new_height = int(mass_width * aspect_ratio)
                mass_size = (mass_width, new_height)
            
            # Sort steps by step number (step1=1, step2=2, ..., step6=6)
            def get_step_number_mass(step_name):
                if step_name.startswith("step"):
                    try:
                        return int(step_name.replace("step", ""))
                    except:
                        return 999  # Unknown steps go to end
                else:
                    return 999
            
            sorted_steps = sorted(valid_steps, key=get_step_number_mass)
            
            # Mass: step1 → position 0, step2 → position 1, step3 → position 2, ...
            position_map = {}
            for idx, step_name in enumerate(sorted_steps):
                position_map[step_name] = idx
            
            # Split steps into slides (max 6 per slide for 2*3 layout)
            slides_steps = []
            current_slide = []
            for step_name in sorted_steps:
                if len(current_slide) < 6:
                    current_slide.append(step_name)
                else:
                    slides_steps.append(current_slide)
                    current_slide = [step_name]
            if current_slide:
                slides_steps.append(current_slide)
            
            # Create slides for Mass
            for slide_idx, slide_steps in enumerate(slides_steps):
                # Check for interruption
                if thread_instance and thread_instance.isInterruptionRequested():
                    log("Generation cancelled by user")
                    return
                if len(slide_steps) == 1:
                    title_text = f"Mass - {slide_steps[0]}"
                else:
                    steps_str = ", ".join(slide_steps)
                    title_text = f"Mass - {steps_str}"
                
                apple_script = f'''
                tell application "Keynote"
                tell document 1
                    set newSlide to make new slide
                    tell newSlide
                        -- Delete default text items (placeholder boxes) if they exist
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
                
                # Insert Mass plots for steps in this slide
                for step_name in slide_steps:
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
                    # Get position index for this step (step1=0, step2=1, step3=2, ...)
                    pos_idx = position_map.get(step_name, 0)
                    # Limit to 0-5 for 2*3 layout
                    pos_idx = min(pos_idx, 5)
                    
                    file_path, file_name = get_file_path(plot_dir, step_name, "Mass", "Mass", file_format)
                    
                    if os.path.exists(file_path):
                        # Calculate position in 2*3 array
                        row = pos_idx // 3  # 0 or 1
                        col = pos_idx % 3   # 0, 1, or 2
                        x, y = mass_positions[row][col]
                        
                        log(f"Inserting Mass {step_name} at position {pos_idx} (row {row}, col {col})")
                        
                        apple_script += f'''
                        set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                        set position of ObjImg to {{{x}, {y}}}
                        set width of ObjImg to {mass_size[0]}
                        set height of ObjImg to {mass_size[1]}
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
    
    ######################
    # Analysis step loop #
    ######################
    for st in selected_steps:
        ###############
        # Object loop #
        ###############
        for obj_group in selected_objects:
            # Get first object name to determine kinematics
            first_obj = obj_group[0]
            
            # Special case: If initial step is selected, only Num_PV is allowed (already processed above)
            if st == "initial":
                if first_obj != "Num_PV":
                    log(f"Skipping {first_obj} for initial step (only Num_PV is allowed in initial step)")
                    continue
            
            # Check if this is an Observable object
            is_observable = first_obj.startswith("O") and first_obj[1:].isdigit()
            
            # Check if this is an afterTopReco object
            after_top_reco_objects = ["Jets", "bJets", "Nu", "AnNu", "bJet", "AnbJet", "Top", "AnTop"]
            is_after_top_reco = first_obj in after_top_reco_objects
            
            # Check if this is Mass object
            is_mass = first_obj == "Mass"
            
            # Filter: Mass object - skip in afterTopReco and Observable steps
            if is_mass:
                if st == "afterTopReco" or st == "Observable":
                    log(f"Skipping {first_obj} for {st} (Mass object not available in this step)")
                    continue
            
            # Filter: Observable objects only for Observable step
            if st == "Observable":
                if not is_observable:
                    log(f"Skipping {first_obj} for Observable step (not an Observable object)")
                    continue
            else:
                # For non-Observable steps, skip Observable objects
                if is_observable:
                    log(f"Skipping {first_obj} for {st} (Observable object)")
                    continue
            
            # Filter: afterTopReco objects only for afterTopReco step
            if st == "afterTopReco":
                if not is_after_top_reco:
                    log(f"Skipping {first_obj} for afterTopReco step (not an afterTopReco object)")
                    continue
            else:
                # For non-afterTopReco steps, skip afterTopReco objects
                if is_after_top_reco:
                    log(f"Skipping {first_obj} for {st} (afterTopReco object)")
                    continue
            
            kinematics = get_kinematics(first_obj)
            
            # Determine arrangement key for custom positions
            num_kinematics = len(kinematics)
            num_objects = len(obj_group)
            arrangement_key = None
            
            log(f"Processing: {first_obj}, kinematics={num_kinematics}, objects={num_objects}")
            log(f"[DEBUG] positionOpt = {positionOpt}")
            log(f"[DEBUG] custom_positions = {custom_positions}")
            log(f"[DEBUG] arrangement_positions = {arrangement_positions}")
            
            # Check if we should use custom mode (either custom_positions has entries or arrangement_positions is provided)
            use_custom_mode = (custom_positions is not None and len(custom_positions) > 0) or (arrangement_positions is not None and len(arrangement_positions) > 0)
            
            # Build arrangement key if in custom mode
            # Use arrangement_positions if available, otherwise use positionOpt
            if use_custom_mode:
                # Build arrangement key (e.g., "2*3_Center")
                # Format: (rows * columns) where rows=objects, columns=kinematics
                if num_objects == 1:
                    # Single object: 1*N (1 row, N columns)
                    if num_kinematics == 0 or num_kinematics == 1:
                        arrangement = "1*1"
                        # Get position from arrangement_positions if available
                        if arrangement_positions and arrangement in arrangement_positions:
                            pos = arrangement_positions[arrangement]
                        else:
                            pos = positionOpt
                        arrangement_key = f"{arrangement}_{pos}"
                        log(f"[DEBUG] Single object with {num_kinematics} kinematics → arrangement_key = {arrangement_key}")
                    elif num_kinematics == 2:
                        arrangement = "1*2"
                        if arrangement_positions and arrangement in arrangement_positions:
                            pos = arrangement_positions[arrangement]
                        else:
                            pos = positionOpt
                        arrangement_key = f"{arrangement}_{pos}"
                        log(f"[DEBUG] Single object with 2 kinematics → arrangement_key = {arrangement_key}")
                    elif num_kinematics == 3:
                        arrangement = "1*3"
                        if arrangement_positions and arrangement in arrangement_positions:
                            pos = arrangement_positions[arrangement]
                        else:
                            pos = positionOpt
                        arrangement_key = f"{arrangement}_{pos}"
                        log(f"[DEBUG] Single object with 3 kinematics → arrangement_key = {arrangement_key}")
                    elif num_kinematics == 4:
                        arrangement = "1*4"
                        if arrangement_positions and arrangement in arrangement_positions:
                            pos = arrangement_positions[arrangement]
                        else:
                            pos = positionOpt
                        arrangement_key = f"{arrangement}_{pos}"
                        log(f"[DEBUG] Single object with 4 kinematics → arrangement_key = {arrangement_key}")
                    # For more kinematics (like 6), fall back to default
                elif num_objects == 2:
                    # Two objects
                    if num_kinematics == 1:
                        # Special case: Nu, AnNu and Observable (O1, O3) should use 1*2 layout (not 2*1)
                        is_observable = first_obj.startswith("O") and first_obj[1:].isdigit()
                        if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                            arrangement = "1*2"
                            if arrangement_positions and arrangement in arrangement_positions:
                                pos = arrangement_positions[arrangement]
                            else:
                                pos = positionOpt
                            arrangement_key = f"{arrangement}_{pos}"
                            log(f"[DEBUG] Nu/AnNu: Using 1*2 layout instead of 2*1 → arrangement_key = {arrangement_key}")
                        elif is_observable or (len(obj_group) == 2 and all(obj.startswith("O") and obj[1:].isdigit() for obj in obj_group)):
                            arrangement = "1*2"
                            if arrangement_positions and arrangement in arrangement_positions:
                                pos = arrangement_positions[arrangement]
                            else:
                                pos = positionOpt
                            arrangement_key = f"{arrangement}_{pos}"
                            log(f"[DEBUG] Observable: Using 1*2 layout instead of 2*1 → arrangement_key = {arrangement_key}")
                        else:
                            arrangement = "2*1"
                            if arrangement_positions and arrangement in arrangement_positions:
                                pos = arrangement_positions[arrangement]
                            else:
                                pos = positionOpt
                            arrangement_key = f"{arrangement}_{pos}"
                    elif num_kinematics == 2:
                        arrangement_key = "2*2_Center"
                    elif num_kinematics == 3:
                        arrangement_key = "2*3_Center"
                    elif num_kinematics == 4:
                        arrangement_key = "2*4_Center"
                elif num_objects == 3:
                    # Three objects
                    if num_kinematics == 2:
                        arrangement_key = "2*3_Center"
                elif num_objects == 4:
                    # Four objects
                    if num_kinematics == 2:
                        arrangement_key = "2*4_Center"
            
            # Get size and positions
            # use_custom_mode is already determined above
            
            # Get arrangement-specific width from arrangement_sizes if available
            arrangement_width = None
            if arrangement_sizes and arrangement_key:
                # Extract arrangement from key (e.g., "1*2_Upper" -> "1*2")
                arrangement_from_key = arrangement_key.split("_")[0] if "_" in arrangement_key else None
                if arrangement_from_key and arrangement_from_key in arrangement_sizes:
                    arrangement_width = arrangement_sizes[arrangement_from_key]
                    log(f"[DEBUG] Using arrangement-specific width for {arrangement_from_key}: {arrangement_width}")
            elif arrangement_sizes:
                # Try to determine arrangement from num_objects and num_kinematics
                if num_objects == 1:
                    if num_kinematics == 0 or num_kinematics == 1:
                        arrangement_from_key = "1*1"
                    elif num_kinematics == 2:
                        arrangement_from_key = "1*2"
                    elif num_kinematics == 3:
                        arrangement_from_key = "1*3"
                    elif num_kinematics == 4:
                        arrangement_from_key = "1*4"
                    else:
                        arrangement_from_key = None
                elif num_objects == 2:
                    if num_kinematics == 1:
                        # Special case: Nu, AnNu and Observable (O1, O3) should use 1*2
                        is_observable = first_obj.startswith("O") and first_obj[1:].isdigit()
                        if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                            arrangement_from_key = "1*2"
                        elif is_observable or (len(obj_group) == 2 and all(obj.startswith("O") and obj[1:].isdigit() for obj in obj_group)):
                            arrangement_from_key = "1*2"
                        else:
                            arrangement_from_key = "2*1"
                    elif num_kinematics == 2:
                        arrangement_from_key = "2*2"
                    elif num_kinematics == 3:
                        arrangement_from_key = "2*3"
                    elif num_kinematics == 4:
                        arrangement_from_key = "2*4"
                    else:
                        arrangement_from_key = None
                else:
                    arrangement_from_key = None
                
                if arrangement_from_key and arrangement_from_key in arrangement_sizes:
                    arrangement_width = arrangement_sizes[arrangement_from_key]
                    log(f"[DEBUG] Using arrangement-specific width for {arrangement_from_key}: {arrangement_width}")
            
            # Fall back to plot_width if arrangement_width is not available
            plot_width_to_use = arrangement_width if arrangement_width is not None else plot_width
            
            if use_custom_mode and arrangement_key:
                log(f"[DEBUG] Looking for custom position key: {arrangement_key}")
                log(f"[DEBUG] Available custom position keys: {list(custom_positions.keys()) if custom_positions else []}")
                
                if custom_positions and arrangement_key in custom_positions:
                    # Use custom positions from settings (user modified positions)
                    log(f"Using custom positions for {arrangement_key}")
                    positions = custom_positions[arrangement_key]
                    # Use default size or custom width (height will be proportional)
                    default_size = get_size_and_positions(num_kinematics, positionOpt=positionOpt, objects_in_group=num_objects)[0]
                    if plot_width_to_use is not None:
                        aspect_ratio = default_size[1] / default_size[0] if default_size[0] > 0 else 1
                        new_height = int(plot_width_to_use * aspect_ratio)
                        size = (plot_width_to_use, new_height)
                    else:
                        size = default_size
                else:
                    # Custom position not found, but we're in custom mode
                    # Use arrangement_positions to get the correct position from default positions
                    log(f"Custom position not found for {arrangement_key}, using arrangement_positions to get default position")
                    # Extract position from arrangement_key (e.g., "1*2_Upper" -> "Upper")
                    if "_" in arrangement_key:
                        pos_from_key = arrangement_key.split("_", 1)[1]
                    else:
                        pos_from_key = positionOpt
                    
                    # Special case: Nu, AnNu (2 objects, 1 kinematic) should use 1*2 layout
                    if num_objects == 2 and num_kinematics == 1:
                        if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                            # Use 1*2 layout for Nu & AnNu
                            size, positions = get_size_and_positions(
                                2,  # Use kinematics_count=2 to get 1*2 layout
                                positionOpt=pos_from_key,
                                objects_in_group=1  # Use objects_in_group=1 to get 1*2 layout
                            )
                            log(f"[DEBUG] Nu/AnNu: Using 1*2 layout (kinematics_count=2, objects_in_group=1)")
                        else:
                            size, positions = get_size_and_positions(
                                num_kinematics, 
                                positionOpt=pos_from_key, 
                                objects_in_group=num_objects
                            )
                    else:
                        size, positions = get_size_and_positions(
                            num_kinematics, 
                            positionOpt=pos_from_key, 
                            objects_in_group=num_objects
                        )
                    # Apply custom plot width if provided (height will be proportional)
                    if plot_width_to_use is not None:
                        # Calculate proportional height based on original aspect ratio
                        aspect_ratio = size[1] / size[0] if size[0] > 0 else 1
                        new_height = int(plot_width_to_use * aspect_ratio)
                        size = (plot_width_to_use, new_height)
                        log(f"Applied custom plot width: {plot_width_to_use}, height: {new_height} (proportional)")
            else:
                # Not in custom mode, or arrangement_key not created
                # But if arrangement_positions exists, use it to get the correct position
                pos_to_use = positionOpt
                if arrangement_positions and len(arrangement_positions) > 0:
                    # Determine arrangement and get position from arrangement_positions
                    if num_objects == 1:
                        if num_kinematics == 0 or num_kinematics == 1:
                            arrangement = "1*1"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use}")
                        elif num_kinematics == 2:
                            arrangement = "1*2"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use}")
                        elif num_kinematics == 3:
                            arrangement = "1*3"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use}")
                        elif num_kinematics == 4:
                            arrangement = "1*4"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use}")
                    elif num_objects == 2 and num_kinematics == 1:
                        # Special case: Nu, AnNu and Observable (O1, O3) should use 1*2 layout (not 2*1)
                        is_observable = first_obj.startswith("O") and first_obj[1:].isdigit()
                        if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                            arrangement = "1*2"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use} (Nu/AnNu special case)")
                        elif is_observable or (len(obj_group) == 2 and all(obj.startswith("O") and obj[1:].isdigit() for obj in obj_group)):
                            arrangement = "1*2"
                            if arrangement in arrangement_positions:
                                pos_to_use = arrangement_positions[arrangement]
                                log(f"Using arrangement_positions: {arrangement} -> {pos_to_use} (Observable special case)")
                
                # Use default positions with position from arrangement_positions if available
                if arrangement_key:
                    log(f"Custom position not found for {arrangement_key}, using default with position: {pos_to_use}")
                else:
                    log(f"No custom arrangement for {num_objects}*{num_kinematics}, using default with position: {pos_to_use}")
                
                # Special case: Nu, AnNu and Observable (O1, O3) (2 objects, 1 kinematic) should use 1*2 layout
                if num_objects == 2 and num_kinematics == 1:
                    is_observable = first_obj.startswith("O") and first_obj[1:].isdigit()
                    if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                        # Use 1*2 layout for Nu & AnNu
                        size, positions = get_size_and_positions(
                            2,  # Use kinematics_count=2 to get 1*2 layout
                            positionOpt=pos_to_use,
                            objects_in_group=1  # Use objects_in_group=1 to get 1*2 layout
                        )
                        log(f"[DEBUG] Nu/AnNu: Using 1*2 layout (kinematics_count=2, objects_in_group=1)")
                    elif is_observable or (len(obj_group) == 2 and all(obj.startswith("O") and obj[1:].isdigit() for obj in obj_group)):
                        # Use 1*2 layout for Observable (O1, O3)
                        size, positions = get_size_and_positions(
                            2,  # Use kinematics_count=2 to get 1*2 layout
                            positionOpt=pos_to_use,
                            objects_in_group=1  # Use objects_in_group=1 to get 1*2 layout
                        )
                        log(f"[DEBUG] Observable: Using 1*2 layout (kinematics_count=2, objects_in_group=1)")
                    else:
                        size, positions = get_size_and_positions(
                            num_kinematics, 
                            positionOpt=pos_to_use, 
                            objects_in_group=num_objects
                        )
                else:
                    size, positions = get_size_and_positions(
                        num_kinematics, 
                        positionOpt=pos_to_use, 
                        objects_in_group=num_objects
                    )
                # Apply custom plot width if provided (height will be proportional)
                if plot_width_to_use is not None:
                    # Calculate proportional height based on original aspect ratio
                    aspect_ratio = size[1] / size[0] if size[0] > 0 else 1
                    new_height = int(plot_width_to_use * aspect_ratio)
                    size = (plot_width_to_use, new_height)
                    log(f"Applied custom plot width: {plot_width_to_use}, height: {new_height} (proportional)")
            # Special case: Num_PV object - all steps in one slide (DRC style: 2*3 layout with step-based positioning)
            if len(obj_group) == 1 and first_obj == "Num_PV":
                # Skip Num_PV for all steps - it will be handled separately
                continue
            
            # Special case: Mass object - all steps in one slide (DRC style: 2*3 layout with step-based positioning)
            if len(obj_group) == 1 and first_obj == "Mass":
                # Skip Mass for all steps - it will be handled separately
                continue
            
            # Special case: Num_Jets object - 1*2 arrangement with Num_Jets and Num_bJets
            # Note: Num_Jets does not exist in afterTopReco step (already filtered out by 199-203 lines above)
            if len(obj_group) == 1 and first_obj == "Num_Jets":
                # Use 1*2 arrangement
                # Determine position to use (from arrangement_positions, custom_positions, or positionOpt)
                pos_to_use = positionOpt
                if arrangement_positions and "1*2" in arrangement_positions:
                    pos_to_use = arrangement_positions["1*2"]
                    log(f"[DEBUG] Num_Jets: Using arrangement_positions['1*2'] = {pos_to_use}")
                
                # Check for custom positions first
                custom_key = f"1*2_{pos_to_use}"
                if custom_positions and custom_key in custom_positions:
                    # Use custom positions
                    num_jets_positions = custom_positions[custom_key]
                    log(f"[DEBUG] Num_Jets: Using custom positions for {custom_key}")
                else:
                    # Use default positions based on pos_to_use
                    if pos_to_use == "Upper":
                        num_jets_positions = positions_[3]  # 1*2 upper
                    elif pos_to_use == "Center":
                        num_jets_positions = positions_[4]  # 1*2 center
                    elif pos_to_use == "Bottom":
                        num_jets_positions = positions_[5]  # 1*2 bottom
                    else:
                        num_jets_positions = positions_[4]  # default: center
                    log(f"[DEBUG] Num_Jets: Using default positions for 1*2_{pos_to_use}")
                
                num_jets_size = size_[1]  # 1*2 레이아웃: size_[1] (500, 485)
                
                # Get arrangement-specific width for 1*2 layout if available
                num_jets_width = None
                if arrangement_sizes and "1*2" in arrangement_sizes:
                    num_jets_width = arrangement_sizes["1*2"]
                elif plot_width is not None:
                    num_jets_width = plot_width
                
                # Apply custom plot width if provided (height will be proportional)
                if num_jets_width is not None:
                    aspect_ratio = num_jets_size[1] / num_jets_size[0] if num_jets_size[0] > 0 else 1
                    new_height = int(num_jets_width * aspect_ratio)
                    num_jets_size = (num_jets_width, new_height)
                
                # Create title text
                title_text = f"Num_Jets - {st}"
                
                apple_script = f'''
                tell application "Keynote"
                tell document 1
                    set newSlide to make new slide
                    tell newSlide
                        -- Delete default text items (placeholder boxes) if they exist
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
                
                # Extract step number for file naming
                # Note: afterTopReco is already skipped above, so we only need to check for "initial"
                if st == "initial":
                    step_num = ""
                else:
                    step_num = st.replace("step", "")
                
                # Insert h_Num_Jets_<step> (left position, col=0)
                num_jets_file_name = f"h_Num_Jets_{step_num}.{file_format}" if step_num else f"h_Num_Jets.{file_format}"
                num_jets_file_path = os.path.join(plot_dir, st, "Num_Jets", num_jets_file_name)
                
                if os.path.exists(num_jets_file_path):
                    x, y = num_jets_positions[0][0]  # 1*2: first column
                    apple_script += f'''
                    set ObjImg1 to make new image with properties {{file:((POSIX file "{num_jets_file_path}") as alias)}}
                    set position of ObjImg1 to {{{x}, {y}}}
                    set width of ObjImg1 to {num_jets_size[0]}
                    set height of ObjImg1 to {num_jets_size[1]}
                    '''
                else:
                    log(f"File not found: {num_jets_file_path}")
                
                # Insert h_Num_bJets_<step> (right position, col=1)
                num_bjets_file_name = f"h_Num_bJets_{step_num}.{file_format}" if step_num else f"h_Num_bJets.{file_format}"
                num_bjets_file_path = os.path.join(plot_dir, st, "Num_Jets", num_bjets_file_name)
                
                if os.path.exists(num_bjets_file_path):
                    x, y = num_jets_positions[0][1]  # 1*2: second column
                    apple_script += f'''
                    set ObjImg2 to make new image with properties {{file:((POSIX file "{num_bjets_file_path}") as alias)}}
                    set position of ObjImg2 to {{{x}, {y}}}
                    set width of ObjImg2 to {num_jets_size[0]}
                    set height of ObjImg2 to {num_jets_size[1]}
                    '''
                else:
                    log(f"File not found: {num_bjets_file_path}")
                
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
                continue  # Skip standard processing for Num_Jets
            
            # Create title text for the slide
            obj_names = " & ".join(obj_group)
            title_text = f"{obj_names} - {st}"
            
            # Create new slide for this object group
            apple_script = f'''
            tell application "Keynote"
            tell document 1
                set newSlide to make new slide
                tell newSlide
                    -- Delete default text items (placeholder boxes) if they exist
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
            # Special case: Observable - always insert O1 and O3
            if st == "Observable" and first_obj.startswith("O"):
                # Always use 1*2 arrangement for O1 and O3
                # Determine position to use (from arrangement_positions, custom_positions, or positionOpt)
                pos_to_use = positionOpt
                if arrangement_positions and "1*2" in arrangement_positions:
                    pos_to_use = arrangement_positions["1*2"]
                    log(f"[DEBUG] Observable: Using arrangement_positions['1*2'] = {pos_to_use}")
                
                # Check for custom positions first
                custom_key = f"1*2_{pos_to_use}"
                if custom_positions and custom_key in custom_positions:
                    # Use custom positions
                    observable_positions = custom_positions[custom_key]
                    log(f"[DEBUG] Observable: Using custom positions for {custom_key}")
                else:
                    # Use default positions based on pos_to_use
                    if pos_to_use == "Upper":
                        observable_positions = positions_[3]  # 1*2 upper
                    elif pos_to_use == "Center":
                        observable_positions = positions_[4]  # 1*2 center
                    elif pos_to_use == "Bottom":
                        observable_positions = positions_[5]  # 1*2 bottom
                    else:
                        observable_positions = positions_[4]  # default: center
                    log(f"[DEBUG] Observable: Using default positions for 1*2_{pos_to_use}")
                
                observable_size = size_[1]  # 1*2 레이아웃: size_[1] (500, 485)
                
                # Get arrangement-specific width for 1*2 layout if available
                observable_width = None
                if arrangement_sizes and "1*2" in arrangement_sizes:
                    observable_width = arrangement_sizes["1*2"]
                elif plot_width is not None:
                    observable_width = plot_width
                
                # Apply custom plot width if provided (height will be proportional)
                if observable_width is not None:
                    aspect_ratio = observable_size[1] / observable_size[0] if observable_size[0] > 0 else 1
                    new_height = int(observable_width * aspect_ratio)
                    observable_size = (observable_width, new_height)
                
                # Always insert O1 and O3
                observable_objects = ["O1", "O3"]
                for col, obj_item in enumerate(observable_objects):
                    file_path, file_name = get_file_path(plot_dir, st, obj_item, kinematics[0], file_format)
                    
                    if os.path.exists(file_path):
                        # 1*2: positions structure is [ [(x1,y1), (x2,y2)] ]
                        x, y = observable_positions[0][col]
                        apple_script += f'''
                        set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                        set position of ObjImg to {{{x}, {y}}}
                        set width of ObjImg to {observable_size[0]}
                        set height of ObjImg to {observable_size[1]}
                        '''
                    else:
                        log(f"File not found: {file_path}")
            else:
                # Standard objects processing
                # Check if this is a Num_* object - skip kinematics loop for afterTopReco step
                is_num_object = first_obj.startswith("Num_")
                if is_num_object and st == "afterTopReco":
                    # Num_* objects don't exist in afterTopReco step, skip processing
                    log(f"Skipping {first_obj} for afterTopReco step (Num_* objects not available in this step)")
                    continue
                
                for row, obj_item in enumerate(obj_group):
                    for col, kin in enumerate(kinematics):
                        file_path, file_name = get_file_path(plot_dir, st, obj_item, kin, file_format)
                        
                        if os.path.exists(file_path):
                            # Handle different position array structures
                            # Special case: Nu, AnNu (2 objects, 1 kinematic) uses 1*2 layout
                            if num_objects == 2 and num_kinematics == 1:
                                if first_obj in ["Nu", "AnNu"] or (len(obj_group) == 2 and all(obj in ["Nu", "AnNu"] for obj in obj_group)):
                                    # 1*2: positions structure is [ [(x1,y1), (x2,y2)] ]
                                    # Access: positions[0][row] (row=0 for Nu, row=1 for AnNu)
                                    x, y = positions[0][row]
                                else:
                                    # 2*1: standard [row][col] access
                                    x, y = positions[row][col]
                            elif num_objects == 1 and num_kinematics == 2:
                                # 1*2: positions structure is [ [(x1,y1), (x2,y2)] ]
                                # Access: positions[0][col]
                                x, y = positions[0][col]
                            elif num_objects == 1 and num_kinematics == 3:
                                # 1*3: positions structure is [ [(x1,y1)], [(x2,y2)], [(x3,y3)] ]
                                # Access: positions[col][0]
                                x, y = positions[col][0]
                            elif num_objects == 1 and num_kinematics == 4:
                                # 1*4: positions structure is [ [(x1,y1)], [(x2,y2)], [(x3,y3)], [(x4,y4)] ]
                                # Access: positions[col][0]
                                x, y = positions[col][0]
                            else:
                                # 2*2, 2*3, 2*4: standard [row][col] access
                                # Structure: [ [(x1,y1), (x2,y2), ...], [...], ... ]
                                x, y = positions[row][col]
                            
                            apple_script += f'''
                            set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                            set position of ObjImg to {{{x}, {y}}}
                            set width of ObjImg to {size[0]}
                            set height of ObjImg to {size[1]}
                            '''
                        else:
                            log(f"File not found: {file_path}")
            # Close slide and execute immediately
            # Close: tell newSlide (16 spaces), tell document 1 (12 spaces), tell application "Keynote" (12 spaces)
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
            if apple_script.strip().endswith("end tell"):
                subprocess.run(["osascript", "-e", apple_script])
            else:
                log("Warning: Incomplete AppleScript detected, skipping execution")
                return
            apple_script = ""  # Reset for next slide
    
    print("=" * 71)
    print("|| Inserted slides into Keynote. ---> Fill in the title and text!!!! ||")
    print("=" * 71)


##############################################
# CPV: Systematic comparison mode            #
##############################################
def insert_pdfs_into_slide_systematic(output_file, input_dir, selected_objects, selected_steps, sample_dirs, sample_names, file_format="pdf", log_callback=None, plot_width=None, thread_instance=None):
    """
    Insert systematic comparison plots (one kinematic per slide, multiple samples)
    
    Args:
        output_file (str): Path to output Keynote file
        input_dir (str): Base input directory
        selected_objects (list): List of object groups (e.g., [["Lep1", "Lep2"]])
        selected_steps (list): List of steps (e.g., ["step5"])
        sample_dirs (list): Auto-generated from categorizedMC + suffix (e.g., ["TTbar_Signal_Central_vs_Up_vs_Down"])
        sample_names (list): Display names from categorizedMC (e.g., ["TTbar Signal"])
        file_format (str): File format ("pdf" or "png")
        log_callback (callable): Optional callback function for logging messages
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    abs_output = os.path.abspath(output_file)
    
    # Systematic layout: 2×3 for 6 samples
    size = size_[3]  # 2*3 레이아웃: size_[3] (323, 313)
    # Apply custom plot width if provided (height will be proportional)
    if plot_width is not None:
        aspect_ratio = size[1] / size[0] if size[0] > 0 else 1
        new_height = int(plot_width * aspect_ratio)
        size = (plot_width, new_height)
        log(f"Applied custom plot width for systematic: {plot_width}, height: {new_height} (proportional)")
    positions = positions_[6]  # 6th index: systematic positions
    
    # Initialize AppleScript
    apple_script = f'''
    tell application "Keynote"
    activate
    set theDoc to open (POSIX file "{abs_output}")
    '''
    
    ######################
    # Analysis step loop #
    ######################
    for st in selected_steps:
        # Check for interruption
        if thread_instance and thread_instance.isInterruptionRequested():
            log("Generation cancelled by user")
            return
        
        ###############
        # Object loop #
        ###############
        for obj_group in selected_objects:
            # Check for interruption
            if thread_instance and thread_instance.isInterruptionRequested():
                log("Generation cancelled by user")
                return
            
            for obj_item in obj_group:
                # Get kinematics for this object
                kinematics = get_kinematics(obj_item)
                
                ##################
                # Kinematic loop #
                ##################
                for kin in kinematics:
                    # Check for interruption
                    if thread_instance and thread_instance.isInterruptionRequested():
                        log("Generation cancelled by user")
                        return
                    # Create new slide (no title for systematic mode)
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
                    '''
                    
                    ###############
                    # Sample loop #
                    ###############
                    for sample_idx, (sample_dir, sample_name) in enumerate(zip(sample_dirs, sample_names)):
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
                        file_path, file_name = get_file_path_systematic(
                            input_dir, sample_dir, st, obj_item, kin, file_format
                        )
                        
                        if os.path.exists(file_path):
                            # Calculate position in 2×3 array
                            row = sample_idx // 3  # 0 or 1
                            col = sample_idx % 3   # 0, 1, or 2
                            x, y = positions[row][col]
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
                    if apple_script.strip().endswith("end tell"):
                        subprocess.run(["osascript", "-e", apple_script])
                    else:
                        log("Warning: Incomplete AppleScript detected, skipping execution")
                        return
                    apple_script = ""  # Reset for next slide
    
    print("=" * 71)
    print("|| Inserted slides into Keynote. ---> Fill in the title and text!!!! ||")
    print("=" * 71)
