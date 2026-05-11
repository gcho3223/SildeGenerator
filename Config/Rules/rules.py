"""
Rules Module
File path generation rules and layout matching logic

================================================================================
MODULE ORGANIZATION
================================================================================

COMMON FUNCTIONS (CPV & DRC):
    - get_kinematics()              : Map object name to kinematic variables
    - get_size_and_positions()      : Map kinematics count to plot size/position

CPV MODE FUNCTIONS:
    - get_file_path()               : Generate file path for CPV plots
    - get_file_path_systematic()    : Generate file path for CPV systematic plots

DRC MODE FUNCTIONS:
    - get_file_path_drc()           : Generate file path for DRC plots

================================================================================
"""
import sys
import os

# Try relative import first, then absolute import
try:
    from ..config import size_, positions_
except ImportError:
    from Config.config import size_, positions_


################################################################################
#                                                                              #
#                          CPV MODE FUNCTIONS                                  #
#                    (File path generation for CPV mode)                       #
#                                                                              #
################################################################################

def get_file_path(plot_dir, step_name, obj_name, kin_name, file_format="pdf"):
    """
    Generate file path for plot based on naming convention
    Args:
        plot_dir (str): Base directory for plots
        step_name (str): Step name (e.g., "step1", "initial", "Observable")
        obj_name (str): Object name (e.g., "Lep1", "MET")
        kin_name (str): Kinematic variable name (e.g., "Pt", "Eta")
        file_format (str): File format ("pdf" or "png")
    Returns:
        tuple: (file_path, file_name)
    """
    # Extract step number
    if step_name == "initial" or step_name == "afterTopReco":
        step_num = ""
    else:
        step_num = step_name.replace("step", "")
    
    # Generate file name based on object type
    if step_name == "Observable" and obj_name.startswith("O"):
        # Observable: h_Reco_CPO1_ReRange.pdf (directly in Observable directory)
        obj_num = obj_name[1:]  # "O1" -> "1"
        file_name = f"h_Reco_CPO{obj_num}_ReRange.{file_format}"
        file_path = os.path.join(plot_dir, step_name, file_name)
    elif obj_name == "Num_PV":
        # For initial step, use "initial" as suffix instead of empty string
        if step_name == "initial":
            file_name = f"h_{obj_name}_initial.{file_format}"
        else:
            file_name = f"h_{obj_name}_{step_num}.{file_format}"
        file_path = os.path.join(plot_dir, step_name, "PV", file_name)
    elif obj_name == "Num_Jets":  # Single plot: Num_Jets
        file_name = f"h_{obj_name}_{step_num}.{file_format}"
        file_path = os.path.join(plot_dir, step_name, "Num_Jets", file_name)
    elif obj_name in ["Jets", "bJets"]:  # Jets, bJets (Num_Jets, Num_bJets)
        # v1.0과 동일: Num_Jets 디렉토리 사용
        file_name = f"h_Num_{obj_name}_{step_num}.{file_format}" if step_num else f"h_Num_{obj_name}.{file_format}"
        file_path = os.path.join(plot_dir, step_name, "Num_Jets", file_name)
    elif obj_name == "Mass":
        file_name = f"h_DiLep_Mass_{step_num}.{file_format}"
        file_path = os.path.join(plot_dir, step_name, "Mass", file_name)
    else:
        # Standard format: h_<object>_<kinematic>_<step>.pdf
        if step_name == "initial" or step_name == "afterTopReco":
            file_name = f"h_{obj_name}_{kin_name.lower()}.{file_format}"
        else:
            file_name = f"h_{obj_name}_{kin_name.lower()}_{step_num}.{file_format}"
        # Use lowercase kinematic name for directory
        file_path = os.path.join(plot_dir, step_name, kin_name.lower(), file_name)
    
    return file_path, file_name


################################################################################
#                                                                              #
#                          DRC MODE FUNCTIONS                                  #
#                    (File path generation for DRC mode)                       #
#                                                                              #
################################################################################

def get_file_path_drc(base_dir, program, channel, energy):
    """
    Generate file path for DRC plots
    Args:
        base_dir (str): Base directory for plots
        program (str): DRC program name (e.g., "proton_Rot")
        channel (str): Channel name (e.g., "S_LCATTcor", "C_S_overlay_LCATTcor")
        energy (str): Energy name (e.g., "energy20_run20")
    Returns:
        tuple: (file_path, file_name)
    """
    # S/C overlay channel: directory is "C_S_overlay" and the file name
    # convention is "eDep_C_S_<s_mode>_<energy>.pdf".
    if channel.startswith("C_S_overlay"):
        # "C_S_overlay_LCATTcor" -> s_mode = "LCATTcor"
        # "C_S_overlay"          -> s_mode = "" (handled but unlikely)
        s_mode = channel[len("C_S_overlay"):].lstrip("_")
        if s_mode:
            file_name = f"eDep_C_S_{s_mode}_{energy}.pdf"
        else:
            file_name = f"eDep_C_S_{energy}.pdf"
        file_path = os.path.join(base_dir, program, "C_S_overlay", file_name)
        return file_path, file_name

    file_name = f"eDep_{channel}_{energy}.pdf"
    file_path = os.path.join(base_dir, program, channel, file_name)
    return file_path, file_name


##############################################
# CPV: Systematic comparison mode            #
##############################################
def get_file_path_systematic(input_dir, sample_dir, step_name, obj_name, kin_name, file_format="pdf"):
    """
    Generate file path for systematic comparison plots
    Uses same structure as normal CPV mode
    
    Args:
        input_dir (str): Base input directory
        sample_dir (str): Sample directory name (e.g., "TTbar_Signal_Central_vs_Up_vs_Down")
        step_name (str): Step name (e.g., "step5")
        obj_name (str): Object name (e.g., "Lep1")
        kin_name (str): Kinematic variable (e.g., "Pt")
        file_format (str): File format ("pdf" or "png")
    Returns:
        tuple: (file_path, file_name)
    """
    # Extract step number
    if step_name == "initial" or step_name == "afterTopReco":
        step_num = ""
    else:
        step_num = step_name.replace("step", "")
    
    # Generate file name based on object type (same as get_file_path)
    if step_name == "Observable" and obj_name.startswith("O"):
        obj_num = obj_name[1:]
        file_name = f"h_Reco_CPO{obj_num}_ReRange.{file_format}"
        file_path = os.path.join(input_dir, sample_dir, step_name, file_name)
    elif obj_name == "Num_PV":
        file_name = f"h_{obj_name}_{step_num}.{file_format}"
        file_path = os.path.join(input_dir, sample_dir, step_name, kin_name, file_name)
    elif obj_name in ["Jets", "bJets"]:
        file_name = f"h_Num_{obj_name}_{step_num}.{file_format}"
        file_path = os.path.join(input_dir, sample_dir, step_name, kin_name, file_name)
    elif obj_name == "Mass":
        file_name = f"h_DiLep_Mass_{step_num}.{file_format}"
        file_path = os.path.join(input_dir, sample_dir, step_name, "Mass", file_name)
    else:
        # Standard format: input_dir/sample_dir/step/kinematic/filename.pdf
        if step_name == "initial" or step_name == "afterTopReco":
            file_name = f"h_{obj_name}_{kin_name.lower()}.{file_format}"
        else:
            file_name = f"h_{obj_name}_{kin_name.lower()}_{step_num}.{file_format}"
        # Use lowercase kinematic name for directory
        file_path = os.path.join(input_dir, sample_dir, step_name, kin_name.lower(), file_name)
    
    return file_path, file_name


################################################################################
#                                                                              #
#                          COMMON FUNCTIONS                                    #
#                    (Used by both CPV and DRC modes)                          #
#                                                                              #
################################################################################

def get_kinematics(object_name):
    """
    return kinematic variables list according to the object name
    Args:
        object_name (str): object name (e.g. "Lep1", "MET", "Top" etc.)
    Returns:
        list: kinematic variables list (e.g. ["Pt", "Eta", "Phi"])
    """
    if object_name == "Num_PV":
        return ["PV"]
    elif object_name == "Mass":
        return ["Mass"]
    elif object_name == "Num_Jets":  # Single plot: Num_Jets
        return ["Num_Jets"]
    elif object_name in ["Jets", "bJets"]:  # Jets, bJets (Num_Jets, Num_bJets)
        return ["Num_Jets"]  # v1.0과 동일: ["Num_Jets"] 사용
    elif object_name in ["Jet1", "Jet2", "bJet1", "bJet2"]:  # Individual jets
        return ["Pt", "Eta", "Phi"]
    elif object_name == "MET":
        return ["Pt", "Phi"]
    elif object_name in ["Nu", "AnNu"]:  # Nu, AnNu
        return ["Energy"]
    elif object_name in ["bJet", "AnbJet"]:  # bJet, AnbJet
        return ["Pt", "Eta", "Phi", "Energy"]
    elif object_name in ["Top", "AnTop"]:  # Top, AnTop
        return ["Pt", "Rapidity", "Phi", "Mass"]
    elif object_name.startswith("O"):  # Observable (O1, O3, ...)
        return ["Observable"]
    else:  # Lep1, Lep2, Jet1, Jet2, bJet1, bJet2 etc.
        return ["Pt", "Eta", "Phi"]


###########################################
# Map kinematics count to size & position #
###########################################
def get_size_and_positions(kinematics_count, positionOpt=None, objects_in_group=1):
    """
    return size and positions of plots according to the number of kinematics and options
    Args:
        kinematics_count (int): number of kinematics variables (0~6)
        positionOpt (str, optional): position option ("Center", "Upper", "Bottom")
        objects_in_group (int, optional): number of objects in a group (default: 1)
    Returns:
        tuple: (size, positions)
            - size: (width, height) tuple
            - positions: 2D list [[row1_positions], [row2_positions], ...]
    """
    print(f"[DEBUG get_size_and_positions] kinematics_count={kinematics_count}, positionOpt={positionOpt}, objects_in_group={objects_in_group}")
    
    if kinematics_count == 0 or kinematics_count == 1:
        if objects_in_group == 1:
            # Single plot (Num_PV, Mass, Num_Jets, Energy, etc.) - 1*1 레이아웃
            size = size_[1]  # 1*1 레이아웃: size_[1] (500, 485)
            if positionOpt == "Upper":
                positions = positions_[0]  # 1*1 upper : 0th index
                print(f"[DEBUG] Using 1*1 Upper → positions_[0]")
            elif positionOpt == "Center":
                positions = positions_[1]  # 1*1 center : 1st index
                print(f"[DEBUG] Using 1*1 Center → positions_[1]")
            elif positionOpt == "Bottom":
                positions = positions_[2]  # 1*1 bottom : 2nd index
                print(f"[DEBUG] Using 1*1 Bottom → positions_[2]")
            else:
                positions = positions_[1]  # default: center
                print(f"[DEBUG] Using 1*1 default Center → positions_[1]")
        else:
            # 2*1 arrangement (multiple objects, 1 kinematic each)
            # Note: Nu/AnNu is handled specially and uses 1*2 layout instead
            size = size_[1]  # 2*1 레이아웃: size_[1] (500, 485)
            positions = positions_[12]  # 2*1 : 12th index
            print(f"[DEBUG] Using 2*1 → positions_[12]")
    elif kinematics_count == 2:
        if objects_in_group == 1:
            # 1*2 arrangement (single object, 2 kinematics) - Num_Jets, O1/O3 등
            size = size_[1]  # 1*2 레이아웃: size_[1] (500, 485)
            if positionOpt == "Upper":
                positions = positions_[3]  # 1*2 upper : 3rd index
            elif positionOpt == "Center":
                positions = positions_[4]  # 1*2 center : 4th index
            elif positionOpt == "Bottom":
                positions = positions_[5]  # 1*2 bottom : 5th index
            else:
                positions = positions_[4]  # default: center
        else:
            # 2*2 arrangement (multiple objects, 2 kinematics each)
            size = size_[2]  # 2*2 레이아웃: size_[2] (300, 291)
            positions = positions_[13]  # 2*2 : 13th index
    elif kinematics_count == 3:
        if objects_in_group == 1:
            # 1*3 arrangement (single object, 3 kinematics)
            size = size_[0]
            if positionOpt == "Upper":
                positions = positions_[6]  # 1*3 upper : 6th index
            elif positionOpt == "Center":
                positions = positions_[7]  # 1*3 center : 7th index
            elif positionOpt == "Bottom":
                positions = positions_[8]  # 1*3 bottom : 8th index
            else:
                positions = positions_[7]  # default: center
        else:
            # 2*3 arrangement (multiple objects, 3 kinematics each)
            size = size_[3]  # 2*3 레이아웃: size_[3] (323, 313)
            positions = positions_[14]  # 2*3 : 14th index (CPV)
    elif kinematics_count == 4:
        if objects_in_group == 1:
            # 1*4 arrangement (single object, 4 kinematics, e.g., Top only or AnTop only)
            size = size_[4]  # 1*4 레이아웃: size_[4] (250, 243)
            if positionOpt == "Upper":
                positions = positions_[9]  # 1*4 upper : 9th index
            elif positionOpt == "Center":
                positions = positions_[10]  # 1*4 center : 10th index
            elif positionOpt == "Bottom":
                positions = positions_[11]  # 1*4 bottom : 11th index
            else:
                positions = positions_[10]  # default: center
        else:
            # 2*4 arrangement (multiple objects, 4 kinematics each, e.g., Top & AnTop)
            size = size_[4]  # 2*4 레이아웃: size_[4] (250, 243)
            positions = positions_[16]  # 2*4 : 16th index
    elif kinematics_count == 5:
        # 2*3 arrangement (always multiple objects)
        size = size_[3]  # 2*3 레이아웃: size_[3] (323, 313)
        positions = positions_[14]  # 2*3 : 14th index (CPV)
    elif kinematics_count == 6:
        # 2*4 arrangement (always multiple objects)
        size = size_[4]  # 2*4 레이아웃: size_[4] (250, 243)
        positions = positions_[16]  # 2*4 : 16th index
    else:
        print(f"Error: Invalid number of kinematics: {kinematics_count}")
        sys.exit(1)
    
    return size, positions