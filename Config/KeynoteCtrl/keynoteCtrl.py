"""
Keynote Control Module - Common Functions
This module provides common functions to control Keynote using AppleScript
Used by both CPV and DRC modes

================================================================================
MODULE ORGANIZATION
================================================================================

COMMON FUNCTIONS (CPV & DRC):
    - create_keynote_file()       : Create new Keynote file
    - add_slide_with_title()      : Add slide with title
    - prepare_keynote()           : Alias for create_keynote_file()

CPV MODE FUNCTIONS:
    See: keynoteCtrl_cpv.py
    - insert_pdfs_into_slide()              : Insert CPV normal mode plots
    - insert_pdfs_into_slide_systematic()   : Insert CPV systematic comparison plots

DRC MODE FUNCTIONS:
    See: keynoteCtrl_drc.py
    - insert_pdfs_into_slide_drc()                      : Insert DRC energy plots
    - insert_pdfs_into_slide_drc_resol_linearity()     : Insert DRC resolution/linearity plots

================================================================================
"""
import os
import subprocess
import sys

# Import CPV and DRC mode functions for backward compatibility
try:
    from .keynoteCtrl_cpv import insert_pdfs_into_slide, insert_pdfs_into_slide_systematic
    from .keynoteCtrl_drc import insert_pdfs_into_slide_drc, insert_pdfs_into_slide_drc_resol_linearity
except ImportError:
    from Config.KeynoteCtrl.keynoteCtrl_cpv import insert_pdfs_into_slide, insert_pdfs_into_slide_systematic
    from Config.KeynoteCtrl.keynoteCtrl_drc import insert_pdfs_into_slide_drc, insert_pdfs_into_slide_drc_resol_linearity


################################################################################
#                                                                              #
#                          COMMON FUNCTIONS                                    #
#                    (Used by both CPV and DRC modes)                          #
#                                                                              #
################################################################################

def create_keynote_file(output_file, theme="White"):
    """
    Create a new Keynote file without template
    Args:
        output_file (str): Path to output Keynote file
        theme (str): Keynote theme name (default: "White")
    Returns:
        str: Absolute path to the created Keynote file
    """
    abs_output_path = os.path.abspath(output_file)
    
    # Check if output file exists and delete it
    if os.path.exists(abs_output_path):
        print(f"Existing output file found. Deleting: {abs_output_path}")
        try:
            os.remove(abs_output_path)
            print("File deleted successfully.")
        except PermissionError:
            print("Permission denied. Keynote might be using the file.")
            print("Please close Keynote and try again.")
            sys.exit(1)
        except Exception as e:
            print(f"Error deleting file: {e}")
            sys.exit(1)
    
    # Create new Keynote document using AppleScript
    apple_script = f'''
    tell application "Keynote"
        activate
        set theDoc to make new document with properties {{document theme:theme "{theme}"}}
        
        -- Delete the default first slide
        tell theDoc
            if (count of slides) > 0 then
                delete slide 1
            end if
        end tell
        -- Save the document
        save theDoc in POSIX file "{abs_output_path}"
    end tell
    '''
    
    subprocess.run(["osascript", "-e", apple_script])
    print(f"Created new Keynote file: {abs_output_path}")
    
    return abs_output_path


def add_slide_with_title(keynote_file, title):
    """
    Add a new slide with title to Keynote file
    Args:
        keynote_file (str): Path to Keynote file
        title (str): Title text for the slide
    """
    abs_file_path = os.path.abspath(keynote_file)
    
    apple_script = f'''
    tell application "Keynote"
        open POSIX file "{abs_file_path}"
        tell front document
            set newSlide to make new slide with properties {{base layout:layout "Blank"}}
            tell newSlide
                set body text to "{title}"
            end tell
        end tell
        save front document
    end tell
    '''
    
    subprocess.run(["osascript", "-e", apple_script])


# Alias for backward compatibility
prepare_keynote = create_keynote_file
