import os
import subprocess
import shutil

#################
# Configuration #
#################
template_file = "./KinematicTemplate.key"
output_file = "../../../Kinematics.key"

base_dir = "/Users/gcho/Desktop/"
plot_dir = base_dir + "UL2016PreVFP/MuMu/"

step = ["step0", "step1", "step2", "step3", "step4", "step5", "step6"]
kinematics = ["pT", "eta", "phi"]
objects = [
    ["Jet1", "Jet2"],
    ["Lep1", "Lep2"],
]

def prepare_keynote(template_file, output_file):
    shutil.copy(template_file, output_file)
    print(f"Copied template: {template_file} → {output_file}")

def insert_pdfs_into_slide(output_file, plot_dir):
    ##################################################
    # 1. set up the plot size and positions as 2 * 3 #
    ##################################################
    size = (323, 313)  # (Width, Height)
    positions = [
        [(44, 166), (363, 166), (681, 166)],  # Jet1 or Lep1
        [(44, 452), (363, 452), (681, 452)]   # Jet2 or Lep2
    ]
    ##################################################################
    # 2. set up AppleScript to control Keynote: open the output file #
    ##################################################################
    abs_output = os.path.abspath(output_file)  # 출력 파일의 절대경로 사용
    apple_script = f'''
    tell application "Keynote"
    activate
    set theDoc to open (POSIX file "{abs_output}")
    '''
    ######################
    # Analysis step loop #
    ######################
    for st in step:
        step_num = st.replace("step", "")
        ###############
        # Object loop #
        ###############
        for obj in objects:
            ##############################################
            # 3. make a new slide with Kinematics layout #
            ##############################################
            apple_script += '''
            tell theDoc
            set ObjectSlide to make new slide at the end of theDoc with properties {base slide:(first master slide whose name is "Kinematics")}
            tell ObjectSlide
            '''
            for row, obj in enumerate(obj):
                for col, kin in enumerate(kinematics):
                    ######################################################
                    # file name format: h_<object><kinematic>_<step>.pdf #
                    # example: h_Jet1pt_0.pdf                            #
                    ######################################################
                    file_name = f"h_{obj}{kin.lower()}_{step_num}.pdf"
                    file_path = os.path.join(plot_dir, st, kin, file_name) ### plot file path ###
                    # print(file_path) # for debugging
                    if os.path.exists(file_path):
                        x, y = positions[row][col]
                        apple_script += f'''
                        set ObjImg to make new image with properties {{file:((POSIX file "{file_path}") as alias)}}
                        set position of ObjImg to {{{x}, {y}}}
                        set height of ObjImg to {size[1]}
                        set width of ObjImg to {size[0]}
                        '''
                    else:
                        print(f"File not found: {file_path}")
            ### close apple script for setting up the plots ###
            apple_script += '''
                end tell
            end tell
            '''
    ### close apple script for making slides ###
    apple_script += '''
    end tell
    '''
    
    #####################################################
    # Execute AppleScript to insert slides into Keynote #
    #####################################################
    subprocess.run(["osascript", "-e", apple_script])
    print("=======================================================================")
    print("|| Inserted slides into Keynote. ---> Fill in the title and text!!!! ||")
    print("=======================================================================")

### copy template and save as new file ###
prepare_keynote(template_file, output_file)

### insert plots into slides ###
insert_pdfs_into_slide(output_file, plot_dir)