#################
# Common Configuration
#################
import os

#########################################
# Mode selection: 'cpv' or 'drc'        #
#########################################
MODE = "cpv"  # default mode

########################
# kinematics variables #
########################
kinematics = [
    ["PV"],                             # Num_PV
    ["Mass"],                           # Mass
    ["Num_Jets"],                       # Jets, bJets
    ["Pt", "Eta", "Phi"],               # Lep1,2, Jet1,2, bJet1,2
    ["PT", "Phi"],                      # MET
    ["Energy"],                         # Nu, AnNu
    ["Pt", "Eta", "Phi", "Energy"],     # bJet, AnbJet
    ["Pt", "Rapidity", "Phi", "Mass"],  # Top, AnTop
    ["Observable"],                     # O1, O3
]

#############
# plot size #
#############
size_ = [
    (314, 305), # 1*3, 2*1 레이아웃
    (500, 485), # 1*1, 1*2
    (300, 291), # 2*2 레이아웃
    (323, 313), # 2*3 레이아웃
    (250, 243), # 1*4, 2*4 레이아웃
]

##################
# plot positions #
##################
positions_ = [
    # 1*1
    [ [(262, 58)]], # upper  : 0th index
    [ [(262, 141)]], # center : 1st index
    [ [(262, 283)]], # bottom : 2nd index
    # 1*2
    [ [(20, 153), (504, 153)] ], # upper  : 3rd index
    [ [(20, 256), (504, 256)] ], # center : 4th index
    [ [(20, 452), (504, 452)] ], # bottom : 5th index
    # 1*3
    [ [(35, 79)], [(358, 79)], [(689, 79)] ], # upper  : 6th index
    [ [(35, 256)], [(358, 256)], [(689, 256)] ], # center : 7th index
    [ [(35, 452)], [(358, 452)], [(689, 452)] ], # bottom : 8th index
    # 1*4
    [ [(39, 153)], [(285, 153)], [(528, 153)], [(770, 153)] ], # upper  : 9th index
    [ [(39, 256)], [(285, 256)], [(528, 256)], [(770, 256)] ], # center : 10th index
    [ [(39, 452)], [(285, 452)], [(528, 452)], [(770, 452)] ], # bottom : 11th index
    # 2*1 : 12th index (for Nu/AnNu, etc.) - 세로로 2개
    [
        [(356, 148)],  # 위쪽
        [(356, 478)],  # 아래쪽
    ],
    # 2*2 : 13th index
    [
        [(212, 157), (538, 157)],
        [(212, 446), (538, 446)],
    ],
    # 2*3 : 14th index (CPV)
    [
        [(49, 161), (363, 161), (681, 161)],
        [(49, 455), (363, 455), (681, 455)],
    ],
    # 2*3 : 15th index (CPV - systematic)
    [
        [(63, 162), (377, 162), (692, 162)],
        [(63, 475), (377, 475), (692, 475)],
    ],
    # 2*4 : 16th index
    [
        [(39, 212), (285, 212), (528, 212), (770, 212)],
        [(39, 453), (285, 453), (528, 453), (770, 453)],
    ],
    # 2*3 : 17th index (DRC - energy plots)
    [
        [(46, 148), (356, 148), (671, 148)],
        [(46, 478), (356, 478), (671, 478)],
    ],
]

# Note: get_output_file() and get_plot_dir() functions are not used
# They were kept for potential future use but are currently unused
# Each mode (CPV/DRC) handles its own output_file and plot_dir directly
