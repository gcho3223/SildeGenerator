#################
# CPV Mode Configuration
#################
import os

#########################################
# CPV Mode Configuration                #
#########################################
cpv_config = {
    "output_file": "/Users/gcho/Desktop/ttbar_cpv.key",
    "input_dir": "/Users/gcho/Desktop/overlay",
    "runPeriod": "UL2018",
    "channel": "MuMu",
    "systematic_suffix": "_Central_vs_Up_vs_Down",  # Suffix for systematic sample directories
    # Display names for slide titles and sample identification
    "categorizedMC": [
        "TTbar_Signal", "DY", "SingleTop",
        "TTbarOther", "TTV", "Diboson"
    ],
}
# Auto-generate systematic_samples from categorizedMC
cpv_config["systematic_samples"] = [
    name + cpv_config["systematic_suffix"] 
    for name in cpv_config["categorizedMC"]
]