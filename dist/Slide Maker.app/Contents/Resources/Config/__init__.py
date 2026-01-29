#!/usr/bin/env python3
"""
Configuration Module for Keynote Slide Generator
Contains configuration settings, file path rules, and Keynote control functions
"""

from Config.config import (
    MODE, kinematics, size_, positions_
)
from Config.config_cpv import (
    cpv_config
)
from Config.config_drc import (
    drc_config
)
from Config.Rules import (
    get_file_path, get_file_path_systematic, get_file_path_drc,
    get_kinematics, get_size_and_positions
)
from Config.KeynoteCtrl import (
    create_keynote_file, add_slide_with_title, prepare_keynote,
    insert_pdfs_into_slide, insert_pdfs_into_slide_systematic,
    insert_pdfs_into_slide_drc, insert_pdfs_into_slide_drc_resol_linearity
)

__all__ = [
    "MODE", "kinematics", "size_", "positions_",
    "cpv_config",
    "drc_config",
    "get_file_path", "get_file_path_systematic", "get_file_path_drc",
    "get_kinematics", "get_size_and_positions",
    "create_keynote_file", "add_slide_with_title", "prepare_keynote",
    "insert_pdfs_into_slide", "insert_pdfs_into_slide_systematic",
    "insert_pdfs_into_slide_drc", "insert_pdfs_into_slide_drc_resol_linearity"
]
