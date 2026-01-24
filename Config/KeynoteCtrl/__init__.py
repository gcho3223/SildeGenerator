#!/usr/bin/env python3
"""
Keynote Control Module for Keynote Slide Generator
Contains functions to control Keynote using AppleScript
"""

from Config.KeynoteCtrl.keynoteCtrl import (
    create_keynote_file, add_slide_with_title, prepare_keynote,
    insert_pdfs_into_slide, insert_pdfs_into_slide_systematic,
    insert_pdfs_into_slide_drc, insert_pdfs_into_slide_drc_resol_linearity
)
from Config.KeynoteCtrl.keynoteCtrl_usrDefined import (
    insert_pdfs_into_slide_usrDefined
)

__all__ = [
    "create_keynote_file", "add_slide_with_title", "prepare_keynote",
    "insert_pdfs_into_slide", "insert_pdfs_into_slide_systematic",
    "insert_pdfs_into_slide_drc", "insert_pdfs_into_slide_drc_resol_linearity",
    "insert_pdfs_into_slide_usrDefined"
]
