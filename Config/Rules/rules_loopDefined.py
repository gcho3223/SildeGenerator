"""
Loop-defined Mode Rules - Dynamic path generation for user-defined templates
"""
import os
from typing import Dict, List, Tuple, Optional, Any
from Config.Rules.loopDefined_template_engine import (
    TemplateParser, DynamicPathGenerator, 
    create_template_parser, validate_template,
    scan_and_match_files
)


def get_loopDefined_file_path(base_dir: str,
                        path_template: str,
                        filename_template: str,
                        variables: Dict[str, any],
                        file_format: str = "pdf") -> Optional[str]:
    """
    Generate file path from user-defined templates
    
    Args:
        base_dir: Base directory
        path_template: Path template (e.g., "{sample}/{channel}/{object}/{step}")
        filename_template: Filename template (e.g., "h_{object}_{kinematic}_{step}.{format}")
        variables: Dictionary of variable values
        file_format: File format
    
    Returns:
        Full file path if file exists, None otherwise
    """
    # Validate templates
    path_valid, path_error = validate_template(path_template)
    if not path_valid:
        raise ValueError(f"Invalid path template: {path_error}")
    
    filename_valid, filename_error = validate_template(filename_template)
    if not filename_valid:
        raise ValueError(f"Invalid filename template: {filename_error}")
    
    # Create parser and generate path
    parser = create_template_parser(path_template, filename_template)
    variables["format"] = file_format
    file_path = parser.build_path(base_dir, variables)
    
    # Check if file exists
    if os.path.exists(file_path):
        return file_path
    
    return None


def find_loopDefined_files(base_dir: str,
                     path_template: str,
                     filename_template: str,
                     variable_combinations: Dict[str, List[any]],
                     file_format: str = "pdf") -> List[Tuple[str, Dict[str, any]]]:
    """
    Find all files matching user-defined templates
    
    Args:
        base_dir: Base directory
        path_template: Path template
        filename_template: Filename template
        variable_combinations: Dict mapping variable names to lists of possible values
        file_format: File format
    
    Returns:
        List of (file_path, variables_used) tuples for existing files
    """
    # Validate templates
    path_valid, path_error = validate_template(path_template)
    if not path_valid:
        raise ValueError(f"Invalid path template: {path_error}")
    
    filename_valid, filename_error = validate_template(filename_template)
    if not filename_valid:
        raise ValueError(f"Invalid filename template: {filename_error}")
    
    # Create parser and generator
    parser = create_template_parser(path_template, filename_template)
    generator = DynamicPathGenerator(parser)
    
    # Find existing files
    existing_files = generator.find_existing_files(
        base_dir, variable_combinations, file_format
    )
    
    return existing_files


def get_loopDefined_file_path_with_fallback(base_dir: str,
                                      path_template: str,
                                      filename_template: str,
                                      variables: Dict[str, any],
                                      file_format: str = "pdf",
                                      search_depth: int = 5) -> Optional[str]:
    """
    Try template path first, then fallback to pattern search
    
    Args:
        base_dir: Base directory
        path_template: Path template
        filename_template: Filename template
        variables: Dictionary of variable values
        file_format: File format
        search_depth: Max depth for pattern search fallback
    
    Returns:
        Full file path if found, None otherwise
    """
    # Try template path first
    file_path = get_loopDefined_file_path(
        base_dir, path_template, filename_template, variables, file_format
    )
    
    if file_path:
        return file_path
    
    # Fallback: Pattern search
    parser = create_template_parser(path_template, filename_template)
    generator = DynamicPathGenerator(parser)
    
    # Convert filename template to search pattern
    search_pattern = filename_template
    for var_name in parser.filename_variables:
        if var_name in variables:
            search_pattern = search_pattern.replace(
                f"{{{var_name}}}", str(variables[var_name])
            )
        else:
            search_pattern = search_pattern.replace(f"{{{var_name}}}", "*")
    
    matches = generator.search_by_pattern(base_dir, search_pattern, search_depth)
    
    if matches:
        return matches[0]  # Return first match
    
    return None


def find_loopDefined_files_scan(base_dir: str,
                                filename_template: str,
                                variable_combinations: Dict[str, List[any]],
                                file_format: str = "pdf",
                                log_callback=None,
                                path_template: str = "") -> Tuple[List[Tuple[str, Dict[str, any]]], Dict[str, any], List[Tuple[str, Dict[str, Any], str]]]:
    """
    Find files using scan mode: recursively scan directory and match against filename template and path template
    
    Args:
        base_dir: Base directory to scan
        filename_template: Filename template (e.g., "Pion_{energy}GeV_noRotation_Edist_{ch}.{format}")
        variable_combinations: Dict mapping variable names to allowed values
            e.g., {"energy": ["20", "40"], "ch": ["DR", "S_C"]}
        file_format: File format (pdf, png, etc.)
        log_callback: Optional callback function for logging
        path_template: Optional path template (e.g., "{e}GeV/TotalEdep_SF/{channel}")
    
    Returns:
        Tuple of:
        - List of (file_path, variables_used) tuples for matched files only
        - Statistics dict with scan results
        - List of (file_path, parsed_variables, status) tuples for all results
    """
    # Validate filename template
    filename_valid, filename_error = validate_template(filename_template)
    if not filename_valid:
        raise ValueError(f"Invalid filename template: {filename_error}")
    
    # Scan and match files (including path template)
    matched_results, stats = scan_and_match_files(
        base_dir, filename_template, file_format, variable_combinations, log_callback, path_template
    )
    
    # Filter to only matched files (exclude "Not found" and "Ambiguous")
    matched_files = [
        (file_path, parsed_vars) 
        for file_path, parsed_vars, status in matched_results 
        if status == "Matched" and file_path is not None
    ]
    
    return matched_files, stats, matched_results
