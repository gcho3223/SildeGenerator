"""
User-defined Template System for Dynamic Path Generation
Supports user-defined path and filename templates with variable depth
"""
import os
import re
import glob
from typing import Dict, List, Tuple, Optional, Any
from itertools import product


class TemplateParser:
    """Parse and process path/filename templates"""
    
    def __init__(self, path_template: str, filename_template: str):
        """
        Initialize template parser
        
        Args:
            path_template: e.g., "{base_dir}/{sample}/{channel}/{object}/{step}"
            filename_template: e.g., "h_{object}_{kinematic}_{step}.{format}"
        """
        self.path_template = path_template
        self.filename_template = filename_template
        
        # Extract variables from templates
        self.path_variables = self._extract_variables(path_template)
        self.filename_variables = self._extract_variables(filename_template)
        
        # All unique variables needed
        self.all_variables = list(set(self.path_variables + self.filename_variables))
    
    def _extract_variables(self, template: str) -> List[str]:
        """Extract variable names from template (e.g., {variable})"""
        pattern = r'\{([^}]+)\}'
        return re.findall(pattern, template)
    
    def build_path(self, base_dir: str, variables: Dict[str, Any]) -> str:
        """
        Build full path from template and variables
        
        Args:
            base_dir: Base directory
            variables: Dictionary of variable values
        """
        # Replace variables in path template
        path = self.path_template
        for var in self.path_variables:
            if var in variables:
                path = path.replace(f"{{{var}}}", str(variables[var]))
            else:
                # Variable not provided - remove the variable placeholder from path
                # This handles cases where object is not in template but was in old code
                path = path.replace(f"{{{var}}}/", "").replace(f"/{{{var}}}", "").replace(f"{{{var}}}", "")
        
        # Replace variables in filename template
        filename = self.filename_template
        for var in self.filename_variables:
            if var in variables:
                filename = filename.replace(f"{{{var}}}", str(variables[var]))
            else:
                # Variable not provided - this should not happen for filename, but handle it
                # Remove variable placeholder (but this might break filename structure)
                filename = filename.replace(f"{{{var}}}", "")
        
        # Combine path and filename
        # If path is empty (no subdirectories), os.path.join handles it correctly
        if path.strip():
            # Check if base_dir already ends with the path to avoid duplication
            # e.g., base_dir = "/path/to/KPS", path = "KPS" -> avoid "/path/to/KPS/KPS"
            base_dir_normalized = os.path.normpath(base_dir)
            path_normalized = os.path.normpath(path)
            
            # Check if base_dir ends with path (case-sensitive)
            if base_dir_normalized.endswith(path_normalized):
                # Remove the duplicate part from base_dir
                # e.g., "/path/to/KPS" + "KPS" -> "/path/to" + "KPS"
                base_dir_parts = base_dir_normalized.split(os.sep)
                path_parts = path_normalized.split(os.sep)
                
                # Check if last N parts of base_dir match path_parts
                if len(base_dir_parts) >= len(path_parts):
                    # Check if the last parts match
                    if base_dir_parts[-len(path_parts):] == path_parts:
                        # Remove the matching parts from base_dir
                        base_dir_normalized = os.sep.join(base_dir_parts[:-len(path_parts)])
                        if not base_dir_normalized:
                            base_dir_normalized = os.sep  # Root directory
            
            full_path = os.path.join(base_dir_normalized, path, filename)
        else:
            # Path template is empty - files are directly in base_dir
            full_path = os.path.join(base_dir, filename)
        return full_path
    
    def get_missing_variables(self, provided_variables: Dict[str, Any]) -> List[str]:
        """Get list of variables that are required but not provided"""
        missing = []
        for var in self.all_variables:
            if var not in provided_variables or provided_variables[var] is None:
                missing.append(var)
        return missing


class DynamicPathGenerator:
    """Generate file paths dynamically based on template and variable combinations"""
    
    def __init__(self, template_parser: TemplateParser):
        self.parser = template_parser
    
    def generate_paths(self, 
                      base_dir: str,
                      variable_combinations: Dict[str, List[Any]],
                      file_format: str = "pdf") -> List[Tuple[str, Dict[str, Any]]]:
        """
        Generate all possible file paths from variable combinations
        
        Args:
            base_dir: Base directory
            variable_combinations: Dict mapping variable names to lists of possible values
                e.g., {
                    "sample": ["dy", "ttbar"],
                    "channel": ["MuMu", "ee"],
                    "object": ["Lep1", "Lep2"],
                    "step": ["step1", "step2"],
                    "kinematic": ["pT", "eta"]
                }
            file_format: File format (pdf, png, etc.)
        
        Returns:
            List of tuples: (file_path, variables_used)
        """
        # Add format to variables
        variable_combinations["format"] = [file_format]
        
        # Filter to only include variables that are in the template
        relevant_vars = {
            k: v for k, v in variable_combinations.items() 
            if k in self.parser.all_variables
        }
        
        # Generate all combinations using itertools.product
        var_names = list(relevant_vars.keys())
        var_values = [relevant_vars[name] for name in var_names]
        
        results = []
        for combination in product(*var_values):
            variables = dict(zip(var_names, combination))
            file_path = self.parser.build_path(base_dir, variables)
            results.append((file_path, variables))
        
        return results
    
    def find_existing_files(self, 
                           base_dir: str,
                           variable_combinations: Dict[str, List[Any]],
                           file_format: str = "pdf") -> List[Tuple[str, Dict[str, Any]]]:
        """
        Generate paths and filter to only existing files
        
        Args:
            base_dir: Base directory
            variable_combinations: Dict mapping variable names to lists of possible values
            file_format: File format
        
        Returns:
            List of tuples: (file_path, variables_used) for existing files only
        """
        all_paths = self.generate_paths(base_dir, variable_combinations, file_format)
        existing = [(path, vars) for path, vars in all_paths if os.path.exists(path)]
        return existing
    
    def search_by_pattern(self,
                         base_dir: str,
                         filename_pattern: str,
                         max_depth: int = 5) -> List[str]:
        """
        Search for files matching pattern (fallback method)
        
        Args:
            base_dir: Base directory to search
            filename_pattern: Filename pattern with wildcards
                e.g., "h_*_pT_*.pdf" or "h_{object}_*_{step}.pdf"
            max_depth: Maximum directory depth to search
        
        Returns:
            List of matching file paths
        """
        # Convert template variables to wildcards for glob
        pattern = filename_pattern
        # Replace {variable} with * for glob matching
        pattern = re.sub(r'\{[^}]+\}', '*', pattern)
        
        search_pattern = os.path.join(base_dir, "**", pattern)
        matches = glob.glob(search_pattern, recursive=True)
        
        # Filter by depth
        filtered = []
        base_depth = len(base_dir.split(os.sep))
        for match in matches:
            match_depth = len(match.split(os.sep))
            if match_depth - base_depth <= max_depth:
                filtered.append(match)
        
        return filtered


def create_template_parser(path_template: str, filename_template: str) -> TemplateParser:
    """Factory function to create TemplateParser"""
    return TemplateParser(path_template, filename_template)


def validate_template(template: str) -> Tuple[bool, Optional[str]]:
    """
    Validate template syntax
    
    Returns:
        (is_valid, error_message)
    """
    try:
        # Check for balanced braces
        open_count = template.count('{')
        close_count = template.count('}')
        if open_count != close_count:
            return False, f"Unbalanced braces: {open_count} open, {close_count} close"
        
        # Check for empty braces
        if re.search(r'\{\s*\}', template):
            return False, "Empty braces found in template"
        
        # Check for nested braces (not supported)
        if re.search(r'\{[^}]*\{', template):
            return False, "Nested braces not supported"
        
        return True, None
    except Exception as e:
        return False, f"Template validation error: {str(e)}"


def template_to_regex(filename_template: str, file_format: str) -> Tuple[re.Pattern, List[str], Dict[Tuple[str, int], str]]:
    """
    Convert filename template to regex pattern for matching and parsing
    
    Args:
        filename_template: Filename template (e.g., "Pion_{energy}GeV_noRotation_Edist_{ch}.{format}")
        file_format: File format (e.g., "pdf", "png")
    
    Returns:
        Tuple of (compiled_regex_pattern, list_of_variable_names, var_instance_to_group)
        - compiled_regex_pattern: Compiled regex pattern for matching filenames
        - list_of_variable_names: List of unique variable names in the template
        - var_instance_to_group: Dict mapping (var_name, instance_index) to group name
          (used for handling duplicate variable names in template)
        
        Note: Callers can unpack as many values as needed:
        - regex_pattern, template_variables = template_to_regex(...)  # Gets first 2
        - regex_pattern, template_variables, var_map = template_to_regex(...)  # Gets all 3
    
    Example:
        template = "Pion_{energy}GeV_noRotation_Edist_{ch}.{format}"
        -> regex: r"Pion_(?P<energy>[^_]+)GeV_noRotation_Edist_(?P<ch>[^.]+)\.png"
    """
    # Replace {format} with actual format
    pattern = filename_template.replace("{format}", file_format)
    
    # Extract variable names (excluding format)
    # Find all variable instances first, then get unique variable names
    var_pattern = r'\{([^}]+)\}'
    all_var_matches = re.finditer(var_pattern, pattern)
    variables_with_positions = [(m.group(1), m.start()) for m in all_var_matches if m.group(1) != "format"]
    # Get unique variable names (preserving order of first occurrence)
    seen_vars = set()
    variables = []
    for var_name, _ in variables_with_positions:
        if var_name not in seen_vars:
            variables.append(var_name)
            seen_vars.add(var_name)
    
    # Build regex pattern by analyzing template structure
    # We need to determine what comes after each variable to set appropriate pattern
    # Strategy: For variables that are followed by a dot, match until dot (allows underscores)
    #           For variables that are followed by underscore, we need to match until the next literal part
    #           For variables that are followed by another variable, use non-greedy match with lookahead
    #           For variables at the end, match until dot (file extension)
    
    # Track variable instances to handle duplicate variable names
    # Map each variable instance to a unique group name
    var_instance_to_group = {}  # Maps (var_name, instance_index) -> group_name
    
    # Find all variable instances with their positions
    var_instances = []  # List of (var_name, position, instance_index, escaped_pattern)
    for var in variables:
        # Find all occurrences of this variable
        var_pos = 0
        instance_idx = 0
        while True:
            var_pos = pattern.find(f"{{{var}}}", var_pos)
            if var_pos == -1:
                break
            # Create unique group name: var_name_instance_index
            group_name = f"{var}_{instance_idx}"
            var_instance_to_group[(var, instance_idx)] = group_name
            # Store escaped pattern for this specific position
            escaped_var = re.escape(f"{{{var}}}")
            var_instances.append((var, var_pos, instance_idx, escaped_var, group_name))
            var_pos += len(f"{{{var}}}")
            instance_idx += 1
    
    # Sort by position (forward order) to process from start to end
    var_instances.sort(key=lambda x: x[1])
    
    # Build regex by directly constructing it character by character
    # This avoids issues with multiple replacements
    regex_chars = []
    i = 0
    var_idx = 0
    
    while i < len(pattern):
        # Check if we're at a variable position
        if var_idx < len(var_instances):
            var, var_pos, instance_idx, escaped_var, group_name = var_instances[var_idx]
            if i == var_pos:
                # We're at a variable - determine replacement
                after_var_pos = var_pos + len(f"{{{var}}}")
                after_var = pattern[after_var_pos:]
                
                # Determine the replacement pattern based on what follows
                if after_var.startswith('.'):
                    replacement = f"(?P<{group_name}>[^.]+)"
                elif after_var.startswith('{'):
                    replacement = f"(?P<{group_name}>[a-zA-Z0-9]+?)"
                elif after_var.startswith('_'):
                    # Variable is followed by underscore
                    # Find the next literal part (including the underscore) to use as lookahead
                    # This allows variable values to contain underscores (e.g., "S_LCATTcor")
                    next_literal_start = 0  # Include the underscore
                    next_literal = None
                    
                    # Look for the next literal part (non-variable text)
                    while next_literal_start < len(after_var):
                        if after_var[next_literal_start] == '{':
                            # Next variable found - use everything up to this point as lookahead
                            if next_literal_start > 0:
                                next_literal = after_var[:next_literal_start]
                            break
                        elif after_var[next_literal_start] not in ['_', '.']:
                            # Found a literal character - take from start (including underscore) to here + a few more chars
                            # Take enough characters to make it unique (at least 3-4 chars)
                            end_pos = min(next_literal_start + 4, len(after_var))
                            # But stop before next variable if any
                            var_pos = after_var.find('{', next_literal_start)
                            if var_pos != -1 and var_pos < end_pos:
                                end_pos = var_pos
                            next_literal = after_var[:end_pos]
                            break
                        next_literal_start += 1
                    
                    if next_literal and len(next_literal) > 0:
                        escaped_literal = re.escape(next_literal)
                        replacement = f"(?P<{group_name}>.*?)(?={escaped_literal})"
                    else:
                        # Fallback: match until next underscore or dot (won't work for values with underscore)
                        replacement = f"(?P<{group_name}>[^_]+)"
                else:
                    replacement = f"(?P<{group_name}>[^_.]+)"
                
                # Add replacement and skip the variable
                regex_chars.extend(list(replacement))
                i += len(f"{{{var}}}")
                var_idx += 1
                continue
        
        # Regular character - escape it
        char = pattern[i]
        if char in '.*+?^$[]{}|()':
            regex_chars.append('\\')
            regex_chars.append(char)
        else:
            regex_chars.append(char)
        i += 1
    
    regex_pattern = ''.join(regex_chars)
    
    # Compile regex
    try:
        compiled_regex = re.compile(regex_pattern)
    except re.error as e:
        # Debug: print the problematic pattern
        print(f"Error compiling regex: {e}")
        print(f"Pattern: {regex_pattern}")
        print(f"Pattern length: {len(regex_pattern)}")
        raise
    
    # Return regex, unique variable names (without duplicates), and mapping info
    # The mapping will be used later to merge duplicate variable instances
    unique_variables = list(set(variables))
    
    return compiled_regex, unique_variables, var_instance_to_group


def scan_and_match_files(base_dir: str,
                         filename_template: str,
                         file_format: str,
                         variable_combinations: Dict[str, List[Any]],
                         log_callback=None) -> Tuple[List[Tuple[str, Dict[str, Any], str]], Dict[str, Any]]:
    """
    Scan directory recursively and match files against filename template
    
    Args:
        base_dir: Base directory to scan
        filename_template: Filename template (e.g., "Pion_{energy}GeV_noRotation_Edist_{ch}.{format}")
        file_format: File format (pdf, png, etc.)
        variable_combinations: Dict mapping variable names to allowed values
            e.g., {"energy": ["20", "40"], "ch": ["DR", "S_C"]}
        log_callback: Optional callback function for logging
    
    Returns:
        Tuple of:
        - List of (file_path, parsed_variables, status) tuples
          status can be: "Matched", "Not found", "Ambiguous"
        - Statistics dict with counts
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)
    
    # Convert template to regex
    regex_pattern, template_variables, var_instance_to_group = template_to_regex(filename_template, file_format)
    
    log(f"Scanning directory: {base_dir}")
    log(f"Filename template: {filename_template}")
    log(f"File format: {file_format}")
    log(f"Template variables: {template_variables}")
    log(f"Allowed variable combinations: {variable_combinations}")
    log(f"Generated regex pattern: {regex_pattern.pattern}")
    
    # Scan for files with matching extension
    all_files = []
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith(f".{file_format}"):
                full_path = os.path.join(root, file)
                all_files.append(full_path)
    
    log(f"Found {len(all_files)} files with .{file_format} extension")
    
    # Check if there are consecutive variables (e.g., {obj}{kin})
    # If so, we need to handle boundary resolution using allowed values
    consecutive_vars = []
    for i, var in enumerate(template_variables):
        if i < len(template_variables) - 1:
            # Check if this variable is followed immediately by another variable
            var_pos = filename_template.find(f"{{{var}}}")
            if var_pos != -1:
                after_pos = var_pos + len(f"{{{var}}}")
                if after_pos < len(filename_template) and filename_template[after_pos] == '{':
                    # This variable is followed by another variable
                    next_var = template_variables[i+1]
                    consecutive_vars.append((var, next_var))
    
    # Match files against template
    matched_files = []  # List of (file_path, parsed_vars, status)
    file_key_to_paths = {}  # Map (var1, var2, ...) -> list of file paths
    
    for file_path in all_files:
        filename = os.path.basename(file_path)
        
        # If there are consecutive variables, we need to use a different matching strategy
        # because the regex pattern may not match correctly due to ambiguous boundaries
        if consecutive_vars:
            # For consecutive variables, extract the combined string directly
            # without relying on regex groups
            resolved = False
            parsed_vars = {}
            
            for var1, var2 in consecutive_vars:
                # Find where these variables appear in the template
                var1_pos = filename_template.find(f"{{{var1}}}")
                var2_pos = filename_template.find(f"{{{var2}}}")
                if var1_pos != -1 and var2_pos != -1 and var2_pos == var1_pos + len(f"{{{var1}}}"):
                    # They are consecutive - find the combined match in filename
                    # Get the literal parts before and after
                    before_var1 = filename_template[:var1_pos]
                    after_var2 = filename_template[var2_pos + len(f"{{{var2}}}"):]
                    
                    # Replace {format} in before/after if needed
                    before_var1 = before_var1.replace("{format}", file_format)
                    after_var2 = after_var2.replace("{format}", file_format)
                    
                    # Replace other variables in after_var2 with wildcard patterns
                    # This allows us to match even if we haven't extracted those variables yet
                    import re as re_module
                    var_pattern_in_after = r'\{([^}]+)\}'
                    for other_var_match in re_module.finditer(var_pattern_in_after, after_var2):
                        other_var = other_var_match.group(1)
                        if other_var != "format" and other_var not in [var1, var2]:
                            # Replace with a pattern that matches the variable value
                            # Use a pattern that matches until the next literal part (underscore, dot, etc.)
                            after_var2 = after_var2.replace(f"{{{other_var}}}", r"([^_\.]+)")
                    
                    # Find the position in filename
                    before_escaped = re.escape(before_var1)
                    # after_var2 may now contain regex patterns, so we need to be careful
                    # Escape literal parts but keep the regex patterns
                    after_escaped = after_var2  # Already has regex patterns if variables were replaced
                    
                    # Extract the combined string
                    pattern_combined = f"{before_escaped}(.*?){after_escaped}"
                    match_combined = re.search(pattern_combined, filename)
                    if match_combined:
                        combined_str = match_combined.group(1)
                        
                        # Try to split combined_str using allowed values
                        var1_allowed = variable_combinations.get(var1, [])
                        var2_allowed = variable_combinations.get(var2, [])
                        
                        # Try each combination of allowed values
                        found_split = False
                        for v1 in var1_allowed:
                            v1_str = str(v1).strip()
                            if combined_str.startswith(v1_str):
                                remaining = combined_str[len(v1_str):]
                                for v2 in var2_allowed:
                                    v2_str = str(v2).strip()
                                    # Check if remaining exactly matches v2
                                    if remaining == v2_str:
                                        # Found a valid split
                                        parsed_vars[var1] = v1
                                        parsed_vars[var2] = v2
                                        found_split = True
                                        resolved = True
                                        break
                                    # Also check if remaining starts with v2 and the next char is a separator
                                    elif remaining.startswith(v2_str):
                                        next_char_pos = len(v2_str)
                                        if next_char_pos < len(remaining):
                                            next_char = remaining[next_char_pos]
                                            # If next char is underscore, dot, or end of string, it's valid
                                            if next_char in ['_', '.'] or next_char_pos == len(remaining):
                                                parsed_vars[var1] = v1
                                                parsed_vars[var2] = v2
                                                found_split = True
                                                resolved = True
                                                break
                                        else:
                                            # End of string
                                            parsed_vars[var1] = v1
                                            parsed_vars[var2] = v2
                                            found_split = True
                                            resolved = True
                                            break
                                if found_split:
                                    break
                        
                        if not found_split:
                            resolved = False
                            break
                    else:
                        resolved = False
                        break
                else:
                    resolved = False
                    break
            
            # Extract remaining variables that are not part of consecutive pairs
            if resolved:
                # Build a working template by replacing already-extracted consecutive variables
                # with their actual values from the filename
                working_template = filename_template
                for var1, var2 in consecutive_vars:
                    if var1 in parsed_vars and var2 in parsed_vars:
                        combined_placeholder = f"{{{var1}}}{{{var2}}}"
                        combined_value = str(parsed_vars[var1]) + str(parsed_vars[var2])
                        working_template = working_template.replace(combined_placeholder, combined_value)
                
                # Now extract all remaining variables dynamically from the template
                remaining_vars = [var for var in template_variables if var not in parsed_vars]
                
                for var in remaining_vars:
                    # Find variable position in the working template
                    var_pos = working_template.find(f"{{{var}}}")
                    if var_pos != -1:
                        # Get literal parts before and after this variable
                        before_var = working_template[:var_pos]
                        after_var = working_template[var_pos + len(f"{{{var}}}"):]
                        
                        # Replace {format} if needed
                        before_var = before_var.replace("{format}", file_format)
                        after_var = after_var.replace("{format}", file_format)
                        
                        # Find the position in filename
                        before_escaped = re.escape(before_var)
                        after_escaped = re.escape(after_var)
                        
                        # Extract the variable value
                        pattern_var = f"{before_escaped}(.*?){after_escaped}"
                        match_var = re.search(pattern_var, filename)
                        if match_var:
                            var_value = match_var.group(1)
                            parsed_vars[var] = var_value
                
                # Also try regex as fallback for any remaining variables that weren't extracted
                match = regex_pattern.search(filename)
                if match:
                    other_vars_dict = match.groupdict()
                    # Merge duplicate variable instances (e.g., e_0, e_1 -> e)
                    other_vars = {}
                    for group_name, value in other_vars_dict.items():
                        # Extract variable name from group name (e.g., "e_0" -> "e")
                        if '_' in group_name:
                            # Check if it's a numbered instance (e.g., "e_0", "e_1")
                            parts = group_name.rsplit('_', 1)
                            if len(parts) == 2 and parts[1].isdigit():
                                var_name = parts[0]
                                # Use the first value encountered for duplicate variables
                                if var_name not in other_vars:
                                    other_vars[var_name] = value
                            else:
                                # Not a numbered instance, use as is
                                other_vars[group_name] = value
                        else:
                            # No underscore, use as is
                            other_vars[group_name] = value
                    # Merge with parsed_vars, but don't overwrite already extracted variables
                    for k, v in other_vars.items():
                        if k not in parsed_vars:
                            parsed_vars[k] = v
        else:
            # No consecutive variables - use normal regex matching
            match = regex_pattern.match(filename)
            if match:
                parsed_vars_dict = match.groupdict()
                # Merge duplicate variable instances (e.g., e_0, e_1 -> e)
                parsed_vars = {}
                for group_name, value in parsed_vars_dict.items():
                    # Extract variable name from group name (e.g., "e_0" -> "e")
                    if '_' in group_name:
                        # Check if it's a numbered instance (e.g., "e_0", "e_1")
                        parts = group_name.rsplit('_', 1)
                        if len(parts) == 2 and parts[1].isdigit():
                            var_name = parts[0]
                            # Use the first value encountered for duplicate variables
                            # (all instances should have the same value anyway)
                            if var_name not in parsed_vars:
                                parsed_vars[var_name] = value
                        else:
                            # Not a numbered instance, use as is
                            parsed_vars[group_name] = value
                    else:
                        # No underscore, use as is
                        parsed_vars[group_name] = value
                resolved = True
            else:
                resolved = False
        
        if not resolved or not parsed_vars:
            continue
        
        # Check if all parsed variables are in allowed lists
        # Use case-insensitive comparison to handle variations like "Comb" vs "COMB"
        is_valid = True
        for var_name, var_value in parsed_vars.items():
            if var_name in variable_combinations:
                allowed_values = variable_combinations[var_name]
                # Convert to strings for comparison (case-insensitive)
                parsed_value_lower = str(var_value).lower().strip()
                allowed_str_lower = [str(v).lower().strip() for v in allowed_values]
                
                # Try exact match first, then case-insensitive match
                if str(var_value).strip() not in [str(v).strip() for v in allowed_values]:
                    if parsed_value_lower not in allowed_str_lower:
                        is_valid = False
                        break
                    else:
                        # Case-insensitive match found - use the original allowed value
                        # Find the matching allowed value (preserving original case)
                        matched_allowed = None
                        for allowed_val in allowed_values:
                            if str(allowed_val).lower().strip() == parsed_value_lower:
                                matched_allowed = allowed_val
                                break
                        if matched_allowed is not None:
                            # Update parsed_vars with the original case from allowed_values
                            parsed_vars[var_name] = matched_allowed
        
        if is_valid:
                # Create key from variable values (excluding format)
                key_vars = {k: v for k, v in parsed_vars.items() if k != "format"}
                # Sort keys for consistent ordering
                key_tuple = tuple(sorted(key_vars.items()))
                
                if key_tuple not in file_key_to_paths:
                    file_key_to_paths[key_tuple] = []
                file_key_to_paths[key_tuple].append((file_path, parsed_vars))
    
    # Determine status for each unique variable combination
    # Generate all expected combinations from variable_combinations
    from itertools import product
    
    expected_combinations = []
    if template_variables:
        var_names = [v for v in template_variables if v in variable_combinations]
        var_value_lists = [variable_combinations[v] for v in var_names]
        for combo in product(*var_value_lists):
            combo_dict = dict(zip(var_names, combo))
            expected_combinations.append(combo_dict)
    
    # Build result list
    result_files = []
    matched_keys = set()
    
    # Process matched files
    for key_tuple, paths in file_key_to_paths.items():
        if len(paths) == 1:
            # Unique match
            file_path, parsed_vars = paths[0]
            result_files.append((file_path, parsed_vars, "Matched"))
            matched_keys.add(key_tuple)
        else:
            # Ambiguous: multiple files for same variable combination
            for file_path, parsed_vars in paths:
                result_files.append((file_path, parsed_vars, "Ambiguous"))
            matched_keys.add(key_tuple)
    
    # Mark missing combinations as "Not found"
    # Use case-insensitive comparison for key matching
    matched_keys_normalized = set()
    for key_tuple in matched_keys:
        # Normalize keys to lowercase for comparison
        normalized_key = tuple(sorted((k, str(v).lower().strip()) for k, v in key_tuple))
        matched_keys_normalized.add(normalized_key)
    
    for combo_dict in expected_combinations:
        key_tuple = tuple(sorted(combo_dict.items()))
        # Normalize for comparison
        normalized_key = tuple(sorted((k, str(v).lower().strip()) for k, v in key_tuple))
        if normalized_key not in matched_keys_normalized:
            # Create a dummy entry with None path
            result_files.append((None, combo_dict, "Not found"))
    
    # Calculate statistics
    stats = {
        "total_scanned": len(all_files),
        "matched_count": len([r for r in result_files if r[2] == "Matched"]),
        "not_found_count": len([r for r in result_files if r[2] == "Not found"]),
        "ambiguous_count": len([r for r in result_files if r[2] == "Ambiguous"]),
        "ambiguous_groups": {}
    }
    
    # Group ambiguous files by variable combination
    ambiguous_by_key = {}
    for file_path, parsed_vars, status in result_files:
        if status == "Ambiguous":
            key_vars = {k: v for k, v in parsed_vars.items() if k != "format"}
            key_tuple = tuple(sorted(key_vars.items()))
            if key_tuple not in ambiguous_by_key:
                ambiguous_by_key[key_tuple] = []
            ambiguous_by_key[key_tuple].append(file_path)
    
    stats["ambiguous_groups"] = ambiguous_by_key
    
    return result_files, stats
