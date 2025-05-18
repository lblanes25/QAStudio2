"""
Excel Formula Utilities for QA Analytics Framework.

This module provides utility functions for working with Excel formulas in the
QA Analytics framework. It includes formula conversion, validation, parsing,
and error handling utilities to support the ExcelFormulaProcessor.

These utilities make it easier to work with Excel formulas throughout the framework,
providing consistent handling of formula syntax, references, and errors.
"""

import re
from typing import Dict, List, Set, Tuple, Optional, Any
import pandas as pd

# Set up logging
from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()

# Excel error codes and messages
EXCEL_ERROR_CODES = {
    "#NULL!": "You specified an invalid intersection of two ranges",
    "#DIV/0!": "Division by zero",
    "#VALUE!": "Wrong type of operand or function argument",
    "#REF!": "Invalid cell reference",
    "#NAME?": "Excel doesn't recognize a name",
    "#NUM!": "Invalid number in a formula or function",
    "#N/A": "Value not available to a function or formula"
}

# Common Excel function translations for documentation
EXCEL_FUNCTION_DESCRIPTIONS = {
    "SUM": "Add values",
    "AVERAGE": "Calculate the average of values",
    "COUNT": "Count numeric values",
    "COUNTA": "Count non-empty values",
    "MAX": "Find the maximum value",
    "MIN": "Find the minimum value",
    "IF": "Conditional logic",
    "AND": "Logical AND operation",
    "OR": "Logical OR operation",
    "NOT": "Logical NOT operation",
    "VLOOKUP": "Vertical lookup in a table",
    "HLOOKUP": "Horizontal lookup in a table",
    "INDEX": "Get value at position in range",
    "MATCH": "Find position in range",
    "IFERROR": "Handle errors in formulas",
    "ISBLANK": "Check if value is blank",
    "ISTEXT": "Check if value is text",
    "ISNUMBER": "Check if value is number",
    "TODAY": "Current date",
    "NOW": "Current date and time",
    "DATE": "Create a date value",
    "DATEVALUE": "Convert text to date value",
    "TEXT": "Format value as text",
    "LEFT": "Extract characters from the left",
    "RIGHT": "Extract characters from the right",
    "MID": "Extract characters from the middle",
    "LEN": "Get text length",
    "FIND": "Find text position (case-sensitive)",
    "SEARCH": "Find text position (not case-sensitive)",
    "TRIM": "Remove extra spaces",
    "UPPER": "Convert to uppercase",
    "LOWER": "Convert to lowercase",
    "PROPER": "Convert to proper case"
}


def is_valid_excel_formula(formula: str) -> bool:
    """
    Check if a string is a valid Excel formula syntax (without executing it).
    
    This performs basic syntax validation, checking for:
    - Starts with equals sign
    - Balanced parentheses and quotation marks
    - Valid cell references
    
    Args:
        formula: Excel formula to validate
        
    Returns:
        bool: True if the formula appears to be valid syntax
    """
    # Empty formulas are not valid
    if not formula or not formula.strip():
        return False
    
    # Formulas should start with equals sign
    if not formula.strip().startswith("="):
        return False
    
    # Remove equals sign for further checks
    formula_content = formula.strip()[1:]
    
    # Check for balanced parentheses
    open_parens = formula_content.count("(")
    close_parens = formula_content.count(")")
    if open_parens != close_parens:
        return False
    
    # Check for balanced quotation marks (ignoring escaped quotes)
    in_quote = False
    for i, char in enumerate(formula_content):
        if char == '"' and (i == 0 or formula_content[i-1] != '\\'):
            in_quote = not in_quote
    
    if in_quote:
        # Unclosed quote
        return False
    
    # Check for obvious errors like empty functions
    if re.search(r'\(\s*\)', formula_content):
        return False
    
    # Additional validation could be done here...
    
    # If no issues found, formula seems valid
    return True


def extract_cell_references(formula: str) -> List[str]:
    """
    Extract all cell references from an Excel formula.
    
    This function handles:
    - A1-style references (e.g., A1, $B$2, C$3)
    - Range references (e.g., A1:B10)
    - Named references are not extracted as they are not direct cell references
    
    Args:
        formula: Excel formula to analyze
        
    Returns:
        List of cell references found in the formula
    """
    # Remove string literals as they might contain patterns that look like references
    formula_without_strings = remove_string_literals(formula)
    
    # Regex pattern for A1-style references, including absolute references with $
    cell_pattern = r'(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,7})(?![A-Za-z0-9_])'
    
    # Regex pattern for range references (e.g., A1:B10)
    range_pattern = r'(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,7}:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,7})'
    
    # Find all cell and range references
    cell_refs = re.findall(cell_pattern, formula_without_strings)
    range_refs = re.findall(range_pattern, formula_without_strings)
    
    # Combine and deduplicate
    all_refs = list(set(cell_refs + range_refs))
    
    return all_refs


def remove_string_literals(formula: str) -> str:
    """
    Remove string literals from a formula to help with parsing.
    
    This replaces string literals with placeholders to prevent false positives
    when parsing cell references or function names.
    
    Args:
        formula: Excel formula to process
        
    Returns:
        Formula with string literals replaced by placeholders
    """
    result = ""
    in_string = False
    escape_next = False
    
    for char in formula:
        if char == '"' and not escape_next:
            in_string = not in_string
            result += char  # Keep quotes in result
        elif in_string:
            if char == '\\':
                escape_next = True
            else:
                escape_next = False
            result += '_'  # Replace string content with underscore
        else:
            result += char
    
    return result


def extract_column_names(formula: str) -> Set[str]:
    """
    Extract column names from a formula assuming they're used directly.
    
    This function handles column names that:
    - Are used directly in the formula
    - Are enclosed in square brackets (Excel's non-alphanumeric field syntax)
    - Have back-ticks around them (alternate syntax for fields with spaces)
    
    Args:
        formula: Excel formula to analyze
        
    Returns:
        Set of column names found in the formula
    """
    column_names = set()
    
    # Remove string literals first to avoid false positives
    formula_without_strings = remove_string_literals(formula)
    
    # Match column names in brackets (Excel's syntax for field names, especially with spaces)
    # e.g., [Column Name] or [Column_Name]
    bracket_pattern = r'\[([^\[\]]+)\]'
    for match in re.finditer(bracket_pattern, formula_without_strings):
        column_names.add(match.group(1))
    
    # Match column names in back-ticks (alternate syntax for fields with spaces)
    # e.g., `Column Name` or `Column_Name`
    backtick_pattern = r'`([^`]+)`'
    for match in re.finditer(backtick_pattern, formula_without_strings):
        column_names.add(match.group(1))
    
    # Try to match direct column name references
    # This is more complex and may have false positives
    # We need to remove Excel functions and operators first
    
    # First, get a version without known Excel functions
    clean_formula = formula_without_strings
    
    # Remove common Excel functions
    function_pattern = r'\b([A-Z][A-Za-z0-9\.]+)\s*\('
    functions = re.findall(function_pattern, clean_formula)
    
    for func in functions:
        clean_formula = re.sub(r'\b' + re.escape(func) + r'\s*\(', 'FUNC(', clean_formula)
    
    # Remove cell references
    cell_refs = extract_cell_references(clean_formula)
    for ref in cell_refs:
        clean_formula = clean_formula.replace(ref, "CELL")
    
    # Now try to identify potential column names (words not adjacent to parentheses)
    # This approach isn't perfect and might need refinement for specific cases
    word_pattern = r'\b([A-Za-z][A-Za-z0-9_]*)\b'
    potential_columns = re.findall(word_pattern, clean_formula)
    
    # Filter out obvious non-column names
    excluded_words = {
        # Excel operators and constants
        "TRUE", "FALSE", "NULL", "NA", "PI", "AND", "OR", "NOT", "IF",
        "THEN", "ELSE", "FUNC", "CELL", "ERROR", 
    }
    
    for word in potential_columns:
        if (word not in excluded_words and 
            not word.upper() in EXCEL_FUNCTION_DESCRIPTIONS.keys()):
            column_names.add(word)
    
    return column_names


def column_index_to_letter(index: int) -> str:
    """
    Convert a column index to Excel column letter (1=A, 2=B, etc.).
    
    Args:
        index: 1-based column index
        
    Returns:
        Excel column letter(s)
    """
    if index < 1:
        raise ValueError("Column index must be positive")
    
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    
    return result


def column_letter_to_index(column_letter: str) -> int:
    """
    Convert Excel column letter to index (A=1, B=2, etc.).
    
    Args:
        column_letter: Excel column letter(s)
        
    Returns:
        1-based column index
    """
    column_letter = column_letter.upper()
    result = 0
    
    for char in column_letter:
        result = result * 26 + (ord(char) - 64)
    
    return result


def a1_to_rc(a1_ref: str, row_offset: int = 0, col_offset: int = 0) -> str:
    """
    Convert A1-style reference to R1C1-style reference.
    
    Args:
        a1_ref: A1-style reference (e.g., A1, $B$2)
        row_offset: Row offset for relative references
        col_offset: Column offset for relative references
        
    Returns:
        R1C1-style reference
    """
    # Handle range references (e.g., A1:B10)
    if ":" in a1_ref:
        start, end = a1_ref.split(":")
        return f"{a1_to_rc(start, row_offset, col_offset)}:{a1_to_rc(end, row_offset, col_offset)}"
    
    # Extract column and row parts, handling absolute references
    match = re.match(r'(\$?)([A-Za-z]+)(\$?)([1-9][0-9]*)', a1_ref)
    if not match:
        return a1_ref  # Return as-is if not a valid A1 reference
        
    col_abs, col_str, row_abs, row_str = match.groups()
    col_idx = column_letter_to_index(col_str)
    row_idx = int(row_str)
    
    # Create R1C1 reference
    r_part = f"R{row_idx}" if row_abs else f"R[{row_idx - row_offset}]"
    c_part = f"C{col_idx}" if col_abs else f"C[{col_idx - col_offset}]"
    
    return f"{r_part}{c_part}"


def rc_to_a1(rc_ref: str, row_pos: int = 1, col_pos: int = 1) -> str:
    """
    Convert R1C1-style reference to A1-style reference.
    
    Args:
        rc_ref: R1C1-style reference (e.g., R1C1, R[-1]C[2])
        row_pos: Current row position for relative references
        col_pos: Current column position for relative references
        
    Returns:
        A1-style reference
    """
    # Handle range references
    if ":" in rc_ref:
        start, end = rc_ref.split(":")
        return f"{rc_to_a1(start, row_pos, col_pos)}:{rc_to_a1(end, row_pos, col_pos)}"
    
    # Extract row and column parts
    r_match = re.search(r'R(\[([+-]?\d+)\]|(\d+))', rc_ref)
    c_match = re.search(r'C(\[([+-]?\d+)\]|(\d+))', rc_ref)
    
    if not r_match or not c_match:
        return rc_ref  # Return as-is if not a valid R1C1 reference
    
    # Process row part
    r_rel, r_abs = r_match.group(2), r_match.group(3)
    if r_rel:  # Relative reference [n]
        row_idx = row_pos + int(r_rel)
        row_abs = ""
    else:  # Absolute reference
        row_idx = int(r_abs)
        row_abs = "$"
    
    # Process column part
    c_rel, c_abs = c_match.group(2), c_match.group(3)
    if c_rel:  # Relative reference [n]
        col_idx = col_pos + int(c_rel)
        col_abs = ""
    else:  # Absolute reference
        col_idx = int(c_abs)
        col_abs = "$"
    
    col_str = column_index_to_letter(col_idx)
    return f"{col_abs}{col_str}{row_abs}{row_idx}"


def convert_formula_to_rc(formula: str, row: int = 1, col: int = 1) -> str:
    """
    Convert an A1-style formula to R1C1-style.
    
    Args:
        formula: Excel formula with A1-style references
        row: Row position for conversion context
        col: Column position for conversion context
        
    Returns:
        Formula with R1C1-style references
    """
    if not formula.startswith("="):
        return formula
    
    # Remove string literals, as we don't want to modify text in quotes
    literals = {}
    formula_no_strings = extract_string_literals(formula, literals)
    
    # Find all A1 references
    a1_refs = extract_cell_references(formula_no_strings)
    
    # Sort by length (descending) to avoid replacing parts of longer references
    a1_refs.sort(key=len, reverse=True)
    
    # Replace each A1 reference with its R1C1 equivalent
    result = formula_no_strings
    for a1_ref in a1_refs:
        rc_ref = a1_to_rc(a1_ref, row - 1, col - 1)
        # Use word boundaries to avoid partial replacements
        result = re.sub(r'\b' + re.escape(a1_ref) + r'\b', rc_ref, result)
    
    # Restore string literals
    result = restore_string_literals(result, literals)
    
    return result


def convert_formula_to_a1(formula: str, row: int = 1, col: int = 1) -> str:
    """
    Convert an R1C1-style formula to A1-style.
    
    Args:
        formula: Excel formula with R1C1-style references
        row: Row position for conversion context
        col: Column position for conversion context
        
    Returns:
        Formula with A1-style references
    """
    if not formula.startswith("="):
        return formula
    
    # Remove string literals, as we don't want to modify text in quotes
    literals = {}
    formula_no_strings = extract_string_literals(formula, literals)
    
    # Find all R1C1 references
    # Pattern for R1C1 references
    rc_pattern = r'R(\[([+-]?\d+)\]|(\d+))C(\[([+-]?\d+)\]|(\d+))'
    rc_refs = re.findall(rc_pattern, formula_no_strings)
    
    # Replace each R1C1 reference with its A1 equivalent
    result = formula_no_strings
    for rc_match in rc_refs:
        r_rel, r_abs, c_rel, c_abs = rc_match[1], rc_match[2], rc_match[4], rc_match[5]
        
        # Reconstruct the original R1C1 reference
        if r_rel:
            r_part = f"R[{r_rel}]"
        else:
            r_part = f"R{r_abs}"
            
        if c_rel:
            c_part = f"C[{c_rel}]"
        else:
            c_part = f"C{c_abs}"
            
        rc_ref = f"{r_part}{c_part}"
        
        # Convert to A1
        a1_ref = rc_to_a1(rc_ref, row, col)
        
        # Replace in the formula
        result = result.replace(rc_ref, a1_ref)
    
    # Restore string literals
    result = restore_string_literals(result, literals)
    
    return result


def extract_string_literals(text: str, literals_dict: Optional[Dict[str, str]] = None) -> str:
    """
    Extract string literals from text and replace with placeholders.
    
    Args:
        text: Text to process
        literals_dict: Dictionary to store extracted literals
        
    Returns:
        Text with string literals replaced by placeholders
    """
    if literals_dict is None:
        literals_dict = {}
    
    result = ""
    in_string = False
    current_string = ""
    i = 0
    
    while i < len(text):
        char = text[i]
        
        if char == '"' and (i == 0 or text[i-1] != '\\'):
            if in_string:
                # End of string, store it
                placeholder = f"__STRING{len(literals_dict)}__"
                literals_dict[placeholder] = current_string
                result += placeholder
                current_string = ""
            else:
                # Start of string
                current_string = ""
            in_string = not in_string
            i += 1
        elif in_string:
            if char == '\\' and i + 1 < len(text) and text[i+1] == '"':
                # Escaped quote
                current_string += '"'
                i += 2
            else:
                current_string += char
                i += 1
        else:
            result += char
            i += 1
    
    return result


def restore_string_literals(text: str, literals_dict: Dict[str, str]) -> str:
    """
    Restore string literals from placeholders.
    
    Args:
        text: Text with placeholders
        literals_dict: Dictionary with extracted literals
        
    Returns:
        Text with string literals restored
    """
    result = text
    
    for placeholder, literal in literals_dict.items():
        result = result.replace(placeholder, f'"{literal}"')
    
    return result


def adapt_formula_for_row(formula: str, source_row: int, target_row: int) -> str:
    """
    Adapt a formula for a different row.
    
    This converts the formula to R1C1 format and back to A1 for the target row,
    which keeps column references aligned correctly.
    
    Args:
        formula: Original Excel formula
        source_row: Original row number
        target_row: Target row number
        
    Returns:
        Adapted formula for the target row
    """
    if not formula or not formula.startswith("="):
        return formula
    
    # Convert to R1C1 from source context
    rc_formula = convert_formula_to_rc(formula, source_row, 1)
    
    # Convert back to A1 in target context
    return convert_formula_to_a1(rc_formula, target_row, 1)


def get_excel_formula_description(formula: str) -> str:
    """
    Generate a human-readable description of an Excel formula.
    
    Args:
        formula: Excel formula to describe
        
    Returns:
        Human-readable description of what the formula does
    """
    if not formula or not formula.startswith("="):
        return "Not a valid Excel formula"
    
    formula_content = formula.strip()[1:]  # Remove equals sign
    
    # Extract main function for simple formulas
    main_function_match = re.match(r'([A-Z][A-Za-z0-9\.]+)\(', formula_content)
    
    if main_function_match:
        main_function = main_function_match.group(1).upper()
        if main_function in EXCEL_FUNCTION_DESCRIPTIONS:
            # Special handling for common formula patterns
            if main_function == "IF":
                return _describe_if_formula(formula)
            elif main_function in ["AND", "OR"]:
                return _describe_logical_formula(formula, main_function)
            else:
                return f"{EXCEL_FUNCTION_DESCRIPTIONS[main_function]} function"
    
    # For more complex formulas, extract fields and operations
    fields = extract_column_names(formula)
    operations = []
    
    # Look for common operations
    if "+" in formula_content:
        operations.append("addition")
    if "-" in formula_content and not re.search(r'[0-9]-[0-9]', formula_content):
        operations.append("subtraction")
    if "*" in formula_content:
        operations.append("multiplication")
    if "/" in formula_content:
        operations.append("division")
    if ">" in formula_content or "<" in formula_content or "=" in formula_content:
        operations.append("comparison")
    
    if fields and operations:
        fields_str = ", ".join([f"'{f}'" for f in fields])
        ops_str = ", ".join(operations)
        return f"Formula using {fields_str} with {ops_str} operations"
    
    if fields:
        fields_str = ", ".join([f"'{f}'" for f in fields])
        return f"Formula referencing {fields_str}"
    
    # Generic fallback
    return "Complex Excel formula"


def _describe_if_formula(formula: str) -> str:
    """
    Generate description for an IF formula.
    
    Args:
        formula: IF formula to describe
        
    Returns:
        Human-readable description
    """
    # Remove equals sign
    formula_content = formula.strip()[1:]
    
    # Extract condition part
    condition_match = re.search(r'IF\s*\((.+?),', formula_content)
    if not condition_match:
        return "Conditional logic formula"
    
    condition = condition_match.group(1).strip()
    
    # Check for common comparison patterns
    comparison_match = re.search(r'([A-Za-z0-9_\[\]`\']+)\s*(<=|>=|<>|=|<|>)\s*(.+)', condition)
    if comparison_match:
        left = comparison_match.group(1)
        op = comparison_match.group(2)
        right = comparison_match.group(3)
        
        op_text = {
            "=": "equals",
            "<>": "does not equal",
            ">": "is greater than",
            "<": "is less than",
            ">=": "is greater than or equal to",
            "<=": "is less than or equal to"
        }.get(op, op)
        
        return f"Check if '{left}' {op_text} {right}"
    
    # If no specific pattern found
    return f"Conditional logic based on: {condition}"


def _describe_logical_formula(formula: str, function: str) -> str:
    """
    Generate description for AND or OR formula.
    
    Args:
        formula: Logical formula to describe
        function: "AND" or "OR"
        
    Returns:
        Human-readable description
    """
    conditions = []
    formula_content = formula.strip()[1:]
    
    # Try to extract individual conditions
    args_match = re.search(rf'{function}\s*\((.+)\)', formula_content)
    if args_match:
        args_str = args_match.group(1)
        # This is a simple split and won't handle nested functions correctly
        # For a complete solution, a proper formula parser would be needed
        args = []
        current_arg = ""
        paren_level = 0
        
        for char in args_str:
            if char == ',' and paren_level == 0:
                args.append(current_arg.strip())
                current_arg = ""
            else:
                if char == '(':
                    paren_level += 1
                elif char == ')':
                    paren_level -= 1
                current_arg += char
        
        if current_arg:
            args.append(current_arg.strip())
        
        if len(args) > 0:
            conditions = args
    
    if conditions:
        if len(conditions) <= 2:
            conditions_str = " and ".join(conditions) if function == "AND" else " or ".join(conditions)
            return f"Check if {conditions_str}"
        else:
            return f"Check if {len(conditions)} conditions are {'all' if function == 'AND' else 'any'} true"
    
    return f"Logical {function.lower()} operation"


def convert_excel_errors_to_none(value: Any) -> Any:
    """
    Convert Excel error values to None, keeping all other values as is.
    
    Args:
        value: Value to check for Excel errors
        
    Returns:
        None if the value is an Excel error, otherwise the original value
    """
    if isinstance(value, str) and value in EXCEL_ERROR_CODES:
        return None
    return value


def get_excel_error_description(error_value: str) -> str:
    """
    Get description for Excel error value.
    
    Args:
        error_value: Excel error string (e.g., "#DIV/0!")
        
    Returns:
        Description of the error
    """
    return EXCEL_ERROR_CODES.get(error_value, "Unknown Excel error")


def get_formula_dependencies(formula: str, df: pd.DataFrame) -> List[str]:
    """
    Get column dependencies for a formula in the context of a DataFrame.
    
    This identifies which columns from the DataFrame are used in the formula.
    
    Args:
        formula: Excel formula to analyze
        df: DataFrame context
        
    Returns:
        List of column names from the DataFrame that are used in the formula
    """
    # Get all potential column names from formula
    potential_columns = extract_column_names(formula)
    
    # Filter to only include columns that exist in the DataFrame
    df_columns = set(df.columns)
    dependencies = [col for col in potential_columns if col in df_columns]
    
    return sorted(dependencies)


def simplify_formula(formula: str) -> str:
    """
    Attempt to simplify a complex Excel formula.
    
    This performs basic simplifications:
    - Remove redundant parentheses
    - Simplify TRUE/FALSE constants in logical operations
    - Consolidate nested IF statements where possible
    
    Args:
        formula: Excel formula to simplify
        
    Returns:
        Simplified formula
    """
    if not formula or not formula.startswith("="):
        return formula
    
    result = formula
    
    # Remove redundant parentheses - e.g. =((A1)) to =(A1)
    redundant_pattern = r'\(\s*\(([^()]+)\)\s*\)'
    while re.search(redundant_pattern, result):
        result = re.sub(redundant_pattern, r'(\1)', result)
    
    # Simplify TRUE/FALSE constants in logical operations
    # e.g. =AND(A1=B1,TRUE) to =A1=B1
    result = re.sub(r'AND\s*\(([^,]+),\s*TRUE\s*\)', r'\1', result)
    result = re.sub(r'AND\s*\(TRUE\s*,\s*([^,]+)\)', r'\1', result)
    result = re.sub(r'OR\s*\(([^,]+),\s*FALSE\s*\)', r'\1', result)
    result = re.sub(r'OR\s*\(FALSE\s*,\s*([^,]+)\)', r'\1', result)
    
    # Replace OR(cond,TRUE) with TRUE and AND(cond,FALSE) with FALSE
    result = re.sub(r'OR\s*\([^,]+,\s*TRUE\s*\)', r'TRUE', result)
    result = re.sub(r'OR\s*\(TRUE\s*,\s*[^,]+\)', r'TRUE', result)
    result = re.sub(r'AND\s*\([^,]+,\s*FALSE\s*\)', r'FALSE', result)
    result = re.sub(r'AND\s*\(FALSE\s*,\s*[^,]+\)', r'FALSE', result)
    
    # Simplify IF(condition,TRUE,FALSE) to just condition
    result = re.sub(r'IF\s*\(([^,]+),\s*TRUE\s*,\s*FALSE\s*\)', r'\1', result)
    
    # Simplify IF(NOT(condition),TRUE,FALSE) to NOT(condition)
    result = re.sub(r'IF\s*\(NOT\s*\(([^()]+)\)\s*,\s*TRUE\s*,\s*FALSE\s*\)', r'NOT(\1)', result)
    
    # Keep the equals sign
    return result


def validate_excel_formula(formula: str) -> Tuple[bool, Optional[str]]:
    """
    Perform comprehensive validation of Excel formula syntax.

    This function checks for:
    - Proper formula syntax (starts with equals sign)
    - Balanced parentheses, brackets, and quotes
    - Valid function names
    - Common syntax errors

    Args:
        formula: Excel formula to validate

    Returns:
        Tuple of (is_valid, error_message)
        - is_valid: True if formula appears valid
        - error_message: Specific error message if invalid, None if valid
    """
    # Check for empty formula
    if not formula or not formula.strip():
        return False, "Formula is empty"

    # Ensure formula starts with equals sign
    formula_content = formula.strip()
    if not formula_content.startswith("="):
        return False, "Formula must start with equals sign (=)"

    # Remove equals sign for further checks
    formula_content = formula_content[1:]

    # Check for balanced parentheses
    open_count = 0
    for char in formula_content:
        if char == '(':
            open_count += 1
        elif char == ')':
            open_count -= 1
            if open_count < 0:
                return False, "Unbalanced parentheses - too many closing parentheses"

    if open_count > 0:
        return False, f"Unbalanced parentheses - missing {open_count} closing parenthesis"

    # Check for balanced square brackets (used for table references)
    open_count = 0
    for char in formula_content:
        if char == '[':
            open_count += 1
        elif char == ']':
            open_count -= 1
            if open_count < 0:
                return False, "Unbalanced square brackets - too many closing brackets"

    if open_count > 0:
        return False, f"Unbalanced square brackets - missing {open_count} closing bracket"

    # Check for balanced quotes
    in_quote = False
    for i, char in enumerate(formula_content):
        if char == '"' and (i == 0 or formula_content[i - 1] != '\\'):
            in_quote = not in_quote

    if in_quote:
        return False, "Unbalanced quotes - unclosed quote"

    # Check for balanced backticks (used for field names with spaces)
    in_backtick = False
    for i, char in enumerate(formula_content):
        if char == '`':
            in_backtick = not in_backtick

    if in_backtick:
        return False, "Unbalanced backticks - unclosed field reference"

    # Check for empty function calls like SUM()
    if re.search(r'\w+\(\s*\)', formula_content):
        return False, "Empty function arguments detected - e.g., SUM()"

    # Check for invalid references like A0 (row 0 doesn't exist)
    cell_refs = re.findall(r'([A-Za-z]+)([0-9]+)', formula_content)
    for col, row in cell_refs:
        if int(row) < 1:
            return False, f"Invalid cell reference {col}{row} - rows must be 1 or greater"

    # Check for obvious syntax errors
    syntax_errors = [
        (r',,', "Multiple consecutive commas"),
        (r'\(\)', "Empty parentheses"),
        (r'=\s*$', "Formula with only equals sign"),
        (r'[+\-*/]\s*[+\-*/]', "Consecutive operators"),
        (r'[+\-*/]\s*\)', "Operator before closing parenthesis"),
    ]

    for pattern, error in syntax_errors:
        if re.search(pattern, formula):
            return False, f"Syntax error: {error}"

    # Extract and validate function names
    function_matches = re.findall(r'([A-Za-z][A-Za-z0-9\.]*)\(', formula_content)

    # List of common Excel functions - can be expanded
    known_functions = set(EXCEL_FUNCTION_DESCRIPTIONS.keys())
    known_functions.update([
        # Additional functions
        "IF", "AND", "OR", "NOT", "TRUE", "FALSE",
        "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN",
        "VLOOKUP", "HLOOKUP", "INDEX", "MATCH",
        "DATE", "NOW", "TODAY", "EOMONTH", "YEAR", "MONTH", "DAY",
        "IFERROR", "IFNA", "ISBLANK", "ISTEXT", "ISNUMBER", "ISERROR",
        "LEFT", "RIGHT", "MID", "LEN", "FIND", "SEARCH", "REPLACE", "SUBSTITUTE",
        "CONCATENATE", "CONCAT", "TEXTJOIN", "TRIM", "UPPER", "LOWER", "PROPER",
        "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "INT", "MOD", "RAND", "RANDBETWEEN"
    ])

    unknown_functions = [f for f in function_matches if f.upper() not in known_functions]

    if unknown_functions:
        # This is just a warning, not an error, since it could be a custom function
        # We'll log it but not fail validation
        logger.warning(f"Formula contains potentially unknown functions: {', '.join(unknown_functions)}")

    # Additional check for circular references (e.g., =A1+B1 in cell A1)
    # This is more of a runtime issue than a syntax issue, but worth warning about

    # If we get here, the formula syntax appears valid
    return True, None

def check_formula_compatibility(formula: str, df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """
    Check if a formula is compatible with a DataFrame.
    
    This function checks:
    - If all referenced columns exist in the DataFrame
    - If obvious type incompatibilities exist (e.g., text operations on numeric columns)
    
    Args:
        formula: Excel formula to check
        df: DataFrame to check against
        
    Returns:
        Tuple of (is_compatible, list of issues)
    """
    issues = []
    
    # Get column dependencies
    dependencies = get_formula_dependencies(formula, df)
    
    # Check if all dependencies exist in DataFrame
    potential_columns = extract_column_names(formula)
    missing_columns = [col for col in potential_columns if col not in df.columns]
    
    if missing_columns:
        issues.append(f"Formula references columns not in data: {', '.join(missing_columns)}")
    
    # Check for basic type compatibility for common cases
    formula_lower = formula.lower()
    
    # Text functions on numeric columns
    text_functions = ['left', 'right', 'mid', 'len', 'search', 'find', 'text']
    for func in text_functions:
        if f"{func}(" in formula_lower:
            # Check numeric columns used with text functions
            for col in dependencies:
                if pd.api.types.is_numeric_dtype(df[col]):
                    issues.append(f"Text function '{func}' used with numeric column '{col}'")
    
    # Math functions on text columns
    math_functions = ['sum', 'average', 'round', 'int', 'abs', 'sqrt']
    for func in math_functions:
        if f"{func}(" in formula_lower:
            # Check text columns used with math functions
            for col in dependencies:
                if pd.api.types.is_string_dtype(df[col]):
                    issues.append(f"Math function '{func}' used with text column '{col}'")
    
    # Date functions on non-date columns
    date_functions = ['year', 'month', 'day', 'weekday', 'date', 'datedif']
    for func in date_functions:
        if f"{func}(" in formula_lower:
            # Check non-date columns used with date functions
            for col in dependencies:
                if not pd.api.types.is_datetime64_any_dtype(df[col]):
                    issues.append(f"Date function '{func}' used with non-date column '{col}'")
    
    is_compatible = len(issues) == 0
    
    return is_compatible, issues


def generate_excel_formula_template(template_name: str) -> str:
    """
    Generate an Excel formula template for common validation patterns.
    
    Args:
        template_name: Name of the template to generate
        
    Returns:
        Excel formula template
    """
    templates = {
        # Simple validation templates
        "not_blank": "=NOT(ISBLANK({field}))",
        "is_number": "=ISNUMBER({field})",
        "is_text": "=ISTEXT({field})",
        "positive": "={field}>0",
        "not_zero": "={field}<>0",
        "within_range": "=AND({field}>={min}, {field}<={max})",
        
        # Date validation templates
        "valid_date": "=ISNUMBER(DATEVALUE({field}))",
        "date_not_future": "={field}<=TODAY()",
        "date_not_past": "={field}>=TODAY()",
        "date_within_days": "=AND({field}>=TODAY()-{days_before}, {field}<=TODAY()+{days_after})",
        
        # Text validation templates
        "min_length": "=LEN({field})>={length}",
        "max_length": "=LEN({field})<={length}",
        "starts_with": "=LEFT({field}, {len})={text}",
        "contains": "=ISNUMBER(SEARCH({text}, {field}))",
        
        # Logical validation templates
        "conditional_required": "=IF({condition}, NOT(ISBLANK({field})), TRUE)",
        "mutually_exclusive": "=OR(AND(NOT(ISBLANK({field1})), ISBLANK({field2})), AND(ISBLANK({field1}), NOT(ISBLANK({field2}))))",
        "required_together": "=OR(AND(NOT(ISBLANK({field1})), NOT(ISBLANK({field2}))), AND(ISBLANK({field1}), ISBLANK({field2})))",
        
        # Comparison validation templates
        "equals": "={field1}={field2}",
        "not_equals": "={field1}<>{field2}",
        "greater_than": "={field1}>{field2}",
        "less_than": "={field1}<{field2}",
        "date_after": "={date1}>{date2}",
        "date_before": "={date1}<{date2}",
        
        # Audit-specific validation templates
        "segregation_of_duties": "={submitter}<>{approver}",
        "approval_sequence": "={submit_date}<={approval_date}",
        "risk_assessment": "=IF(ISBLANK({risk_field}), TRUE, {risk_field}>{threshold})",
        "third_party_check": "=IF(NOT(ISBLANK({vendor_field})), NOT(ISBLANK({assessment_field})), TRUE)",
    }
    
    if template_name in templates:
        return templates[template_name]
    else:
        available = ", ".join(sorted(templates.keys()))
        return f"# Template '{template_name}' not found. Available templates: {available}"


def create_formula_documentation(formula: str, df: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """
    Create comprehensive documentation for a formula.
    
    Args:
        formula: Excel formula to document
        df: Optional DataFrame for context
        
    Returns:
        Dictionary with formula documentation
    """
    if not formula or not formula.startswith("="):
        return {"error": "Not a valid Excel formula"}
    
    # Extract core information
    description = get_excel_formula_description(formula)
    columns = extract_column_names(formula)
    
    # Check if formula uses common Excel functions
    functions_used = []
    for func in EXCEL_FUNCTION_DESCRIPTIONS.keys():
        if re.search(rf'\b{func}\s*\(', formula, re.IGNORECASE):
            functions_used.append(func)
    
    # Check dependencies if DataFrame provided
    dependencies = []
    if df is not None:
        dependencies = get_formula_dependencies(formula, df)
        # Check compatibility
        is_compatible, issues = check_formula_compatibility(formula, df)
    else:
        is_compatible = None
        issues = []
    
    # Create documentation
    documentation = {
        "formula": formula,
        "description": description,
        "columns_referenced": sorted(columns),
        "functions_used": sorted(functions_used),
        "simplified_version": simplify_formula(formula)
    }
    
    if df is not None:
        documentation.update({
            "data_dependencies": dependencies,
            "compatible_with_data": is_compatible,
            "compatibility_issues": issues
        })
    
    return documentation


# Example usage
if __name__ == "__main__":
    # Test some utility functions
    formula = "=IF(A1>0, IF(B1<C1, \"Valid\", \"Invalid\"), \"N/A\")"
    print(f"Formula: {formula}")
    print(f"Description: {get_excel_formula_description(formula)}")
    print(f"Simplified: {simplify_formula(formula)}")
    print(f"Column names: {extract_column_names(formula)}")
    print(f"Cell references: {extract_cell_references(formula)}")
    
    # Test template generation
    print("\nTemplate example:")
    template = generate_excel_formula_template("approval_sequence")
    print(template)
    
    # Test formula to R1C1 conversion
    print("\nR1C1 conversion:")
    rc_formula = convert_formula_to_rc(formula, 5, 2)
    print(f"R1C1 version: {rc_formula}")
    print(f"Back to A1: {convert_formula_to_a1(rc_formula, 5, 2)}")

