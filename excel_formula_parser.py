"""
Excel Formula Parser for QA Analytics Framework

This module provides functionality to translate Excel-style formulas into
pandas expressions that can be evaluated against DataFrames.

Example:
    "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    
    Translates to:
    
    "(df['Submitter'] != df['Approver']) & (df['Submit Date'] <= df['TL Date'])"
"""

import re
import logging
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any, Union, Set

# Set up logging
from logging_config import setup_logging
logger = setup_logging()


class ExcelFormulaParser:
    """
    Parser to convert Excel-style formulas to pandas expressions.

    This class handles the translation of familiar Excel syntax to
    pandas operations that can be safely evaluated against DataFrames.
    """

    def __init__(self):
        """Initialize the Excel formula parser."""
        # Define operator mappings
        self.operator_map = {
            '=': '==',
            '<>': '!=',
            'AND': '&',
            'OR': '|',
            'NOT': '~',
            '>': '>',
            '>=': '>=',
            '<': '<',
            '<=': '<=',
            'IN': '.isin',
        }

        # Define Excel function mappings
        self.function_map = {
            'ISBLANK': 'pd.isna',
            'ISNUMBER': 'pd.to_numeric({}, errors="coerce").notna()',
            'ISTEXT': '{}.apply(lambda x: isinstance(x, str))',
            'TODAY': 'pd.Timestamp.today()',
            'NOW': 'pd.Timestamp.now()',
        }

    def parse(self, formula: str) -> Tuple[str, List[str]]:
        """
        Parse an Excel-style formula and convert it to a pandas expression.

        Args:
            formula: Excel-style formula string

        Returns:
            Tuple containing:
            - Parsed pandas expression
            - List of field names used in the formula
        """
        if not formula:
            logger.error("Empty formula provided")
            return "", []

        # Track fields for validation and documentation
        fields_used = []

        # First, handle backtick-quoted field names
        formula, backtick_fields = self._extract_backtick_fields(formula)
        fields_used.extend(backtick_fields)

        # Tokenize the formula into distinct parts
        tokens = self._tokenize(formula)

        # Process tokens and identify field names
        processed_tokens = []
        i = 0

        while i < len(tokens):
            token = tokens[i]

            # Handle operators
            if token.upper() in self.operator_map:
                processed_tokens.append(self.operator_map[token.upper()])

            # Handle functions
            elif token.upper() in self.function_map and i + 1 < len(tokens) and tokens[i+1] == '(':
                func_name = token.upper()
                processed_tokens.append(self._process_function(func_name, tokens, i, fields_used))
                # Skip to the end of the function
                paren_count = 1
                i += 2  # Skip past function name and opening paren
                while i < len(tokens) and paren_count > 0:
                    if tokens[i] == '(':
                        paren_count += 1
                    elif tokens[i] == ')':
                        paren_count -= 1
                    i += 1
                i -= 1  # Adjust for the outer loop increment

            # Handle other tokens (field names, literals, etc.)
            else:
                processed_token = self._process_token(token, fields_used)
                processed_tokens.append(processed_token)

            i += 1

        # Handle logical operators to ensure proper precedence
        parsed_formula = self._apply_precedence(processed_tokens)

        logger.info(f"Parsed formula: {formula} -> {parsed_formula}")
        return parsed_formula, fields_used

    def _extract_backtick_fields(self, formula: str) -> Tuple[str, List[str]]:
        """
        Extract backtick-quoted field names from formula.

        Args:
            formula: Excel-style formula

        Returns:
            Tuple of (formula with placeholders, list of field names)
        """
        backtick_fields = []
        backtick_pattern = r'`([^`]+)`'

        # Find all backtick-quoted fields
        matches = list(re.finditer(backtick_pattern, formula))

        # Replace each field with a placeholder
        modified_formula = formula
        for i, match in enumerate(reversed(matches)):  # Process in reverse to avoid index issues
            field_name = match.group(1)
            backtick_fields.append(field_name)

            # Replace with field name (without backticks)
            start, end = match.span()
            modified_formula = modified_formula[:start] + field_name + modified_formula[end:]

        return modified_formula, backtick_fields

    def _tokenize(self, formula: str) -> List[str]:
        """
        Break the formula into tokens for parsing.

        Args:
            formula: Excel-style formula

        Returns:
            List of formula tokens
        """
        # Build a pattern to match operators, function names, parentheses,
        # string literals, and identifiers
        pattern = r'(AND|OR|NOT|IN|<=|>=|<>|<|>|=|\(|\)|,|"[^"]*"|\'[^\']*\'|\b[A-Za-z][A-Za-z0-9_]*\b|\d+(?:\.\d+)?)'

        # Tokenize the formula
        tokens = re.findall(pattern, formula, re.IGNORECASE)

        # Clean up tokens (remove leading/trailing whitespace)
        return [token.strip() for token in tokens]

    def _process_function(self, func_name: str, tokens: List[str], start_idx: int, fields_used: List[str]) -> str:
        """
        Process an Excel function and its arguments.

        Args:
            func_name: Name of the function
            tokens: List of all tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Processed function string
        """
        # Get the template for this function
        template = self.function_map[func_name]

        # For simple function substitutions (like ISBLANK -> pd.isna)
        if "{}" not in template:
            return template

        # For functions that need argument processing
        # Extract the argument from between parentheses
        arg_tokens = []
        paren_count = 0
        i = start_idx + 1  # Start after function name

        while i < len(tokens):
            token = tokens[i]

            if token == '(':
                paren_count += 1
                if paren_count > 1:  # Only add if not the opening paren
                    arg_tokens.append(token)
            elif token == ')':
                paren_count -= 1
                if paren_count == 0:  # End of function
                    break
                arg_tokens.append(token)
            else:
                arg_tokens.append(token)

            i += 1

        # Process the argument tokens
        arg_str = ""
        for token in arg_tokens:
            if token.upper() in self.operator_map:
                arg_str += self.operator_map[token.upper()]
            else:
                arg_str += self._process_token(token, fields_used)

        # Return the formatted function
        return template.format(arg_str)

    def _process_token(self, token: str, fields_used: List[str]) -> str:
        """
        Process a token and identify if it's a field name.

        Args:
            token: Token to process
            fields_used: List to track field names

        Returns:
            Processed token string
        """
        # Handle string literals
        if (token.startswith('"') and token.endswith('"')) or (token.startswith("'") and token.endswith("'")):
            return token

        # Handle numeric literals
        if token.replace('.', '', 1).isdigit():
            return token

        # Handle operators
        if token in self.operator_map:
            return self.operator_map[token]

        # Handle other symbols
        if token in "(),.+-*/":
            return token

        # Assume it's a field name
        if token not in fields_used and re.match(r'^[A-Za-z][A-Za-z0-9_\s]*$', token):
            fields_used.append(token)

        # Return as a DataFrame reference
        return f"df['{token}']"

    def _apply_precedence(self, tokens: List[str]) -> str:
        """
        Apply operator precedence to ensure correct evaluation.

        Args:
            tokens: List of processed tokens

        Returns:
            Formula string with correct precedence
        """
        # Join tokens into a string
        formula = ''.join(tokens)

        # Handle logical operators
        for op in ['&', '|']:
            # Find all occurrences of the operator
            pattern = r'([^&|()]+)' + re.escape(op) + r'([^&|()]+)'
            matches = list(re.finditer(pattern, formula))

            # Replace each occurrence with properly parenthesized version
            for match in reversed(matches):  # Process in reverse to avoid index issues
                left, right = match.group(1), match.group(2)

                # Skip if already parenthesized
                if (left.startswith('(') and left.endswith(')') and
                    right.startswith('(') and right.endswith(')')):
                    continue

                # Parenthesize the operands if needed
                if not (left.startswith('(') and left.endswith(')')):
                    left = f"({left})"

                if not (right.startswith('(') and right.endswith(')')):
                    right = f"({right})"

                # Replace in formula
                start, end = match.span()
                formula = formula[:start] + f"{left} {op} {right}" + formula[end:]

        # Wrap the entire formula in parentheses if it's not already
        if not (formula.startswith('(') and formula.endswith(')')):
            formula = f"({formula})"

        return formula

    def test_formula(self, formula: str, data: pd.DataFrame) -> Tuple[bool, pd.Series, str]:
        """
        Test a formula against sample data.

        Args:
            formula: Excel-style formula
            data: DataFrame to test against

        Returns:
            Tuple of (success, result series, error message)
        """
        try:
            # Parse the formula
            parsed_formula, fields_used = self.parse(formula)

            # Check that all fields exist in the data
            for field in fields_used:
                if field not in data.columns:
                    return False, None, f"Field '{field}' not found in data"

            # Create safe evaluation environment
            restricted_globals = {"__builtins__": {}}
            safe_locals = {"df": data, "pd": pd, "np": np}

            # Evaluate the formula
            result = eval(parsed_formula, restricted_globals, safe_locals)

            # Ensure result is a boolean Series
            if not isinstance(result, pd.Series):
                return False, None, f"Formula did not return a Series (got {type(result).__name__})"

            # Convert to boolean if needed
            if result.dtype != bool:
                try:
                    result = result.astype(bool)
                except:
                    return False, None, f"Could not convert result to boolean (dtype: {result.dtype})"

            return True, result, None

        except Exception as e:
            logger.error(f"Error testing formula: {e}")
            return False, None, str(e)


# Example usage (for testing)
if __name__ == "__main__":
    # Test formula
    test_formula = "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    parser = ExcelFormulaParser()
    result, fields = parser.parse(test_formula)
    print(f"Original: {test_formula}")
    print(f"Parsed: {result}")
    print(f"Fields used: {fields}")

    # Test with sample data
    sample_data = pd.DataFrame({
        'Submitter': ['John', 'Mary', 'John', 'Bob'],
        'Approver': ['Alice', 'John', 'John', 'Charlie'],
        'Submit Date': pd.to_datetime(['2025-01-01', '2025-02-01', '2025-03-01', '2025-04-01']),
        'TL Date': pd.to_datetime(['2025-01-05', '2025-01-15', '2025-02-01', '2025-05-01'])
    })

    success, result, error = parser.test_formula(test_formula, sample_data)
    if success:
        print("\nTest Results:")
        print(pd.DataFrame({
            'Submitter': sample_data['Submitter'],
            'Approver': sample_data['Approver'],
            'Submit Date': sample_data['Submit Date'],
            'TL Date': sample_data['TL Date'],
            'Result': result
        }))
    else:
        print(f"Error: {error}")