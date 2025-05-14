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
        formula, backtick_fields, placeholder_map = self._extract_backtick_fields(formula)
        fields_used.extend(backtick_fields)

        # Tokenize the formula into distinct parts
        tokens = self._tokenize(formula)

        # Process tokens and identify field names
        processed_tokens = []
        i = 0

        while i < len(tokens):
            token = tokens[i]

            # Handle IN operator
            if token.upper() == 'IN':
                # Process the IN operator and its arguments
                in_expr = self._process_in_operator(tokens, i, fields_used)
                processed_tokens.append(in_expr)

                # Skip past the IN expression
                paren_count = 0
                while i < len(tokens):
                    if tokens[i] == '(':
                        paren_count += 1
                    elif tokens[i] == ')':
                        paren_count -= 1
                        if paren_count == 0:
                            break
                    i += 1

            # Handle functions
            elif token.upper() in self.function_map and i + 1 < len(tokens) and tokens[i+1] == '(':
                func_name = token.upper()
                func_expr = self._process_function(func_name, tokens, i, fields_used)
                processed_tokens.append(func_expr)

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

            # Handle operators
            elif token.upper() in self.operator_map:
                processed_tokens.append(self.operator_map[token.upper()])

            # Handle other tokens (field names, literals, etc.)
            else:
                processed_token = self._process_token(token, fields_used)
                processed_tokens.append(processed_token)

            i += 1

        # Join tokens and handle precedence
        parsed_formula = self._apply_precedence(processed_tokens)

        # Replace placeholders with actual field names
        for placeholder, field_name in placeholder_map.items():
            parsed_formula = parsed_formula.replace(f"df['{placeholder}']", f"df['{field_name}']")

        logger.info(f"Parsed formula: {formula} -> {parsed_formula}")
        return parsed_formula, fields_used

    def _extract_backtick_fields(self, formula: str) -> Tuple[str, List[str], Dict[str, str]]:
        """
        Extract backtick-quoted field names from formula.

        Args:
            formula: Excel-style formula

        Returns:
            Tuple of (formula with placeholders, list of field names, placeholder map)
        """
        backtick_fields = []
        backtick_pattern = r'`([^`]+)`'
        placeholder_map = {}

        # Find all backtick-quoted fields
        matches = list(re.finditer(backtick_pattern, formula))

        # Replace each field with a placeholder that won't be broken by tokenization
        modified_formula = formula

        for i, match in enumerate(reversed(matches)):  # Process in reverse to avoid index issues
            field_name = match.group(1)
            backtick_fields.append(field_name)

            # Create a placeholder without spaces
            placeholder = f"__FIELD_{i}__"
            placeholder_map[placeholder] = field_name

            # Replace in formula
            start, end = match.span()
            modified_formula = modified_formula[:start] + placeholder + modified_formula[end:]

        return modified_formula, backtick_fields, placeholder_map

    def _tokenize(self, formula: str) -> List[str]:
        """
        Break the formula into tokens for parsing.

        Args:
            formula: Excel-style formula

        Returns:
            List of formula tokens
        """
        # Build a pattern to match operators, function names, parentheses,
        # string literals, and identifiers (including our placeholders)
        pattern = r'(AND|OR|NOT|IN|<=|>=|<>|<|>|=|\(|\)|,|"[^"]*"|\'[^\']*\'|\b[A-Za-z][A-Za-z0-9_]*\b|\d+(?:\.\d+)?|__FIELD_\d+__)'

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
        # Special handling for ISBLANK
        if func_name == 'ISBLANK':
            # Find the argument (field name)
            field_name = None
            i = start_idx + 2  # Skip past ISBLANK and opening parenthesis

            while i < len(tokens) and tokens[i] != ')':
                if tokens[i] not in '(),+-*/' and not tokens[i].upper() in self.operator_map:
                    field_name = tokens[i]
                    if field_name not in fields_used:
                        fields_used.append(field_name)
                    break
                i += 1

            if not field_name:
                logger.error(f"Could not find field name for {func_name}")
                return "pd.Series(False, index=df.index)"  # Fallback

            return f"pd.isna(df['{field_name}'])"

        # Handle other functions that need argument substitution
        if func_name in self.function_map:
            template = self.function_map[func_name]

            # If the template doesn't need argument substitution
            if "{}" not in template:
                return template

            # Extract arguments
            arg_tokens = []
            paren_count = 0
            i = start_idx + 1  # Start after function name

            # Skip to opening parenthesis
            while i < len(tokens) and tokens[i] != '(':
                i += 1

            i += 1  # Skip past opening parenthesis

            # Collect argument tokens
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_count += 1
                    arg_tokens.append(token)
                elif token == ')':
                    if paren_count == 0:
                        break  # End of function
                    paren_count -= 1
                    arg_tokens.append(token)
                else:
                    arg_tokens.append(token)

                i += 1

            # Process argument tokens
            arg_str = ""
            for token in arg_tokens:
                if token.upper() in self.operator_map:
                    arg_str += self.operator_map[token.upper()]
                else:
                    arg_str += self._process_token(token, fields_used)

            return template.format(arg_str)

        # Default case (shouldn't reach here)
        logger.warning(f"Unhandled function: {func_name}")
        return f"{func_name.lower()}"

    def _process_in_operator(self, tokens: List[str], in_idx: int, fields_used: List[str]) -> str:
        """
        Process the IN operator and its list of values.

        Args:
            tokens: List of tokens
            in_idx: Index of the IN token
            fields_used: List to track field names

        Returns:
            Processed IN expression
        """
        # Get the field name (token before IN)
        if in_idx == 0:
            logger.error("IN operator has no preceding field")
            return ".isin([])"

        field_name = tokens[in_idx - 1]
        if field_name not in fields_used:
            fields_used.append(field_name)

        # Find the list values
        values = []
        i = in_idx + 1  # Start after IN

        # Skip to opening parenthesis
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error("No opening parenthesis found after IN")
            return f"df['{field_name}'].isin([])"

        i += 1  # Skip past opening parenthesis

        # Collect values until closing parenthesis
        while i < len(tokens):
            token = tokens[i]

            if token == ')':
                break
            elif token == ',':
                i += 1
                continue

            # Add the value
            values.append(token)
            i += 1

        # Format the values list
        values_str = ", ".join(values)

        return f"df['{field_name}'].isin([{values_str}])"

    def _process_token(self, token: str, fields_used: List[str]) -> str:
        """
        Process a token and identify if it's a field name.

        Args:
            token: Token to process
            fields_used: List to track field names

        Returns:
            Processed token string
        """
        # Handle placeholder fields
        if token.startswith('__FIELD_') and token.endswith('__'):
            return f"df['{token}']"  # Will be replaced with actual field name later

        # Handle string literals
        if (token.startswith('"') and token.endswith('"')) or (token.startswith("'") and token.endswith("'")):
            return token

        # Handle numeric literals
        if token.replace('.', '', 1).isdigit():
            return token

        # Handle parentheses and punctuation
        if token in "(),.+-*/":
            return token

        # Handle operators
        if token.upper() in self.operator_map:
            return self.operator_map[token.upper()]

        # Assume it's a field name if not recognized as anything else
        if token not in fields_used:
            fields_used.append(token)

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

        # Handle logical operators by ensuring proper parenthesization
        for op in ['&', '|']:
            # Find expressions connected by this operator
            parts = []
            current = ""
            paren_count = 0

            for char in formula:
                if char == '(':
                    paren_count += 1
                    current += char
                elif char == ')':
                    paren_count -= 1
                    current += char
                elif char == op and paren_count == 0:
                    # Found operator at top level
                    parts.append(current)
                    current = ""
                else:
                    current += char

            if current:
                parts.append(current)

            # If we found parts separated by the operator
            if len(parts) > 1:
                # Ensure each part is parenthesized
                for i in range(len(parts)):
                    part = parts[i].strip()
                    if not (part.startswith('(') and part.endswith(')')):
                        parts[i] = f"({part})"

                # Reconstruct the formula
                formula = f" {op} ".join(parts)

        # Ensure the entire formula is parenthesized
        if not (formula.startswith('(') and formula.endswith(')')):
            formula = f"({formula})"

        return formula

    def test_formula(self, formula: str, data: pd.DataFrame) -> Tuple[bool, Optional[pd.Series], Optional[str]]:
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
            missing_fields = [field for field in fields_used if field not in data.columns]
            if missing_fields:
                return False, None, f"Fields not found in data: {', '.join(missing_fields)}"

            # Create safe evaluation environment
            restricted_globals = {"__builtins__": {}}
            safe_locals = {"df": data, "pd": pd, "np": np}

            # Log the formula for debugging
            logger.debug(f"Evaluating formula: {parsed_formula}")

            # Evaluate the formula
            result = eval(parsed_formula, restricted_globals, safe_locals)

            # Ensure result is a boolean Series
            if not isinstance(result, pd.Series):
                return False, None, f"Formula did not return a Series (got {type(result).__name__})"

            # Ensure all results are boolean
            if result.dtype != bool:
                try:
                    result = result.astype(bool)
                except Exception as e:
                    return False, None, f"Could not convert result to boolean: {str(e)}"

            return True, result, None

        except Exception as e:
            error_msg = str(e)
            logger.error(f"Custom formula failed: {error_msg}, Formula: {formula}")
            return False, None, error_msg


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