"""
Excel Formula Parser Function Enhancement

This module extends the existing ExcelFormulaParser with comprehensive support for
additional Excel functions including:

1. Logical Functions: IF, IFS, SWITCH
2. Date Functions: DATEDIF, EDATE, DATEVALUE, YEARFRAC
3. String Functions: LEFT, RIGHT, MID, TRIM, CONCATENATE
4. Aggregation Functions: COUNTIF, SUMIF, AVERAGEIF
5. Information Functions: ISBLANK, ISNUMBER, ISTEXT, ISERROR

Usage:
    parser = ExcelFormulaParser()
    parsed_formula, fields = parser.parse("IF(Risk_Level=\"High\", Days_Open<=30, Days_Open<=90)")
"""

import re
import logging
import pandas as pd
import numpy as np
import datetime
import sys
from dateutil.relativedelta import relativedelta
from typing import Dict, List, Optional, Tuple, Any, Union, Set, Callable

# Unicode Character Replacement
sys.stdout.reconfigure(encoding='utf-8')

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
        """Initialize the Excel formula parser with extended function support."""
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

        # Define basic Excel function mappings
        self.simple_function_map = {
            'ISBLANK': 'pd.isna',
            'ISNUMBER': 'pd.to_numeric({}, errors="coerce").notna()',
            'ISTEXT': '{}.apply(lambda x: isinstance(x, str))',
            'TODAY': 'pd.Timestamp.today()',
            'NOW': 'pd.Timestamp.now()',
        }

        # Define complex functions that need special processing
        self.complex_functions = {
            'IF': self._process_if_function,
            'IFS': self._process_ifs_function,
            'SWITCH': self._process_switch_function,
            'DATEDIF': self._process_datedif_function,
            'EDATE': self._process_edate_function,
            'DATEVALUE': self._process_datevalue_function,
            'YEARFRAC': self._process_yearfrac_function,
            'LEFT': self._process_left_function,
            'RIGHT': self._process_right_function,
            'MID': self._process_mid_function,
            'TRIM': self._process_trim_function,
            'CONCATENATE': self._process_concatenate_function,
            'COUNTIF': self._process_countif_function,
            'SUMIF': self._process_sumif_function,
            'AVERAGEIF': self._process_averageif_function,
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
        # Track placeholders to remove from fields_used later
        placeholders = []

        # First, handle backtick-quoted field names
        formula, backtick_fields, placeholder_map = self._extract_backtick_fields(formula)
        fields_used.extend(backtick_fields)
        placeholders.extend(placeholder_map.keys())

        # Tokenize the formula into distinct parts
        tokens = self._tokenize(formula)

        # Process multi-word field names
        tokens = self._pre_process_multi_word_fields(tokens)

        # Process tokens and identify field names
        processed_tokens = []
        i = 0

        while i < len(tokens):
            token = tokens[i]

            # Handle complex functions first
            if token.upper() in self.complex_functions and i + 1 < len(tokens) and tokens[i+1] == '(':
                func_name = token.upper()
                handler = self.complex_functions[func_name]

                # Call the appropriate function handler
                func_expr, new_i, func_fields = handler(tokens, i, fields_used)

                # Add new fields to our tracked fields list
                for field in func_fields:
                    if field not in fields_used:
                        fields_used.append(field)

                processed_tokens.append(func_expr)
                i = new_i

            # Handle IN operator
            elif token.upper() == 'IN':
                # Process the IN operator and its arguments
                in_expr, in_fields = self._process_in_operator(tokens, i, fields_used)
                processed_tokens.append(in_expr)

                # Add new fields to our tracked fields list
                for field in in_fields:
                    if field not in fields_used:
                        fields_used.append(field)

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

            # Handle simple functions
            elif token.upper() in self.simple_function_map and i + 1 < len(tokens) and tokens[i+1] == '(':
                func_name = token.upper()
                func_expr, new_i, func_fields = self._process_simple_function(func_name, tokens, i, fields_used)

                # Add new fields to our tracked fields list
                for field in func_fields:
                    if field not in fields_used:
                        fields_used.append(field)

                processed_tokens.append(func_expr)
                i = new_i

            # Handle operators
            elif token.upper() in self.operator_map:
                processed_tokens.append(self.operator_map[token.upper()])

            # Handle other tokens (field names, literals, etc.)
            else:
                processed_token, token_fields = self._process_token(token, fields_used)
                processed_tokens.append(processed_token)

                # Add new fields to our tracked fields list
                for field in token_fields:
                    if field not in fields_used:
                        fields_used.append(field)

            i += 1

        # Join tokens and handle precedence
        parsed_formula = self._apply_precedence(processed_tokens)

        # Replace placeholders with actual field names
        for placeholder, field_name in placeholder_map.items():
            parsed_formula = parsed_formula.replace(f"df['{placeholder}']", f"df['{field_name}']")

        # Clean up fields_used - remove placeholders
        fields_used = [field for field in fields_used if field not in placeholders]

        logger.info(f"Parsed formula: {formula} -> {parsed_formula}")
        return parsed_formula, fields_used

    def _pre_process_multi_word_fields(self, tokens: List[str]) -> List[str]:
        """
        Pre-process tokens to handle multi-word field names.

        This merges consecutive tokens that are likely part of the same
        multi-word field name.

        Args:
            tokens: List of tokens from the tokenizer

        Returns:
            Processed token list with multi-word fields merged
        """
        if not tokens:
            return tokens

        result = []
        i = 0

        while i < len(tokens):
            current_token = tokens[i]

            # Look ahead for potential multi-word field
            if (i + 2 < len(tokens) and  # Need at least 3 tokens: field1 field2 (operator or parenthesis)
                # Check if current token is a potential field name
                not current_token.upper() in self.operator_map and
                not current_token in "(),.+-*/" and
                not current_token.startswith('"') and
                not current_token.startswith("'") and
                not current_token.replace('.', '', 1).isdigit() and
                # Check if next token is a potential field name continuation
                not tokens[i+1].upper() in self.operator_map and
                not tokens[i+1] in "(),.+-*/" and
                not tokens[i+1].startswith('"') and
                not tokens[i+1].startswith("'") and
                not tokens[i+1].replace('.', '', 1).isdigit() and
                # Check if token after that is a proper terminator (operator, parenthesis, etc.)
                (tokens[i+2].upper() in self.operator_map or
                 tokens[i+2] in "(),.+-*/")
                ):

                # Combine current and next token as a multi-word field
                combined_token = current_token + " " + tokens[i+1]
                result.append(combined_token)
                i += 2  # Skip the next token since we've combined it

            else:
                # Just add the current token as is
                result.append(current_token)
                i += 1

        return result

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
        Break the formula into distinct parts for parsing.

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

    def _process_token(self, token: str, fields_used: List[str]) -> Tuple[str, List[str]]:
        """
        Process a token and identify if it's a field name.

        Args:
            token: Token to process
            fields_used: List to track field names

        Returns:
            Tuple of (processed token string, new fields found)
        """
        new_fields = []

        # Handle placeholder fields
        if token.startswith('__FIELD_') and token.endswith('__'):
            new_fields.append(token)
            return f"df['{token}']", new_fields

        # Handle string literals
        if (token.startswith('"') and token.endswith('"')) or (token.startswith("'") and token.endswith("'")):
            return token, new_fields

        # Handle numeric literals
        if token.replace('.', '', 1).isdigit():
            return token, new_fields

        # Handle parentheses and punctuation
        if token in "(),.+-*/":
            return token, new_fields

        # Handle operators
        if token.upper() in self.operator_map:
            return self.operator_map[token.upper()], new_fields

        # Assume it's a field name if not recognized as anything else
        new_fields.append(token)
        return f"df['{token}']", new_fields

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

    def _process_simple_function(self, func_name: str, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a simple Excel function and its arguments.

        Args:
            func_name: Name of the function
            tokens: List of all tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Special handling for ISBLANK
        if func_name == 'ISBLANK':
            # Find the argument (field name)
            field_name = None
            i = start_idx + 2  # Skip past ISBLANK and opening parenthesis

            while i < len(tokens) and tokens[i] != ')':
                if tokens[i] not in '(),+-*/' and not tokens[i].upper() in self.operator_map:
                    field_name = tokens[i]
                    if field_name not in fields_used:
                        new_fields.append(field_name)
                    break
                i += 1

            if not field_name:
                logger.error(f"Could not find field name for {func_name}")
                return "pd.Series(False, index=df.index)", start_idx + 2, new_fields

            return f"pd.isna(df['{field_name}'])", i + 1, new_fields  # +1 to skip past closing parenthesis

        # Handle other functions that need argument substitution
        if func_name in self.simple_function_map:
            template = self.simple_function_map[func_name]

            # If the template doesn't need argument substitution
            if "{}" not in template:
                return template, start_idx + 3, new_fields  # +3 to skip past function name, ( and )

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
                    processed_token, token_fields = self._process_token(token, fields_used)
                    arg_str += processed_token
                    new_fields.extend(token_fields)

            return template.format(arg_str), i + 1, new_fields  # +1 to skip past closing parenthesis

        # Default case (shouldn't reach here)
        logger.warning(f"Unhandled function: {func_name}")
        return f"{func_name.lower()}", start_idx + 1, new_fields

    def _process_in_operator(self, tokens: List[str], in_idx: int, fields_used: List[str]) -> Tuple[str, List[str]]:
        """
        Process the IN operator and its list of values.

        Args:
            tokens: List of tokens
            in_idx: Index of the IN token
            fields_used: List to track field names

        Returns:
            Tuple of (processed IN expression, new fields found)
        """
        new_fields = []

        # Get the field name (tokens before IN)
        if in_idx == 0:
            logger.error("IN operator has no preceding field")
            return ".isin([])", new_fields

        # Handle multi-word field names (not in backticks)
        # Look for consecutive identifiers before the IN token
        field_name_parts = []
        field_idx = in_idx - 1

        # Add the token immediately before IN
        field_name_parts.insert(0, tokens[field_idx])

        # Check if we should look for more parts of a multi-word field
        while field_idx > 0:
            prev_token = tokens[field_idx - 1]

            # If the previous token is an identifier (not an operator, parenthesis, etc.)
            # and not already a fully-qualified field reference
            if (not prev_token.upper() in self.operator_map and
                prev_token not in "(),+-*/" and
                not prev_token.startswith("df['") and
                not prev_token.startswith("__FIELD_")):

                # Insert at the beginning to maintain order
                field_name_parts.insert(0, prev_token)
                field_idx -= 1
            else:
                # Stop if we hit something that's not part of a field name
                break

        # Join the parts to form the complete field name
        field_name = " ".join(field_name_parts)

        # Add to fields_used if not already present
        if field_name not in fields_used:
            new_fields.append(field_name)

        # Find the list values
        values = []
        i = in_idx + 1  # Start after IN

        # Skip to opening parenthesis
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error("No opening parenthesis found after IN")
            return f"df['{field_name}'].isin([])", new_fields

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

        return f"df['{field_name}'].isin([{values_str}])", new_fields

    # === Logical Function Handlers ===

    def _process_if_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process an IF function.

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract the three arguments: condition, true_value, false_value
        condition, true_value, false_value, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 3, fields_used
        )

        new_fields.extend(arg_fields)

        # Use numpy.where to implement the IF function
        result = f"np.where({condition}, {true_value}, {false_value})"

        return result, end_idx, new_fields

    def _process_ifs_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process an IFS function (multiple IF conditions).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Skip to opening parenthesis
        i = start_idx + 1
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error("No opening parenthesis found for IFS function")
            return "pd.Series(np.nan, index=df.index)", i, new_fields

        i += 1  # Skip past opening parenthesis

        # Extract all condition/value pairs
        conditions = []
        values = []

        while i < len(tokens):
            # Extract condition
            condition_tokens = []
            paren_level = 0

            # Collect tokens for the condition (until comma or end of function)
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    condition_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function without finding a comma - error
                        logger.error("Unexpected end of IFS function")
                        return "pd.Series(np.nan, index=df.index)", i, new_fields
                    paren_level -= 1
                    condition_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # Found the comma separator
                    break
                else:
                    condition_tokens.append(token)

                i += 1

            # Process the condition tokens
            condition = ""
            for token in condition_tokens:
                if token.upper() in self.operator_map:
                    condition += self.operator_map[token.upper()]
                else:
                    processed_token, token_fields = self._process_token(token, fields_used)
                    condition += processed_token
                    new_fields.extend(token_fields)

            conditions.append(condition)

            # Skip the comma
            if i < len(tokens) and tokens[i] == ',':
                i += 1
            else:
                # Error - expected a comma
                logger.error("Expected comma in IFS function")
                return "pd.Series(np.nan, index=df.index)", i, new_fields

            # Extract value
            value_tokens = []
            paren_level = 0

            # Collect tokens for the value (until comma or end of function)
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    value_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function
                        i += 1  # Skip past closing parenthesis
                        break
                    paren_level -= 1
                    value_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # Found the comma separator for the next condition
                    break
                else:
                    value_tokens.append(token)

                i += 1

            # Process the value tokens
            value = ""
            for token in value_tokens:
                if token.upper() in self.operator_map:
                    value += self.operator_map[token.upper()]
                else:
                    processed_token, token_fields = self._process_token(token, fields_used)
                    value += processed_token
                    new_fields.extend(token_fields)

            values.append(value)

            # If we ended with a closing parenthesis, we're done
            if i > 0 and tokens[i-1] == ')':
                break

            # If we ended with a comma, continue to the next condition/value pair
            if i < len(tokens) and tokens[i] == ',':
                i += 1
            else:
                # Error - expected a comma or closing parenthesis
                logger.error("Expected comma or closing parenthesis in IFS function")
                return "pd.Series(np.nan, index=df.index)", i, new_fields

        # Ensure we have at least one condition/value pair
        if not conditions or not values or len(conditions) != len(values):
            logger.error("Invalid IFS function arguments")
            return "pd.Series(np.nan, index=df.index)", i, new_fields

        # Build the nested numpy.where expression
        result = "pd.Series(np.nan, index=df.index)"  # Default if no conditions match

        # Build from the last condition to the first (bottom up)
        for j in range(len(conditions) - 1, -1, -1):
            result = f"np.where({conditions[j]}, {values[j]}, {result})"

        return result, i, new_fields

    def _process_switch_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a SWITCH function.

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Skip to opening parenthesis
        i = start_idx + 1
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error("No opening parenthesis found for SWITCH function")
            return "pd.Series(np.nan, index=df.index)", i, new_fields

        i += 1  # Skip past opening parenthesis

        # Extract the expression to switch on
        expression_tokens = []
        paren_level = 0

        # Collect tokens for the expression (until comma)
        while i < len(tokens):
            token = tokens[i]

            if token == '(':
                paren_level += 1
                expression_tokens.append(token)
            elif token == ')':
                if paren_level == 0:
                    # End of function without finding a comma - error
                    logger.error("Unexpected end of SWITCH function")
                    return "pd.Series(np.nan, index=df.index)", i, new_fields
                paren_level -= 1
                expression_tokens.append(token)
            elif token == ',' and paren_level == 0:
                # Found the comma separator
                i += 1  # Skip past comma
                break
            else:
                expression_tokens.append(token)

            i += 1

        # Process the expression tokens
        expression = ""
        for token in expression_tokens:
            if token.upper() in self.operator_map:
                expression += self.operator_map[token.upper()]
            else:
                processed_token, token_fields = self._process_token(token, fields_used)
                expression += processed_token
                new_fields.extend(token_fields)

        # Extract value/result pairs and default value (if present)
        values = []
        results = []
        default_result = None

        while i < len(tokens):
            # Extract value to match
            value_tokens = []
            paren_level = 0

            # Collect tokens for the value (until comma)
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    value_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function
                        i += 1  # Skip past closing parenthesis

                        # If we have an odd number of tokens, the last one is the default
                        if len(values) == len(results):
                            # Process the value as the default result
                            default_result = ""
                            for token in value_tokens:
                                if token.upper() in self.operator_map:
                                    default_result += self.operator_map[token.upper()]
                                else:
                                    processed_token, token_fields = self._process_token(token, fields_used)
                                    default_result += processed_token
                                    new_fields.extend(token_fields)
                        else:
                            # This is a value, not a result - error
                            logger.error("Unexpected end of SWITCH function - value without result")
                            return "pd.Series(np.nan, index=df.index)", i, new_fields

                        break
                    paren_level -= 1
                    value_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # Found the comma separator
                    i += 1  # Skip past comma
                    break
                else:
                    value_tokens.append(token)

                i += 1

            # If we ended with a closing parenthesis, we're done
            if i > 0 and tokens[i-1] == ')':
                break

            # Process the value tokens
            value = ""
            for token in value_tokens:
                if token.upper() in self.operator_map:
                    value += self.operator_map[token.upper()]
                else:
                    processed_token, token_fields = self._process_token(token, fields_used)
                    value += processed_token
                    new_fields.extend(token_fields)

            values.append(value)

            # Extract result
            result_tokens = []
            paren_level = 0

            # Collect tokens for the result (until comma or end of function)
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    result_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function
                        i += 1  # Skip past closing parenthesis
                        break
                    paren_level -= 1
                    result_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # Found the comma separator
                    i += 1  # Skip past comma
                    break
                else:
                    result_tokens.append(token)

                i += 1

            # Process the result tokens
            result = ""
            for token in result_tokens:
                if token.upper() in self.operator_map:
                    result += self.operator_map[token.upper()]
                else:
                    processed_token, token_fields = self._process_token(token, fields_used)
                    result += processed_token
                    new_fields.extend(token_fields)

            results.append(result)

            # If we ended with a closing parenthesis, we're done
            if i > 0 and tokens[i-1] == ')':
                break

        # Ensure we have at least one value/result pair
        if not values or not results or len(values) != len(results):
            logger.error("Invalid SWITCH function arguments")
            return "pd.Series(np.nan, index=df.index)", i, new_fields

        # Use default result if provided, otherwise use NaN
        if default_result is None:
            default_result = "pd.Series(np.nan, index=df.index)"

        # Build the nested numpy.where expression for the SWITCH
        result = default_result

        # Build from the last condition to the first (bottom up)
        for j in range(len(values) - 1, -1, -1):
            result = f"np.where({expression} == {values[j]}, {results[j]}, {result})"

        return result, i, new_fields

    # === Date Function Handlers ===

    def _process_datedif_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a DATEDIF function (date difference).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract the three arguments: start_date, end_date, unit
        start_date, end_date, unit, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 3, fields_used
        )

        new_fields.extend(arg_fields)

        # Implement different units (Y=years, M=months, D=days, YM=months excluding years, etc.)
        unit_str = unit.strip("'\"")

        # Ensure dates are datetime objects
        start_parsed = f"pd.to_datetime({start_date}, errors='coerce')"
        end_parsed = f"pd.to_datetime({end_date}, errors='coerce')"

        # Implement different unit calculations
        if unit_str.upper() == 'Y':  # Years
            result = f"(({end_parsed} - {start_parsed}).dt.days / 365.25).astype(int)"
        elif unit_str.upper() == 'M':  # Months
            result = f"((({end_parsed}.dt.year - {start_parsed}.dt.year) * 12) + ({end_parsed}.dt.month - {start_parsed}.dt.month))"
        elif unit_str.upper() == 'D':  # Days
            result = f"({end_parsed} - {start_parsed}).dt.days"
        elif unit_str.upper() == 'YM':  # Months excluding years
            result = f"(({end_parsed}.dt.month - {start_parsed}.dt.month) % 12)"
        elif unit_str.upper() == 'MD':  # Days excluding months and years
            # This is more complex - need to calculate days between same day in the month
            result = f"np.minimum(({end_parsed}.dt.day), pd.to_datetime({end_parsed}.dt.year.astype(str) + '-' + {end_parsed}.dt.month.astype(str) + '-' + {start_parsed}.dt.day.astype(str), errors='coerce').dt.day) - {start_parsed}.dt.day"
        elif unit_str.upper() == 'YD':  # Days excluding years
            # Days in the year (day of year)
            result = f"(({end_parsed}.dt.dayofyear - {start_parsed}.dt.dayofyear) % 365)"
        else:
            # Default to days if unit is not recognized
            logger.warning(f"Unrecognized DATEDIF unit: {unit_str}, defaulting to days")
            result = f"({end_parsed} - {start_parsed}).dt.days"

        return result, end_idx, new_fields

    def _process_edate_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process an EDATE function (date plus/minus months).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract the two arguments: date, months
        date, months, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 2, fields_used
        )

        new_fields.extend(arg_fields)

        # Parse the date and add months using dateutil.relativedelta
        result = f"pd.to_datetime({date}, errors='coerce') + pd.to_timedelta({months} * 30, unit='d')"

        # More accurate month calculation with relativedelta
        result = f"pd.to_datetime({date}, errors='coerce').apply(lambda x: x + relativedelta(months=int({months})) if pd.notna(x) else pd.NaT)"

        return result, end_idx, new_fields

    def _process_datevalue_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a DATEVALUE function (converts string to date).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract the one argument: date_text
        date_text, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 1, fields_used
        )

        new_fields.extend(arg_fields)

        # Convert to pandas datetime
        result = f"pd.to_datetime({date_text}, errors='coerce')"

        return result, end_idx, new_fields

    def _process_yearfrac_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a YEARFRAC function (fraction of year between dates).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: start_date, end_date, [basis]
        args, end_idx, arg_fields = self._extract_variable_args(
            tokens, start_idx + 1, 2, 3, fields_used
        )

        new_fields.extend(arg_fields)

        start_date = args[0]
        end_date = args[1]

        # Default basis is 0 (US 30/360)
        basis = args[2] if len(args) > 2 else "0"

        # Parse dates
        start_parsed = f"pd.to_datetime({start_date}, errors='coerce')"
        end_parsed = f"pd.to_datetime({end_date}, errors='coerce')"

        # Implement different basis calculations
        # 0 or missing = US 30/360
        # 1 = Actual/actual
        # 2 = Actual/360
        # 3 = Actual/365
        # 4 = European 30/360

        # Use a lambda function to handle the different basis options
        result = f"""(lambda start, end, basis: 
            (end - start).dt.days / 365 if basis == 1 else
            (end - start).dt.days / 360 if basis == 2 else
            (end - start).dt.days / 365 if basis == 3 else
            ((end.dt.year - start.dt.year) * 360 + 
             (end.dt.month - start.dt.month) * 30 + 
             np.minimum(end.dt.day, 30) - np.minimum(start.dt.day, 30)) / 360
        )({start_parsed}, {end_parsed}, {basis})"""

        return result, end_idx, new_fields

    # === String Function Handlers ===

    def _process_left_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a LEFT function (leftmost characters).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: text, num_chars
        text, num_chars, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 2, fields_used
        )

        new_fields.extend(arg_fields)

        # Use pandas str accessor
        result = f"{text}.astype(str).str[:int({num_chars})]"

        return result, end_idx, new_fields

    def _process_right_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a RIGHT function (rightmost characters).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: text, num_chars
        text, num_chars, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 2, fields_used
        )

        new_fields.extend(arg_fields)

        # Use pandas str accessor with negative indexing
        result = f"{text}.astype(str).str[-int({num_chars}):]"

        return result, end_idx, new_fields

    def _process_mid_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a MID function (middle characters).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: text, start_pos, num_chars
        text, start_pos, num_chars, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 3, fields_used
        )

        new_fields.extend(arg_fields)

        # Adjust for 1-based indexing in Excel vs. 0-based in Python
        adjusted_start = f"int({start_pos}) - 1"

        # Use pandas str accessor with slicing
        result = f"{text}.astype(str).str[{adjusted_start}:{adjusted_start} + int({num_chars})]"

        return result, end_idx, new_fields

    def _process_trim_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a TRIM function (remove spaces).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: text
        text, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 1, fields_used
        )

        new_fields.extend(arg_fields)

        # Use pandas str accessor
        result = f"{text}.astype(str).str.strip()"

        return result, end_idx, new_fields

    def _process_concatenate_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a CONCATENATE function (join strings).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract variable number of text arguments
        args, end_idx, arg_fields = self._extract_variable_args(
            tokens, start_idx + 1, 1, None, fields_used
        )

        new_fields.extend(arg_fields)

        # Convert all arguments to strings and concatenate
        result = " + ".join([f"{arg}.astype(str)" for arg in args])

        return result, end_idx, new_fields

    # === Aggregation Function Handlers ===

    def _process_countif_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a COUNTIF function (count cells meeting criteria).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # This is a bit complex in pandas as COUNTIF operates on ranges,
        # but we're working with single-row expressions. For our use case,
        # we'll implement it as either 0 or 1 for the current row.

        # Extract arguments: range, criteria
        range_expr, criteria, end_idx, arg_fields = self._extract_function_args(
            tokens, start_idx + 1, 2, fields_used
        )

        new_fields.extend(arg_fields)

        # Parse criteria - could be a value, comparison, or wildcard
        if criteria.startswith('"') or criteria.startswith("'"):
            # String criteria - remove quotes
            criteria_str = criteria.strip("'\"")

            # Check if it's a comparison operator
            if criteria_str.startswith(('=', '>', '<', '>=', '<=', '<>')):
                operator = criteria_str[0]
                if criteria_str.startswith(('>=', '<=', '<>')):
                    operator = criteria_str[:2]
                    value = criteria_str[2:]
                else:
                    value = criteria_str[1:]

                # Map to Python operator
                if operator == '=':
                    comparison = f"{range_expr} == {value}"
                elif operator == '<>':
                    comparison = f"{range_expr} != {value}"
                else:
                    comparison = f"{range_expr} {operator} {value}"
            elif '*' in criteria_str:
                # Wildcard match
                pattern = criteria_str.replace('*', '.*')
                comparison = f"{range_expr}.astype(str).str.match(r'{pattern}')"
            else:
                # Exact match
                comparison = f"{range_expr} == '{criteria_str}'"
        else:
            # Numeric or field criteria
            comparison = f"{range_expr} == {criteria}"

        # Result is 1 if condition is true, 0 otherwise
        result = f"np.where({comparison}, 1, 0)"

        return result, end_idx, new_fields

    def _process_sumif_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process a SUMIF function (sum cells meeting criteria).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: range, criteria, [sum_range]
        args, end_idx, arg_fields = self._extract_variable_args(
            tokens, start_idx + 1, 2, 3, fields_used
        )

        new_fields.extend(arg_fields)

        range_expr = args[0]
        criteria = args[1]
        sum_range = args[2] if len(args) > 2 else range_expr

        # Parse criteria - similar to COUNTIF
        if criteria.startswith('"') or criteria.startswith("'"):
            # String criteria - remove quotes
            criteria_str = criteria.strip("'\"")

            # Check if it's a comparison operator
            if criteria_str.startswith(('=', '>', '<', '>=', '<=', '<>')):
                operator = criteria_str[0]
                if criteria_str.startswith(('>=', '<=', '<>')):
                    operator = criteria_str[:2]
                    value = criteria_str[2:]
                else:
                    value = criteria_str[1:]

                # Map to Python operator
                if operator == '=':
                    comparison = f"{range_expr} == {value}"
                elif operator == '<>':
                    comparison = f"{range_expr} != {value}"
                else:
                    comparison = f"{range_expr} {operator} {value}"
            elif '*' in criteria_str:
                # Wildcard match
                pattern = criteria_str.replace('*', '.*')
                comparison = f"{range_expr}.astype(str).str.match(r'{pattern}')"
            else:
                # Exact match
                comparison = f"{range_expr} == '{criteria_str}'"
        else:
            # Numeric or field criteria
            comparison = f"{range_expr} == {criteria}"

        # Result is sum_range if condition is true, 0 otherwise
        result = f"np.where({comparison}, {sum_range}, 0)"

        return result, end_idx, new_fields

    def _process_averageif_function(self, tokens: List[str], start_idx: int, fields_used: List[str]) -> Tuple[str, int, List[str]]:
        """
        Process an AVERAGEIF function (average cells meeting criteria).

        Args:
            tokens: List of tokens
            start_idx: Starting index of the function name
            fields_used: List to track field names

        Returns:
            Tuple of (processed function string, new position index, new fields found)
        """
        new_fields = []

        # Extract arguments: range, criteria, [average_range]
        args, end_idx, arg_fields = self._extract_variable_args(
            tokens, start_idx + 1, 2, 3, fields_used
        )

        new_fields.extend(arg_fields)

        range_expr = args[0]
        criteria = args[1]
        avg_range = args[2] if len(args) > 2 else range_expr

        # Parse criteria - similar to COUNTIF and SUMIF
        if criteria.startswith('"') or criteria.startswith("'"):
            # String criteria - remove quotes
            criteria_str = criteria.strip("'\"")

            # Check if it's a comparison operator
            if criteria_str.startswith(('=', '>', '<', '>=', '<=', '<>')):
                operator = criteria_str[0]
                if criteria_str.startswith(('>=', '<=', '<>')):
                    operator = criteria_str[:2]
                    value = criteria_str[2:]
                else:
                    value = criteria_str[1:]

                # Map to Python operator
                if operator == '=':
                    comparison = f"{range_expr} == {value}"
                elif operator == '<>':
                    comparison = f"{range_expr} != {value}"
                else:
                    comparison = f"{range_expr} {operator} {value}"
            elif '*' in criteria_str:
                # Wildcard match
                pattern = criteria_str.replace('*', '.*')
                comparison = f"{range_expr}.astype(str).str.match(r'{pattern}')"
            else:
                # Exact match
                comparison = f"{range_expr} == '{criteria_str}'"
        else:
            # Numeric or field criteria
            comparison = f"{range_expr} == {criteria}"

        # For average, we need to calculate sum and count
        # Result is avg_range if condition is true, 0 otherwise for sum
        # and 1 if condition is true, 0 otherwise for count
        # Then divide sum by count (or return 0 if count is 0)
        result = (
            f"np.where("
            f"  np.sum(np.where({comparison}, 1, 0)) > 0, "
            f"  np.sum(np.where({comparison}, {avg_range}, 0)) / np.sum(np.where({comparison}, 1, 0)), "
            f"  0"
            f")"
        )

        return result, end_idx, new_fields

    # === Helper Methods for Function Processing ===

    def _extract_function_args(self, tokens: List[str], start_idx: int, num_args: int, fields_used: List[str]) -> Tuple[str, ...]:
        """
        Extract a fixed number of function arguments.

        Args:
            tokens: List of tokens
            start_idx: Starting index (usually after function name)
            num_args: Number of arguments to extract
            fields_used: List to track field names

        Returns:
            Tuple of arguments as strings, ending index, and new fields found
        """
        args = []
        new_fields = []

        # Skip to opening parenthesis
        i = start_idx
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error(f"No opening parenthesis found for function")
            placeholder_args = ["pd.Series(np.nan, index=df.index)"] * num_args
            return (*placeholder_args, i, new_fields)

        i += 1  # Skip past opening parenthesis

        # Extract each argument
        for arg_num in range(num_args):
            arg_tokens = []
            paren_level = 0

            # Collect tokens for this argument
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    arg_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function
                        break
                    paren_level -= 1
                    arg_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # End of this argument
                    break
                else:
                    arg_tokens.append(token)

                i += 1

            # Process the argument tokens
            if not arg_tokens and arg_num < num_args - 1:
                # Missing argument (but not the last one)
                logger.error(f"Missing argument {arg_num + 1} for function")
                args.append("pd.Series(np.nan, index=df.index)")
            else:
                # Process tokens into a string
                arg = ""
                for token in arg_tokens:
                    if token.upper() in self.operator_map:
                        arg += self.operator_map[token.upper()]
                    else:
                        processed_token, token_fields = self._process_token(token, fields_used)
                        arg += processed_token
                        new_fields.extend(token_fields)

                args.append(arg)

            # Skip comma between arguments
            if i < len(tokens) and tokens[i] == ',':
                i += 1

            # If we reached the end of function, but we haven't gotten all arguments,
            # add placeholders for missing arguments
            if i < len(tokens) and tokens[i] == ')' and arg_num < num_args - 1:
                for _ in range(arg_num + 1, num_args):
                    logger.error(f"Missing argument {_ + 1} for function")
                    args.append("pd.Series(np.nan, index=df.index)")
                break

        # Skip closing parenthesis
        if i < len(tokens) and tokens[i] == ')':
            i += 1

        # Ensure we have the right number of arguments
        while len(args) < num_args:
            logger.error(f"Missing argument {len(args) + 1} for function")
            args.append("pd.Series(np.nan, index=df.index)")

        # Return all arguments as separate items in the tuple, plus the ending index
        return (*args, i, new_fields)

    def _extract_variable_args(self, tokens: List[str], start_idx: int, min_args: int, max_args: Optional[int], fields_used: List[str]) -> Tuple[List[str], int, List[str]]:
        """
        Extract a variable number of function arguments.

        Args:
            tokens: List of tokens
            start_idx: Starting index (usually after function name)
            min_args: Minimum number of arguments required
            max_args: Maximum number of arguments (None for unlimited)
            fields_used: List to track field names

        Returns:
            Tuple of (list of arguments as strings, ending index, and new fields found)
        """
        args = []
        new_fields = []

        # Skip to opening parenthesis
        i = start_idx
        while i < len(tokens) and tokens[i] != '(':
            i += 1

        if i >= len(tokens):
            logger.error(f"No opening parenthesis found for function")
            return [f"pd.Series(np.nan, index=df.index)"] * min_args, i, new_fields

        i += 1  # Skip past opening parenthesis

        # Extract arguments until closing parenthesis or max args reached
        arg_count = 0
        while i < len(tokens) and (max_args is None or arg_count < max_args):
            # Check if we've reached the end of the function
            if tokens[i] == ')':
                i += 1  # Skip closing parenthesis
                break

            # Extract this argument
            arg_tokens = []
            paren_level = 0

            # Collect tokens for this argument
            while i < len(tokens):
                token = tokens[i]

                if token == '(':
                    paren_level += 1
                    arg_tokens.append(token)
                elif token == ')':
                    if paren_level == 0:
                        # End of function
                        break
                    paren_level -= 1
                    arg_tokens.append(token)
                elif token == ',' and paren_level == 0:
                    # End of this argument
                    break
                else:
                    arg_tokens.append(token)

                i += 1

            # Process the argument tokens
            arg = ""
            for token in arg_tokens:
                if token.upper() in self.operator_map:
                    arg += self.operator_map[token.upper()]
                else:
                    processed_token, token_fields = self._process_token(token, fields_used)
                    arg += processed_token
                    new_fields.extend(token_fields)

            args.append(arg)
            arg_count += 1

            # Skip comma between arguments
            if i < len(tokens) and tokens[i] == ',':
                i += 1

        # Skip closing parenthesis if we didn't already
        if i < len(tokens) and tokens[i] == ')':
            i += 1

        # Ensure we have the minimum number of arguments
        if len(args) < min_args:
            logger.error(f"Not enough arguments for function, expected at least {min_args}, got {len(args)}")
            while len(args) < min_args:
                args.append("pd.Series(np.nan, index=df.index)")

        return args, i, new_fields

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
            safe_locals = {
                "df": data,
                "pd": pd,
                "np": np,
                "datetime": datetime,
                "relativedelta": relativedelta
            }

            # Log the formula for debugging
            logger.debug(f"Evaluating formula: {parsed_formula}")

            # Evaluate the formula
            result = eval(parsed_formula, restricted_globals, safe_locals)

            # Ensure result is a boolean Series
            if not isinstance(result, pd.Series):
                return False, None, f"Formula did not return a Series (got {type(result).__name__})"

            # For numeric operations, convert to boolean based on non-zero values
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
    test_formula = "IF(Risk_Level=\"High\", Days_Open<=30, Days_Open<=90)"
    parser = ExcelFormulaParser()
    result, fields = parser.parse(test_formula)
    print(f"Original: {test_formula}")
    print(f"Parsed: {result}")
    print(f"Fields used: {fields}")

    # Test with sample data
    sample_data = pd.DataFrame({
        'Risk_Level': ['High', 'Medium', 'Low', 'High'],
        'Days_Open': [20, 40, 100, 35]
    })

    success, result, error = parser.test_formula(test_formula, sample_data)
    if success:
        print("\nTest Results:")
        print(pd.DataFrame({
            'Risk_Level': sample_data['Risk_Level'],
            'Days_Open': sample_data['Days_Open'],
            'Result': result
        }))
    else:
        print(f"Error: {error}")