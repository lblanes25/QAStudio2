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
import ast
import logging
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any, Union

# Set up logging
from logging_config import setup_logging
logger = setup_logging()


class ExcelFormulaParser:
    """
    Parser to convert Excel-style formulas to pandas expressions.
    
    This class handles the translation of familiar Excel syntax to
    pandas operations that can be safely evaluated against DataFrames.
    """
    
    # Define operator mappings
    OPERATOR_MAP = {
        '=': '==',
        '<>': '!=',
        'AND': '&',
        'OR': '|',
        'NOT': '~',
    }
    
    # Define Excel function mappings
    FUNCTION_MAP = {
        'ISBLANK': 'pd.isna',
        'ISNUMBER': 'pd.to_numeric({}, errors="coerce").notna()',
        'ISTEXT': 'pd.Series([], dtype="object").dtype == {}.dtype',
        'TODAY': 'pd.Timestamp.today()',
        # More functions will be added here
    }
    
    def __init__(self):
        """Initialize the Excel formula parser."""
        # Compiled regex patterns
        self.field_pattern = re.compile(r'`([^`]+)`|([a-zA-Z][a-zA-Z0-9_]*)')
        self.operator_pattern = re.compile(r'(<>|=|<=|>=|<|>|\+|-|\*|/|\(|\))')
        self.function_pattern = re.compile(r'\b([A-Z][A-Z0-9_]*)\(')
        
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
        
        # Process the formula in steps
        cleaned_formula = self._clean_formula(formula)
        tokenized_formula = self._tokenize(cleaned_formula)
        parsed_formula = self._parse_tokens(tokenized_formula, fields_used)
        
        logger.info(f"Parsed formula: {formula} -> {parsed_formula}")
        return parsed_formula, fields_used
    
    def _clean_formula(self, formula: str) -> str:
        """
        Clean and normalize the formula for processing.
        
        Args:
            formula: Original Excel-style formula
            
        Returns:
            Cleaned formula string
        """
        # TODO: Implement formula cleaning
        # - Remove extra whitespace
        # - Standardize line breaks
        # - Handle special characters
        
        return formula.strip()
    
    def _tokenize(self, formula: str) -> List[str]:
        """
        Break the formula into tokens for parsing.
        
        Args:
            formula: Cleaned Excel-style formula
            
        Returns:
            List of formula tokens
        """
        # Process formula to separate operators, fields, and literals
        tokens = []
        i = 0
        
        while i < len(formula):
            char = formula[i]
            
            # Skip whitespace
            if char.isspace():
                i += 1
                continue
                
            # Handle backtick-enclosed field names
            if char == '`':
                # Find the closing backtick
                end = formula.find('`', i + 1)
                if end == -1:
                    raise ValueError(f"Unclosed backtick in formula: {formula}")
                    
                # Extract the field name
                field_name = formula[i:end+1]
                tokens.append(field_name)
                i = end + 1
                
            # Handle string literals
            elif char == '"' or char == "'":
                # Find the closing quote
                end = formula.find(char, i + 1)
                if end == -1:
                    raise ValueError(f"Unclosed string in formula: {formula}")
                    
                # Extract the string literal
                string_literal = formula[i:end+1]
                tokens.append(string_literal)
                i = end + 1
                
            # Handle operators
            elif char in "=<>!&|+-*/()," or formula[i:i+2] in ("<=", ">=", "<>", "==", "!=", "AND", "OR"):
                # Check for multi-character operators
                if i < len(formula) - 1:
                    op2 = formula[i:i+2]
                    if op2 in ("<=", ">=", "<>", "==", "!="):
                        tokens.append(op2)
                        i += 2
                        continue
                    
                    # Handle logical operators (case-insensitive)
                    word = ""
                    j = i
                    while j < len(formula) and formula[j].isalpha():
                        word += formula[j]
                        j += 1
                    
                    if word.upper() == "AND":
                        tokens.append("AND")
                        i += 3
                        continue
                    elif word.upper() == "OR":
                        tokens.append("OR")
                        i += 2
                        continue
                    elif word.upper() == "NOT":
                        tokens.append("NOT")
                        i += 3
                        continue
                    elif word.upper() == "IN":
                        tokens.append("IN")
                        i += 2
                        continue
                
                # Single character operator
                tokens.append(char)
                i += 1
                
            # Handle functions and field names
            elif char.isalpha() or char == '_':
                # Extract the identifier
                j = i
                while j < len(formula) and (formula[j].isalnum() or formula[j] == '_'):
                    j += 1
                
                identifier = formula[i:j]
                tokens.append(identifier)
                i = j
                
            # Handle numbers
            elif char.isdigit() or char == '.':
                # Extract the number
                j = i
                while j < len(formula) and (formula[j].isdigit() or formula[j] == '.'):
                    j += 1
                
                number = formula[i:j]
                tokens.append(number)
                i = j
                
            else:
                # Unknown character
                raise ValueError(f"Unknown character in formula: {char}")
        
        return tokens
    
    def _parse_tokens(self, tokens: List[str], fields_used: List[str]) -> str:
        """
        Parse tokenized formula into pandas expression.
        
        Args:
            tokens: List of formula tokens
            fields_used: List to collect field names (modified in-place)
            
        Returns:
            Parsed pandas expression
        """
        if not tokens:
            return ""
            
        # Process tokens
        result = []
        i = 0
        
        while i < len(tokens):
            token = tokens[i]
            
            # Handle operators
            if token in self.OPERATOR_MAP:
                result.append(self.OPERATOR_MAP[token])
                
            # Handle parentheses and other symbols
            elif token in "()+-*/,":
                result.append(token)
                
            # Handle backtick-enclosed field names
            elif token.startswith('`') and token.endswith('`'):
                field_name = token[1:-1]  # Remove backticks
                result.append(f"df['{field_name}']")
                if field_name not in fields_used:
                    fields_used.append(field_name)
                
            # Handle string literals
            elif (token.startswith('"') and token.endswith('"')) or (token.startswith("'") and token.endswith("'")):
                result.append(token)
                
            # Handle functions
            elif token.upper() in self.FUNCTION_MAP and i + 1 < len(tokens) and tokens[i+1] == '(':
                func_name = token.upper()
                result.append(self.FUNCTION_MAP[func_name])
                
                # Skip the opening parenthesis as it's handled by the function mapping
                i += 1
                
            # Handle IN operator
            elif token.upper() == "IN":
                result.append(".isin")
                
            # Handle field names (identifiers)
            elif token.isalnum() or token.startswith('_'):
                # Check if this is a field name (not a keyword or function)
                if token.upper() not in ("AND", "OR", "NOT", "IN") and not token.upper() in self.FUNCTION_MAP:
                    result.append(f"df['{token}']")
                    if token not in fields_used:
                        fields_used.append(token)
                else:
                    # Handle any missed keywords
                    if token.upper() in self.OPERATOR_MAP:
                        result.append(self.OPERATOR_MAP[token.upper()])
                    else:
                        result.append(token)
            
            # Handle numbers
            elif token.replace('.', '', 1).isdigit():
                result.append(token)
                
            else:
                # Unknown token
                raise ValueError(f"Unknown token in formula: {token}")
                
            i += 1
        
        # Join all tokens
        parsed = ''.join(result)
        
        # Wrap in parentheses for safety
        if not (parsed.startswith('(') and parsed.endswith(')')):
            parsed = f"({parsed})"
            
        return parsed
    
    def validate_formula(self, formula: str, available_fields: List[str] = None) -> Tuple[bool, Optional[str]]:
        """
        Validate that a formula is correctly formed and uses available fields.
        
        Args:
            formula: Excel-style formula to validate
            available_fields: Optional list of available field names in the data
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # TODO: Implement formula validation
        # - Check syntax correctness
        # - Validate all referenced fields exist if available_fields is provided
        # - Validate functions are properly formed
        # - Check for balanced parentheses
        
        return True, None  # Placeholder
    
    def test_formula(self, formula: str, sample_data: pd.DataFrame) -> Tuple[bool, pd.Series, Optional[str]]:
        """
        Test a formula against sample data to verify its behavior.
        
        Args:
            formula: Excel-style formula to test
            sample_data: DataFrame to test against
            
        Returns:
            Tuple of (success, result_series, error_message)
        """
        if not formula:
            return False, pd.Series(False, index=sample_data.index), "Empty formula"
            
        try:
            # First validate the formula
            valid, error_msg = self.validate_formula(formula, sample_data.columns.tolist())
            if not valid:
                return False, pd.Series(False, index=sample_data.index), error_msg
            
            # Parse the formula
            parsed_formula, fields_used = self.parse(formula)
            
            # Check that all required fields exist in sample data
            missing_fields = [field for field in fields_used if field not in sample_data.columns]
            if missing_fields:
                return False, pd.Series(False, index=sample_data.index), f"Fields not found in data: {', '.join(missing_fields)}"
            
            # Create a safe environment for evaluation
            restricted_globals = {"__builtins__": {}}
            safe_locals = {
                "df": sample_data,
                "pd": pd,
                "np": np
            }
            
            # Evaluate the formula
            result = eval(parsed_formula, restricted_globals, safe_locals)
            
            # Validate the result type
            if not isinstance(result, pd.Series):
                return False, pd.Series(False, index=sample_data.index), f"Formula did not return a Series (got {type(result).__name__})"
                
            # Try to convert to boolean if not already
            if result.dtype != bool:
                try:
                    result = result.astype(bool)
                except:
                    return False, pd.Series(False, index=sample_data.index), f"Formula result cannot be converted to boolean (dtype: {result.dtype})"
            
            # Calculate quick stats for debugging
            total = len(result)
            passing = result.sum()
            failing = total - passing
            
            logger.info(f"Formula test: {passing} of {total} records pass ({passing/total*100:.1f}% success rate)")
                
            return True, result, None
            
        except Exception as e:
            logger.error(f"Error testing formula: {e}")
            return False, pd.Series(False, index=sample_data.index), f"Formula test error: {str(e)}"
    
    def explain_formula(self, formula: str) -> str:
        """
        Convert a formula to plain English explanation.
        
        Args:
            formula: Excel-style formula
            
        Returns:
            Plain English explanation of the formula
        """
        # TODO: Implement formula explanation
        # - Convert operators to words
        # - Explain what the formula is checking
        # - Format in readable English
        
        return f"This formula checks if the values meet certain conditions."


def parse_excel_formula(formula: str) -> str:
    """
    Convenience function to parse an Excel formula to a pandas expression.
    
    Args:
        formula: Excel-style formula string
        
    Returns:
        Pandas expression string
    """
    parser = ExcelFormulaParser()
    result, _ = parser.parse(formula)
    return result


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
    import pandas as pd
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
