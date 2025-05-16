"""
Excel Formula to Pandas Expression Converter

This module provides a robust converter for transforming Excel-style formulas
into pandas expressions that can be evaluated against a DataFrame.
It uses xlcalculator for parsing and implements a visitor pattern to traverse
the syntax tree.

Example:
    converter = ExcelToPandasConverter()
    pandas_expr = converter.convert("IF(AND(Status=\"Active\", Value>100), \"High\", \"Low\")")
"""

import re
from typing import Dict, List, Set, Tuple, Any, Union, Optional
import ast
import pandas as pd
import numpy as np

# Import xlcalculator components
try:
    from xlcalculator import ModelCompiler, Model, Evaluator
    from xlcalculator.xlfunctions import xl
    import xlcalculator.xlfunctions as xlfunctions
except ImportError:
    raise ImportError(
        "xlcalculator is required for this module. "
        "Install it using: pip install xlcalculator"
    )


class ColumnMapper:
    """
    Handles bidirectional mapping between DataFrame column names and 
    Excel-compatible identifiers.
    
    This class is responsible for:
    1. Converting DataFrame column names to Excel-safe identifiers
    2. Maintaining a mapping to translate back to original column names
    3. Handling columns with spaces and special characters
    """
    
    def __init__(self):
        """Initialize the column mapper."""
        self.df_to_excel = {}  # Maps DataFrame column name to Excel identifier
        self.excel_to_df = {}  # Maps Excel identifier to DataFrame column name
        
    def register_columns(self, columns: List[str]) -> None:
        """
        Register DataFrame column names and create Excel-safe identifiers.
        
        Args:
            columns: List of DataFrame column names to register
        """
        for col in columns:
            if col not in self.df_to_excel:
                # Create Excel-safe identifier
                excel_name = self._create_excel_name(col)
                
                # Ensure uniqueness by adding a suffix if needed
                base_name = excel_name
                counter = 1
                while excel_name in self.excel_to_df:
                    excel_name = f"{base_name}_{counter}"
                    counter += 1
                
                # Store the mapping both ways
                self.df_to_excel[col] = excel_name
                self.excel_to_df[excel_name] = col
    
    def _create_excel_name(self, column_name: str) -> str:
        """
        Create an Excel-safe identifier from a DataFrame column name.
        
        Args:
            column_name: DataFrame column name
            
        Returns:
            Excel-safe identifier
        """
        # Replace spaces and special characters
        safe_name = re.sub(r'[^a-zA-Z0-9_]', '_', column_name)
        
        # Ensure it starts with a letter
        if not safe_name[0].isalpha() and safe_name[0] != '_':
            safe_name = 'col_' + safe_name
            
        return safe_name
    
    def to_excel_name(self, df_column: str) -> str:
        """
        Convert DataFrame column name to Excel identifier.
        
        Args:
            df_column: DataFrame column name
            
        Returns:
            Excel-safe identifier
        """
        if df_column not in self.df_to_excel:
            raise ValueError(f"Column '{df_column}' not registered")
        
        return self.df_to_excel[df_column]
    
    def to_df_name(self, excel_name: str) -> str:
        """
        Convert Excel identifier to DataFrame column name.
        
        Args:
            excel_name: Excel identifier
            
        Returns:
            Original DataFrame column name
        """
        if excel_name not in self.excel_to_df:
            # It might be a literal or non-column reference
            return excel_name
        
        return self.excel_to_df[excel_name]


class ExcelToPandasConverter:
    """
    Converts Excel formulas to pandas expressions using xlcalculator.
    
    This class parses Excel formulas and generates equivalent pandas code
    that can be evaluated against a DataFrame.
    """
    
    def __init__(self):
        """Initialize the converter."""
        self.column_mapper = ColumnMapper()
        self.fields_used = set()  # Set of DataFrame columns used in the formula
        
        # Function mapping from Excel to pandas
        self.function_map = {
            'IF': self._translate_if,
            'AND': self._translate_and,
            'OR': self._translate_or,
            'NOT': self._translate_not,
            'ISBLANK': self._translate_isblank,
            'ISERROR': self._translate_iserror,
            'ISNUMBER': self._translate_isnumber,
            'COUNT': self._translate_count,
            'COUNTIF': self._translate_countif,
            'SUM': self._translate_sum,
            'SUMIF': self._translate_sumif,
            'AVERAGE': self._translate_average,
            'MIN': self._translate_min,
            'MAX': self._translate_max,
            'LEFT': self._translate_left,
            'RIGHT': self._translate_right,
            'MID': self._translate_mid,
            'LEN': self._translate_len,
            'CONCATENATE': self._translate_concatenate,
            'TODAY': self._translate_today,
            'NOW': self._translate_now,
        }
        
        # Operator mapping from Excel to pandas
        self.operator_map = {
            '=': '==',
            '<>': '!=',
            '&': '+',  # String concatenation
        }
    
    def convert(self, formula: str, df_columns: Optional[List[str]] = None) -> Tuple[str, Set[str]]:
        """
        Convert an Excel formula to a pandas expression.
        
        Args:
            formula: Excel formula to convert
            df_columns: Optional list of DataFrame column names
            
        Returns:
            Tuple containing:
            - Pandas expression string
            - Set of DataFrame column names used in the formula
        """
        # Reset fields used
        self.fields_used = set()
        
        # Register columns if provided
        if df_columns:
            self.column_mapper.register_columns(df_columns)
        
        try:
            # Create a mock workbook with the formula
            formula_cell = f'=({formula})'  # Wrap in parentheses for better parsing
            
            # Create a model with the formula
            compiler = ModelCompiler()
            model = compiler.read_excel_formula_cells(formula_cells={
                'Sheet1!A1': formula_cell
            })
            
            # Get the parsed formula
            formula_ast = model.cells['Sheet1!A1'].formula.tokens
            
            # Translate the formula AST to pandas
            pandas_expr = self._translate_node(formula_ast)
            
            # Wrap in a Series constructor to ensure proper DataFrame integration
            final_expr = f"pd.Series({pandas_expr}, index=df.index)"
            
            return final_expr, self.fields_used
            
        except Exception as e:
            # Handle parsing errors
            error_msg = f"Error converting formula: {formula}. {str(e)}"
            raise ValueError(error_msg)
    
    def _translate_node(self, node) -> str:
        """
        Recursively translate a node in the formula AST to pandas code.
        
        Args:
            node: AST node from xlcalculator
            
        Returns:
            Pandas code fragment as a string
        """
        if isinstance(node, list):
            # Handle function calls
            if len(node) > 0 and isinstance(node[0], str):
                function_name = node[0].upper()
                
                if function_name in self.function_map:
                    # Use the specific translator for this function
                    return self.function_map[function_name](node)
                else:
                    # Handle unknown function
                    args = [self._translate_node(arg) for arg in node[1:]]
                    return f"{function_name.lower()}({', '.join(args)})"
            
            # Handle general expressions
            if len(node) == 3:  # Binary operation
                left = self._translate_node(node[0])
                operator = node[1]
                right = self._translate_node(node[2])
                
                # Map Excel operators to pandas operators
                if operator in self.operator_map:
                    operator = self.operator_map[operator]
                
                return f"({left} {operator} {right})"
            
            # Handle other list structures
            return str(node)
            
        elif isinstance(node, (int, float)):
            # Handle numeric literals
            return str(node)
            
        elif isinstance(node, str):
            if node.startswith('"') and node.endswith('"'):
                # String literal
                return node
            elif node in self.column_mapper.excel_to_df:
                # Column reference - convert to df[] access
                df_column = self.column_mapper.to_df_name(node)
                self.fields_used.add(df_column)
                return f"df['{df_column}']"
            else:
                # Could be an Excel name or reference not in our mapping
                # For now, treat as literal value
                if node.isdigit() or (node[0] == '-' and node[1:].isdigit()):
                    return node
                else:
                    # Assume it's a string literal if not a number
                    # Check if it's already quoted
                    if not (node.startswith('"') and node.endswith('"')):
                        return f'"{node}"'
                    return node
        
        # Handle other types
        return str(node)
    
    def _translate_if(self, node) -> str:
        """Translate IF function to numpy.where."""
        if len(node) < 4:
            raise ValueError("IF function requires at least 3 arguments")
        
        condition = self._translate_node(node[1])
        true_value = self._translate_node(node[2])
        false_value = self._translate_node(node[3]) if len(node) > 3 else '"False"'
        
        return f"np.where({condition}, {true_value}, {false_value})"
    
    def _translate_and(self, node) -> str:
        """Translate AND function to bitwise &."""
        if len(node) < 2:
            raise ValueError("AND function requires at least 1 argument")
        
        conditions = [self._translate_node(arg) for arg in node[1:]]
        
        # Ensure each condition is wrapped as a Series for bitwise operations
        conditions = [f"pd.Series({cond}, index=df.index).astype(bool)" for cond in conditions]
        
        return " & ".join(conditions)
    
    def _translate_or(self, node) -> str:
        """Translate OR function to bitwise |."""
        if len(node) < 2:
            raise ValueError("OR function requires at least 1 argument")
        
        conditions = [self._translate_node(arg) for arg in node[1:]]
        
        # Ensure each condition is wrapped as a Series for bitwise operations
        conditions = [f"pd.Series({cond}, index=df.index).astype(bool)" for cond in conditions]
        
        return " | ".join(conditions)
    
    def _translate_not(self, node) -> str:
        """Translate NOT function to unary ~."""
        if len(node) != 2:
            raise ValueError("NOT function requires exactly 1 argument")
        
        condition = self._translate_node(node[1])
        
        # Ensure the condition is wrapped as a Series for bitwise operation
        return f"~pd.Series({condition}, index=df.index).astype(bool)"
    
    def _translate_isblank(self, node) -> str:
        """Translate ISBLANK function to pd.isna."""
        if len(node) != 2:
            raise ValueError("ISBLANK function requires exactly 1 argument")
        
        value = self._translate_node(node[1])
        return f"pd.isna({value})"
    
    def _translate_iserror(self, node) -> str:
        """Translate ISERROR function to custom error checking."""
        if len(node) != 2:
            raise ValueError("ISERROR function requires exactly 1 argument")
        
        value = self._translate_node(node[1])
        # This is a simplified version - a more complete version would handle more error types
        return f"pd.Series([isinstance(x, (ValueError, TypeError, ZeroDivisionError)) for x in {value}], index=df.index)"
    
    def _translate_isnumber(self, node) -> str:
        """Translate ISNUMBER function to pd.to_numeric with error handling."""
        if len(node) != 2:
            raise ValueError("ISNUMBER function requires exactly 1 argument")
        
        value = self._translate_node(node[1])
        return f"pd.to_numeric({value}, errors='coerce').notna()"
    
    def _translate_count(self, node) -> str:
        """Translate COUNT function."""
        if len(node) < 2:
            raise ValueError("COUNT function requires at least 1 argument")
        
        # Count non-NA values across multiple columns/ranges
        ranges = [self._translate_node(arg) for arg in node[1:]]
        counts = [f"{r}.count()" for r in ranges]
        
        return " + ".join(counts)
    
    def _translate_countif(self, node) -> str:
        """Translate COUNTIF function."""
        if len(node) != 3:
            raise ValueError("COUNTIF function requires exactly 2 arguments")
        
        range_expr = self._translate_node(node[1])
        criteria = self._translate_node(node[2])
        
        # Handle various criteria formats
        if criteria.startswith('"') and criteria.endswith('"'):
            # String criteria - remove quotes
            criteria_str = criteria.strip('"')
            
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
            else:
                # Exact match
                comparison = f"{range_expr} == {criteria}"
        else:
            # Numeric or field criteria
            comparison = f"{range_expr} == {criteria}"
        
        # Count True values
        return f"({comparison}).sum()"
    
    def _translate_sum(self, node) -> str:
        """Translate SUM function."""
        if len(node) < 2:
            raise ValueError("SUM function requires at least 1 argument")
        
        # Sum values across multiple columns/ranges
        ranges = [self._translate_node(arg) for arg in node[1:]]
        
        return " + ".join([f"{r}.sum()" for r in ranges])
    
    def _translate_sumif(self, node) -> str:
        """Translate SUMIF function."""
        if len(node) < 3 or len(node) > 4:
            raise ValueError("SUMIF function requires 2 or 3 arguments")
        
        range_expr = self._translate_node(node[1])
        criteria = self._translate_node(node[2])
        
        # If 3 arguments, use the third as sum_range, otherwise use range
        sum_range = self._translate_node(node[3]) if len(node) > 3 else range_expr
        
        # Similar criteria handling as COUNTIF
        if criteria.startswith('"') and criteria.endswith('"'):
            criteria_str = criteria.strip('"')
            if criteria_str.startswith(('=', '>', '<', '>=', '<=', '<>')):
                operator = criteria_str[0]
                if criteria_str.startswith(('>=', '<=', '<>')):
                    operator = criteria_str[:2]
                    value = criteria_str[2:]
                else:
                    value = criteria_str[1:]
                
                if operator == '=':
                    comparison = f"{range_expr} == {value}"
                elif operator == '<>':
                    comparison = f"{range_expr} != {value}"
                else:
                    comparison = f"{range_expr} {operator} {value}"
            else:
                comparison = f"{range_expr} == {criteria}"
        else:
            comparison = f"{range_expr} == {criteria}"
        
        # Sum values where condition is True
        return f"({sum_range}[{comparison}]).sum()"
    
    def _translate_average(self, node) -> str:
        """Translate AVERAGE function."""
        if len(node) < 2:
            raise ValueError("AVERAGE function requires at least 1 argument")
        
        ranges = [self._translate_node(arg) for arg in node[1:]]
        
        # Concatenate Series for multi-range average
        if len(ranges) == 1:
            return f"{ranges[0]}.mean()"
        else:
            return f"pd.concat([{', '.join(ranges)}]).mean()"
    
    def _translate_min(self, node) -> str:
        """Translate MIN function."""
        if len(node) < 2:
            raise ValueError("MIN function requires at least 1 argument")
        
        ranges = [self._translate_node(arg) for arg in node[1:]]
        
        if len(ranges) == 1:
            return f"{ranges[0]}.min()"
        else:
            return f"pd.concat([{', '.join(ranges)}]).min()"
    
    def _translate_max(self, node) -> str:
        """Translate MAX function."""
        if len(node) < 2:
            raise ValueError("MAX function requires at least 1 argument")
        
        ranges = [self._translate_node(arg) for arg in node[1:]]
        
        if len(ranges) == 1:
            return f"{ranges[0]}.max()"
        else:
            return f"pd.concat([{', '.join(ranges)}]).max()"
    
    def _translate_left(self, node) -> str:
        """Translate LEFT function."""
        if len(node) < 2 or len(node) > 3:
            raise ValueError("LEFT function requires 1 or 2 arguments")
        
        text = self._translate_node(node[1])
        num_chars = self._translate_node(node[2]) if len(node) > 2 else "1"
        
        return f"({text}.astype(str).str[:int({num_chars})])"
    
    def _translate_right(self, node) -> str:
        """Translate RIGHT function."""
        if len(node) < 2 or len(node) > 3:
            raise ValueError("RIGHT function requires 1 or 2 arguments")
        
        text = self._translate_node(node[1])
        num_chars = self._translate_node(node[2]) if len(node) > 2 else "1"
        
        return f"({text}.astype(str).str[-int({num_chars}):])"
    
    def _translate_mid(self, node) -> str:
        """Translate MID function."""
        if len(node) != 4:
            raise ValueError("MID function requires exactly 3 arguments")
        
        text = self._translate_node(node[1])
        start_pos = self._translate_node(node[2])
        num_chars = self._translate_node(node[3])
        
        # Adjust for 1-based indexing in Excel
        adjusted_start = f"(int({start_pos}) - 1)"
        
        return f"({text}.astype(str).str[{adjusted_start}:({adjusted_start} + int({num_chars}))])"
    
    def _translate_len(self, node) -> str:
        """Translate LEN function."""
        if len(node) != 2:
            raise ValueError("LEN function requires exactly 1 argument")
        
        text = self._translate_node(node[1])
        return f"({text}.astype(str).str.len())"
    
    def _translate_concatenate(self, node) -> str:
        """Translate CONCATENATE function."""
        if len(node) < 2:
            raise ValueError("CONCATENATE function requires at least 1 argument")
        
        texts = [self._translate_node(arg) for arg in node[1:]]
        
        # Convert all arguments to strings and concatenate
        texts = [f"({text}).astype(str)" for text in texts]
        
        return " + ".join(texts)
    
    def _translate_today(self, node) -> str:
        """Translate TODAY function."""
        return "pd.Timestamp.today().normalize()"
    
    def _translate_now(self, node) -> str:
        """Translate NOW function."""
        return "pd.Timestamp.now()"


# Demo and example usage
def demo_converter():
    """Demonstrate the converter with examples."""
    import pandas as pd
    import numpy as np
    
    # Create sample DataFrame
    data = {
        "Status": ["Active", "Inactive", "Active", "On Hold"],
        "Value": [120, 80, 200, 50],
        "Risk Level": ["High", "Low", "Medium", "Low"],
        "Start Date": pd.to_datetime(["2023-01-15", "2022-11-10", "2023-03-22", "2023-02-05"]),
        "Helper-KPA Contains Key TLM Third Party": ["Yes", "No", "Yes", "No"],
        "PRIMARY TLM THIRD PARTY ENGAGEMENT": [None, "ABC Corp", "XYZ Inc", None]
    }
    df = pd.DataFrame(data)
    
    # Print sample data
    print("Sample DataFrame:")
    print(df)
    print()
    
    # Create converter
    converter = ExcelToPandasConverter()
    
    # Register DataFrame columns
    converter.column_mapper.register_columns(df.columns)
    
    # Example formulas to convert
    formulas = [
        'IF(Status="Active", "Yes", "No")',
        'IF(AND(Value>100, Risk Level="High"), "Critical", "Normal")',
        'IF(OR(Value<60, Status="Inactive"), "Review", "Pass")',
        'IF(ISBLANK(PRIMARY TLM THIRD PARTY ENGAGEMENT), "Missing", "Complete")',
        'IF(AND(Helper-KPA Contains Key TLM Third Party="Yes", ISBLANK(PRIMARY TLM THIRD PARTY ENGAGEMENT)), "DNC", "GC")'
    ]
    
    # Convert each formula and evaluate
    for formula in formulas:
        print(f"\nExcel Formula: {formula}")
        
        # Convert the formula
        pandas_expr, fields = converter.convert(formula, df.columns)
        
        print(f"Pandas Expression: {pandas_expr}")
        print(f"Fields Used: {fields}")
        
        try:
            # Safely evaluate the expression with the DataFrame
            result = eval(pandas_expr, {"__builtins__": {}}, {"df": df, "pd": pd, "np": np})
            
            # Print the result
            print("Result:")
            print(result)
        except Exception as e:
            print(f"Error evaluating expression: {e}")


if __name__ == "__main__":
    # Run the demo when the module is executed directly
    demo_converter()
