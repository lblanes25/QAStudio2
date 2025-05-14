"""
Excel formula validation rule for the QA Analytics Framework.

This extension adds custom formula support to the ValidationRules class,
allowing users to define validations using Excel-style formulas.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional
import logging

from excel_formula_parser import ExcelFormulaParser

# Get existing logger
logger = logging.getLogger("qa_analytics")


class CustomFormulaValidation:
    """
    Add-on to ValidationRules that provides custom Excel formula support.
    
    This class can be integrated with the existing ValidationRules class to
    add support for Excel-style formulas in the validation framework.
    """
    
    @staticmethod
    def custom_formula(df: pd.DataFrame, params: Dict) -> pd.Series:
        """
        Execute a user-defined Excel formula against the dataframe.
        
        Args:
            df: DataFrame containing the data to validate
            params: Dictionary with formula parameters:
                - formula: Pandas expression (parsed from original_formula)
                - original_formula: Original Excel-style formula
                
        Returns:
            Series with True for rows that conform, False for non-conforming
        """
        try:
            # Get formula and original formula
            formula = params.get('formula')
            original = params.get('original_formula', 'Unknown formula')
            
            if not formula:
                logger.error("Missing formula parameter")
                return pd.Series(False, index=df.index)
                
            # Use safe evaluation approach
            restricted_globals = {"__builtins__": {}}
            safe_locals = {"df": df, "pd": pd, "np": np}
            
            # Execute formula
            result = eval(formula, restricted_globals, safe_locals)
            
            # Ensure result is a boolean Series
            if not isinstance(result, pd.Series):
                logger.error(f"Formula did not return a Series: {original}")
                return pd.Series(False, index=df.index)
                
            if result.dtype != bool:
                logger.error(f"Formula did not return boolean values: {original}")
                return pd.Series(False, index=df.index)
                
            return result
            
        except Exception as e:
            logger.error(f"Custom formula failed: {e}, Formula: {params.get('original_formula', 'Unknown')}")
            return pd.Series(False, index=df.index)


# Example implementation in ValidationRules class:
"""
from validation_rules import ValidationRules

# Add the custom_formula method to ValidationRules
ValidationRules.custom_formula = CustomFormulaValidation.custom_formula

# Now ValidationRules has a custom_formula method that can be used in configurations
"""


# Example of integration with existing data processor:
def process_custom_formula_rule(rule_config: Dict, data_processor) -> None:
    """
    Process a custom formula rule in the enhanced data processor.
    
    This function demonstrates how to process a custom formula rule in the
    enhanced_data_processor.py file.
    
    Args:
        rule_config: Rule configuration from YAML
        data_processor: EnhancedDataProcessor instance
    """
    # Get the original Excel formula
    original_formula = rule_config.get('parameters', {}).get('original_formula')
    
    if not original_formula:
        logger.error("Missing original_formula parameter in custom_formula rule")
        return
    
    # Parse the formula
    parser = ExcelFormulaParser()
    parsed_formula, fields_used = parser.parse(original_formula)
    
    # Update the rule configuration with the parsed formula
    rule_config['parameters']['formula'] = parsed_formula
    
    # Add fields used to the documentation for reference
    rule_config['metadata'] = {
        'fields_used': fields_used
    }
    
    # Log the transformation
    logger.info(f"Processed custom formula: {original_formula} -> {parsed_formula}")


# Example of how a custom formula rule would look in a YAML configuration:
"""
validations:
  - rule: custom_formula
    description: "Submitter must be different from approvers and proper approval sequence"
    parameters:
      original_formula: "Submitter <> Approver AND `Submit Date` <= `TL Date`"
      # formula will be populated during processing
"""


# Example of testing a custom formula with sample data:
def test_custom_formula(original_formula: str, sample_data: pd.DataFrame) -> Dict:
    """
    Test a custom Excel formula with sample data.
    
    Args:
        original_formula: Excel-style formula to test
        sample_data: Sample DataFrame to test against
        
    Returns:
        Dictionary with test results
    """
    parser = ExcelFormulaParser()
    parsed_formula, fields_used = parser.parse(original_formula)
    
    # Validate all required fields exist
    missing_fields = [field for field in fields_used if field not in sample_data.columns]
    if missing_fields:
        return {
            'success': False,
            'error': f"Formula references fields not in the data: {', '.join(missing_fields)}",
            'parsed_formula': parsed_formula,
            'fields_used': fields_used
        }
    
    try:
        # Create a rule configuration
        rule_config = {
            'parameters': {
                'formula': parsed_formula,
                'original_formula': original_formula
            }
        }
        
        # Execute the custom formula rule
        result = CustomFormulaValidation.custom_formula(sample_data, rule_config['parameters'])
        
        # Get passing and failing examples
        passing = sample_data[result].head(3) if not result.empty else pd.DataFrame()
        failing = sample_data[~result].head(3) if not (~result).empty else pd.DataFrame()
        
        # Calculate statistics
        total_records = len(sample_data)
        passing_count = result.sum()
        failing_count = total_records - passing_count
        passing_pct = (passing_count / total_records * 100) if total_records > 0 else 0
        
        return {
            'success': True,
            'parsed_formula': parsed_formula,
            'fields_used': fields_used,
            'total_records': total_records,
            'passing_count': int(passing_count),
            'failing_count': int(failing_count),
            'passing_percentage': f"{passing_pct:.1f}%",
            'passing_examples': passing.to_dict('records'),
            'failing_examples': failing.to_dict('records')
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'parsed_formula': parsed_formula,
            'fields_used': fields_used
        }


# Example usage
if __name__ == "__main__":
    # Create sample data
    sample_data = pd.DataFrame({
        'Submitter': ['John', 'Mary', 'John', 'Bob'],
        'Approver': ['Alice', 'John', 'John', 'Charlie'],
        'Submit Date': pd.to_datetime(['2025-01-01', '2025-02-01', '2025-03-01', '2025-04-01']),
        'TL Date': pd.to_datetime(['2025-01-05', '2025-01-15', '2025-02-01', '2025-05-01'])
    })
    
    # Test a custom formula
    test_formula = "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    result = test_custom_formula(test_formula, sample_data)
    
    if result['success']:
        print(f"Formula: {test_formula}")
        print(f"Parsed: {result['parsed_formula']}")
        print(f"Fields used: {', '.join(result['fields_used'])}")
        print(f"Results: {result['passing_count']} of {result['total_records']} records pass ({result['passing_percentage']})")
        
        if result['passing_examples']:
            print("\nPassing Examples:")
            for idx, example in enumerate(result['passing_examples'], 1):
                print(f"  {idx}. {example}")
        
        if result['failing_examples']:
            print("\nFailing Examples:")
            for idx, example in enumerate(result['failing_examples'], 1):
                print(f"  {idx}. {example}")
    else:
        print(f"Error: {result['error']}")
