"""
Test module for Excel Formula feature in QA Analytics Framework.

This script demonstrates the Excel Formula parsing and validation features
by running a series of test cases.
"""

import pandas as pd
import numpy as np
import logging
from excel_formula_parser import ExcelFormulaParser
from custom_formula_validation import CustomFormulaValidation, test_custom_formula

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("excel_formula_test")


def create_test_data():
    """Create test data for demonstration."""
    return pd.DataFrame({
        'Submitter': ['John', 'Mary', 'John', 'Bob', 'Alice'],
        'Approver': ['Alice', 'John', 'John', 'Charlie', 'Bob'],
        'Submit Date': pd.to_datetime(['2025-01-01', '2025-02-01', '2025-03-01', '2025-04-01', '2025-05-01']),
        'TL Date': pd.to_datetime(['2025-01-05', '2025-01-15', '2025-02-01', '2025-05-01', '2025-04-15']),
        'Risk Level': ['High', 'Medium', 'Low', 'Critical', 'N/A'],
        'Value': [100, 200, 50, 500, 0],
        'Complete': [True, True, False, True, False]
    })


def run_parser_tests():
    """Run tests for the Excel Formula Parser."""
    logger.info("=== EXCEL FORMULA PARSER TESTS ===")
    
    # Create parser instance
    parser = ExcelFormulaParser()
    
    # Test formulas
    test_formulas = [
        # Basic comparison operators
        ("Submitter = Approver", "(df['Submitter'] == df['Approver'])"),
        ("Submitter <> Approver", "(df['Submitter'] != df['Approver'])"),
        ("Value > 100", "(df['Value'] > 100)"),
        ("Value >= 100", "(df['Value'] >= 100)"),
        ("Value < 200", "(df['Value'] < 200)"),
        ("Value <= 200", "(df['Value'] <= 200)"),
        
        # Logical operators
        ("Submitter <> Approver AND Value > 100", 
         "((df['Submitter'] != df['Approver']) & (df['Value'] > 100))"),
        ("Submitter = Approver OR Value > 100", 
         "((df['Submitter'] == df['Approver']) | (df['Value'] > 100))"),
        ("NOT Complete", "(~(df['Complete']))"),
        
        # Parentheses
        ("(Submitter <> Approver) AND (Value > 100)", 
         "(((df['Submitter'] != df['Approver'])) & ((df['Value'] > 100)))"),
        
        # Field names with spaces
        ("`Submit Date` <= `TL Date`", "(df['Submit Date'] <= df['TL Date'])"),
        
        # Combined complex formula
        ("Submitter <> Approver AND `Submit Date` <= `TL Date` AND Value > 100",
         "((df['Submitter'] != df['Approver']) & (df['Submit Date'] <= df['TL Date']) & (df['Value'] > 100))"),
        
        # Functions
        ("ISBLANK(Approver)", "(pd.isna(df['Approver']))"),
        ("NOT ISBLANK(Approver)", "(~pd.isna(df['Approver']))"),
        ("Risk Level IN (\"High\", \"Medium\")", "(df['Risk Level'].isin([\"High\", \"Medium\"]))"),
    ]
    
    # Run tests
    for original, expected in test_formulas:
        try:
            # With a fully implemented parser, this would use the actual parse method
            parsed, fields = parser.parse(original)
            
            # For demonstration, we'll compare with expected results
            # In a real implementation, we would use the actual parsed result
            if parsed == expected:
                logger.info(f"✓ {original} -> {parsed}")
            else:
                logger.info(f"✗ {original} -> {parsed}")
                logger.info(f"  Expected: {expected}")
            
            logger.info(f"  Fields used: {fields}")
            logger.info("---")
            
        except Exception as e:
            logger.error(f"✗ Error parsing {original}: {e}")
            logger.info("---")

def run_validation_tests():
    """Run tests for the Custom Formula Validation."""
    logger.info("\n=== CUSTOM FORMULA VALIDATION TESTS ===")
    
    # Create test data
    test_data = create_test_data()
    logger.info(f"Test data created with {len(test_data)} records")
    
    # Create parser instance
    parser = ExcelFormulaParser()
    
    # Test formulas
    test_formulas = [
        "Submitter <> Approver",
        "`Submit Date` <= `TL Date`",
        "Value > 100",
        "Risk Level = \"High\"",
        "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    ]
    
    # Run tests
    for formula in test_formulas:
        logger.info(f"\nTesting formula: {formula}")
        
        try:
            # Parse the formula
            parsed_formula, fields_used = parser.parse(formula)
            
            logger.info(f"Parsed: {parsed_formula}")
            logger.info(f"Fields: {fields_used}")
            
            # Create rule parameters
            params = {
                'formula': parsed_formula,
                'original_formula': formula
            }
            
            # Execute the validation using CustomFormulaValidation
            result = CustomFormulaValidation.custom_formula(test_data, params)
            
            # Print results
            passing = result.sum()
            total = len(result)
            logger.info(f"Results: {passing} of {total} records pass ({passing/total*100:.1f}%)")
            
            # Show some examples
            if not result.empty:
                passing_examples = test_data[result].head(2)
                if not passing_examples.empty:
                    logger.info("\nPassing Examples:")
                    logger.info(passing_examples[fields_used].to_string())
                
                failing_examples = test_data[~result].head(2)
                if not failing_examples.empty:
                    logger.info("\nFailing Examples:")
                    logger.info(failing_examples[fields_used].to_string())
            
        except Exception as e:
            logger.error(f"Error testing formula: {e}")


def run_config_example():
    """Show example of formula integration in YAML configuration."""
    logger.info("\n=== CONFIGURATION EXAMPLE ===")
    
    yaml_config = """
analytic_id: 99
analytic_name: 'Custom Formula Demo'
analytic_description: 'Demonstrates Excel-style formula validation'

# Data source configuration
data_source:
  name: 'approval_workflow'
  required_fields:
    - 'Submitter'
    - 'Approver'
    - 'Submit Date'
    - 'TL Date'
    - 'Risk Level'
    - 'Value'

# Validations with custom formula
validations:
  - rule: custom_formula
    description: 'Submitter cannot approve their own work and approval sequence must be followed'
    parameters:
      original_formula: 'Submitter <> Approver AND `Submit Date` <= `TL Date`'
      formula: "(df['Submitter'] != df['Approver']) & (df['Submit Date'] <= df['TL Date'])"
    metadata:
      fields_used: ['Submitter', 'Approver', 'Submit Date', 'TL Date']
  
  - rule: segregation_of_duties
    description: 'Traditional segregation of duties validation for comparison'
    parameters:
      submitter_field: 'Submitter'
      approver_fields: ['Approver']

thresholds:
  error_percentage: 5.0
  rationale: 'Industry standard allows for up to 5% error rate.'

reporting:
  group_by: 'Approver'
  summary_fields: ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage']
  detail_required: True
"""
    logger.info("Sample YAML configuration with Excel formula:")
    logger.info(yaml_config)


def run_test_utils_example():
    """Demonstrate the test_custom_formula utility function."""
    logger.info("\n=== TEST UTILITY EXAMPLE ===")
    
    # Create test data
    test_data = create_test_data()
    
    # Test a formula
    formula = "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    logger.info(f"Testing formula: {formula}")
    
    # For demo purposes, we'll use a simplified version
    parser = ExcelFormulaParser()
    fields_used = ["Submitter", "Approver", "Submit Date", "TL Date"]
    
    # Manually create the parsed formula (since parser is not fully implemented)
    parsed_formula = "(df['Submitter'] != df['Approver']) & (df['Submit Date'] <= df['TL Date'])"
    
    try:
        # In a real implementation, we'd use:
        # test_result = test_custom_formula(formula, test_data)
        
        # For the demo, create a similar result structure
        restricted_globals = {"__builtins__": {}}
        safe_locals = {"df": test_data, "pd": pd, "np": np}
        result = eval(parsed_formula, restricted_globals, safe_locals)
        
        # Calculate statistics
        total_records = len(test_data)
        passing_count = result.sum()
        failing_count = total_records - passing_count
        passing_pct = (passing_count / total_records * 100) if total_records > 0 else 0
        
        test_result = {
            'success': True,
            'parsed_formula': parsed_formula,
            'fields_used': fields_used,
            'total_records': total_records,
            'passing_count': int(passing_count),
            'failing_count': int(failing_count),
            'passing_percentage': f"{passing_pct:.1f}%",
            'passing_examples': test_data[result].head(2).to_dict('records'),
            'failing_examples': test_data[~result].head(2).to_dict('records')
        }
        
        # Display results similar to the UI
        logger.info(f"Parsed formula: {test_result['parsed_formula']}")
        logger.info(f"Fields used: {', '.join(test_result['fields_used'])}")
        logger.info(f"Total records: {test_result['total_records']}")
        logger.info(f"Passing: {test_result['passing_count']} ({test_result['passing_percentage']})")
        logger.info(f"Failing: {test_result['failing_count']}")
        
        if test_result['passing_examples']:
            logger.info("\nPassing Examples:")
            for idx, example in enumerate(test_result['passing_examples'], 1):
                logger.info(f"  {idx}. {example}")
        
        if test_result['failing_examples']:
            logger.info("\nFailing Examples:")
            for idx, example in enumerate(test_result['failing_examples'], 1):
                logger.info(f"  {idx}. {example}")
                
    except Exception as e:
        logger.error(f"Error testing formula: {e}")


def main():
    """Run all demonstration test cases."""
    logger.info("EXCEL FORMULA ENHANCEMENT DEMO")
    logger.info("============================")
    
    run_parser_tests()
    run_validation_tests()
    run_config_example()
    run_test_utils_example()
    
    logger.info("\nDemo complete. This shows how the Excel Formula Enhancement would work")
    logger.info("in the QA Analytics Framework. The actual parser implementation would")
    logger.info("replace the hardcoded examples used in this demonstration.")


if __name__ == "__main__":
    main()
