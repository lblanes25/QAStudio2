"""
Test module for Excel Formula feature in QA Analytics Framework.

This script demonstrates the Excel Formula parsing and validation features
by running a series of test cases.
"""

import pandas as pd
import numpy as np
import logging
from excel_formula_parser import ExcelFormulaParser
from custom_formula_validation import test_custom_formula, format_test_results

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
        ("Submitter = Approver", "df['Submitter']==df['Approver']"),
        ("Submitter <> Approver", "df['Submitter']!=df['Approver']"),
        ("Value > 100", "df['Value']>100"),
        ("Value >= 100", "df['Value']>=100"),
        ("Value < 200", "df['Value']<200"),
        ("Value <= 200", "df['Value']<=200"),

        # Logical operators
        ("Submitter <> Approver AND Value > 100",
         "(df['Submitter']!=df['Approver']) & (df['Value']>100)"),
        ("Submitter = Approver OR Value > 100",
         "(df['Submitter']==df['Approver']) | (df['Value']>100)"),
        ("NOT Complete", "~df['Complete']"),

        # Parentheses
        ("(Submitter <> Approver) AND (Value > 100)",
         "((df['Submitter']!=df['Approver'])) & ((df['Value']>100))"),

        # Field names with spaces
        ("`Submit Date` <= `TL Date`", "df['Submit Date']<=df['TL Date']"),

        # Combined complex formula
        ("Submitter <> Approver AND `Submit Date` <= `TL Date` AND Value > 100",
         "(df['Submitter']!=df['Approver']) & (df['Submit Date']<=df['TL Date']) & (df['Value']>100)"),

        # Functions
        ("ISBLANK(Approver)", "pd.isna(df['Approver'])"),
        ("NOT ISBLANK(Approver)", "~pd.isna(df['Approver'])"),
        ("Risk Level IN (\"High\", \"Medium\")", "df['Risk Level'].isin([\"High\", \"Medium\"])"),
    ]

    # Run tests
    for original, expected in test_formulas:
        try:
            # Parse the formula
            parsed, fields = parser.parse(original)

            # Check if simplified versions match (ignoring whitespace and exact parentheses)
            simplified_parsed = parsed.replace(" ", "").replace("(", "").replace(")", "")
            simplified_expected = expected.replace(" ", "").replace("(", "").replace(")", "")

            success = simplified_parsed == simplified_expected

            if success:
                logger.info(f"✓ {original} -> {parsed}")
            else:
                logger.info(f"✗ {original} -> {parsed}")
                logger.info(f"  Expected: {expected}")
                logger.info(f"  Fields used: {fields}")

            logger.info("---")

        except Exception as e:
            logger.error(f"✗ Error parsing {original}: {e}")
            logger.info("---")


def run_formula_tests():
    """Run tests for formula evaluation."""
    logger.info("\n=== FORMULA EVALUATION TESTS ===")

    # Create test data
    test_data = create_test_data()
    logger.info(f"Test data created with {len(test_data)} records")

    # Test formulas with expected pass counts
    test_cases = [
        ("Submitter <> Approver", 3),  # 3 records should pass
        ("Value > 100", 3),  # 3 records have Value > 100
        ("`Submit Date` <= `TL Date`", 3),  # 3 records have Submit Date <= TL Date
        ("Submitter <> Approver AND Value > 100", 2),  # 2 records should pass both conditions
        ("Risk Level = \"High\"", 1),  # 1 record has Risk Level = "High"
        ("Risk Level IN (\"High\", \"Medium\")", 2),  # 2 records have Risk Level in ["High", "Medium"]
        ("NOT ISBLANK(Value)", 5),  # All 5 records have non-blank Value
    ]

    for formula, expected_passing in test_cases:
        logger.info(f"\nTesting formula: {formula}")
        logger.info(f"Expected passing: {expected_passing} records")

        # Test the formula
        results = test_custom_formula(formula, test_data)

        if results['success']:
            logger.info(format_test_results(results))

            # Verify the expected passing count
            actual_passing = results['passing_count']
            if actual_passing == expected_passing:
                logger.info(f"✓ Pass count verified: {actual_passing} records")
            else:
                logger.info(f"✗ Pass count mismatch: Expected {expected_passing}, got {actual_passing}")
        else:
            logger.error(f"Error testing formula: {results.get('error')}")


def run_config_example():
    """Show example of formula integration in YAML configuration."""
    logger.info("\n=== CONFIGURATION EXAMPLE ===")

    # Get a sample formula
    parser = ExcelFormulaParser()
    formula = "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    parsed_formula, fields_used = parser.parse(formula)

    yaml_config = f"""
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
      original_formula: '{formula}'
      formula: "{parsed_formula}"
    metadata:
      fields_used: {fields_used}
  
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


def main():
    """Run all demonstration test cases."""
    logger.info("EXCEL FORMULA ENHANCEMENT DEMO")
    logger.info("============================")

    # Run parser tests
    run_parser_tests()

    # Run formula tests
    run_formula_tests()

    # Show config example
    run_config_example()

    logger.info("\nDemo complete. This shows how the Excel Formula Enhancement works")
    logger.info("in the QA Analytics Framework. The implementation handles comparison")
    logger.info("operators, logical operators, functions, and complex expressions.")


if __name__ == "__main__":
    main()