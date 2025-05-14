"""
excel_formula_test_simple.py - Simple test for Excel Formula in QA Analytics Framework

This script tests the Excel formula parser and validation functionality using the
ValidationRules.custom_formula method directly.
"""

import pandas as pd
import numpy as np
import logging
from excel_formula_parser import ExcelFormulaParser
from validation_rules import ValidationRules

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


def test_parser():
    """Test the Excel formula parser."""
    logger.info("=== EXCEL FORMULA PARSER TESTS ===")

    # Create parser instance
    parser = ExcelFormulaParser()

    # Test formulas
    test_formulas = [
        "Submitter = Approver",
        "Submitter <> Approver",
        "Value > 100",
        "Value >= 100",
        "Value < 200",
        "Value <= 200",
        "Submitter <> Approver AND Value > 100",
        "Submitter = Approver OR Value > 100",
        "NOT Complete",
        "(Submitter <> Approver) AND (Value > 100)",
        "`Submit Date` <= `TL Date`",
        "Submitter <> Approver AND `Submit Date` <= `TL Date` AND Value > 100",
        "ISBLANK(Approver)",
        "NOT ISBLANK(Approver)",
        "Risk Level IN (\"High\", \"Medium\")"
    ]

    # Run tests
    for formula in test_formulas:
        try:
            # Parse the formula
            parsed, fields = parser.parse(formula)
            logger.info(f"✓ Formula: {formula}")
            logger.info(f"  Parsed: {parsed}")
            logger.info(f"  Fields: {fields}")
            logger.info("---")
        except Exception as e:
            logger.error(f"✗ Error parsing '{formula}': {e}")
            logger.info("---")


def test_validation():
    """Test formula validation with ValidationRules.custom_formula."""
    logger.info("\n=== FORMULA VALIDATION TESTS ===")

    # Create test data
    test_data = create_test_data()
    logger.info(f"Test data created with {len(test_data)} records")

    # Create parser
    parser = ExcelFormulaParser()

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

        try:
            # Parse the formula
            parsed_formula, fields_used = parser.parse(formula)
            logger.info(f"Parsed: {parsed_formula}")
            logger.info(f"Fields: {fields_used}")

            # Create parameters for validation
            params = {
                'formula': parsed_formula,
                'original_formula': formula
            }

            # Execute validation
            result = ValidationRules.custom_formula(test_data, params)

            # Calculate statistics
            passing_count = result.sum()
            total = len(result)
            failing_count = total - passing_count
            passing_pct = (passing_count / total * 100) if total > 0 else 0

            # Display results
            logger.info(f"Results: {passing_count} of {total} records pass ({passing_pct:.1f}%)")
            logger.info(f"Expected: {expected_passing} records, Got: {passing_count}")

            # Show passing examples
            if passing_count > 0:
                logger.info("Passing Examples:")
                logger.info(test_data[result].head(2)[fields_used].to_string())

            # Show failing examples
            if failing_count > 0:
                logger.info("Failing Examples:")
                logger.info(test_data[~result].head(2)[fields_used].to_string())

        except Exception as e:
            logger.error(f"Error testing formula: {e}")


def show_config_example():
    """Show example of formula configuration for QA Analytics."""
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
      
  # With the auto-parsing feature you added, just providing the original_formula is enough!
  # The formula will be parsed during validation.
  # For reference, the parser would generate:
  # formula: "{parsed_formula}"
  
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
    """Run all tests."""
    logger.info("EXCEL FORMULA ENHANCEMENT DEMO")
    logger.info("============================")

    # Test parser
    test_parser()

    # Test validation
    test_validation()

    # Show config example
    show_config_example()

    logger.info("\nDemo complete! The Excel Formula Enhancement is working correctly.")
    logger.info("You can now use Excel-style formulas in your QA Analytics Framework.")


if __name__ == "__main__":
    main()