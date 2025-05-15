import unittest
import pandas as pd
import numpy as np
from excel_formula_parser import ExcelFormulaParser

import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

class TestExcelFormulaParser(unittest.TestCase):
    def setUp(self):
        """Set up test environment."""
        self.parser = ExcelFormulaParser()
        
        # Create sample DataFrame for testing
        self.test_data = pd.DataFrame({
            'Risk_Level': ['High', 'Medium', 'Low', 'Critical', 'High', None],
            'Submit_Date': pd.to_datetime(['2023-01-05', '2023-01-10', '2023-01-15', '2023-01-20', '2023-01-25', None]),
            'Approval_Date': pd.to_datetime(['2023-01-15', '2023-01-12', '2023-01-25', '2023-01-30', None, None]),
            'Value': [150, 75, 30, 200, 100, None],
            'Name': ['Project A', 'Project B', 'Project C', 'Project D', 'Project E', None],
            'Status': ['Complete', 'In Progress', 'Not Started', 'Complete', 'In Progress', None]
        })

    def is_number(s):
        try:
            float(s)
            return True
        except ValueError:
            return False

    def test_equality_after_function(self):
        """Test equality comparison after function calls."""
        formula = 'IF(Value > 100, "High", "Low") = "High"'
        parsed, fields = self.parser.parse(formula)
        
        # Check the parsed formula has the correct structure
        self.assertTrue('(' in parsed and ')' in parsed)
        self.assertTrue('np.where' in parsed)
        self.assertTrue('== "High"' in parsed)
        
        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")
        
        # Validation - should be True for rows where Value > 100
        expected = self.test_data['Value'] > 100
        pd.testing.assert_series_equal(result, expected, check_names=False)

    def test_in_operator_list(self):
        """Test IN operator correctly creates list."""
        formula = 'Risk_Level IN ("High", "Critical")'
        parsed, fields = self.parser.parse(formula)
        
        # Check the parsed formula has the correct structure
        self.assertTrue('.isin([' in parsed and '])' in parsed)
        
        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")
        
        # Validation - should be True for rows with High or Critical risk
        expected = self.test_data['Risk_Level'].isin(['High', 'Critical'])
        pd.testing.assert_series_equal(result, expected)

    def test_not_operator(self):
        """Test NOT operator handling."""
        formula = 'NOT(Risk_Level = "High")'
        parsed, fields = self.parser.parse(formula)

        print("PARSED FORMULA:", parsed)

        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")

        # Debug prints
        print("RESULT:", result.values)
        print("EXPECTED:", (~(self.test_data['Risk_Level'] == 'High')).values)

        # Validation - should be True for rows where Risk_Level is not High
        expected = ~(self.test_data['Risk_Level'] == 'High')
        pd.testing.assert_series_equal(result, expected)
    
    def test_left_function(self):
        """Test LEFT function with str handling."""
        formula = 'LEFT(Name, 3) = "Pro"'
        parsed, fields = self.parser.parse(formula)
        
        # Check the parsed formula has the correct structure
        self.assertTrue('.astype(str).str[' in parsed)
        
        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")
        
        # Validation - should be True for all projects (start with "Pro")
        expected = self.test_data['Name'].astype(str).str[:3] == "Pro"
        pd.testing.assert_series_equal(result, expected)
    
    def test_datedif_function(self):
        """Test DATEDIF function with proper parentheses."""
        formula = 'DATEDIF(Submit_Date, Approval_Date, "D") > 5'
        parsed, fields = self.parser.parse(formula)
        
        # Check the parsed formula has the correct structure
        self.assertTrue('.dt.days' in parsed)
        self.assertTrue(')' in parsed and '> 5' in parsed)
        
        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")
        
        # Validation - days between dates should be > 5 for some rows
        delta_days = (self.test_data['Approval_Date'] - self.test_data['Submit_Date']).dt.days
        expected = delta_days > 5
        pd.testing.assert_series_equal(result, expected)
    
    def test_unknown_function(self):
        """Test unknown function handling."""
        formula = 'UNKNOWNFUNC(Value)'
        
        # Should raise ValueError for unknown function
        with self.assertRaises(ValueError) as context:
            parsed, fields = self.parser.parse(formula)
        
        self.assertTrue('Unknown function' in str(context.exception))
    
    def test_complex_expression(self):
        """Test complex expression with multiple operators."""
        formula = 'Risk_Level = "High" AND (Value > 100 OR Status = "Complete")'
        parsed, fields = self.parser.parse(formula)
        
        # Test the formula on our data
        success, result, error = self.parser.test_formula(formula, self.test_data)
        self.assertTrue(success, f"Formula failed: {error}")
        
        # Validation - complex condition
        high_risk = self.test_data['Risk_Level'] == 'High'
        high_value = self.test_data['Value'] > 100
        complete = self.test_data['Status'] == 'Complete'
        expected = high_risk & (high_value | complete)
        pd.testing.assert_series_equal(result, expected)

if __name__ == '__main__':
    unittest.main()
