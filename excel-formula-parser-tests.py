import unittest
import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
from excel_formula_parser import ExcelFormulaParser

class TestExcelFormulaParser(unittest.TestCase):
    """Test cases for the ExcelFormulaParser class"""

    def setUp(self):
        """Set up test environment before each test method"""
        self.parser = ExcelFormulaParser()
        
        # Create sample data for testing formulas
        self.sample_data = pd.DataFrame({
            'Risk_Level': ['High', 'Medium', 'Low', 'High', 'Critical'],
            'Days_Open': [20, 40, 100, 35, 5],
            'Value': [150, 75, 200, 50, 300],
            'Submit_Date': pd.to_datetime(['2025-01-01', '2025-01-15', '2025-02-01', '2025-03-01', '2025-04-01']),
            'Approval_Date': pd.to_datetime(['2025-01-05', '2025-01-20', '2025-02-15', '2025-03-10', '2025-04-02']),
            'Text_Field': ['ABC123', 'DEF456', 'GHI789', 'JKL012', 'MNO345'],
            'Owner': ['John', 'Sarah', 'John', 'Mike', 'Sarah'],
            'Third_Party': ['Vendor A', '', 'Vendor B, Vendor C', '', 'Vendor D']
        })
    
    def test_basic_comparison(self):
        """Test basic comparison operators"""
        tests = [
            # Format: (formula, expected_true_count)
            ("Value > 100", 3),  # 150, 200, 300
            ("Value < 100", 2),  # 75, 50
            ("Value = 150", 1),  # Just the first row
            ("Value <> 150", 4),  # All except first row
            ("Value >= 150", 3),  # 150, 200, 300
            ("Value <= 75", 2),   # 75, 50
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_in_operator(self):
        """Test IN operator for membership tests"""
        tests = [
            # Format: (formula, expected_true_count)
            ("Risk_Level IN (\"High\", \"Critical\")", 3),  # Rows 0, 3, 4
            ("Owner IN (\"John\")", 2),  # Rows 0, 2
            ("Value IN (50, 75, 100)", 2),  # Rows 1, 3
            ("Risk_Level IN (\"Medium\", \"Low\")", 2),  # Rows 1, 2
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_if_function(self):
        """Test IF function for conditional logic"""
        tests = [
            # Format: (formula, expected_true_count)
            ("IF(Risk_Level=\"High\", Days_Open<=30, Days_Open<=90)", 4),  # All except row 2
            ("IF(Value>100, \"High\", \"Low\") = \"High\"", 3),  # Rows 0, 2, 4
            ("IF(Owner=\"John\", Value>100, Value<100) = True", 3),  # Rows 0, 2, 3
            ("IF(Third_Party=\"\", \"N/A\", \"Has TP\") = \"N/A\"", 2),  # Rows 1, 3
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_logical_operators(self):
        """Test logical operators (AND, OR, NOT)"""
        tests = [
            # Format: (formula, expected_true_count)
            ("Risk_Level = \"High\" AND Value > 100", 1),  # Just row 0
            ("Risk_Level = \"High\" OR Value > 100", 5),   # All rows
            ("NOT(Risk_Level = \"High\")", 3),             # Rows 1, 2, 4
            ("Risk_Level = \"High\" AND NOT(Value < 100)", 1),  # Just row 0
            ("(Value > 100 AND Risk_Level = \"High\") OR Risk_Level = \"Critical\"", 2),  # Rows 0, 4
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_date_functions(self):
        """Test date functions and operations"""
        tests = [
            # Format: (formula, expected_true_count)
            ("Submit_Date <= Approval_Date", 5),  # All should be true
            ("(Approval_Date - Submit_Date).dt.days <= 10", 3),  # Rows 0, 1, 4
            ("DATEDIF(Submit_Date, Approval_Date, \"D\") > 10", 2),  # Rows 2, 3
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_string_functions(self):
        """Test string manipulation functions"""
        tests = [
            # Format: (formula, expected_true_count)
            ("LEFT(Text_Field, 3) = \"ABC\"", 1),  # Just row 0
            ("RIGHT(Text_Field, 3) = \"789\"", 1),  # Just row 2
            ("TRIM(Text_Field) = Text_Field", 5),   # All rows
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_backtick_fields(self):
        """Test fields with spaces using backtick notation"""
        # Create data with space in column name
        data_with_spaces = self.sample_data.copy()
        data_with_spaces['Risk Level'] = data_with_spaces['Risk_Level']
        data_with_spaces['Approval Date'] = data_with_spaces['Approval_Date']
        
        tests = [
            # Format: (formula, expected_true_count)
            ("`Risk Level` = \"High\"", 2),  # Rows 0, 3
            ("`Risk Level` IN (\"High\", \"Critical\")", 3),  # Rows 0, 3, 4
            ("IF(`Risk Level`=\"High\", `Approval Date` > \"2025-02-01\", False)", 1),  # Just row 3
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, data_with_spaces)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_complex_formulas(self):
        """Test complex combinations of functions and operators"""
        tests = [
            # Format: (formula, expected_true_count)
            (
                "IF(Risk_Level IN (\"High\", \"Critical\"), "
                "   Days_Open<=30 AND Value>100, "
                "   Days_Open<=90 AND Value<200)",
                3  # Calculation for each row needs manual verification
            ),
            (
                "IF(ISBLANK(Third_Party), \"N/A\", "
                "   IF(Risk_Level IN (\"High\", \"Critical\"), \"DNC\", \"GC\")) = \"N/A\"",
                2  # Rows 1, 3
            ),
            (
                "NOT(Risk_Level IN (\"High\", \"Critical\")) OR "
                "(Submit_Date <= Approval_Date AND Days_Open <= 30)",
                5  # All rows
            ),
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")
    
    def test_error_conditions(self):
        """Test error handling for invalid formulas"""
        tests = [
            # Format: (formula, expected_error_substring)
            ("IN (\"High\", \"Medium\")", "IN operator has no preceding field"),
            ("Risk_Level = ", "invalid syntax"),
            ("IF(Risk_Level=\"High\")", "Unexpected end of IF function"),
            ("UNKNOWNFUNC(Value)", "name 'UNKNOWNFUNC' is not defined"),
        ]
        
        for formula, expected_error in tests:
            with self.subTest(formula=formula):
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertFalse(success, "Expected formula to fail but it succeeded")
                self.assertIn(expected_error, error.lower(), 
                             f"Expected error to contain '{expected_error}', got '{error}'")

    def test_ifs_function(self):
        """Test IFS function (multiple conditions)"""
        # For IFS, we need to test with actual values returned, not just True/False
        formula = (
            "IFS("
            "   Risk_Level=\"Critical\", \"Urgent\", "
            "   Risk_Level=\"High\", \"Important\", "
            "   Risk_Level=\"Medium\", \"Normal\", "
            "   Risk_Level=\"Low\", \"Low Priority\""
            ") = \"Important\""
        )
        
        parsed, fields = self.parser.parse(formula)
        success, result, error = self.parser.test_formula(formula, self.sample_data)
        
        self.assertTrue(success, f"Formula failed: {error}")
        self.assertEqual(result.sum(), 2, "Expected 2 'Important' results")

    def test_validation_rule_formulas(self):
        """Test formulas that would be used in real validation rules"""
        tests = [
            # Segregation of duties check
            ("Owner <> \"John\" OR Risk_Level <> \"High\"", 4),  # All except row 0
            
            # Approval sequence check
            ("Submit_Date <= Approval_Date", 5),  # All rows
            
            # Risk level validation
            ("IF(Third_Party <> \"\", Risk_Level IN (\"High\", \"Critical\"), Risk_Level = \"Low\")", 2),  # Rows 0, 2
            
            # Days threshold validation
            ("IF(Risk_Level = \"Critical\", Days_Open <= 7, "
             "IF(Risk_Level = \"High\", Days_Open <= 30, Days_Open <= 90))", 4),  # All except row 2
             
            # Complex business rule
            ("(Risk_Level IN (\"High\", \"Critical\") AND Value > 100) OR "
             "(Risk_Level = \"Medium\" AND Value <= 100)", 2),  # Rows 0, 4
        ]
        
        for formula, expected_count in tests:
            with self.subTest(formula=formula):
                parsed, fields = self.parser.parse(formula)
                success, result, error = self.parser.test_formula(formula, self.sample_data)
                
                self.assertTrue(success, f"Formula failed: {error}")
                self.assertEqual(result.sum(), expected_count, 
                                f"Expected {expected_count} True results, got {result.sum()}")

if __name__ == '__main__':
    unittest.main()
