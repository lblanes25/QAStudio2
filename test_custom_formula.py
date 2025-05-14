"""
custom_formula_validation.py

This module provides functionality to test Excel-style formulas against data
and generate examples for the Excel Formula UI component.

It includes functions to:
1. Test formulas against real or generated sample data
2. Create meaningful examples of passing and failing records
3. Analyze formula behavior and provide statistics
4. Generate recommendations for formula improvement
"""

import os
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple, Any, Union
import logging
import datetime
import random
from excel_formula_parser import ExcelFormulaParser

# Set up logging
logger = logging.getLogger("qa_analytics")


def test_custom_formula(formula: str, data: pd.DataFrame) -> Dict:
    """
    Test a custom Excel formula against a DataFrame and return detailed results.

    Args:
        formula: Excel-style formula string
        data: DataFrame to test the formula against

    Returns:
        Dict containing test results including:
        - success: Whether the test was successful
        - error: Error message if test failed
        - total_records: Number of records tested
        - passing_count: Number of records that pass the formula
        - failing_count: Number of records that fail the formula
        - passing_percentage: Percentage of records that pass (formatted string)
        - failing_percentage: Percentage of records that fail (formatted string)
        - passing_examples: List of sample passing records (as dicts)
        - failing_examples: List of sample failing records (as dicts)
        - field_statistics: Statistics about fields used in the formula
        - formula_complexity: Analysis of formula complexity
        - recommendations: Suggestions for improving the formula
    """
    try:
        # Initialize parser
        parser = ExcelFormulaParser()

        # Parse the formula
        parsed_formula, fields_used = parser.parse(formula)

        # Validate fields exist in data
        missing_fields = [field for field in fields_used if field not in data.columns]
        if missing_fields:
            return {
                'success': False,
                'error': f"Fields not found in data: {', '.join(missing_fields)}",
                'total_records': len(data),
                'passing_count': 0,
                'failing_count': 0,
                'passing_percentage': "0%",
                'failing_percentage': "0%",
                'passing_examples': [],
                'failing_examples': []
            }

        # Create parameters for validation
        params = {
            'formula': parsed_formula,
            'original_formula': formula
        }

        # Use restricted execution for safety
        restricted_globals = {"__builtins__": {}}
        safe_locals = {"df": data, "pd": pd, "np": np}

        # Execute the formula
        result = eval(parsed_formula, restricted_globals, safe_locals)

        # Convert result to boolean Series if needed
        if not isinstance(result, pd.Series):
            return {
                'success': False,
                'error': f"Formula did not return a Series (got {type(result).__name__})",
                'total_records': len(data),
                'passing_count': 0,
                'failing_count': 0,
                'passing_percentage': "0%",
                'failing_percentage': "0%",
                'passing_examples': [],
                'failing_examples': []
            }

        # Ensure boolean results
        if result.dtype != bool:
            try:
                result = result.astype(bool)
            except Exception as e:
                return {
                    'success': False,
                    'error': f"Could not convert result to boolean: {str(e)}",
                    'total_records': len(data),
                    'passing_count': 0,
                    'failing_count': 0,
                    'passing_percentage': "0%",
                    'failing_percentage': "0%",
                    'passing_examples': [],
                    'failing_examples': []
                }

        # Calculate statistics
        total_records = len(data)
        passing_count = result.sum()
        failing_count = total_records - passing_count

        passing_pct = (passing_count / total_records * 100) if total_records > 0 else 0
        failing_pct = (failing_count / total_records * 100) if total_records > 0 else 0

        # Generate examples
        passing_examples = _generate_examples(data, result, fields_used, True)
        failing_examples = _generate_examples(data, ~result, fields_used, False)

        # Generate field statistics
        field_statistics = _analyze_fields(data, fields_used, result)

        # Analyze formula complexity
        complexity = _analyze_formula_complexity(formula)

        # Generate recommendations
        recommendations = _generate_recommendations(formula, data, result, field_statistics)

        # Return comprehensive results
        return {
            'success': True,
            'error': None,
            'total_records': total_records,
            'passing_count': int(passing_count),
            'failing_count': int(failing_count),
            'passing_percentage': f"{passing_pct:.1f}%",
            'failing_percentage': f"{failing_pct:.1f}%",
            'passing_examples': passing_examples,
            'failing_examples': failing_examples,
            'field_statistics': field_statistics,
            'formula_complexity': complexity,
            'recommendations': recommendations
        }

    except Exception as e:
        # Handle any errors during testing
        error_msg = str(e)
        logger.error(f"Error testing formula: {error_msg}")
        return {
            'success': False,
            'error': error_msg,
            'total_records': len(data) if data is not None else 0,
            'passing_count': 0,
            'failing_count': 0,
            'passing_percentage': "0%",
            'failing_percentage': "0%",
            'passing_examples': [],
            'failing_examples': []
        }


def _generate_examples(
    data: pd.DataFrame,
    mask: pd.Series,
    fields: List[str],
    is_passing: bool
) -> List[Dict]:
    """
    Generate example records based on a mask.

    Args:
        data: DataFrame containing records
        mask: Boolean mask indicating which records to include
        fields: List of fields to include in examples
        is_passing: Whether these are passing examples (True) or failing examples (False)

    Returns:
        List of dictionaries, each representing an example record
    """
    # If no matching records, return empty list
    if mask.sum() == 0:
        return []

    # Limit to 5 examples maximum
    filtered_data = data[mask].head(5)

    # Determine which columns to include
    # Always include the fields used in the formula
    include_columns = fields.copy()

    # Add a few extra columns that might be relevant
    extra_columns = [col for col in data.columns if col not in include_columns]
    if extra_columns:
        # Add up to 3 extra columns
        for col in extra_columns[:min(3, len(extra_columns))]:
            include_columns.append(col)

    # Convert to list of dictionaries
    examples = []
    for _, row in filtered_data.iterrows():
        example = {}
        for col in include_columns:
            # Format dates nicely
            if pd.api.types.is_datetime64_dtype(data[col]) or isinstance(row[col], (pd.Timestamp, datetime.datetime)):
                example[col] = row[col].strftime("%Y-%m-%d") if not pd.isna(row[col]) else None
            else:
                # Convert any non-serializable types to strings
                value = row[col]
                if isinstance(value, (np.int64, np.float64)):
                    value = float(value) if isinstance(value, np.float64) else int(value)
                elif isinstance(value, pd.Timestamp):
                    value = value.strftime("%Y-%m-%d")
                elif not isinstance(value, (str, int, float, bool, type(None))):
                    value = str(value)
                example[col] = value

        # Add a status field
        example['_status'] = "Pass" if is_passing else "Fail"

        examples.append(example)

    return examples


def _analyze_fields(data: pd.DataFrame, fields: List[str], result: pd.Series) -> Dict:
    """
    Analyze fields used in the formula to provide insights.

    Args:
        data: DataFrame containing records
        fields: List of fields used in the formula
        result: Boolean Series with formula results

    Returns:
        Dictionary with field statistics and insights
    """
    field_stats = {}

    for field in fields:
        field_info = {
            'type': str(data[field].dtype),
            'missing_values': int(data[field].isna().sum()),
            'missing_percentage': f"{(data[field].isna().sum() / len(data) * 100):.1f}%",
            'unique_values': int(data[field].nunique()),
        }

        # Add type-specific statistics
        if pd.api.types.is_numeric_dtype(data[field]):
            field_info.update({
                'min': float(data[field].min()) if not data[field].empty and not pd.isna(data[field].min()) else None,
                'max': float(data[field].max()) if not data[field].empty and not pd.isna(data[field].max()) else None,
                'mean': float(data[field].mean()) if not data[field].empty and not pd.isna(data[field].mean()) else None,
                'median': float(data[field].median()) if not data[field].empty and not pd.isna(data[field].median()) else None,
            })
        elif pd.api.types.is_datetime64_dtype(data[field]):
            if not data[field].empty and not pd.isna(data[field].min()):
                field_info['min_date'] = data[field].min().strftime("%Y-%m-%d")
            else:
                field_info['min_date'] = None

            if not data[field].empty and not pd.isna(data[field].max()):
                field_info['max_date'] = data[field].max().strftime("%Y-%m-%d")
            else:
                field_info['max_date'] = None
        else:
            # For string/object columns, show common values
            value_counts = data[field].value_counts()
            if not value_counts.empty:
                top_values = value_counts.head(3).to_dict()
                # Convert any non-serializable keys
                field_info['common_values'] = {str(k): int(v) for k, v in top_values.items()}

        # Calculate correlation with formula result
        if pd.api.types.is_numeric_dtype(data[field]):
            try:
                correlation = data[field].corr(result.astype(int))
                if not pd.isna(correlation):
                    field_info['correlation_to_result'] = round(correlation, 2)
            except:
                pass

        field_stats[field] = field_info

    return field_stats


def _analyze_formula_complexity(formula: str) -> Dict:
    """
    Analyze the complexity of the formula.

    Args:
        formula: The Excel-style formula string

    Returns:
        Dictionary with complexity metrics
    """
    # Count operators
    operator_count = sum(1 for op in ['=', '<>', '>', '<', '>=', '<=', 'AND', 'OR', 'NOT']
                        if op in formula.upper())

    # Count functions
    function_count = sum(1 for func in ['ISBLANK', 'ISNUMBER', 'ISTEXT', 'TODAY', 'NOW', 'IN']
                         if func in formula.upper())

    # Count field references
    field_count = formula.count('`')//2 + sum(1 for c in formula if c.isalnum() and c not in ('AND', 'OR', 'NOT'))

    # Count parentheses pairs
    paren_count = min(formula.count('('), formula.count(')'))

    # Determine complexity level
    complexity_level = "Simple"
    if operator_count + function_count > 3 or paren_count > 2:
        complexity_level = "Moderate"
    if operator_count + function_count > 5 or paren_count > 4:
        complexity_level = "Complex"

    return {
        'operator_count': operator_count,
        'function_count': function_count,
        'field_count': field_count,
        'parentheses_count': paren_count,
        'complexity_level': complexity_level,
        'length': len(formula)
    }


def _generate_recommendations(
    formula: str,
    data: pd.DataFrame,
    result: pd.Series,
    field_stats: Dict
) -> List[str]:
    """
    Generate recommendations for improving the formula.

    Args:
        formula: The Excel-style formula
        data: DataFrame containing records
        result: Boolean Series with formula results
        field_stats: Field statistics from _analyze_fields

    Returns:
        List of recommendation strings
    """
    recommendations = []

    # Check pass/fail ratio
    passing_ratio = result.mean()
    if passing_ratio == 0:
        recommendations.append("⚠️ No records pass this formula. Check if it's too restrictive.")
    elif passing_ratio == 1:
        recommendations.append("⚠️ All records pass this formula. Check if it's too permissive.")
    elif passing_ratio < 0.05:
        recommendations.append("⚠️ Very few records pass this formula (<5%). Consider making it less restrictive.")
    elif passing_ratio > 0.95:
        recommendations.append("⚠️ Almost all records pass this formula (>95%). Consider making it more restrictive.")

    # Check for missing values in fields
    for field, stats in field_stats.items():
        missing_pct = float(stats['missing_percentage'].strip('%'))
        if missing_pct > 10:
            recommendations.append(f"⚠️ Field '{field}' has {stats['missing_percentage']} missing values. Consider handling nulls with ISBLANK().")

    # Check formula complexity
    if "Complex" in formula:
        recommendations.append("⚠️ This is a complex formula. Consider breaking it into multiple simpler validations.")

    # Look for common patterns that could be optimized
    if "AND NOT ISBLANK" in formula.upper():
        recommendations.append("💡 Using 'AND NOT ISBLANK(field)' pattern. Consider extracting this to a separate validation rule.")

    if formula.upper().count(" AND ") > 2:
        recommendations.append("💡 Formula uses multiple AND conditions. Consider using multiple validation rules instead.")

    # Add general recommendations if list is empty
    if not recommendations:
        recommendations.append("✓ Formula looks good! No specific recommendations needed.")

    return recommendations


def generate_sample_data(fields_used: List[str], record_count: int = 100) -> pd.DataFrame:
    """
    Generate sample data for given fields.

    Args:
        fields_used: List of field names that will be used in formulas
        record_count: Number of records to generate

    Returns:
        DataFrame with sample data
    """
    data = {}

    for field in fields_used:
        # Determine field type based on name
        if "date" in field.lower() or "time" in field.lower():
            # Generate dates
            base_date = pd.Timestamp('2025-01-01')
            data[field] = [base_date + pd.Timedelta(days=random.randint(0, 100))
                         for _ in range(record_count)]

        elif any(keyword in field.lower() for keyword in ["amount", "value", "cost", "price", "total"]):
            # Generate numeric values
            data[field] = np.random.uniform(10, 1000, record_count).round(2)

        elif any(keyword in field.lower() for keyword in ["count", "number", "quantity", "qty"]):
            # Generate integer values
            data[field] = np.random.randint(1, 100, record_count)

        elif any(keyword in field.lower() for keyword in ["percentage", "rate", "ratio", "pct"]):
            # Generate percentage values
            data[field] = np.random.uniform(0, 100, record_count).round(2)

        elif any(keyword in field.lower() for keyword in ["flag", "indicator", "complete", "done"]):
            # Generate boolean/status values
            data[field] = np.random.choice([True, False], record_count)

        elif any(keyword in field.lower() for keyword in ["status", "state", "stage", "phase"]):
            # Generate status values
            statuses = ["New", "In Progress", "Complete", "On Hold", "Cancelled"]
            data[field] = np.random.choice(statuses, record_count)

        elif any(keyword in field.lower() for keyword in ["priority", "severity", "risk", "impact"]):
            # Generate priority/risk levels
            levels = ["Low", "Medium", "High", "Critical", "N/A"]
            data[field] = np.random.choice(levels, record_count)

        elif "submitter" in field.lower() or "creator" in field.lower() or "author" in field.lower():
            # Generate submitter names
            names = ["John Smith", "Jane Doe", "Robert Johnson", "Emily Wilson",
                    "Michael Brown", "Sarah Davis", "David Miller", "Emma Garcia"]
            data[field] = np.random.choice(names, record_count)

        elif "approver" in field.lower() or "reviewer" in field.lower() or "manager" in field.lower():
            # Generate approver names
            names = ["Alex Wong", "Maria Rodriguez", "James Taylor", "Elizabeth Clark",
                    "Thomas Lee", "Olivia Martin", "William White", "Sophia Moore"]
            data[field] = np.random.choice(names, record_count)

        else:
            # Default to text field with names
            values = [f"Value-{i}" for i in range(1, 11)]
            data[field] = np.random.choice(values, record_count)

    # Create DataFrame
    df = pd.DataFrame(data)

    # Add some NULL values for realism (about 5% of fields)
    for field in fields_used:
        # Skip date fields for simplicity
        if "date" not in field.lower():
            null_indices = np.random.choice(
                record_count,
                size=max(1, int(record_count * 0.05)),
                replace=False
            )
            df.loc[null_indices, field] = np.nan

    return df


def find_potential_fields(formula: str, all_fields: List[str]) -> List[str]:
    """
    Find potential fields that might be relevant for a formula.

    Args:
        formula: The Excel-style formula
        all_fields: List of all available fields

    Returns:
        List of field names that might be relevant
    """
    # First, parse the formula to get explicitly used fields
    parser = ExcelFormulaParser()
    try:
        _, explicit_fields = parser.parse(formula)
    except:
        # If parsing fails, extract potential fields heuristically
        words = set(formula.replace('`', ' ').replace('(', ' ').replace(')', ' ').split())
        explicit_fields = [word for word in words
                          if word not in ['AND', 'OR', 'NOT', 'IN', 'ISBLANK', 'ISNUMBER', 'ISTEXT']]

    # Find related fields
    related_fields = []

    for field in explicit_fields:
        # Find fields with similar names
        field_lower = field.lower()
        for other_field in all_fields:
            if other_field not in explicit_fields and other_field not in related_fields:
                other_lower = other_field.lower()

                # Check for common prefixes/suffixes
                if (field_lower.startswith(other_lower) or
                    other_lower.startswith(field_lower) or
                    field_lower.endswith(other_lower) or
                    other_lower.endswith(field_lower)):
                    related_fields.append(other_field)

                # Check for date pairs
                if 'date' in field_lower and 'date' in other_lower:
                    related_fields.append(other_field)

                # Check for common field pairs
                if ('submitter' in field_lower and 'approver' in other_lower) or \
                   ('approver' in field_lower and 'submitter' in other_lower):
                    related_fields.append(other_field)

    # Combine explicit and related fields, removing duplicates
    all_relevant = explicit_fields + related_fields
    return list(dict.fromkeys(all_relevant))  # Remove duplicates while preserving order


def suggest_formula_improvements(formula: str, data: pd.DataFrame) -> List[Dict]:
    """
    Suggest improvements for the formula based on the data.

    Args:
        formula: The Excel-style formula
        data: DataFrame containing records

    Returns:
        List of suggestion dictionaries with 'original', 'improved', and 'explanation'
    """
    suggestions = []

    # Parse the formula
    parser = ExcelFormulaParser()
    try:
        parsed, fields = parser.parse(formula)
    except:
        # If parsing fails, return a suggestion to fix the syntax
        return [{
            'original': formula,
            'improved': None,
            'explanation': "The formula has syntax errors. Please check for typos or missing operators."
        }]

    # Check for missing ISBLANK checks
    for field in fields:
        if data[field].isna().any():
            missing_pct = data[field].isna().mean() * 100

            if missing_pct > 5 and f"ISBLANK({field})" not in formula and f"NOT ISBLANK({field})" not in formula:
                # Original formula doesn't handle nulls, suggest improvement
                improved = formula

                # For simple equalities, suggest adding NOT ISBLANK
                if f"{field} = " in formula or f"{field}=" in formula:
                    improved = f"NOT ISBLANK({field}) AND ({formula})"
                    suggestions.append({
                        'original': formula,
                        'improved': improved,
                        'explanation': f"Field '{field}' has {missing_pct:.1f}% missing values. Added NULL check."
                    })

    # Check for common patterns that could be simplified
    if " = \"\"" in formula or " =\"\"" in formula:
        improved = formula.replace(" = \"\"", " ISBLANK").replace(" =\"\"", " ISBLANK")
        suggestions.append({
            'original': formula,
            'improved': improved,
            'explanation': "Replaced '= \"\"' with ISBLANK for clearer empty string checking."
        })

    # Check date comparisons without time consideration
    date_fields = [field for field in fields
                  if field in data.columns and pd.api.types.is_datetime64_dtype(data[field])]

    for date_field in date_fields:
        if date_field in formula and "time" not in formula.lower():
            # The formula compares dates but doesn't mention time - might need time truncation
            # This is a common Excel formula issue when comparing dates
            suggestions.append({
                'original': formula,
                'improved': formula,  # Keep the same for now, as pandas handles this well
                'explanation': f"Date comparison with '{date_field}' works as expected. In pandas, dates are compared correctly ignoring time components by default."
            })

    return suggestions


def formula_performance_test(formula: str, data: pd.DataFrame, iterations: int = 10) -> Dict:
    """
    Test the performance of a formula on different data sizes.

    Args:
        formula: The Excel-style formula
        data: DataFrame containing records
        iterations: Number of test iterations

    Returns:
        Dictionary with performance metrics
    """
    import time

    # Parse the formula
    parser = ExcelFormulaParser()
    parsed_formula, fields_used = parser.parse(formula)

    results = {
        'formula_length': len(formula),
        'parsed_length': len(parsed_formula),
        'field_count': len(fields_used),
        'timings': [],
        'record_counts': []
    }

    # Test on different data sizes
    sizes = [100, 1000, 10000] if len(data) >= 10000 else [100, 500, 1000]
    for size in sizes:
        # Skip if data is too small
        if len(data) < size:
            continue

        # Sample the data
        sample = data.sample(size) if len(data) > size else data

        # Create parameters
        params = {
            'formula': parsed_formula,
            'original_formula': formula
        }

        # Time the execution
        start_time = time.time()

        for _ in range(iterations):
            from validation_rules import ValidationRules
            result = ValidationRules.custom_formula(sample, params)

        end_time = time.time()
        avg_time = (end_time - start_time) / iterations * 1000  # in milliseconds

        results['timings'].append(avg_time)
        results['record_counts'].append(size)

    # Calculate metrics
    if results['timings']:
        results['avg_time_per_record'] = sum(
            t/c for t, c in zip(results['timings'], results['record_counts'])
        ) / len(results['timings'])

        results['evaluation_speed'] = 'Fast'
        if results['avg_time_per_record'] > 0.1:  # More than 0.1ms per record
            results['evaluation_speed'] = 'Moderate'
        if results['avg_time_per_record'] > 1.0:  # More than 1ms per record
            results['evaluation_speed'] = 'Slow'

    return results


# Example usage
if __name__ == "__main__":
    # Create sample data
    sample_data = pd.DataFrame({
        'Submitter': ['John', 'Mary', 'John', 'Bob', 'Alice'],
        'Approver': ['Alice', 'John', 'John', 'Charlie', 'Bob'],
        'Submit Date': pd.to_datetime(['2025-01-01', '2025-02-01', '2025-03-01', '2025-04-01', '2025-05-01']),
        'TL Date': pd.to_datetime(['2025-01-05', '2025-01-15', '2025-02-01', '2025-05-01', '2025-04-15']),
        'Risk Level': ['High', 'Medium', 'Low', 'Critical', 'N/A'],
        'Value': [100, 200, 50, 500, 0],
        'Complete': [True, True, False, True, False]
    })

    # Test a formula
    formula = "Submitter <> Approver AND `Submit Date` <= `TL Date`"
    results = test_custom_formula(formula, sample_data)

    # Print results (ensuring JSON serialization)
    import json

    # Use a custom encoder to handle non-serializable types
    class CustomJSONEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, (np.int64, np.float64)):
                return int(obj) if isinstance(obj, np.int64) else float(obj)
            if isinstance(obj, (pd.Timestamp, datetime.datetime)):
                return obj.strftime("%Y-%m-%d %H:%M:%S")
            if isinstance(obj, np.ndarray):
                return obj.tolist()
            return super().default(obj)

    print(json.dumps(results, indent=2, cls=CustomJSONEncoder))