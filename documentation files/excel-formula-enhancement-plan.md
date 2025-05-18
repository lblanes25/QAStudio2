# Excel Formula Enhancement Implementation Plan

## Overview

The Excel Formula Enhancement will allow users (especially those with Excel expertise) to rapidly create analytics by writing validation logic in familiar Excel-style syntax. This feature bridges the gap between Excel-based manual checks and our automated framework, significantly accelerating the analytics creation process.

## Implementation Roadmap

### Phase 1: Core Parser Development (Days 1-5)

#### Day 1-2: Basic Parser Design
- [ ] Create `excel_formula_parser.py` module
- [ ] Implement basic operator translation:
  - `=` → `==`
  - `<>` → `!=`
  - `AND` → `&`
  - `OR` → `|`
  - `NOT` → `~`
- [ ] Implement field name detection and conversion to `df['Field Name']` format
- [ ] Handle basic parentheses and operator precedence
- [ ] Add unit tests for simple formulas

#### Day 3-4: Advanced Formula Support
- [ ] Add support for common Excel functions:
  - `ISBLANK()` → `pd.isna()`
  - `ISNUMBER()` → `pd.to_numeric(df[field], errors='coerce').notna()`
  - `ISERROR()` → Custom error checking
  - `COUNTIF()` → DataFrame filtering and counting
- [ ] Support date comparisons with proper type conversion
- [ ] Implement backtick notation for field names with spaces
- [ ] Add multi-condition formula support (nested if statements)

#### Day 5: Safety and Security
- [ ] Implement safe evaluation approach (restricted globals/locals)
- [ ] Add formula validation to prevent code injection
- [ ] Create formula sanitization functions
- [ ] Document security architecture

### Phase 2: ValidationRules Integration (Days 6-8)

#### Day 6: Extend ValidationRules Class
- [ ] Add `custom_formula` method to `ValidationRules` class:
```python
@staticmethod
def custom_formula(df: pd.DataFrame, params: Dict) -> pd.Series:
    """
    Execute a user-defined formula against the dataframe
    
    Args:
        df: DataFrame to validate
        params: Dictionary with 'formula' key containing the pandas expression
        
    Returns:
        Series with True for rows that pass validation, False otherwise
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
```

#### Day 7: Formula Testing Infrastructure
- [ ] Add formula testing functionality
- [ ] Implement test data sampling for formula validation
- [ ] Create detailed formula error messages
- [ ] Build formula explanation translator (converts formula to plain English)

#### Day 8: Configuration Integration
- [ ] Update `enhanced_data_processor.py` to handle custom formulas
- [ ] Modify YAML parsing to support custom formula rule type
- [ ] Add formula validation during configuration loading
- [ ] Create formula library management functionality

### Phase 3: UI Integration (Days 9-13)

#### Day 9-10: Configuration Wizard Enhancement
- [ ] Add custom formula section to wizard parameters tab
- [ ] Implement formula entry field with syntax highlighting
- [ ] Create real-time formula translation preview
- [ ] Add formula validation with inline error messages

#### Day 11: Formula Testing UI
- [ ] Add "Test Formula" button to configuration wizard
- [ ] Implement formula testing results view
- [ ] Create sample data display with pass/fail highlighting
- [ ] Add formula refinement suggestions

#### Day 12-13: Formula Library UI
- [ ] Create formula template library
- [ ] Add ability to save and load custom formulas
- [ ] Implement formula categorization
- [ ] Add formula search functionality
- [ ] Create formula recommendation system

### Phase 4: Documentation and Training (Days 14-16)

#### Day 14: User Documentation
- [ ] Create Excel formula reference guide
- [ ] Document supported operators and functions
- [ ] Create formula examples for common validation scenarios
- [ ] Document formula limitations and best practices

#### Day 15: Training Materials
- [ ] Develop custom formula tutorial
- [ ] Create step-by-step formula examples
- [ ] Record demo video of formula creation process
- [ ] Build formula conversion cheat sheet

#### Day 16: Integration Testing
- [ ] Perform end-to-end testing with complex formulas
- [ ] Test formula performance with large datasets
- [ ] Validate formula error handling
- [ ] Create formula regression test suite

## Example Use Cases

### 1. Segregation of Duties
**Excel Formula:**
```
Submitter <> Approver AND NOT ISBLANK(Approver)
```
**Translated to:**
```python
(df['Submitter'] != df['Approver']) & (~pd.isna(df['Approver']))
```

### 2. Approval Sequence
**Excel Formula:**
```
Submit_Date <= Approval_Date AND NOT ISBLANK(Approval_Date)
```
**Translated to:**
```python
(df['Submit_Date'] <= df['Approval_Date']) & (~pd.isna(df['Approval_Date']))
```

### 3. Conditional Validation
**Excel Formula:**
```
IF(Risk_Level = "High", Due_Date <= TODAY() - 30, Due_Date <= TODAY() - 90)
```
**Translated to:**
```python
np.where(df['Risk_Level'] == "High", 
         df['Due_Date'] <= (pd.Timestamp.today() - pd.Timedelta(days=30)), 
         df['Due_Date'] <= (pd.Timestamp.today() - pd.Timedelta(days=90)))
```

## Implementation Considerations

### Security Considerations
- Implement strict formula validation to prevent code injection
- Use restricted globals/locals for formula evaluation
- Validate formula output to ensure proper types
- Add formula complexity limits to prevent resource exhaustion

### Performance Optimization
- Cache formula parsing results for repeated use
- Optimize field name detection for large DataFrames
- Add formula execution time tracking
- Implement early termination for expensive operations

### User Experience
- Provide immediate feedback during formula creation
- Highlight formula syntax with appropriate colors
- Show translated formula in Python/pandas syntax
- Explain formula errors in user-friendly language
- Suggest formula improvements

## Success Criteria
1. **Usability**: Users can create formulas without Python knowledge
2. **Accuracy**: Formulas produce correct validation results
3. **Performance**: Formula execution is efficient with large datasets
4. **Robustness**: Error handling gracefully manages formula issues
5. **Documentation**: Formula reference guide is comprehensive
6. **Adoption**: Users create new analytics using formula approach

## Next Steps
1. Establish formula syntax specification
2. Create parser prototype for testing
3. Define function mapping between Excel and pandas
4. Design formula UI components
5. Develop formula testing approach
