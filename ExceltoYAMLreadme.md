# Excel to YAML Data Source Configuration Converter

This tool converts Excel files (XLSX, XLS, CSV) to YAML configuration files for the QA Analytics data source system.

## Features

- Automatically analyzes Excel file structure and content
- Extracts metadata about sheets, columns, and data types
- Identifies potential key columns and relationships between sheets
- Detects validation rules based on data patterns
- Generates a well-structured YAML configuration file ready for QA Analytics

## Requirements

- Python 3.6+
- pandas
- numpy
- pyyaml

Install dependencies:
```bash
pip install pandas numpy pyyaml
```

## Usage

### Command Line

```bash
python excel_to_yaml_converter.py input_file.xlsx [output_file.yaml]
```

If you don't specify an output file, it will use the input filename with a `.yaml` extension.

### As a Module

```python
from excel_to_yaml_converter import ExcelAnalyzer, YAMLGenerator

# Analyze the Excel file
analyzer = ExcelAnalyzer("my_report.xlsx")
metadata = analyzer.analyze()

# Generate YAML configuration
generator = YAMLGenerator(metadata, "my_report.xlsx")
yaml_content = generator.to_yaml()

# Save to file
with open("my_report_config.yaml", 'w') as f:
    f.write(yaml_content)
```

## Input and Output Examples

### Example Input: `approval_data.xlsx`

A sample Excel file with approval workflow data containing:
- Sheet 1: "Workpapers" - List of audit workpapers with submitter, reviewer, approval dates
- Sheet 2: "Users" - List of users with their roles and departments

### Example Output: `approval_data.yaml`

```yaml
data_sources:
  approval_data:
    type: report
    description: XLSX file with 2 sheets and 250 total rows
    version: '1.0'
    owner: QA Analytics
    refresh_frequency: Weekly
    last_updated: '2025-05-10T14:30:22'
    file_type: xlsx
    file_pattern: approval_data_{YYYY}{MM}{DD}.xlsx
    validation_rules:
      - type: row_count_min
        threshold: 100
        description: Should have at least 100 rows
      - type: required_columns
        columns:
          - Workpaper_ID
          - Submitter
          - TL_Approver
          - AL_Approver
          - Submit_Date
        description: Critical columns that must be present
    columns_mapping:
      - source: Workpaper_ID
        target: Workpaper_ID
        data_type: id
      - source: Submitter
        target: Submitter
        data_type: string
      - source: TL_Approver
        target: TL_Approver
        aliases:
          - Reviewer
        data_type: string
      - source: AL_Approver
        target: AL_Approver
        aliases:
          - Final_Approver
        data_type: string
      - source: Submit_Date
        target: Submit_Date
        data_type: date
      - source: TL_Approval_Date
        target: TL_Approval_Date
        data_type: date
      - source: AL_Approval_Date
        target: AL_Approval_Date
        data_type: date
      - source: Status
        target: Status
        data_type: categorical
        valid_values:
          - Complete
          - In Progress
          - Pending
          - Rejected
    components:
      - name: workpapers
        sheet_name: Workpapers
        key_columns:
          - Workpaper_ID
      - name: users
        sheet_name: Users
        key_columns:
          - User_ID
        join_to: workpapers
        join_key: User_ID

analytics_mapping:
  - data_source: approval_data
    analytics: []  # Fill in with relevant analytics IDs
```

## How It Works

1. **ExcelAnalyzer** - Performs detailed analysis of the Excel file:
   - Examines file metadata (size, modification date)
   - Analyzes each sheet for structure and content
   - Studies each column to infer data types and patterns
   - Detects relationships between sheets
   - Identifies potential validation rules

2. **YAMLGenerator** - Creates a structured YAML configuration:
   - Generates appropriate data source name
   - Defines file pattern for matching similar files
   - Creates column mappings with data types and aliases
   - Builds validation rules based on data patterns
   - Organizes multi-sheet files into components with relationships

## Post-Processing Steps

After generating the YAML file, you should:

1. Review and adjust the configuration as needed
2. Add appropriate analytics IDs to the `analytics_mapping` section
3. Verify the `key_columns` selections are correct
4. Fine-tune any validation rules
5. Place the YAML file in your QA Analytics `configs` directory

## Customization

The tool makes educated guesses about many aspects of your data. For best results:

- Adjust the `refresh_frequency` to match your actual data refresh schedule
- Verify and update the `owner` field with the appropriate team
- Add or modify `validation_rules` based on your specific requirements
- Refine the `file_pattern` to match your naming conventions

## Troubleshooting

- **Incorrect Data Types**: The tool uses heuristics to infer data types. Review and correct if needed.
- **Missing Relationships**: Some relationships might not be detected automatically. Add them manually.
- **Complex Excel Files**: Very complex Excel files with pivots, formulas, or custom formatting might need manual adjustment.
