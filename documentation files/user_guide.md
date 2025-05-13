# Enhanced QA Analytics Framework: User Guide

## Table of Contents
1. [Introduction](#introduction)
2. [Getting Started](#getting-started)
3. [Running Existing Analytics](#running-existing-analytics)
4. [Creating New Analytics](#creating-new-analytics)
5. [Testing Analytics](#testing-analytics)
6. [Scheduling Automated Runs](#scheduling-automated-runs)
7. [Managing Reference Data](#managing-reference-data)
8. [Troubleshooting](#troubleshooting)
9. [Glossary](#glossary)

## Introduction

### About This Guide
This guide provides instructions for using the Enhanced QA Analytics Framework. It is designed for both regular users who need to run analytics and power users who create new analytics configurations.

### System Overview
The Enhanced QA Analytics Framework automates the validation of audit compliance, replacing a previously manual process. The system:

- Applies validation rules to audit data files
- Identifies compliant and non-compliant records
- Generates compliance reports
- Automates regular processing

### Key Features
- **No-code configuration**: Create new analytics without programming knowledge
- **Template-based approach**: Use pre-built templates for common validation patterns
- **Interactive testing**: Test configurations with sample data
- **Automated scheduling**: Schedule regular execution and reporting
- **Email notifications**: Receive results automatically

### User Roles

**Basic User**
- Run existing analytics
- View and interpret reports
- Basic troubleshooting

**Power User**
- Create new analytics from templates
- Customize validation rules
- Test configurations
- Set up scheduled jobs

**Administrator**
- Manage reference data
- Monitor system performance
- Configure email settings
- Maintain templates

## Getting Started

### System Requirements
- Windows operating system
- 4GB RAM minimum (8GB recommended)
- 500MB free disk space
- Network access to shared data locations

### Installation
The application is typically installed by your IT department. If you need to install it yourself:

1. Extract the provided zip file to your desired location
2. Run `setup.bat` to configure the environment
3. Create a desktop shortcut to `enhanced_qa_analytics.py`

### First Launch
1. Double-click the application shortcut
2. The main application window will appear with several tabs
3. The application will create necessary directories if they don't exist

### Understanding the Interface
The application has a tabbed interface:

- **Run Analytics**: Execute existing analytics configurations
- **Configuration Wizard**: Create new analytics using templates
- **Testing**: Test analytics with sample data
- **Scheduler**: Set up automated execution
- **Data Sources**: Manage data source connections
- **Reference Data**: View and update reference data

## Running Existing Analytics

Running an existing analytic is the most common task and requires just a few steps:

### Step 1: Select an Analytic
1. Go to the "Run Analytics" tab
2. Select the QA-ID from the dropdown list (e.g., "77 - Audit Test Workpaper Approvals")
3. You'll see information about the selected analytic

### Step 2: Select Source Data
1. Click "Browse..." next to the "Source Data File" field
2. Navigate to and select your Excel data file
3. The file path will appear in the field

### Step 3: Choose Output Location
1. Click "Browse..." next to the "Output Directory" field
2. Select where you want reports to be saved
3. By default, this is set to an "output" folder in the application directory

### Step 4: Run the Analysis
1. Click the "Run Analysis" button
2. A progress bar will show the processing status
3. The status log will show details of the process

### Step 5: Review Results
1. When processing completes, a success message will appear
2. Navigate to the output directory to find your reports:
   - Main report with summary and details
   - Individual reports by group (if enabled)

### Understanding Reports
The generated reports contain:

- **Summary sheet**: Overall statistics and results by group
- **Detail sheet**: Line-by-line validation results
- **Configuration sheet**: Settings and parameters used

Results are marked as:
- **GC**: Generally Conforms
- **DNC**: Does Not Conform
- **PC**: Partially Conforms

## Creating New Analytics

Creating a new analytic using the Configuration Wizard doesn't require programming knowledge:

### Step 1: Start the Wizard
1. Go to the "Configuration Wizard" tab
2. Browse the available templates
3. Select a template that matches your needs

### Step 2: Select a Template
1. Review the template details and validation rules
2. Ensure the template includes the validation rules you need
3. Click "Next" to proceed to basic configuration

### Step 3: Configure Basic Settings
1. Provide an Analytics ID (numeric identifier)
2. Enter a descriptive Analytics Name
3. Add a brief description of what this analytic validates
4. Select a data source (if registered)
5. Set the error threshold percentage
6. Specify the field to group results by
7. Click "Next" to continue

### Step 4: Configure Template Parameters
1. Fill in all required parameters (marked with *)
2. Use the "Quick Fill" option to apply example values
3. Customize parameters to match your data structure
4. Click "Next" to preview the configuration

### Step 5: Review and Save
1. Review the generated configuration
2. Make any necessary adjustments by going back
3. Click "Save Configuration"
4. The configuration is saved to the `configs` directory

### Template Selection Tips
Choose a template based on the validation type:

- **Approval Workflow**: For sequences of approvals with role validation
- **Risk Assessment**: For risk rating validations and assessments
- **Control Testing**: For control design and operating effectiveness
- **Issue Management**: For issue tracking and remediation

### Parameter Configuration Tips
For better results:

- Match field names exactly to your data source columns
- Use descriptive names for easier maintenance
- Consider using the data source registry for standard field mapping
- Test configurations thoroughly before using in production

## Testing Analytics

The Testing environment allows you to validate configurations before using them:

### Step 1: Select Analytics to Test
1. Go to the "Testing" tab
2. Select the Analytics ID to test from the dropdown
3. Click "Load Configuration"

### Step 2: Choose Test Data Approach
There are two approaches:

**Generate Sample Data**:
1. Select "Generate Sample Data"
2. Specify the number of records to generate
3. Set the error percentage for testing
4. The system will generate appropriate test data

**Use Existing File**:
1. Select "Use Existing File"
2. Click "Browse..." to select a data file
3. Choose a real data file or previously created sample

### Step 3: Run the Test
1. Click "Run Test"
2. The system will process the data
3. Results will appear in the tabs below

### Step 4: Review Test Results
The results are shown in several tabs:

**Summary Tab**:
- Shows overall statistics
- Displays results by group
- Highlights threshold violations

**Detail Tab**:
- Shows line-by-line validation results
- Can be filtered to show GC, DNC, or PC records
- Identifies specific validation failures

**Sample Data Tab**:
- Shows the data used for testing
- Useful for verification and review

### Step 5: Export or Refine
1. Click "Export Sample Data" to save the test data
2. Click "Export Results" to save test results
3. Click "Generate Report" to create a standard report
4. If issues are found, adjust the configuration and retest

### Testing Best Practices
- Test with various error rates to ensure detection
- Verify all validation rules are working as expected
- Test with edge cases and special scenarios
- Save sample data sets for regression testing

## Scheduling Automated Runs

Set up automated processing to run analytics on a regular schedule:

### Step 1: Create a Job
1. Go to the "Scheduler" tab
2. Click "New Job" to create a new scheduled job

### Step 2: Configure Job Settings
1. Fill in the Job ID (unique identifier)
2. Select the Analytics ID to run
3. Choose the schedule type:
   - **Daily**: Run every day
   - **Weekly**: Run on a specific day of the week
   - **Monthly**: Run on a specific day of the month
4. Set the time to run (24-hour format, e.g., "08:00")
5. For weekly, select the day of the week
6. For monthly, select the day of the month

### Step 3: Specify Data Source
1. Enter the data file pattern (e.g., "data/*.xlsx")
2. The scheduler will use the most recent file matching this pattern
3. Click "Browse..." to help construct the pattern

### Step 4: Configure Email Notifications (Optional)
1. Check "Send Email Notification" to enable
2. Add recipient email addresses
3. Check "Generate Individual Reports" if needed

### Step 5: Save and Activate
1. Click "Save Job" to store the configuration
2. Click "Start Scheduler" to activate all jobs
3. The status will show "Scheduler is running"

### Managing Jobs
- **View**: Select a job to view its configuration
- **Edit**: Make changes and click "Save Job"
- **Delete**: Select a job and click "Delete Job"
- **Run Now**: Run a job immediately for testing

### Email Configuration
To enable email notifications:

1. Go to the "Configuration" tab in the scheduler
2. Under "Email Settings", check "Enable Email Notifications"
3. Enter SMTP server details:
   - SMTP Server (e.g., "smtp.example.com")
   - SMTP Port (typically 587 for TLS)
   - Username and Password
   - From Address and Admin Address
4. Click "Test Email" to verify settings
5. Click "Save Email Config" when done

### Schedule Configuration
To set default scheduling options:

1. Go to the "Schedule Settings" tab
2. Set the default time and day
3. Specify the output directory for automated runs
4. Click "Save Schedule Config"

## Managing Reference Data

Reference data is used for validations like title-based approval:

### Viewing Reference Data
1. Go to the "Reference Data" tab
2. The list shows all configured reference data:
   - Name
   - Format (Dictionary or DataFrame)
   - Version
   - Last Modified date
   - Row Count
   - Freshness status

### Updating Reference Data
1. Select a reference data entry
2. Click "Update Reference File"
3. Select the new data file
4. Confirm the update
5. The system records the update in the audit log

### Viewing Update History
1. Click "View Update History"
2. A dialog shows all reference data updates:
   - Timestamp
   - User
   - Action
   - Previous and new versions

### Reference Data Freshness
Reference data has a configured "freshness" period:

- **Fresh**: Updated within the configured period
- **Stale**: Not updated recently
- **Not Loaded**: Not yet loaded into the system

Stale reference data generates warnings during processing.

## Troubleshooting

### Common Issues and Solutions

**Issue**: Application fails to start
- **Solution**: Verify Python installation, check for error logs in the logs directory

**Issue**: "Missing required columns" error when running analytics
- **Solution**: Verify that your source data has all required columns
- **Solution**: Check for column name mismatches or aliases in the data source registry

**Issue**: "Failed to load reference data" error
- **Solution**: Verify that reference data files exist in the correct location
- **Solution**: Check file format and column names

**Issue**: All records show as "DNC" (Does Not Conform)
- **Solution**: Review validation rules in the configuration
- **Solution**: Check for data formatting issues (especially dates)
- **Solution**: Test with the Testing environment to identify specific issues

**Issue**: Scheduler not running jobs
- **Solution**: Verify the scheduler is started (Status: "Scheduler is running")
- **Solution**: Check job configurations, especially data file patterns
- **Solution**: Verify system time matches scheduled time

**Issue**: Email notifications not being sent
- **Solution**: Check email configuration (server, port, credentials)
- **Solution**: Verify recipient email addresses
- **Solution**: Check for firewalls blocking SMTP traffic

### Logging
The application maintains detailed logs:

- **Application log**: `qa_analytics.log` in the application directory
- **Reference data audit log**: `logs/reference_data_audit.json`

To enable more detailed logging:
1. Edit `logging_config.py`
2. Change `level=logging.INFO` to `level=logging.DEBUG`
3. Restart the application

### Getting Help
If you encounter issues not covered in this guide:

1. Check the detailed log files
2. Contact your administrator or designated super user
3. Refer to the knowledge base on the internal portal

**Validation Rule**: Specific check applied to data to determine compliance

**Scheduler**: Component that automates the execution of analytics on a regular schedule

**Job**: A scheduled task to run an analytic automatically

**Data Source Pattern**: Pattern used to find the latest data file for automated runs

**Threshold**: Maximum acceptable percentage of non-conforming records

**Segregation of Duties**: Validation rule ensuring the submitter is not also an approver

**Approval Sequence**: Validation rule checking that approvals happened in the correct order

**Title-based Approval**: Validation rule verifying approvers have appropriate titles

**Third-party Risk Validation**: Rule checking risk ratings for third-party vendors

**Error Percentage**: Percentage of records that do not conform to validation rules

## Advanced Topics

### Custom Validation Rules

The system includes several pre-built validation rules:

1. **Segregation of Duties**
   - Purpose: Ensure submitter is not also an approver
   - Parameters:
     - `submitter_field`: Field containing the submitter name
     - `approver_fields`: List of fields containing approver names

2. **Approval Sequence**
   - Purpose: Verify approvals happened in the correct order
   - Parameters:
     - `date_fields_in_order`: List of date fields that should be in sequence

3. **Title-based Approval**
   - Purpose: Verify approvers have appropriate job titles
   - Parameters:
     - `approver_field`: Field containing the approver name
     - `allowed_titles`: List of acceptable titles
     - `title_reference`: Reference data containing employee titles

4. **Third-party Risk Validation**
   - Purpose: Ensure risk ratings are assigned when third parties are present
   - Parameters:
     - `third_party_field`: Field containing third-party information
     - `risk_level_field`: Field containing the risk rating

5. **Enumeration Validation**
   - Purpose: Verify field values match an allowed list
   - Parameters:
     - `field_name`: Field to check
     - `valid_values`: List of acceptable values

For additional custom validation needs, contact your administrator.

### Data Source Registry

The Data Source Registry centralizes data source definitions for consistency:

#### Benefits
- Standardized column mappings across analytics
- Automatic aliasing for column name variations
- Centralized validation rules for data quality
- Easy reference across multiple analytics

#### Viewing Registry
1. Go to the "Data Sources" tab
2. View the list of registered data sources
3. Click "View Details" to see the full configuration

#### Using in Analytics
When creating a new analytic:
1. Select a registered data source in the Configuration Wizard
2. The system will automatically apply column mappings
3. Required field names will be pre-populated

### Batch Processing

For processing multiple analytics at once:

1. Create a batch file (e.g., `run_batch.bat`) with commands:
   ```
   python enhanced_qa_analytics.py -a 77 -s data\file1.xlsx -o output
   python enhanced_qa_analytics.py -a 78 -s data\file2.xlsx -o output
   ```

2. Schedule the batch file using Windows Task Scheduler

3. Alternatively, create multiple jobs in the Scheduler tab

### Report Customization

To customize report appearance:

1. Modify the template files in the `templates` directory:
   - `report_template.xlsx`: Main report template
   - `individual_template.xlsx`: Individual report template

2. Adjust formatting, colors, and layouts as needed

3. Do not remove or rename sheets or key cells

### Performance Optimization

For processing large datasets:

1. Increase available memory:
   - Edit `enhanced_qa_analytics.bat`
   - Modify `-Xmx` parameter to increase heap size

2. Use data chunking:
   - Split large files into smaller batches
   - Process each batch separately
   - Combine results using the reporting tools

3. Schedule during off-hours:
   - Use the Scheduler for overnight processing
   - Set email notifications to receive results

## Tutorial: Complete Workflow Example

This tutorial walks through a complete workflow from creating to scheduling an analytic.

### Scenario

You need to validate that all audit workpapers follow proper approval procedures:
- Submitter cannot be also an approver (segregation of duties)
- Team Lead must approve before Audit Leader (approval sequence)
- Audit Leader must have appropriate job title (title-based approval)

### Step 1: Create the Analytics Configuration

1. Go to the "Configuration Wizard" tab
2. Select the "Approval Workflow" template
3. Click "Next" to proceed to basic configuration

4. Enter basic settings:
   - Analytics ID: `85`
   - Analytics Name: `Q2 Workpaper Approvals`
   - Description: `Validates Q2 audit workpaper approval process`
   - Data Source: `audit_workpaper_approvals` (if registered)
   - Error Threshold: `5.0`
   - Group By: `AL approver`

5. Configure template parameters:
   - submitter_field: `TW submitter`
   - first_approver_field: `TL approver`
   - second_approver_field: `AL approver`
   - submission_date_field: `Submit Date`
   - first_approval_date_field: `TL Approval Date`
   - second_approval_date_field: `AL Approval Date`
   - title_reference: `HR_Titles`
   - second_approver_allowed_titles: `["Audit Leader", "Executive Auditor", "Audit Manager"]`

6. Review the configuration and click "Save Configuration"

### Step 2: Test the Configuration

1. Go to the "Testing" tab
2. Select "85 - Q2 Workpaper Approvals"
3. Click "Load Configuration"

4. Select "Generate Sample Data"
   - Set Number of Records: `100`
   - Set Error Percentage: `20`

5. Click "Run Test"

6. Review results:
   - Check Summary tab for overall compliance
   - Examine Detail tab to see which records failed
   - Verify that all validation rules are working correctly

7. If issues are found, return to the Configuration Wizard to adjust settings

### Step 3: Run with Real Data

1. Go to the "Run Analytics" tab
2. Select "85 - Q2 Workpaper Approvals"

3. Click "Browse..." and select your data file:
   - `Q2_Audit_Workpapers.xlsx`

4. Set output directory:
   - `output/Q2_2025`

5. Click "Run Analysis"

6. When processing completes, review the generated reports:
   - Open `output/Q2_2025/QA_85_Main_[timestamp].xlsx`
   - Check the Summary sheet for overall compliance
   - Review any DNC records in the Detail sheet

### Step 4: Schedule Regular Execution

1. Go to the "Scheduler" tab
2. Click "New Job"

3. Configure job settings:
   - Job ID: `q2_approvals_weekly`
   - Analytics ID: `85`
   - Schedule Type: `weekly`
   - Time: `08:00`
   - Day: `monday`

4. Set data file pattern:
   - `Q:\Audit\Workpapers\Weekly\*.xlsx`

5. Enable email notifications:
   - Check "Send Email Notification"
   - Add recipient: `audit.team@example.com`
   - Check "Generate Individual Reports"

6. Click "Save Job"

7. Click "Start Scheduler"

The system will now automatically run this analysis every Monday at 8:00 AM, using the most recent data file matching the pattern, and email results to the audit team.

## Best Practices

### Data Preparation

For best results:

1. **Consistent Naming**
   - Use consistent column names across files
   - Match column names to configuration or use data source registry

2. **Data Formatting**
   - Format dates consistently (preferably YYYY-MM-DD)
   - Remove extra spaces from text fields
   - Ensure numeric fields contain only numbers

3. **File Organization**
   - Use descriptive file names with dates
   - Keep files in organized folders by type/period
   - Archive old files regularly

### Configuration Management

To maintain a clean configuration library:

1. **Naming Conventions**
   - Use sequential Analytics IDs
   - Include period or category in Analytics Name
   - Add detailed descriptions

2. **Version Control**
   - Document changes in a change log
   - Create backups before significant changes
   - Consider using a file versioning system

3. **Validation Selection**
   - Only include necessary validations
   - Set appropriate thresholds based on risk
   - Test thoroughly before deployment

### Automation Strategy

For efficient automated processing:

1. **Scheduling Frequency**
   - Match to data update frequency
   - Avoid peak system usage times
   - Consider dependencies between analytics

2. **Notification Management**
   - Only notify relevant stakeholders
   - Include summary statistics in email
   - Set up issue-based alerts for exceptions

3. **Output Organization**
   - Use consistent folder structure
   - Include date/time in filenames
   - Archive older reports periodically

## System Administration

This section is primarily for administrators but provides useful context for all users.

### Directory Structure

The application uses the following directory structure:

```
enhanced_qa_analytics/
├── configs/              # Configuration files
│   ├── data_sources.yaml   # Data source registry
│   ├── reference_data.yaml # Reference data config
│   ├── scheduler.yaml      # Scheduler configuration
│   └── qa_XX.yaml          # Analytics configurations
├── templates/            # Template files
│   ├── approval_workflow.yaml
│   ├── risk_assessment.yaml
│   └── metadata.yaml
├── ref_data/             # Reference data files
├── output/               # Report output
├── logs/                 # Log files
└── temp/                 # Temporary files
```

### Configuration Files

Key configuration files:

1. **data_sources.yaml**
   - Central registry of data sources
   - Column mappings and validations
   - Used for standardization

2. **reference_data.yaml**
   - Defines reference data sources
   - Sets freshness parameters
   - Controls validation lookups

3. **scheduler.yaml**
   - Email configuration
   - Job definitions
   - Schedule settings

4. **qa_XX.yaml**
   - Individual analytics configurations
   - Created by Configuration Wizard
   - One file per analytic

### Backup Procedures

Regular backups should include:

1. **Configuration Directory**
   - `configs/` directory with all YAML files
   - Preserves all analytics configurations

2. **Reference Data**
   - `ref_data/` directory
   - Critical for title-based validations

3. **Templates**
   - `templates/` directory
   - Foundation for analytics creation

## Appendix A: Keyboard Shortcuts

The application supports the following keyboard shortcuts:

- **Ctrl+R**: Run the selected analytic
- **Ctrl+T**: Switch to Testing tab
- **Ctrl+W**: Switch to Configuration Wizard tab
- **Ctrl+S**: Switch to Scheduler tab
- **F1**: Show help
- **F5**: Refresh data source or reference data
- **Esc**: Cancel current operation

## Appendix B: Command Line Usage

The application can be run from the command line:

```
python enhanced_qa_analytics.py [options]
```

Options:
- `-a, --analytic-id`: Analytic ID to run
- `-s, --source-file`: Source data file path
- `-o, --output-dir`: Output directory
- `-i, --individual-reports`: Generate individual reports
- `--gui`: Start in GUI mode (default)

Example:
```
python enhanced_qa_analytics.py -a 77 -s data/approvals.xlsx -o output/may
```

## Appendix C: Validation Rule Parameters

Detailed parameters for each validation rule:

### Segregation of Duties
```yaml
rule: segregation_of_duties
parameters:
  submitter_field: Field containing submitter name
  approver_fields: 
    - First approver field
    - Second approver field
    # Add more approver fields as needed
```

### Approval Sequence
```yaml
rule: approval_sequence
parameters:
  date_fields_in_order:
    - First date field
    - Second date field
    - Third date field
    # Add more date fields in expected sequence
```

### Title-based Approval
```yaml
rule: title_based_approval
parameters:
  approver_field: Field with approver name
  allowed_titles:
    - "Title 1"
    - "Title 2"
    - "Title 3"
  title_reference: Reference data name
```

### Third-party Risk Validation
```yaml
rule: third_party_risk_validation
parameters:
  third_party_field: Field with third-party info
  risk_level_field: Field with risk rating
```

### Enumeration Validation
```yaml
rule: enumeration_validation
parameters:
  field_name: Field to validate
  valid_values:
    - "Value 1"
    - "Value 2"
    - "Value 3"
```

## Support and Contact Information

For additional assistance:

- **Email**: qa.support@example.com
- **Internal Portal**: https://intranet.example.com/qa-analytics
- **Help Desk**: Extension 1234

Super Users:
- Jane Smith (Finance) - ext. 5678
- John Davis (Audit) - ext. 9012

---

*This document was last updated on May 13, 2025.*
