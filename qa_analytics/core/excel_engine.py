"""
Excel Formula Engine for QA Analytics Framework.

This module provides functionality to process pandas DataFrames with Excel formulas
using Excel automation via win32com. It allows using Excel's native formula engine
for validation rules, eliminating the need to translate Excel formulas to pandas.

The ExcelFormulaProcessor class handles creating temporary Excel files, applying
formulas to data, and extracting results back into pandas DataFrames, while
ensuring proper cleanup of Excel resources.
"""

import os
import tempfile
import time
import uuid
import pandas as pd
from typing import Dict, List, Any, Optional, Tuple
import pythoncom
import pywintypes

# Handle import of win32com components with try/except to provide better error messages
try:
    from win32com.client import Dispatch, constants, gencache
    import win32com.client as win32
except ImportError:
    raise ImportError(
        "win32com is required for Excel formula processing. "
        "Please install pywin32 package using 'pip install pywin32'."
    )

# Set up logging
from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()


class ExcelFormulaProcessor:
    """
    Processes pandas DataFrames using Excel formulas via Excel automation.

    This class creates a bridge between pandas and Excel, allowing the use of
    native Excel formulas for data validation. It handles initializing Excel,
    creating temporary workbooks, applying formulas, and cleaning up resources.
    """

    def __init__(self, visible: bool = False, add_to_mru: bool = False):
        """
        Initialize the Excel Formula Processor.

        Args:
            visible: If True, Excel application will be visible during processing.
                    Default is False for headless operation.
            add_to_mru: If True, temporary files will be added to Excel's Recently
                    Used list. Default is False.
        """
        self.visible = visible
        self.add_to_mru = add_to_mru
        self.excel_app = None
        self.workbook = None
        self.worksheet = None
        self.temp_file = None
        self.initialized = False
        self.error_state = False
        
        # Performance settings
        self.calculation_mode = None  # Will store Excel's original calculation mode
        
        # Column mapping between DataFrame and Excel (1-indexed)
        self.column_map = {}  # Maps DataFrame column name to Excel column index

    def initialize_excel(self) -> bool:
        """
        Initialize Excel application via COM.
        
        This method starts Excel in the background, configures it for automation,
        and prepares it for processing.
        
        Returns:
            bool: True if Excel was successfully initialized, False otherwise
        """
        if self.initialized:
            return True
            
        try:
            logger.info("Initializing Excel application")
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            # Create Excel application object
            self.excel_app = Dispatch("Excel.Application")
            self.excel_app.Visible = self.visible
            self.excel_app.DisplayAlerts = False
            self.excel_app.ScreenUpdating = False
            self.excel_app.EnableEvents = False
            
            # Store original calculation mode and set to manual for performance
            self.calculation_mode = self.excel_app.Calculation
            self.excel_app.Calculation = constants.xlCalculationManual
            
            self.initialized = True
            return True
            
        except Exception as e:
            logger.error(f"Failed to initialize Excel: {e}")
            self.error_state = True
            self.cleanup()
            return False

    def __enter__(self):
        """
        Enter method for context manager support.

        Allows using the processor in a with statement:
        with ExcelFormulaProcessor() as processor:
            # processor will be automatically cleaned up

        Returns:
            Self for context manager usage
        """
        if not self.initialized:
            self.initialize_excel()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Exit method for context manager support.

        Ensures proper cleanup when leaving a with block.

        Args:
            exc_type: Exception type if an exception occurred
            exc_val: Exception value if an exception occurred
            exc_tb: Exception traceback if an exception occurred
        """
        self.cleanup()

        # Don't suppress exceptions
        return False

    def process_data_with_formulas(
        self, 
        df: pd.DataFrame, 
        formulas: Dict[str, str]
    ) -> Tuple[Optional[pd.DataFrame], List[str]]:
        """
        Process a DataFrame with Excel formulas.
        
        Creates a temporary Excel file with the data, applies formulas, and
        extracts results back into a pandas DataFrame.
        
        Args:
            df: Pandas DataFrame to process
            formulas: Dictionary of column_name: formula pairs
                      Formulas should be in Excel A1 notation
        
        Returns:
            Tuple of (result_df, warnings):
                result_df: DataFrame with formula results added as columns
                warnings: List of warning messages
        """
        warnings = []
        
        if self.error_state or not df.shape[0]:
            logger.error("Cannot process data: Excel error state or empty DataFrame")
            return None, ["Cannot process data: Excel is in error state or DataFrame is empty"]
        
        if not self.initialized and not self.initialize_excel():
            return None, ["Failed to initialize Excel"]
            
        try:
            # Create a temporary Excel file
            if not self._create_temp_workbook():
                return None, ["Failed to create temporary Excel workbook"]
                
            # Write data to Excel
            if not self._write_dataframe_to_excel(df):
                return None, ["Failed to write data to Excel worksheet"]
                
            # Apply formulas
            formula_results = {}
            for column_name, formula in formulas.items():
                result_range, result_warnings = self._apply_formula(formula, column_name)
                if result_range is not None:
                    formula_results[column_name] = result_range
                warnings.extend(result_warnings)
                
            # Read results back into DataFrame
            result_df = df.copy()
            
            # Add formula result columns to the DataFrame
            for column_name, result_range in formula_results.items():
                try:
                    values = self._convert_range_to_list(result_range)
                    if len(values) == len(df):
                        result_df[column_name] = values
                    else:
                        logger.warning(f"Formula result length mismatch: {len(values)} vs {len(df)}")
                        warnings.append(f"Formula result for '{column_name}' has incorrect length")
                except Exception as e:
                    logger.error(f"Error extracting formula results for {column_name}: {e}")
                    warnings.append(f"Could not extract formula results for '{column_name}'")
                    
            # Calculate once at the end for all formulas
            try:
                self.excel_app.Calculate()
            except Exception as e:
                logger.error(f"Error during Excel calculation: {e}")
                warnings.append(f"Error during final calculation: {str(e)}")
                
            return result_df, warnings
            
        except Exception as e:
            logger.error(f"Error processing data with formulas: {e}")
            self.error_state = True
            return None, [f"Error processing data with formulas: {str(e)}"]
            
        finally:
            # Save file for debugging if visible mode
            if self.visible and self.workbook:
                try:
                    debug_path = os.path.join(
                        tempfile.gettempdir(), 
                        f"qa_debug_{str(uuid.uuid4())[:8]}.xlsx"
                    )
                    self.workbook.SaveAs(debug_path)
                    logger.info(f"Saved debug workbook to {debug_path}")
                except Exception as e:
                    logger.warning(f"Could not save debug workbook: {e}")
                    
            # Proper cleanup if not in permanent visible mode
            if not self.visible:
                self.cleanup_workbook()

    def _create_temp_workbook(self) -> bool:
        """
        Create a temporary Excel workbook.
        
        Returns:
            bool: True if workbook was created successfully, False otherwise
        """
        try:
            # Create a temporary file
            fd, self.temp_file = tempfile.mkstemp(suffix='.xlsx')
            os.close(fd)
            
            # Create a new workbook
            self.workbook = self.excel_app.Workbooks.Add()
            
            # Use the first worksheet
            self.worksheet = self.workbook.Worksheets(1)
            self.worksheet.Name = "Data"
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to create temporary workbook: {e}")
            self.error_state = True
            return False

    def _write_dataframe_to_excel(self, df: pd.DataFrame) -> bool:
        """
        Write DataFrame to Excel worksheet.
        
        Creates column headers and writes data to the Excel worksheet.
        
        Args:
            df: Pandas DataFrame to write
            
        Returns:
            bool: True if data was written successfully, False otherwise
        """
        if not self.worksheet:
            return False
            
        try:
            # Add column headers in row 1
            columns = df.columns.tolist()
            for col_idx, col_name in enumerate(columns, 1):
                cell = self.worksheet.Cells(1, col_idx)
                cell.Value = col_name
                # Store mapping between DataFrame column name and Excel column index
                self.column_map[col_name] = col_idx
                
            # Check if there's data to write
            if df.shape[0] > 0:
                # Convert DataFrame to array of values
                values = df.values.tolist()
                
                # Insert data row by row
                for row_idx, row_data in enumerate(values, 2):  # Start from row 2 (after headers)
                    for col_idx, value in enumerate(row_data, 1):
                        cell = self.worksheet.Cells(row_idx, col_idx)
                        
                        # Handle different data types
                        if pd.isna(value):
                            cell.Value = None
                        elif isinstance(value, pd.Timestamp) or isinstance(value, pd.Period):
                            # Convert pandas timestamp to datetime
                            try:
                                cell.Value = value.to_pydatetime()
                            except:
                                cell.Value = str(value)
                        else:
                            cell.Value = value
                            
            # Auto-fit columns for better visibility (when in visible mode)
            if self.visible:
                self.worksheet.UsedRange.Columns.AutoFit()
                
            return True
            
        except Exception as e:
            logger.error(f"Failed to write DataFrame to Excel: {e}")
            self.error_state = True
            return False

    def _apply_formula(
        self, 
        formula: str, 
        result_column_name: str
    ) -> Tuple[Optional[Any], List[str]]:
        """
        Apply a formula to data in Excel and return the result range.
        
        Args:
            formula: Excel formula in A1 notation
            result_column_name: Name for the result column
            
        Returns:
            Tuple of (result_range, warnings):
                result_range: Excel Range object with formula results
                warnings: List of warning messages
        """
        warnings = []
        
        if not self.worksheet:
            return None, ["Worksheet not available"]
            
        try:
            # Get the last column for results
            last_col = len(self.column_map) + 1
            
            # Add header for result column
            self.worksheet.Cells(1, last_col).Value = result_column_name
            
            # Store in column map
            self.column_map[result_column_name] = last_col
            
            # Get row count (excluding header)
            row_count = self.worksheet.UsedRange.Rows.Count - 1
            if row_count <= 0:
                return None, ["No data rows to apply formula"]
                
            # Apply formula to first cell
            first_cell = self.worksheet.Cells(2, last_col)
            
            try:
                # Set formula for the first cell
                first_cell.Formula = formula
                
                # Check for errors in the first cell formula
                self.excel_app.Calculate()
                if self._is_error_cell(first_cell):
                    error_info = self._get_error_info(first_cell)
                    logger.warning(f"Formula error in first cell: {error_info}")
                    warnings.append(f"Formula '{formula}' resulted in error: {error_info}")
                    
                # Fill formula down to all rows if first cell is valid
                if row_count > 1:
                    try:
                        source_range = self.worksheet.Range(
                            self.worksheet.Cells(2, last_col),
                            self.worksheet.Cells(2, last_col)
                        )
                        target_range = self.worksheet.Range(
                            self.worksheet.Cells(2, last_col),
                            self.worksheet.Cells(row_count + 1, last_col)
                        )
                        source_range.AutoFill(target_range, constants.xlFillDefault)
                    except Exception as e:
                        logger.warning(f"Could not auto-fill formula: {e}")
                        warnings.append(f"Could not apply formula to all rows: {str(e)}")
                        
                # Get the range with results
                result_range = self.worksheet.Range(
                    self.worksheet.Cells(2, last_col),
                    self.worksheet.Cells(row_count + 1, last_col)
                )
                
                return result_range, warnings
                
            except Exception as e:
                logger.error(f"Error applying formula '{formula}': {e}")
                return None, [f"Error applying formula '{formula}': {str(e)}"]
                
        except Exception as e:
            logger.error(f"Unexpected error applying formula: {e}")
            return None, [f"Unexpected error applying formula: {str(e)}"]

    def _convert_range_to_list(self, excel_range: Any) -> List[Any]:
        """
        Convert an Excel range to a Python list.
        
        Args:
            excel_range: Excel Range object
            
        Returns:
            List of values from the range
        """
        values = []
        
        try:
            # Get values from Excel range as a tuple of tuples
            range_values = excel_range.Value
            
            # Convert to list based on the structure
            if isinstance(range_values, tuple):
                # Single cell
                if not isinstance(range_values[0], tuple):
                    values = [range_values]
                else:
                    # Range of cells
                    for row in range_values:
                        if isinstance(row, tuple) and len(row) > 0:
                            values.append(row[0])
                        else:
                            values.append(None)
            else:
                # Empty range or unexpected format
                values = [None] * excel_range.Rows.Count
                
            # Convert Excel error values to None
            values = [None if self._is_excel_error(v) else v for v in values]
                
        except Exception as e:
            logger.error(f"Error converting Excel range to list: {e}")
            # Return a list of None values with the correct length
            values = [None] * excel_range.Rows.Count
            
        return values

    def _is_excel_error(self, value: Any) -> bool:
        """
        Check if a value is an Excel error code.
        
        Args:
            value: Value to check
            
        Returns:
            bool: True if the value is an Excel error code
        """
        # Common Excel error codes
        error_codes = [
            constants.xlErrDiv0,   # Division by zero
            constants.xlErrNA,     # Value not available
            constants.xlErrName,   # Name error
            constants.xlErrNull,   # Null value error
            constants.xlErrNum,    # Number error
            constants.xlErrRef,    # Reference error
            constants.xlErrValue   # Value error
        ]
        
        try:
            # Check if value is a COM error object
            if isinstance(value, int) and value in error_codes:
                return True
        except:
            pass
            
        return False

    def _is_error_cell(self, cell: Any) -> bool:
        """
        Check if a cell contains an error.
        
        Args:
            cell: Excel cell object
            
        Returns:
            bool: True if the cell contains an error
        """
        try:
            return cell.Errors.Item(constants.xlErrDiv0).Value or \
                   cell.Errors.Item(constants.xlErrNA).Value or \
                   cell.Errors.Item(constants.xlErrName).Value or \
                   cell.Errors.Item(constants.xlErrNull).Value or \
                   cell.Errors.Item(constants.xlErrNum).Value or \
                   cell.Errors.Item(constants.xlErrRef).Value or \
                   cell.Errors.Item(constants.xlErrValue).Value
        except:
            # If we can't check errors properly, assume it's not an error
            return False

    def _get_error_info(self, cell: Any) -> str:
        """
        Get information about an error in a cell.
        
        Args:
            cell: Excel cell object
            
        Returns:
            str: Description of the error
        """
        try:
            error_type = "Unknown"
            error_types = {
                constants.xlErrDiv0: "Division by zero",
                constants.xlErrNA: "Value not available",
                constants.xlErrName: "Name error",
                constants.xlErrNull: "Null value error",
                constants.xlErrNum: "Number error",
                constants.xlErrRef: "Reference error",
                constants.xlErrValue: "Value error"
            }
            
            for error_code, error_desc in error_types.items():
                if cell.Errors.Item(error_code).Value:
                    error_type = error_desc
                    break
                    
            return f"Excel error: {error_type}"
            
        except:
            return "Unidentified Excel error"

    def cleanup_workbook(self) -> None:
        """
        Clean up the workbook resources with enhanced error handling.

        This closes the current workbook and deletes the temporary file,
        with multiple safeguards to ensure resources are released.
        """
        try:
            # Close workbook without saving - with enhanced error handling
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=False)
                except Exception as e:
                    logger.warning(f"Error closing workbook: {e}")
                finally:
                    self.workbook = None
                    self.worksheet = None

            # Delete temporary file with verification
            if self.temp_file and os.path.exists(self.temp_file):
                try:
                    # Sometimes file may be locked, retry a few times
                    max_attempts = 3
                    for attempt in range(max_attempts):
                        try:
                            os.unlink(self.temp_file)
                            break
                        except PermissionError:
                            if attempt < max_attempts - 1:
                                time.sleep(0.5)  # Wait before retry
                            else:
                                raise
                except Exception as e:
                    logger.warning(f"Error deleting temporary file {self.temp_file}: {e}")
                self.temp_file = None

            # Reset column mapping
            self.column_map = {}

        except Exception as e:
            logger.warning(f"Error during workbook cleanup: {e}")

    def cleanup(self) -> None:
        """
        Clean up all Excel resources with enhanced safeguards.

        This closes the workbook and quits the Excel application,
        with multiple levels of error handling to ensure resources
        are released even in exceptional circumstances.
        """
        try:
            # First cleanup workbook
            self.cleanup_workbook()

            # Then quit Excel
            if self.excel_app:
                try:
                    # Restore original calculation mode
                    if self.calculation_mode is not None:
                        try:
                            self.excel_app.Calculation = self.calculation_mode
                        except Exception as e:
                            logger.warning(f"Error restoring calculation mode: {e}")

                    # Restore settings
                    try:
                        self.excel_app.DisplayAlerts = True
                        self.excel_app.ScreenUpdating = True
                        self.excel_app.EnableEvents = True
                    except Exception as e:
                        logger.warning(f"Error restoring Excel settings: {e}")

                    # Quit Excel (if not in visible mode)
                    if not self.visible:
                        try:
                            # Force garbage collection before quitting Excel
                            # This helps release any COM references
                            import gc
                            gc.collect()

                            # Set a reasonable timeout for quitting
                            quit_timeout = 5  # seconds
                            quit_start = time.time()

                            self.excel_app.Quit()

                            # Wait for Excel to actually quit
                            while time.time() - quit_start < quit_timeout:
                                try:
                                    # If we can get a property, Excel is still running
                                    test = self.excel_app.Visible
                                    time.sleep(0.5)
                                except:
                                    # Excel has quit
                                    break
                        except Exception as e:
                            logger.warning(f"Error quitting Excel: {e}")

                    self.excel_app = None

                except Exception as e:
                    logger.warning(f"Error cleaning up Excel application: {e}")
                    self.excel_app = None

                # Uninitialize COM
                try:
                    pythoncom.CoUninitialize()
                except Exception as e:
                    logger.warning(f"Error uninitializing COM: {e}")

                self.initialized = False
                self.error_state = False

        except Exception as e:
            logger.warning(f"Error during Excel cleanup: {e}")

            # Last resort - try to terminate Excel forcefully
            try:
                if self.excel_app:
                    import win32process
                    import win32api

                    # Get Excel's process ID and terminate it
                    try:
                        pid = win32process.GetWindowThreadProcessId(
                            self.excel_app.Hwnd)[1]
                        handle = win32api.OpenProcess(1, 0, pid)
                        win32api.TerminateProcess(handle, 0)
                        win32api.CloseHandle(handle)
                        logger.warning("Forcefully terminated Excel process")
                    except:
                        pass

                    self.excel_app = None
            except:
                pass

            # Reset state regardless of errors
            self.initialized = False
            self.error_state = True

    def __del__(self):
        """
        Destructor to ensure all Excel resources are cleaned up.
        """
        self.cleanup()


def ensure_excel_closed(force=False) -> bool:
    """
    Utility function to ensure Excel processes are properly closed.

    This is a safety net function that can be called at program exit or
    after batch processing to make sure no orphaned Excel processes remain.

    Args:
        force: If True, will attempt to forcefully terminate Excel processes
               that don't respond to graceful shutdown requests

    Returns:
        bool: True if all Excel instances were closed successfully
    """
    success = True

    try:
        # Try multiple approaches to ensure Excel is really closed
        excel_closed = False

        # First try the COM approach
        try:
            # Try to connect to any running Excel instances
            excel = win32.GetActiveObject("Excel.Application")

            # If we got here, Excel is running - attempt to close gracefully
            try:
                logger.info("Found running Excel instance, attempting to close gracefully")
                excel.DisplayAlerts = False

                # Close all workbooks
                open_workbooks = excel.Workbooks.Count
                for i in range(open_workbooks, 0, -1):  # Count backwards to avoid index issues
                    try:
                        excel.Workbooks(i).Close(SaveChanges=False)
                    except:
                        pass

                # Quit Excel and release references
                excel.Quit()
                excel = None

                # Force garbage collection to release COM references
                import gc
                gc.collect()

                # Note success
                logger.info("Successfully closed Excel via COM")
                excel_closed = True
            except Exception as e:
                logger.warning(f"Error when trying to close Excel gracefully: {e}")
                success = False

        except pywintypes.com_error:
            # No active Excel instance found via COM
            excel_closed = True

        # If COM approach failed or if force=True, try process-based approach
        if not excel_closed or force:
            try:
                import subprocess
                import platform

                # Different commands based on OS
                if platform.system() == 'Windows':
                    # Check if Excel is running
                    output = subprocess.check_output('tasklist /FI "IMAGENAME eq EXCEL.EXE"', shell=True).decode()

                    if 'EXCEL.EXE' in output:
                        logger.warning("Excel still running, attempting to terminate via taskkill")

                        # Try graceful termination first
                        try:
                            subprocess.call('taskkill /IM EXCEL.EXE', shell=True)
                            logger.info("Terminated Excel processes with taskkill")
                        except:
                            # If that fails and force=True, use /F to force termination
                            if force:
                                try:
                                    subprocess.call('taskkill /F /IM EXCEL.EXE', shell=True)
                                    logger.warning("Forcefully terminated Excel processes")
                                except Exception as e:
                                    logger.error(f"Failed to forcefully terminate Excel: {e}")
                                    success = False
                elif platform.system() == 'Darwin':  # macOS
                    # Check if Excel is running
                    try:
                        output = subprocess.check_output('pgrep "Microsoft Excel"', shell=True).decode()
                        if output.strip():
                            logger.warning("Excel still running on macOS, attempting to terminate")
                            subprocess.call('pkill "Microsoft Excel"', shell=True)
                            logger.info("Terminated Excel processes on macOS")
                    except subprocess.CalledProcessError:
                        # If pgrep returns non-zero, Excel is not running
                        pass
                else:  # Linux or other
                    # Generic approach - less reliable
                    try:
                        subprocess.call('pkill -f excel', shell=True)
                    except:
                        pass
            except Exception as e:
                logger.error(f"Error during process-based Excel termination: {e}")
                success = False

        # Final verification - try COM approach again to see if Excel is still running
        try:
            excel = win32.GetActiveObject("Excel.Application")
            # If we get here, Excel is still running
            logger.warning("Excel is still running after cleanup attempts")
            success = False
        except pywintypes.com_error:
            # Excel is not running - success!
            if not excel_closed:  # If this is the first confirmation
                logger.info("Verified Excel is no longer running")

        return success

    except Exception as e:
        logger.error(f"Error when checking for Excel processes: {e}")
        return False

# Example usage
if __name__ == "__main__":
    # Basic test of Excel Formula Processor
    processor = ExcelFormulaProcessor(visible=True)
    
    # Create test data
    test_data = pd.DataFrame({
        'ID': [1, 2, 3, 4, 5],
        'Value1': [10, 20, 30, 40, 50],
        'Value2': [5, 15, 25, 35, 45],
        'Date': pd.date_range(start='2025-01-01', periods=5)
    })
    
    # Define test formulas
    test_formulas = {
        'Sum': '=Value1 + Value2',
        'IsGreaterThan40': '=IF(Sum > 40, TRUE, FALSE)',
        'DateCheck': '=IF(Date > TODAY(), "Future", "Past")'
    }
    
    # Process data with formulas
    result_df, warnings = processor.process_data_with_formulas(test_data, test_formulas)
    
    # Print results
    if result_df is not None:
        print("Results:")
        print(result_df)
        
    if warnings:
        print("\nWarnings:")
        for warning in warnings:
            print(f"- {warning}")
            
    # Clean up
    processor.cleanup()
