"""
excel_mapper.py
---------------
Maps CSV data into Excel template using Excel COM (win32com).
This preserves ALL images, formatting, and embedded objects perfectly.
Now includes screenshot capture functionality.

CSV Format:
    Row 1: prog, id, <file_id>
    Row 2: X, Y, Z (headers)
    Row 3: X... Y... Z... (G54 - P1)
    Row 4: X... Y... Z... (G55 - P2)
    Row 5: X... Y... Z... (G56 - P3)
    Row 6: X... Y... Z... (G57 - P4)
"""

import logging
from pathlib import Path
from typing import Optional
import time

try:
    from . import config
    from . import screenshot_gui
except ImportError:
    import config
    try:
        import screenshot_gui
    except ImportError:
        screenshot_gui = None
        logging.warning("screenshot_gui module not found - screenshot feature disabled")

try:
    import pandas as pd
except ImportError:
    pd = None
    logging.warning("pandas not installed - Excel mapping will not work")

try:
    import win32com.client as win32
    from pywintypes import com_error
    WIN32COM_AVAILABLE = True
except ImportError:
    win32 = None
    WIN32COM_AVAILABLE = False
    logging.error("win32com not available - Excel mapping requires pywin32")


def map_csv_to_excel(csv_path: Path, template_path: Path, output_path: Path,
                     sheet_name: str, open_excel: bool = False,
                     enable_screenshots: bool = False) -> Optional[Path]:
    """
    Map CSV data into Excel template using Excel COM automation.
    This preserves all images, formatting, and embedded objects.
    
    Args:
        csv_path: Path to CSV file with coordinate data
        template_path: Path to Excel template file
        output_path: Path for output Excel file
        sheet_name: Name of worksheet to update
        open_excel: Whether to keep Excel visible after processing
        enable_screenshots: Whether to open screenshot GUI after mapping
        
    Returns:
        Path to output file on success, None on failure
    """

    # Check dependencies
    if not pd:
        logging.error("Missing pandas. Cannot map CSV to Excel.")
        return None
    
    if not WIN32COM_AVAILABLE:
        logging.error("Missing win32com (pywin32). Cannot map CSV to Excel.")
        logging.error("Install with: pip install pywin32")
        return None

    # Validate input files
    if not csv_path.exists():
        logging.error(f"CSV file not found: {csv_path}")
        return None
        
    if not template_path.exists():
        logging.error(f"Template file not found: {template_path}")
        return None

    excel_app = None
    wb = None
    
    try:
        # Read CSV data
        df = pd.read_csv(csv_path, header=None, encoding='utf-8', keep_default_na=False)

        # Validate CSV structure
        if df.shape[0] < 2:
            logging.error(f"CSV has insufficient rows: {csv_path}")
            return None

        # Extract metadata from CSV row 1
        program = df.iloc[0, 0] if df.shape[1] > 0 else ""
        id_field = df.iloc[0, 1] if df.shape[1] > 1 else ""
        file_id = df.iloc[0, 2] if df.shape[1] > 2 else ""

        logging.info(f"Mapping data - Program: {program}, ID: {id_field}, FileID: {file_id}")

        # Get formatting options from config
        try:
            add_axis_prefix = config.CONFIG.getboolean("FORMATTING", "ADD_AXIS_PREFIX", fallback=True)
            add_plus_sign = config.CONFIG.getboolean("FORMATTING", "ADD_PLUS_SIGN", fallback=True)
            decimal_places = config.CONFIG.getint("FORMATTING", "DECIMAL_PLACES", fallback=-1)
        except:
            add_axis_prefix = True
            add_plus_sign = True
            decimal_places = -1
        
        logging.info(f"Formatting: Prefix={add_axis_prefix}, PlusSign={add_plus_sign}, Decimals={decimal_places}")

        # Work offset mapping (G54->0, G55->1, G56->2, G57->3)
        offset_row_map = {
            'G54': 0,
            'G55': 1,
            'G56': 2,
            'G57': 3
        }

        # === USE EXCEL COM FOR EVERYTHING ===
        logging.info("Using Excel COM to preserve all images and formatting...")
        
        # Get or create Excel instance
        try:
            excel_app = win32.GetObject(Class="Excel.Application")
            logging.debug("Connected to existing Excel instance")
        except:
            excel_app = win32.Dispatch("Excel.Application")
            logging.debug("Created new Excel instance")
        
        # CRITICAL: Keep Excel hidden from the start
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        logging.debug("Excel set to hidden mode from start")
        
        # Open template
        logging.info(f"Opening template: {template_path}")
        wb = excel_app.Workbooks.Open(str(template_path.absolute()))
        
        # Get worksheet
        try:
            ws = wb.Worksheets(sheet_name)
        except:
            logging.error(f"Worksheet '{sheet_name}' not found in template")
            wb.Close(SaveChanges=False)
            return None
        
        logging.info(f"✓ Working with worksheet: {sheet_name}")

        # Process each mapping from config
        mapping_count = 0
        for csv_key, excel_cell in config.CONFIG["EXCEL_MAPPING"].items():
            try:
                value = None
                csv_key_upper = csv_key.upper()

                # Handle special keys
                if csv_key_upper == "PROGRAM_NUMBER":
                    value = file_id if file_id else ""
                elif csv_key_upper == "PROG":
                    value = program if program else ""
                elif csv_key_upper == "ID":
                    value = id_field if id_field else ""
                else:
                    # Parse work offset coordinate (e.g., G54_X, G55_Y)
                    parts = csv_key_upper.split('_')
                    if len(parts) == 2:
                        offset_name, axis = parts
                        
                        # Get row index for this work offset
                        coord_index = offset_row_map.get(offset_name)
                        
                        if coord_index is not None:
                            # Calculate actual CSV row (add 2 for header rows)
                            csv_row = coord_index + 2
                            
                            # Check if row exists
                            if csv_row < len(df):
                                # Get value from appropriate column
                                if axis == 'X' and df.shape[1] > 0:
                                    value = df.iloc[csv_row, 0]
                                elif axis == 'Y' and df.shape[1] > 1:
                                    value = df.iloc[csv_row, 1]
                                elif axis == 'Z' and df.shape[1] > 2:
                                    value = df.iloc[csv_row, 2]
                                
                                # Handle empty strings (leave blank)
                                if value == "" or pd.isna(value):
                                    value = ""
                                    logging.debug(f"  {csv_key}: Empty/blank")
                                elif value is not None:
                                    # Format coordinate value
                                    num_val = None
                                    
                                    if isinstance(value, str):
                                        clean_val = value.lstrip('XYZ')
                                        try:
                                            num_val = float(clean_val)
                                        except (ValueError, TypeError):
                                            pass
                                    else:
                                        try:
                                            num_val = float(value)
                                        except (ValueError, TypeError):
                                            pass
                                    
                                    if num_val is not None:
                                        # Round if configured
                                        if decimal_places >= 0:
                                            num_val = round(num_val, decimal_places)
                                        
                                        # Build formatted string
                                        formatted_value = ""
                                        
                                        if add_axis_prefix:
                                            formatted_value += axis
                                        
                                        if add_plus_sign and num_val >= 0:
                                            formatted_value += '+' + str(num_val)
                                        elif num_val < 0:
                                            formatted_value += str(num_val)
                                        else:
                                            formatted_value += str(num_val)
                                        
                                        value = formatted_value
                                        logging.debug(f"  {csv_key} = {value}")

                # Write value to Excel cell using COM
                if value is not None:
                    # Handle both single cells and ranges (use first cell)
                    target_cell = excel_cell.split(':')[0] if ':' in excel_cell else excel_cell
                    
                    try:
                        ws.Range(target_cell).Value = value
                        mapping_count += 1
                        if value == "":
                            logging.debug(f"Mapped {csv_key} -> {target_cell} = (blank)")
                        else:
                            logging.debug(f"Mapped {csv_key} -> {target_cell} = {value}")
                    except com_error as e:
                        logging.error(f"Error writing to cell {target_cell}: {e}")

            except Exception as e:
                logging.error(f"Error mapping {csv_key} -> {excel_cell}: {e}")

        logging.info(f"✓ Mapped {mapping_count} values to Excel")

        # Save as output file
        logging.info(f"Saving to: {output_path}")
        try:
            # Close the output file if it's already open in Excel
            for i in range(1, excel_app.Workbooks.Count + 1):
                try:
                    wb_check = excel_app.Workbooks(i)
                    if Path(wb_check.FullName).resolve() == output_path.resolve():
                        logging.info(f"Output file already open - closing it first")
                        wb_check.Close(SaveChanges=False)
                        break
                except:
                    pass
            
            # Save the workbook
            wb.SaveAs(str(output_path.absolute()))
            logging.info(f"✓ Excel file saved with all images preserved")
            
            # Close the workbook after saving (whether screenshots enabled or not)
            # The screenshot GUI will reopen it if needed
            try:
                wb.Close(SaveChanges=False)
                logging.debug("Closed workbook after saving")
                wb = None  # Mark as closed
            except:
                pass  # Already closed or disconnected
            
            # For the NEW workflow (ScreenshotGUIWithExcelCreation):
            # Just return the path - screenshot GUI handles everything else
            if not enable_screenshots and not open_excel:
                # Clean exit - file is saved and closed
                return output_path
                
        except com_error as e:
            logging.error(f"Error saving Excel file: {e}")
            logging.error(f"Error saving Excel file: {e}")
            
            # Try alternative: Save to temp location then copy
            try:
                import shutil
                temp_output = output_path.parent / f"~temp_{output_path.name}"
                wb.SaveAs(str(temp_output.absolute()))
                wb.Close(SaveChanges=False)
                
                # If original exists, try to delete it
                if output_path.exists():
                    try:
                        output_path.unlink()
                    except:
                        logging.error(f"Could not delete existing file: {output_path}")
                        return None
                
                # Move temp to final location
                shutil.move(str(temp_output), str(output_path))
                logging.info(f"✓ Excel file saved via temp file")
                
                # Mark workbook as closed
                wb = None
                
                # Return early if no screenshots needed
                if not enable_screenshots:
                    return output_path
                
                # Reopen for screenshot GUI if needed
                if enable_screenshots:
                    wb = excel_app.Workbooks.Open(str(output_path.absolute()))
                    ws = wb.Worksheets(sheet_name)
                    
            except Exception as e2:
                logging.error(f"Alternative save method also failed: {e2}")
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
                return None

        # === SCREENSHOT GUI INTEGRATION (OLD WORKFLOW ONLY) ===
        # This section is for the OLD workflow where screenshots are added AFTER Excel is created
        # The NEW workflow (ScreenshotGUIWithExcelCreation) doesn't use this
        if enable_screenshots and screenshot_gui:
            logging.info("Opening screenshot capture GUI (OLD workflow)...")
            
            # Reopen workbook if it was closed
            if wb is None:
                try:
                    wb = excel_app.Workbooks.Open(str(output_path.absolute()))
                    ws = wb.Worksheets(sheet_name)
                    logging.info("Reopened workbook for screenshot GUI")
                except Exception as e:
                    logging.error(f"Could not reopen workbook: {e}")
                    return output_path
            
            # Make sure we have a valid workbook reference
            try:
                _ = wb.Name
            except:
                logging.info("Reopening workbook for screenshot GUI (was closed)")
                wb = excel_app.Workbooks.Open(str(output_path.absolute()))
                ws = wb.Worksheets(sheet_name)
            
            # Open screenshot GUI
            try:
                screenshot_gui.open_screenshot_gui(wb, sheet_name)
                logging.info("✓ Screenshot GUI completed")
                
                try:
                    wb.Save()
                    logging.info("✓ Workbook saved with screenshots")
                except com_error as e:
                    logging.error(f"Error saving after screenshots: {e}")
                
                # Screenshot GUI handled everything, return
                return output_path
                    
            except Exception as e:
                logging.error(f"Error with screenshot GUI: {e}")
                logging.exception("Full traceback:")
                return None
        elif enable_screenshots and not screenshot_gui:
            logging.warning("Screenshots enabled but screenshot_gui module not available")
            return output_path

        # Keep Excel visible (OLD workflow only - when open_excel=True)
        if open_excel:
            if wb is None:
                # Reopen if needed
                wb = excel_app.Workbooks.Open(str(output_path.absolute()))
                ws = wb.Worksheets(sheet_name)
            
            # Show and maximize Excel
            excel_app.Visible = True
            excel_app.ScreenUpdating = True
            
            try:
                ws.Activate()
            except:
                pass
            
            time.sleep(0.2)
            try:
                import win32gui
                import win32con
                hwnd = excel_app.Hwnd
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                logging.info("✓ Excel maximized and brought to foreground")
            except Exception as e:
                logging.warning(f"Could not maximize Excel: {e}")
            
            logging.info("✓ Excel opened and visible")
        else:
            # Close workbook if it's still open and we don't need to show it
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                    logging.info("✓ Excel workbook closed")
                except:
                    pass

    except Exception as e:
        logging.error(f"Excel mapping error: {e}")
        logging.exception("Full traceback:")
        
        # Cleanup on error
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        
        return None
    
    finally:
        # Cleanup: restore Excel state if we're not keeping it visible
        if excel_app and not open_excel and not enable_screenshots:
            try:
                excel_app.ScreenUpdating = True
                excel_app.DisplayAlerts = True
            except:
                pass

    return output_path