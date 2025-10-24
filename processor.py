"""
processor.py
------------
Main control flow: load config, extract data, map to Excel, notify user.
Now includes screenshot capture functionality.
Works both as a module and standalone script.
"""

import logging
import time
from pathlib import Path
from typing import List

# Support both package and standalone imports
try:
    from . import config, logging_setup, notifications, ncf_parser, excel_mapper, window_detector
except ImportError:
    # Standalone mode - import directly
    import config
    import logging_setup
    import notifications
    import ncf_parser
    import excel_mapper
    import window_detector


def process_ncf_file(ncf_file: Path, template: Path, out_dir: Path, 
                     temp_csv_dir: Path, sheet_name: str, 
                     overwrite_prompt: bool, excel_visible: bool,
                     enable_screenshots: bool = True) -> bool:
    """
    Process a single NCF file: extract coordinates and create Excel report.
    
    Args:
        ncf_file: Path to .NCF file
        template: Path to Excel template
        out_dir: Output directory for Excel files
        temp_csv_dir: Temporary directory for CSV files
        sheet_name: Excel worksheet name
        overwrite_prompt: Whether to prompt before overwriting
        excel_visible: Whether to open Excel after processing
        enable_screenshots: Whether to open screenshot GUI after Excel insertion
        
    Returns:
        True if successful, False otherwise
    """
    import getpass
    
    logging.info(f"Processing: {ncf_file.name}")
    
    # Get username for unique file naming
    username = getpass.getuser()
    
    # Define output paths with username to prevent conflicts
    csv_file = temp_csv_dir / f"{username}_{ncf_file.stem}.csv"
    report_file = out_dir / f"{ncf_file.stem}_{username}.xlsx"

    # Check if report already exists
    if report_file.exists():
        if overwrite_prompt:
            logging.warning(f"Report already exists: {report_file}")
            if not notifications.confirm_overwrite(report_file):
                logging.info(f"Skipped (user declined overwrite): {ncf_file.name}")
                return False
        else:
            logging.info(f"Overwriting existing report: {report_file}")

    # Extract coordinates from NCF to CSV
    if not ncf_parser.extract_coordinates(ncf_file, csv_file):
        logging.error(f"Failed to extract coordinates from {ncf_file.name}")
        return False

    # NEW WORKFLOW: If screenshots enabled, capture FIRST then create Excel
    if enable_screenshots:
        logging.info("Opening screenshot GUI before Excel creation...")
        try:
            try:
                from . import screenshot_gui
            except ImportError:
                import screenshot_gui
            
            # This will open GUI, let user capture, then create Excel when Done clicked
            result = screenshot_gui.capture_then_create_excel(
                csv_file, template, report_file, sheet_name
            )
            
            if not result:
                logging.info("Screenshot capture cancelled or failed")
                return False
            
            logging.info("âœ“ Screenshots captured and Excel created successfully")
            return True
            
        except Exception as e:
            logging.error(f"Error with screenshot workflow: {e}")
            logging.exception("Full traceback:")
            return False
    else:
        # OLD WORKFLOW: No screenshots - just create Excel normally
        result = excel_mapper.map_csv_to_excel(
            csv_file, template, report_file, 
            sheet_name, excel_visible,
            enable_screenshots=False
        )
        
        if not result:
            logging.error(f"Failed to create Excel report for {ncf_file.name}")
            return False
        
        logging.info(f"âœ“ Successfully processed: {ncf_file.name}")
        return True


def main():
    """Main entry point for GibbsCAM processor."""
    
    # Load configuration
    try:
        cfg = config.load_config("config.ini")
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        print("Please ensure config.ini is in the same directory as the executable.")
        time.sleep(5)
        return
    except Exception as e:
        print(f"ERROR loading config: {e}")
        time.sleep(5)
        return

    # Get paths from config
    try:
        log_path = config.get_path("PATHS", "LOG_PATH")
        ncf_dir = config.get_path("PATHS", "NETWORK_PATH")
        template = config.get_path("PATHS", "REPORT_TEMPLATE")
        out_dir = config.get_path("PATHS", "REPORT_OUTPUT_PATH")
        temp_csv_dir = config.get_path("PATHS", "TEMP_CSV_PATH")
        sheet_name = config.get_value("PATHS", "REPORT_SHEET_NAME", "Setup Sheet")
        
        # Get icon path - check local directory first, then network path
        icon_path = None
        try:
            from pathlib import Path
            import sys
            
            # Get resource directory based on execution mode
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller executable
                if hasattr(sys, '_MEIPASS'):
                    bundle_dir = Path(sys._MEIPASS)
                    icon_candidates = [
                        bundle_dir / "Gibbscam.ico",
                        bundle_dir / "GCP_clau" / "Gibbscam.ico",
                    ]
                else:
                    bundle_dir = Path(sys.executable).parent
                    icon_candidates = [bundle_dir / "Gibbscam.ico"]
                
                exe_dir = Path(sys.executable).parent
                icon_candidates.append(exe_dir / "Gibbscam.ico")
                
            else:
                # Running as Python script/module
                if '__file__' in globals():
                    script_dir = Path(__file__).parent
                else:
                    script_dir = Path.cwd()
                
                icon_candidates = [
                    script_dir / "Gibbscam.ico",
                    script_dir.parent / "Gibbscam.ico",
                ]
            
            logging.info("Searching for icon in these locations:")
            for candidate in icon_candidates:
                logging.info(f"  - {candidate}")
                if candidate.exists():
                    icon_path = candidate
                    logging.info(f"âœ“ Found icon: {icon_path}")
                    break
            
            # If not found locally, try network path from config
            if not icon_path:
                try:
                    network_icon = config.get_path("PATHS", "ICON_PATH")
                    logging.info(f"Checking network icon: {network_icon}")
                    
                    if network_icon.exists():
                        icon_path = network_icon
                        logging.info(f"âœ“ Using network icon: {icon_path}")
                except Exception as e:
                    logging.debug(f"Could not load network icon: {e}")
        
        except Exception as e:
            logging.error(f"Icon search error: {e}")
        
        # Log final icon status
        if icon_path:
            logging.info(f"FINAL ICON: {icon_path}")
        else:
            logging.warning("NO ICON FOUND - toasts will show without icon")
            
    except ValueError as e:
        print(f"ERROR: {e}")
        time.sleep(5)
        return

    # Get behavior flags from config
    debug_mode = config.get_flag("BEHAVIOR", "DEBUG_MODE", False)
    toast_start = config.get_flag("BEHAVIOR", "TOAST_STARTUP", True)
    toast_finish = config.get_flag("BEHAVIOR", "TOAST_FINISH", True)
    toast_duration = config.get_int("BEHAVIOR", "TOAST_DURATION", 6)
    exit_delay = config.get_int("BEHAVIOR", "EXIT_DELAY", 2)
    force_gui = config.get_flag("BEHAVIOR", "FORCE_GUI", False)
    show_gui_on_success = config.get_flag("BEHAVIOR", "SHOW_GUI_ON_SUCCESS", False)
    show_gui_on_cancel = config.get_flag("BEHAVIOR", "SHOW_GUI_ON_CANCEL", False)
    excel_visible = config.get_flag("BEHAVIOR", "EXCEL_VISIBLE", True)
    overwrite_prompt = config.get_flag("BEHAVIOR", "OVERWRITE_PROMPT", True)
    force_winotify = config.get_flag("BEHAVIOR", "FORCE_WINOTIFY", False)
    enable_screenshots = config.get_flag("BEHAVIOR", "ENABLE_SCREENSHOTS", True)

    # Setup logging
    log_file = logging_setup.setup_logging(log_path, debug_mode)

    # Log screenshot status
    if enable_screenshots:
        logging.info("Screenshot capture is ENABLED")
    else:
        logging.info("Screenshot capture is DISABLED")

    # Validate critical paths
    if not ncf_dir.exists():
        logging.error(f"NCF directory not found: {ncf_dir}")
        notifications.show_toast("Error", "NCF directory not found", icon_path, toast_duration)
        time.sleep(exit_delay)
        return

    if not template.exists():
        logging.error(f"Excel template not found: {template}")
        notifications.show_toast("Error", "Excel template not found", icon_path, toast_duration)
        time.sleep(exit_delay)
        return

    # Create output directories if needed
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"âœ“ Output directory ready: {out_dir}")
    except Exception as e:
        logging.error(f"Could not create output directory {out_dir}: {e}")
        time.sleep(exit_delay)
        return
    
    try:
        temp_csv_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"âœ“ Temp CSV directory ready: {temp_csv_dir}")
    except Exception as e:
        logging.error(f"Could not create temp CSV directory {temp_csv_dir}: {e}")
        time.sleep(exit_delay)
        return

    # Get notification messages from config
    msg_startup_title = config.get_value("MESSAGES", "TOAST_STARTUP_TITLE", "GibbsCAM Processor")
    msg_startup_text = config.get_value("MESSAGES", "TOAST_STARTUP_MESSAGE", "Processing started...")
    msg_success_title = config.get_value("MESSAGES", "TOAST_SUCCESS_TITLE", "GibbsCAM Processor")
    msg_success_text = config.get_value("MESSAGES", "TOAST_SUCCESS_MESSAGE", "Success! Processed: {filename}")
    msg_failure_title = config.get_value("MESSAGES", "TOAST_FAILURE_TITLE", "GibbsCAM Processor")
    msg_failure_text = config.get_value("MESSAGES", "TOAST_FAILURE_MESSAGE", "Failed to process: {filename}")
    msg_error_no_file_title = config.get_value("MESSAGES", "TOAST_ERROR_NO_FILE_TITLE", "Error")
    msg_error_no_file_text = config.get_value("MESSAGES", "TOAST_ERROR_NO_FILE_MESSAGE", "No GibbsCAM file detected in open windows")
    msg_error_not_found_title = config.get_value("MESSAGES", "TOAST_ERROR_NOT_FOUND_TITLE", "Error")
    msg_error_not_found_text = config.get_value("MESSAGES", "TOAST_ERROR_NOT_FOUND_MESSAGE", "File not found: {filename}")
    msg_error_no_dir_title = config.get_value("MESSAGES", "TOAST_ERROR_NO_DIRECTORY_TITLE", "Error")
    msg_error_no_dir_text = config.get_value("MESSAGES", "TOAST_ERROR_NO_DIRECTORY_MESSAGE", "NCF directory not found")
    msg_error_no_template_title = config.get_value("MESSAGES", "TOAST_ERROR_NO_TEMPLATE_TITLE", "Error")
    msg_error_no_template_text = config.get_value("MESSAGES", "TOAST_ERROR_NO_TEMPLATE_MESSAGE", "Excel template not found")

    # Show startup notification
    if toast_start:
        notifications.show_toast(msg_startup_title, 
                               msg_startup_text, 
                               icon_path, 
                               toast_duration,
                               force_winotify=force_winotify)
        time.sleep(1)  # Give toast time to display

    # Detect active GibbsCAM file from window titles
    logging.info("Detecting active GibbsCAM file from window titles...")
    active_filename = window_detector.get_active_gibbscam_file()
    
    if not active_filename:
        logging.error("No active GibbsCAM file detected in window titles")
        notifications.show_toast("Error", 
                               "No GibbsCAM file detected in open windows", 
                               icon_path, 
                               toast_duration,
                               force_winotify=force_winotify)
        if force_gui:
            notifications.show_error_gui(log_file, "GibbsCAM Processor - Error")
        time.sleep(exit_delay)
        return
    
    logging.info(f"Active file detected: {active_filename}")
    
    # Search for the file in network path
    ncf_file = window_detector.search_ncf_in_network(active_filename, ncf_dir)
    
    if not ncf_file:
        logging.error(f"Could not find {active_filename} in network path")
        notifications.show_toast("Error", 
                               f"File not found: {active_filename}", 
                               icon_path, 
                               toast_duration,
                               force_winotify=force_winotify)
        if force_gui:
            notifications.show_error_gui(log_file, "GibbsCAM Processor - Error")
        time.sleep(exit_delay)
        return
    
    logging.info(f"File located: {ncf_file}")
    
    # Process the single detected file
    logging.info(f"Processing detected file: {ncf_file.name}")
    processed_count = 0
    user_cancelled = False
    
    try:
        if process_ncf_file(ncf_file, template, out_dir, temp_csv_dir,
                          sheet_name, overwrite_prompt, excel_visible,
                          enable_screenshots):
            processed_count = 1
        else:
            # Check if failure was due to user cancelling screenshot capture
            if enable_screenshots:
                user_cancelled = True
                logging.info('User cancelled screenshot capture')
    except Exception as e:
        logging.error(f"Unexpected error processing {ncf_file.name}: {e}")
        logging.exception("Full traceback:")
    
    # Log summary
    logging.info("=" * 60)
    if processed_count > 0:
        logging.info(f"âœ“ Successfully processed: {ncf_file.name}")
    elif user_cancelled:
        logging.warning(f"⚠ Cancelled: {ncf_file.name}")
    else:
        logging.error(f"âœ— Failed to process: {ncf_file.name}")
    logging.info("=" * 60)

    # Show completion notification
    if toast_finish:
        if processed_count > 0:
            msg = f"Success! Processed: {ncf_file.name}"
        else:
            msg = f"Failed to process: {ncf_file.name}"
        
        notifications.show_toast("GibbsCAM Processor", msg, icon_path, toast_duration, force_winotify=force_winotify)
        time.sleep(1)  # Give toast time to display

    # Show GUI if needed
    if force_gui:
        notifications.show_error_gui(log_file, "GibbsCAM Processor - Log")
    elif user_cancelled and show_gui_on_cancel:
        notifications.show_error_gui(log_file, "GibbsCAM Processor - Cancelled")
    elif logging_setup.has_errors():
        notifications.show_error_gui(log_file, "GibbsCAM Processor - Errors Detected")
    elif show_gui_on_success and processed_count > 0:
        notifications.show_success_gui(log_file, processed_count)

    # Wait before exit
    time.sleep(exit_delay)


if __name__ == "__main__":
    main()