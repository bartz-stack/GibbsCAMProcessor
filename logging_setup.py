"""
logging_setup.py
----------------
Configures logging system with file and console output.
Tracks if errors occurred during execution.
"""

import logging
import getpass
import threading
from pathlib import Path
from typing import Optional

# Event flag to detect errors during execution
error_occurred = threading.Event()


class ErrorFlagHandler(logging.Handler):
    """Custom handler that sets error flag when an error is logged."""
    
    def emit(self, record):
        """Called when a log record is emitted."""
        if record.levelno >= logging.ERROR:
            error_occurred.set()


def setup_logging(log_dir: Path, debug_mode: bool = False) -> Path:
    """
    Initialize logging system with file and console handlers.
    
    Args:
        log_dir: Directory for log files
        debug_mode: If True, set level to DEBUG instead of INFO
        
    Returns:
        Path to log file
    """
    
    # Create log directory if needed
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"Warning: Could not create log directory {log_dir}: {e}")
        # Fall back to current directory
        log_dir = Path.cwd()
    
    # Generate username-based log filename
    try:
        username = getpass.getuser()
    except Exception:
        username = "unknown"
    
    log_file = log_dir / f"{username}_gibbscam.log"

    # Clear any existing handlers to avoid duplicates
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)

    # Set logging level
    log_level = logging.DEBUG if debug_mode else logging.INFO

    # Create formatters
    detailed_formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    simple_formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S'
    )

    # File handler - detailed logging
    try:
        file_handler = logging.FileHandler(log_file, mode="w", encoding="utf-8")
        file_handler.setLevel(log_level)
        file_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(file_handler)
    except Exception as e:
        print(f"Warning: Could not create log file handler: {e}")

    # Console handler - simpler format
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(simple_formatter)
    root_logger.addHandler(console_handler)

    # Add custom error flag handler
    root_logger.addHandler(ErrorFlagHandler())

    # Set root logger level
    root_logger.setLevel(log_level)

    # Log startup message
    logging.info("=" * 60)
    logging.info("GibbsCAM Coordinate Processor - Starting")
    logging.info(f"Log file: {log_file}")
    logging.info(f"Log level: {logging.getLevelName(log_level)}")
    logging.info("=" * 60)
    
    return log_file


def has_errors() -> bool:
    """Check if any errors were logged during execution."""
    return error_occurred.is_set()


def reset_error_flag():
    """Reset the error flag (useful for testing or multi-run scenarios)."""
    error_occurred.clear()