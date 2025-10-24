"""
window_detector.py
------------------
Detects the active GibbsCAM filename from virtual.exe process windows.
"""

import re
import logging
from pathlib import Path
from typing import Optional, List
import os

try:
    import win32gui
    import win32process
    import psutil
    WIN32_AVAILABLE = True
except ImportError:
    win32gui = None
    win32process = None
    psutil = None
    WIN32_AVAILABLE = False
    logging.warning("win32gui/psutil not available - window detection disabled")


def get_virtual_exe_windows() -> List[tuple]:
    """
    Get all window handles and titles associated with virtual.exe process.
    
    Returns:
        List of tuples: [(hwnd, title, pid), ...]
    """
    if not WIN32_AVAILABLE:
        logging.error("Required libraries not available for window detection")
        return []
    
    virtual_windows = []
    
    def callback(hwnd, _):
        """Callback function for EnumWindows."""
        if win32gui.IsWindowVisible(hwnd):
            try:
                # Get process ID for this window
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                
                # Get process name
                try:
                    process = psutil.Process(pid)
                    process_name = process.name().lower()
                    
                    # Check if this is virtual.exe
                    if 'virtual.exe' in process_name:
                        title = win32gui.GetWindowText(hwnd).strip()
                        if title:
                            virtual_windows.append((hwnd, title, pid))
                            logging.debug(f"Found virtual.exe window: {title}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                    
            except Exception as e:
                logging.debug(f"Error checking window: {e}")
    
    try:
        win32gui.EnumWindows(callback, None)
        logging.info(f"Found {len(virtual_windows)} virtual.exe window(s)")
    except Exception as e:
        logging.error(f"Error enumerating windows: {e}")
    
    return virtual_windows


def extract_filename_from_title(title: str) -> Optional[str]:
    """
    Extract .NCF filename from window title.
    
    Looks for .vnc or .ncf extensions in title.
    
    Args:
        title: Window title string
        
    Returns:
        Filename string (e.g., "PROGRAM.NCF") or None if not found
    """
    # Look for .vnc or .ncf in window title
    # Pattern matches: "path\file.vnc" or "file.ncf" etc.
    match = re.search(r"([\w\-\s\\/.]+)\.(vnc|ncf)", title, re.IGNORECASE)
    if match:
        # Extract just the filename and ensure .NCF extension
        filename = Path(match.group(1)).name
        ncf_filename = f"{filename}.NCF"
        return ncf_filename
    
    return None


def get_active_gibbscam_file() -> Optional[str]:
    """
    Get the active GibbsCAM filename from virtual.exe process windows.
    
    Strategy:
    1. Find all windows belonging to virtual.exe process
    2. Look for child/secondary windows with .vnc or .ncf in title
    3. Return the first matching filename found
    
    Returns:
        Filename string (e.g., "PROGRAM.NCF") or None if not found
    """
    if not WIN32_AVAILABLE:
        logging.error("Cannot detect GibbsCAM file - missing required libraries")
        logging.error("Install with: pip install pywin32 psutil")
        return None
    
    virtual_windows = get_virtual_exe_windows()
    
    if not virtual_windows:
        logging.error("No virtual.exe process windows found")
        logging.error("Make sure GibbsCAM is running and a file is open")
        return None
    
    logging.info(f"Checking {len(virtual_windows)} virtual.exe window(s) for filename...")
    
    # Check each window for filename
    for hwnd, title, pid in virtual_windows:
        logging.debug(f"  Window: {title}")
        
        filename = extract_filename_from_title(title)
        if filename:
            logging.info(f"✓ Detected active file: {filename}")
            logging.debug(f"  From window title: {title}")
            return filename
    
    # If no direct match, log all titles for debugging
    logging.warning("Could not extract filename from any virtual.exe window")
    logging.info("Virtual.exe window titles found:")
    for hwnd, title, pid in virtual_windows:
        logging.info(f"  - {title}")
    
    return None


def search_ncf_in_network(filename: str, network_path: Path) -> Optional[Path]:
    """
    Search for NCF file in network path.
    
    Args:
        filename: Name of NCF file to find (e.g., "PROGRAM.NCF")
        network_path: Root directory to search
        
    Returns:
        Path to found file or None
    """
    if not filename:
        logging.error("No filename provided to search for")
        return None
    
    logging.info(f"Searching for '{filename}' in {network_path}")
    
    if not network_path.exists():
        logging.error(f"Network path does not exist: {network_path}")
        return None
    
    try:
        # Try direct match first (faster)
        direct_path = network_path / filename
        if direct_path.exists():
            logging.info(f"✓ Found NCF file (direct): {direct_path}")
            return direct_path
        
        # Try case-insensitive search in subdirectories
        logging.info("Searching subdirectories (this may take a moment)...")
        for root, dirs, files in os.walk(network_path):
            # Case-insensitive comparison
            for file in files:
                if file.lower() == filename.lower():
                    found_path = Path(root) / file
                    logging.info(f"✓ Found NCF file: {found_path}")
                    return found_path
            
            # Limit search depth to avoid too long search times
            depth = len(Path(root).relative_to(network_path).parts)
            if depth > 3:  # Only search 3 levels deep
                dirs.clear()  # Don't search deeper
                
    except Exception as e:
        logging.error(f"Error searching network: {e}")
    
    logging.error(f"✗ File '{filename}' not found in network path")
    return None


# Standalone test
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format="[%(levelname)s] %(message)s"
    )
    
    print("\n" + "="*70)
    print("TESTING WINDOW DETECTOR - virtual.exe Process")
    print("="*70 + "\n")
    
    if not WIN32_AVAILABLE:
        print("ERROR: Required libraries not installed")
        print("Install with: pip install pywin32 psutil")
        exit(1)
    
    # Test 1: Find virtual.exe windows
    print("[Test 1] Finding virtual.exe windows...")
    windows = get_virtual_exe_windows()
    if windows:
        print(f"✓ Found {len(windows)} virtual.exe window(s):")
        for hwnd, title, pid in windows:
            print(f"  - PID {pid}: {title}")
    else:
        print("✗ No virtual.exe windows found")
        print("  Make sure GibbsCAM is running")
    
    # Test 2: Extract filename
    print("\n[Test 2] Extracting filename from window titles...")
    for hwnd, title, pid in windows:
        filename = extract_filename_from_title(title)
        if filename:
            print(f"✓ Extracted: {filename}")
            print(f"  From: {title}")
    
    # Test 3: Full detection
    print("\n[Test 3] Complete detection workflow...")
    active_file = get_active_gibbscam_file()
    if active_file:
        print(f"✓ Active GibbsCAM file: {active_file}")
    else:
        print("✗ Could not detect active file")
        print("  Ensure:")
        print("  1. GibbsCAM (virtual.exe) is running")
        print("  2. A .vnc or .ncf file is open")
        print("  3. The filename is visible in the window title")
    
    print("\n" + "="*70)
    print("WINDOW DETECTOR TEST COMPLETE")
    print("="*70 + "\n")