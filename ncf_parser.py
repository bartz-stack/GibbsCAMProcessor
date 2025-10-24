"""
ncf_parser.py
-------------
Extracts coordinate data from GibbsCAM .NCF files and writes to CSV format.
Extracts from G10 L2 lines first, then VZOF arrays as backup.
"""

import re
import csv
import logging
from pathlib import Path
from typing import Optional
from . import config


def extract_coordinates(ncf_file: Path, csv_file: Path) -> Optional[Path]:
    """
    Extract coordinates from .NCF file and write to CSV.
    
    Priority order:
    1. G10 L2 P# lines (work offset commands)
    2. VZOFX/VZOFY/VZOFZ arrays (if G10 not found)
    
    CSV Format:
        Row 1: prog, id, <file_id>
        Row 2: X, Y, Z (headers)
        Row 3: G54 coordinates (P1)
        Row 4: G55 coordinates (P2)
        Row 5: G56 coordinates (P3)
        Row 6: G57 coordinates (P4)
    
    Args:
        ncf_file: Path to input .NCF file
        csv_file: Path to output CSV file
        
    Returns:
        Path to CSV file on success, None on failure
    """
    
    # Regex pattern for G10 L2 lines
    # Matches: G90 G10 L2 P1 X-64.7500 Y-24.7475 Z-21.9675
    g10_pattern = re.compile(
        r'G10\s+L2\s+P(\d+)\s+X(-?\d*\.?\d+)\s+Y(-?\d*\.?\d+)\s+Z(-?\d*\.?\d+)',
        re.IGNORECASE
    )
    
    # Load VZOF regex patterns from config as backup
    try:
        regex_vzofx = re.compile(config.CONFIG["REGEX"]["VZOFX"], re.IGNORECASE)
        regex_vzofy = re.compile(config.CONFIG["REGEX"]["VZOFY"], re.IGNORECASE)
        regex_vzofz = re.compile(config.CONFIG["REGEX"]["VZOFZ"], re.IGNORECASE)
    except KeyError as e:
        logging.warning(f"VZOF regex pattern missing in config [REGEX] section: {e}")
        regex_vzofx = regex_vzofy = regex_vzofz = None

    # Storage for coordinates
    # g10_data: {1: {'X': val, 'Y': val, 'Z': val}, 2: {...}, ...}
    g10_data = {}
    vzof_data = {}
    file_id = None

    # Read NCF file
    try:
        with open(ncf_file, "r", encoding="utf-8", errors="ignore") as f_in:
            lines = f_in.readlines()

            # Extract file ID from second line (common GibbsCAM pattern)
            if len(lines) > 1:
                potential_id = lines[1].strip().strip("()")
                if potential_id and not potential_id.startswith(('G', 'M', 'N', 'T', 'S', 'F')):
                    file_id = potential_id
                    
            # If no file ID found in line 2, try first 5 lines
            if not file_id:
                for i, line in enumerate(lines[:5]):
                    stripped = line.strip().strip("()")
                    if stripped and not stripped.startswith(('G', 'M', 'N', 'T', 'S', 'F', '%', 'O')):
                        file_id = stripped
                        break

            # Extract coordinates from all lines
            for line_num, line in enumerate(lines, 1):
                # Remove comments in parentheses
                clean_line = re.sub(r'\([^)]*\)', '', line).strip()
                
                # === PRIORITY 1: Check for G10 L2 lines ===
                match_g10 = g10_pattern.search(clean_line)
                if match_g10:
                    p_num, x, y, z = match_g10.groups()
                    p_num = int(p_num)
                    g10_data[p_num] = {
                        'X': float(x),
                        'Y': float(y),
                        'Z': float(z)
                    }
                    logging.info(f"Line {line_num}: G10 L2 P{p_num} X={x} Y={y} Z={z}")
                    continue
                
                # === BACKUP: Check for VZOF arrays ===
                if regex_vzofx:
                    match_x = regex_vzofx.search(clean_line)
                    if match_x:
                        idx, val = match_x.groups()
                        idx = int(idx)
                        if idx not in vzof_data:
                            vzof_data[idx] = {}
                        vzof_data[idx]['X'] = float(val)
                        logging.debug(f"Line {line_num}: VZOFX[{idx}] = {val}")
                        continue
                
                if regex_vzofy:
                    match_y = regex_vzofy.search(clean_line)
                    if match_y:
                        idx, val = match_y.groups()
                        idx = int(idx)
                        if idx not in vzof_data:
                            vzof_data[idx] = {}
                        vzof_data[idx]['Y'] = float(val)
                        logging.debug(f"Line {line_num}: VZOFY[{idx}] = {val}")
                        continue
                
                if regex_vzofz:
                    match_z = regex_vzofz.search(clean_line)
                    if match_z:
                        idx, val = match_z.groups()
                        idx = int(idx)
                        if idx not in vzof_data:
                            vzof_data[idx] = {}
                        vzof_data[idx]['Z'] = float(val)
                        logging.debug(f"Line {line_num}: VZOFZ[{idx}] = {val}")

    except FileNotFoundError:
        logging.error(f"NCF file not found: {ncf_file}")
        return None
    except Exception as e:
        logging.error(f"Error reading {ncf_file}: {e}")
        return None

    # Determine which data source to use
    if g10_data:
        logging.info(f"✓ Found {len(g10_data)} G10 L2 work offset(s)")
        coordinate_source = g10_data
        source_name = "G10 L2"
    elif vzof_data:
        logging.info(f"✓ Found {len(vzof_data)} VZOF offset(s)")
        coordinate_source = vzof_data
        source_name = "VZOF"
    else:
        logging.warning(f"No G10 L2 or VZOF coordinates found in {ncf_file}")
        logging.warning("Expected lines like:")
        logging.warning("  G10 L2 P1 X### Y### Z###")
        logging.warning("  or VZOFX[1] = ###")
        return None

    # Build coordinate list
    # Map P numbers to work offsets: P1=G54, P2=G55, P3=G56, P4=G57
    coordinates = []
    
    # Ensure we have data for P1-P4 (or indices 1-4 for VZOF)
    for idx in range(1, 5):  # P1, P2, P3, P4
        if idx in coordinate_source:
            x_val = coordinate_source[idx].get('X', 0.0)
            y_val = coordinate_source[idx].get('Y', 0.0)
            z_val = coordinate_source[idx].get('Z', 0.0)
            coordinates.append((f"X{x_val}", f"Y{y_val}", f"Z{z_val}"))
            logging.info(f"  P{idx}: X={x_val}, Y={y_val}, Z={z_val}")
        else:
            # No data for this offset - leave blank (empty strings)
            coordinates.append(("", "", ""))
            logging.debug(f"  P{idx}: No data found, leaving blank")

    # Write CSV output
    try:
        with open(csv_file, "w", newline="", encoding="utf-8") as f_out:
            writer = csv.writer(f_out)

            # Row 1: metadata with file ID
            writer.writerow(["prog", "id", file_id if file_id else ""])

            # Row 2: column headers
            writer.writerow(["X", "Y", "Z"])

            # Rows 3-6: coordinate data (G54, G55, G56, G57)
            writer.writerows(coordinates)

        logging.info(f"✓ CSV written: {csv_file}")
        logging.info(f"  Source: {source_name}")
        logging.info(f"  Work offsets: {len(coordinates)}")
        return csv_file
        
    except PermissionError:
        logging.error(f"Permission denied writing to {csv_file}")
        return None
    except Exception as e:
        logging.error(f"Error writing CSV {csv_file}: {e}")
        return None


def extract_vzof_offsets(ncf_file: Path) -> dict:
    """
    Extract VZOF offset values from .NCF file (if present).
    
    Returns:
        Dictionary with structure: {'X': {index: value}, 'Y': {}, 'Z': {}}
    """
    offsets = {'X': {}, 'Y': {}, 'Z': {}}
    
    try:
        regex_x = re.compile(config.CONFIG["REGEX"]["VZOFX"], re.IGNORECASE)
        regex_y = re.compile(config.CONFIG["REGEX"]["VZOFY"], re.IGNORECASE)
        regex_z = re.compile(config.CONFIG["REGEX"]["VZOFZ"], re.IGNORECASE)
    except KeyError as e:
        logging.warning(f"VZOF regex pattern missing in config: {e}")
        return offsets

    try:
        with open(ncf_file, "r", encoding="utf-8", errors="ignore") as f_in:
            for line in f_in:
                clean_line = line.strip()
                
                # Check for X offset
                match_x = regex_x.search(clean_line)
                if match_x:
                    idx, val = match_x.groups()
                    offsets['X'][int(idx)] = float(val)
                    continue
                    
                # Check for Y offset
                match_y = regex_y.search(clean_line)
                if match_y:
                    idx, val = match_y.groups()
                    offsets['Y'][int(idx)] = float(val)
                    continue
                    
                # Check for Z offset
                match_z = regex_z.search(clean_line)
                if match_z:
                    idx, val = match_z.groups()
                    offsets['Z'][int(idx)] = float(val)
                    
    except Exception as e:
        logging.error(f"Error extracting VZOF offsets from {ncf_file}: {e}")
    
    return offsets