"""
screenshot_gui.py
-----------------
GUI for capturing screenshots and mapping them to Excel cells.
Allows user to take multiple screenshots, review them, retake if needed,
and automatically insert them into specified Excel cells.

CLEAN VERSION - v74 (with Excel maximize fix on Done!)
"""

import logging
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import List, Optional, Tuple
import tempfile
import shutil

try:
    from PIL import Image, ImageTk, ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    logging.error("PIL/Pillow not available - install with: pip install pillow")

try:
    import win32com.client as win32
    from pywintypes import com_error
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    logging.error("win32com not available - install with: pip install pywin32")


class ScreenshotCapture:
    """Handles screenshot capture with a fixed size selection."""
    
    def __init__(self, width_inches=3.0, height_inches=2.5, dpi=96):
        """
        Initialize screenshot capture with specific dimensions.
        Sized to match Excel cell dimensions better (approximately 3" x 2.5").
        
        Args:
            width_inches: Width in inches (default 3.0")
            height_inches: Height in inches (default 2.5")
            dpi: Screen DPI (default 96)
        """
        self.width_px = int(width_inches * dpi)
        self.height_px = int(height_inches * dpi)
        self.capture_window = None
        self.canvas = None
        self.rect = None
        self.start_x = None
        self.start_y = None
        self.screenshot = None
        
    def start_capture(self, callback):
        """Start the screenshot capture process."""
        if not PIL_AVAILABLE:
            logging.error("Cannot capture screenshot - Pillow not installed")
            return
        
        self.callback = callback
        
        # Create fullscreen transparent overlay
        self.capture_window = tk.Toplevel()
        self.capture_window.attributes('-fullscreen', True)
        self.capture_window.attributes('-alpha', 0.3)
        self.capture_window.attributes('-topmost', True)
        self.capture_window.config(cursor='cross')
        
        # Create canvas
        self.canvas = tk.Canvas(
            self.capture_window,
            bg='black',
            highlightthickness=0
        )
        self.canvas.pack(fill='both', expand=True)
        
        # Instructions
        instructions = (
            f"Click and drag to position the capture area ({self.width_px}x{self.height_px}px)\n"
            "Press ESC to cancel"
        )
        self.canvas.create_text(
            self.capture_window.winfo_screenwidth() // 2,
            50,
            text=instructions,
            fill='white',
            font=('Arial', 14, 'bold')
        )
        
        # Bind events
        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.capture_window.bind('<Escape>', lambda e: self.cancel_capture())
        
    def on_mouse_down(self, event):
        """Handle mouse button press."""
        self.start_x = event.x
        self.start_y = event.y
        
        if self.rect:
            self.canvas.delete(self.rect)
        
        self.rect = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x + self.width_px,
            self.start_y + self.height_px,
            outline='red',
            width=3
        )
        
    def on_mouse_drag(self, event):
        """Handle mouse drag to position rectangle."""
        if self.rect:
            self.canvas.delete(self.rect)
        
        self.start_x = event.x
        self.start_y = event.y
        
        self.rect = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x + self.width_px,
            self.start_y + self.height_px,
            outline='red',
            width=3
        )
        
    def on_mouse_up(self, event):
        """Handle mouse button release - capture the area."""
        if not self.start_x or not self.start_y:
            return
        
        self.capture_window.withdraw()
        self.capture_window.update()
        self.capture_window.after(100, self.perform_capture)
        
    def perform_capture(self):
        """Perform the actual screenshot capture."""
        try:
            bbox = (
                self.start_x,
                self.start_y,
                self.start_x + self.width_px,
                self.start_y + self.height_px
            )
            
            self.screenshot = ImageGrab.grab(bbox)
            self.capture_window.destroy()
            
            if self.callback:
                self.callback(self.screenshot)
                
        except Exception as e:
            logging.error(f"Error capturing screenshot: {e}")
            self.cancel_capture()
    
    def cancel_capture(self):
        """Cancel the capture process."""
        if self.capture_window:
            self.capture_window.destroy()
        if self.callback:
            self.callback(None)


class ScreenshotGUI:
    """Main GUI for managing multiple screenshots."""
    
    def __init__(self, excel_workbook, worksheet_name):
        """Initialize the screenshot GUI."""
        if not PIL_AVAILABLE:
            raise ImportError("Pillow is required for screenshot functionality")
        
        if not WIN32COM_AVAILABLE:
            raise ImportError("pywin32 is required for Excel integration")
        
        self.workbook = excel_workbook
        self.worksheet_name = worksheet_name
        self.screenshots = {}
        self.temp_files = []
        self.temp_dir = Path(tempfile.mkdtemp(prefix="gibbscam_screenshots_"))
        
        # Load position mapping from config
        self.position_mapping = self._load_position_mapping()
        
        # Track visible positions
        max_pos = max(self.position_mapping.keys()) if self.position_mapping else 4
        initial_max = min(4, max_pos)
        self.visible_positions = list(range(1, initial_max + 1))
        self.position_widgets = {}
        
        logging.info(f"Initialization: max_pos={max_pos}, showing positions {self.visible_positions}")
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("GibbsCAM Screenshot Tool")
        
        # Set window size and center it
        window_width = 880
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.minsize(860, 650)
        self.root.configure(bg="#30302e")
        
        self.setup_ui()

        # NEW: align windows so GibbsCAM is behind the GUI at startup
        self.root.after(450, self._align_with_gibbscam_under_gui)
        
    def _load_position_mapping(self):
        """Load screenshot position to cell mapping from config."""
        try:
            try:
                from . import config
            except ImportError:
                import config
            
            mapping = {}
            
            if not config.CONFIG:
                logging.warning("Config not loaded yet - attempting to load")
                try:
                    config.load_config("config.ini")
                except Exception as e:
                    logging.error(f"Failed to load config: {e}")
            
            if config.CONFIG and "SCREENSHOT_MAPPING" in config.CONFIG:
                logging.info("Found SCREENSHOT_MAPPING section in config")
                logging.info(f"Section contains {len(config.CONFIG['SCREENSHOT_MAPPING'])} items")
                
                all_keys = list(config.CONFIG["SCREENSHOT_MAPPING"].keys())
                logging.info(f"All keys in section: {all_keys}")
                
                for key, value in config.CONFIG["SCREENSHOT_MAPPING"].items():
                    logging.info(f"  Processing key: '{key}' = '{value}'")
                    
                    if key.upper().startswith("POSITION_"):
                        try:
                            pos_num = int(key.split("_")[1])
                            mapping[pos_num] = value
                            logging.info(f"    ✓ Loaded: Position {pos_num} -> {value}")
                        except (IndexError, ValueError) as e:
                            logging.warning(f"    ✗ Invalid key format: {key} - {e}")
                    else:
                        logging.info(f"    - Skipping non-position key: {key}")
            else:
                logging.warning("SCREENSHOT_MAPPING section not found in config")
            
            if mapping:
                logging.info(f"✓ Loaded {len(mapping)} screenshot position mappings from config")
                logging.info(f"  Positions configured: {sorted(mapping.keys())}")
                return mapping
            else:
                logging.warning("No screenshot mappings found in config, using defaults")
                
        except Exception as e:
            logging.error(f"Error loading screenshot mapping from config: {e}")
            logging.exception("Full traceback:")
        
        logging.warning("Using hardcoded defaults (4 positions)")
        return {1: "G9", 2: "G31", 3: "A63", 4: "G63"}

    # ---------- NEW HELPER FOR STARTUP Z-ORDER ----------
    def _align_with_gibbscam_under_gui(self):
        """
        Bring GibbsCAM (virtual.exe) to foreground, then lift this Tk window above it.
        Result: GUI is clickable, Gibbs is visible directly behind it.
        """
        try:
            import win32gui
            import win32process
            import psutil

            def find_virtual_exe():
                windows = []
                def callback(hwnd, _):
                    if win32gui.IsWindowVisible(hwnd):
                        try:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            if 'virtual.exe' in psutil.Process(pid).name().lower():
                                windows.append(hwnd)
                        except Exception:
                            pass
                win32gui.EnumWindows(callback, None)
                return windows

            vws = find_virtual_exe()
            if vws:
                hwnd = vws[0]
                win32gui.ShowWindow(hwnd, 9)          # SW_RESTORE
                win32gui.SetForegroundWindow(hwnd)    # Gibbs up
                logging.info("✓ GibbsCAM foregrounded (startup)")

            # Lift GUI above GibbsCAM (temporary topmost toggle)
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(250, lambda: self.root.attributes("-topmost", False))
            logging.info("✓ GUI lifted above GibbsCAM")

        except Exception as e:
            logging.warning(f"Could not align GUI with GibbsCAM on startup: {e}")
    # ----------------------------------------------------
        
    def setup_ui(self):
        """Setup the user interface."""
        # Instructions
        instructions_text = 'Click "Capture" - click and hold then position the red rectangle over the area you want.'
        
        instructions_frame = tk.Frame(self.root, bg="#30302e")
        instructions_frame.pack(fill="x", padx=10, pady=8)
        
        tk.Label(
            instructions_frame,
            text=instructions_text,
            font=("Arial", 9, "bold"),
            bg="#30302e",
            wraplength=850,
            justify="center",
            fg="white"
        ).pack(padx=10, pady=8)
        
        # Main content area with scrollbar
        content_container = tk.Frame(self.root, bg="#30302e")
        content_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Create canvas for scrolling
        self.canvas = tk.Canvas(content_container, bg="#30302e", highlightthickness=0)
        scrollbar = tk.Scrollbar(content_container, orient="vertical", command=self.canvas.yview)
        self.content_frame = tk.Frame(self.canvas, bg="#30302e")
        
        self.content_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind mouse wheel scrolling
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create initial position boxes
        self.update_position_display()
        
        # Setup bottom button bar
        self.setup_bottom_bar()
    
    def setup_bottom_bar(self):
        """Setup bottom button bar - SINGLE FRAME DESIGN."""
        # Remove existing if present
        if hasattr(self, 'button_frame'):
            try:
                self.button_frame.destroy()
            except:
                pass
        
        # Create single dark frame
        self.button_frame = tk.Frame(self.root, bg="#30302e", height=70)
        self.button_frame.pack(fill="x", side="bottom")
        self.button_frame.pack_propagate(False)
        
        # Calculate button visibility
        all_position_numbers = sorted(self.position_mapping.keys())
        max_position_number = all_position_numbers[-1] if all_position_numbers else 0
        highest_visible = max(self.visible_positions) if self.visible_positions else 0
        
        # Add More button
        if highest_visible < max_position_number:
            self.add_more_button = tk.Button(
                self.button_frame,
                text="Add More",
                command=self.add_more_positions,
                font=("Arial", 9, "bold"),
                bg="#0078d4",
                fg="white",
                padx=20,
                pady=8,
                cursor="hand2",
                relief="raised",
                bd=1
            )
            self.add_more_button.place(x=15, y=15)
        
        # Status label
        self.status_label = tk.Label(
            self.button_frame,
            text=f"{len(self.screenshots)}/{len(self.position_mapping)} captured",
            font=("Arial", 10, "bold"),
            bg="#30302e",
            fg="white"
        )
        self.status_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # Cancel button
        cancel_button = tk.Button(
            self.button_frame,
            text="Cancel",
            command=self.cancel,
            font=("Arial", 9, "bold"),
            bg="#e74c3c",
            fg="white",
            padx=25,
            pady=8,
            cursor="hand2",
            relief="raised",
            bd=1
        )
        cancel_button.place(relx=1.0, x=-180, y=15, anchor="nw")
        
        # Done button
        self.done_button = tk.Button(
            self.button_frame,
            text="Done",
            command=self.finish_and_insert,
            font=("Arial", 9, "bold"),
            bg="#28a616",
            fg="white",
            padx=25,
            pady=8,
            cursor="hand2",
            state="disabled" if len(self.screenshots) == 0 else "normal",
            relief="raised",
            bd=1
        )
        self.done_button.place(relx=1.0, x=-85, y=15, anchor="nw")
    
    def update_position_display(self):
        """Update the display of position boxes."""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        self.position_widgets.clear()
        
        row = 0
        col = 0
        
        for position in self.visible_positions:
            if position not in self.position_mapping:
                continue
                
            widgets = self.create_position_box(position, row, col)
            self.position_widgets[position] = widgets
            
            col += 1
            if col >= 2:
                col = 0
                row += 1
        
        self.content_frame.update_idletasks()
    
    def create_position_box(self, position, row, col):
        """Create a position box with capture/redo buttons."""
        box_frame = tk.Frame(
            self.content_frame,
            relief="solid",
            borderwidth=2,
            bg="#3a3a38",
            highlightbackground="#555555",
            highlightthickness=1,
            width=400,
            height=240
        )
        box_frame.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
        box_frame.grid_propagate(False)
        
        self.content_frame.grid_rowconfigure(row, weight=1)
        self.content_frame.grid_columnconfigure(col, weight=1)
        
        # Header
        header = tk.Frame(box_frame, bg="#0076b8", height=28)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text=f"Position {position}",
            font=("Arial", 9, "bold"),
            bg="#0076b8",
            fg="white"
        ).pack(expand=True, pady=5)
        
        # Preview area
        preview_frame = tk.Frame(box_frame, bg="#2a2a28", width=380, height=150)
        preview_frame.pack(fill="x", padx=8, pady=8)
        preview_frame.pack_propagate(False)
        
        preview_label = tk.Label(
            preview_frame,
            text="No screenshot",
            bg="#2a2a28",
            font=("Arial", 8),
            fg="#999999"
        )
        preview_label.pack(expand=True)
        
        # Button area
        button_area = tk.Frame(box_frame, bg="#3a3a38", height=40)
        button_area.pack(fill="x", padx=8, pady=(0, 8))
        button_area.pack_propagate(False)
        
        button_container = tk.Frame(button_area, bg="#3a3a38")
        button_container.pack(expand=True)
        
        # Capture button
        capture_button = tk.Button(
            button_container,
            text="Capture",
            command=lambda p=position: self.capture_screenshot(p),
            font=("Arial", 9, "bold"),
            bg="#28a616",
            fg="white",
            padx=18,
            pady=5,
            cursor="hand2",
            width=9,
            relief="raised",
            bd=1
        )
        capture_button.pack(side="left", padx=4)
        
        # Redo button
        redo_button = tk.Button(
            button_container,
            text="Redo",
            command=lambda p=position: self.capture_screenshot(p),
            font=("Arial", 9),
            bg="#f39c12",
            fg="white",
            padx=18,
            pady=5,
            cursor="hand2",
            width=9,
            state="disabled",
            relief="raised",
            bd=1
        )
        redo_button.pack(side="left", padx=4)
        
        return {
            'box_frame': box_frame,
            'preview_frame': preview_frame,
            'preview_label': preview_label,
            'capture_button': capture_button,
            'redo_button': redo_button
        }
    
    def add_more_positions(self):
        """Show the next set of positions (up to 4 more)."""
        all_positions = sorted(self.position_mapping.keys())
        current_max = max(self.visible_positions) if self.visible_positions else 0
        
        logging.info(f"Add More clicked - current max visible: {current_max}")
        
        next_positions = [p for p in all_positions if p > current_max]
        positions_to_add = next_positions[:4]
        
        if not positions_to_add:
            logging.warning("No more positions to add!")
            return
        
        for pos in positions_to_add:
            if pos not in self.visible_positions:
                self.visible_positions.append(pos)
        
        logging.info(f"Added positions: {positions_to_add}")
        logging.info(f"Now visible: {self.visible_positions}")
        
        self.update_position_display()
        
        # Restore screenshots
        for position, image in self.screenshots.items():
            if position in self.position_widgets:
                self.update_preview(position, image)
                self.position_widgets[position]['redo_button'].config(state="normal")
        
        # Recreate bottom bar
        self.setup_bottom_bar()
        
        logging.info(f"Add More complete. Visible: {self.visible_positions}")
    
    def capture_screenshot(self, position):
        """Start screenshot capture for a specific position."""
        logging.info(f"Starting screenshot capture for Position {position}")
        
        self.root.withdraw()
        
        # Bring GibbsCAM to foreground (and re-force right before starting capture)
        try:
            import win32gui
            import win32process
            import psutil
            
            def find_virtual_exe():
                virtual_windows = []
                def callback(hwnd, _):
                    if win32gui.IsWindowVisible(hwnd):
                        try:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            process = psutil.Process(pid)
                            if 'virtual.exe' in process.name().lower():
                                virtual_windows.append(hwnd)
                        except:
                            pass
                win32gui.EnumWindows(callback, None)
                return virtual_windows
            
            def delayed_start():
                try:
                    v2 = find_virtual_exe()
                    if v2:
                        hwnd2 = v2[0]
                        win32gui.ShowWindow(hwnd2, 9)
                        win32gui.SetForegroundWindow(hwnd2)
                        logging.info("✓ GibbsCAM re-forced to foreground before capture")
                except Exception as ee:
                    logging.warning(f"Could not re-force GibbsCAM before capture: {ee}")
                self._start_capture(position)

            virtual_windows = find_virtual_exe()
            if virtual_windows:
                hwnd = virtual_windows[0]
                win32gui.ShowWindow(hwnd, 9)
                win32gui.SetForegroundWindow(hwnd)
                logging.info("Brought GibbsCAM (virtual.exe) to foreground")
                # Slight delay so Excel can't steal focus; then re-force and capture
                self.root.after(300, delayed_start)
            else:
                logging.warning("No GibbsCAM (virtual.exe) window found")
                self._start_capture(position)
                
        except Exception as e:
            logging.warning(f"Could not bring GibbsCAM to foreground: {e}")
            self._start_capture(position)
    
    def _start_capture(self, position):
        """Internal method to start the actual capture."""
        capturer = ScreenshotCapture()
        capturer.start_capture(lambda img: self.on_screenshot_captured(position, img))
    
    def on_screenshot_captured(self, position, image):
        """Handle captured screenshot."""
        self.root.deiconify()
        
        if image is None:
            logging.info(f"Screenshot capture cancelled for Position {position}")
            return
        
        self.screenshots[position] = image
        logging.info(f"Screenshot captured for Position {position}")
        
        if position in self.position_widgets:
            self.update_preview(position, image)
            self.position_widgets[position]['redo_button'].config(state="normal")
        
        self.update_status()
    
    def update_preview(self, position, image):
        """Update the preview for a position."""
        if position not in self.position_widgets:
            return
            
        widgets = self.position_widgets[position]
        preview_label = widgets['preview_label']
        
        preview_width = 340
        preview_height = 130
        
        img_copy = image.copy()
        img_copy.thumbnail((preview_width, preview_height), Image.Resampling.LANCZOS)
        
        photo = ImageTk.PhotoImage(img_copy)
        
        preview_label.config(image=photo, text="")
        preview_label.image = photo
    
    def update_status(self):
        """Update the status label."""
        count = len(self.screenshots)
        total = len(self.position_mapping)
        
        if hasattr(self, 'status_label') and self.status_label.winfo_exists():
            self.status_label.config(text=f"{count}/{total} captured")
        
        if hasattr(self, 'done_button') and self.done_button.winfo_exists():
            if count > 0:
                self.done_button.config(state="normal")
            else:
                self.done_button.config(state="disabled")
    
    def finish_and_insert(self):
        """Insert all screenshots into Excel and close."""
        if not self.screenshots:
            messagebox.showwarning("No Screenshots", "Please capture at least one screenshot first.")
            return
        
        try:
            ws = self.workbook.Worksheets(self.worksheet_name)
            excel_app = self.workbook.Application
            
            inserted_count = 0
            for position, image in sorted(self.screenshots.items()):
                cell = self.position_mapping.get(position)
                if not cell:
                    logging.warning(f"No cell mapping for Position {position}")
                    continue
                
                temp_file = self.temp_dir / f"position_{position}.png"
                image.save(temp_file, 'PNG')
                self.temp_files.append(temp_file)
                
                logging.info(f"Inserting Position {position} into cell {cell}")
                
                try:
                    cell_range = ws.Range(cell)
                    
                    if cell_range.MergeCells:
                        merge_area = cell_range.MergeArea
                        left = merge_area.Left
                        top = merge_area.Top
                        cell_width = merge_area.Width
                        cell_height = merge_area.Height
                        logging.info(f"  Cell {cell} is merged: {merge_area.Address}")
                    else:
                        left = cell_range.Left
                        top = cell_range.Top
                        cell_width = cell_range.Width
                        cell_height = cell_range.Height
                    
                    picture = ws.Shapes.AddPicture(
                        Filename=str(temp_file.absolute()),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=-1,
                        Height=-1
                    )
                    
                    margin = 8
                    target_width = cell_width - margin
                    target_height = cell_height - margin
                    
                    img_width = picture.Width
                    img_height = picture.Height
                    aspect_ratio = img_width / img_height
                    
                    if target_width / target_height > aspect_ratio:
                        new_height = target_height
                        new_width = new_height * aspect_ratio
                    else:
                        new_width = target_width
                        new_height = new_width / aspect_ratio
                    
                    picture.Width = new_width
                    picture.Height = new_height
                    
                    picture.Left = left + (cell_width - new_width) / 2
                    picture.Top = top + (cell_height - new_height) / 2
                    
                    logging.info(f"✓ Inserted Position {position} at {cell} ({new_width:.0f}x{new_height:.0f}px)")
                    inserted_count += 1
                    
                except com_error as e:
                    logging.error(f"Error inserting image at {cell}: {e}")
            
            # MAXIMIZE Excel FIRST (while hidden), THEN make it visible
            try:
                import time
                import win32gui
                import win32con
                
                # Get Excel window handle
                hwnd = excel_app.Hwnd
                
                # MAXIMIZE while still hidden
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                time.sleep(0.1)
                
                # NOW make it visible
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                
                # Activate worksheet
                self.workbook.Activate()
                ws.Activate()
                
                # Bring to front
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetForegroundWindow(hwnd)
                
                logging.info("✓ Brought Excel to foreground and maximized")
                
            except Exception as e:
                logging.warning(f"Windows API maximize failed: {e}")
                try:
                    # Fallback to COM
                    excel_app.Visible = True
                    excel_app.ScreenUpdating = True
                    self.workbook.Windows(1).WindowState = -4137  # xlMaximized
                    logging.info("✓ Excel maximized via COM fallback")
                except Exception as e2:
                    logging.warning(f"COM maximize fallback also failed: {e2}")
                    
                except Exception as e:
                    logging.warning(f"Windows API maximize failed: {e}")
                    try:
                        # Fallback to COM
                        self.workbook.Windows(1).WindowState = -4137  # xlMaximized
                        excel_app.WindowState = -4137
                        logging.info("✓ Excel maximized via COM fallback")
                    except Exception as e2:
                        logging.warning(f"COM maximize fallback also failed: {e2}")
                        
            except Exception as e:
                logging.warning(f"Could not bring Excel to foreground: {e}")
            
            logging.info(f"Successfully inserted {inserted_count} screenshot(s) into Excel")
            
            self.cleanup()
            self.root.destroy()
            
        except Exception as e:
            logging.error(f"Error inserting screenshots: {e}")
            messagebox.showerror("Error", f"Failed to insert screenshots:\n{e}")
    
    def cancel(self):
        """Cancel and close without inserting."""
        if self.screenshots:
            result = messagebox.askyesno(
                "Cancel",
                "Discard all screenshots and close?",
                icon="warning"
            )
            if not result:
                return
        
        self.cleanup()
        self.root.destroy()
    
    def cleanup(self):
        """Clean up temporary files."""
        try:
            if self.temp_dir.exists():
                shutil.rmtree(self.temp_dir)
                logging.info("✓ Cleaned up temporary screenshot files")
        except Exception as e:
            logging.warning(f"Could not clean up temp files: {e}")
    
    def show(self):
        """Show the GUI and start main loop."""
        self.root.mainloop()


# ============================================================================
# NEW WORKFLOW: Capture screenshots FIRST, then create Excel
# ============================================================================

class ScreenshotGUIWithExcelCreation(ScreenshotGUI):
    """
    Extended GUI that creates Excel file when Done is clicked.
    Inherits from ScreenshotGUI but overrides finish_and_insert.
    """
    
    def __init__(self, csv_path, template_path, output_path, sheet_name):
        """Initialize with CSV and template info (no Excel workbook yet)."""
        if not PIL_AVAILABLE:
            raise ImportError("Pillow is required for screenshot functionality")
        
        if not WIN32COM_AVAILABLE:
            raise ImportError("pywin32 is required for Excel integration")
        
        # Store paths for later Excel creation
        self.csv_path = csv_path
        self.template_path = template_path
        self.output_path = output_path
        
        # No workbook yet!
        self.workbook = None
        self.worksheet_name = sheet_name
        self.screenshots = {}
        self.temp_files = []
        self.temp_dir = Path(tempfile.mkdtemp(prefix="gibbscam_screenshots_"))
        self.success = False
        
        # Load position mapping from config
        self.position_mapping = self._load_position_mapping()
        
        # Track visible positions
        max_pos = max(self.position_mapping.keys()) if self.position_mapping else 4
        initial_max = min(4, max_pos)
        self.visible_positions = list(range(1, initial_max + 1))
        self.position_widgets = {}
        
        logging.info(f"Initialization: max_pos={max_pos}, showing positions {self.visible_positions}")
        
        # Create main window
        self.root = tk.Tk()
        self.root.title("GibbsCAM Screenshot Tool")
        
        # Set window size and center it
        window_width = 880
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.minsize(860, 650)
        self.root.configure(bg="#30302e")
        
        self.setup_ui()

        # Align windows so GibbsCAM is behind the GUI at startup
        self.root.after(450, self._align_with_gibbscam_under_gui)
    
    def finish_and_insert(self):
        """
        Override: Create Excel with data, insert screenshots, then show maximized.
        This runs when user clicks Done button.
        """
        if not self.screenshots:
            messagebox.showwarning("No Screenshots", "Please capture at least one screenshot first.")
            return
        
        try:
            logging.info("Creating Excel file with data and screenshots...")
            
            # Import excel_mapper for the CSV mapping logic
            try:
                from . import excel_mapper
            except ImportError:
                import excel_mapper
            
            # Step 1: Create Excel with data (keep it hidden)
            logging.info("Step 1: Mapping CSV data to Excel...")
            result = excel_mapper.map_csv_to_excel(
                self.csv_path,
                self.template_path,
                self.output_path,
                self.worksheet_name,
                open_excel=False,  # Keep hidden
                enable_screenshots=False  # Don't open GUI again!
            )
            
            if not result:
                logging.error("Failed to create Excel with data")
                messagebox.showerror("Error", "Failed to create Excel file with coordinate data")
                return
            
            logging.info("✓ Excel created with coordinate data")
            
            # Step 2: Open the Excel file we just created
            logging.info("Step 2: Opening Excel to insert screenshots...")
            import win32com.client as win32
            
            excel_app = None
            try:
                excel_app = win32.GetObject(Class="Excel.Application")
            except:
                excel_app = win32.Dispatch("Excel.Application")
            
            excel_app.Visible = False  # Keep hidden
            excel_app.DisplayAlerts = False
            excel_app.ScreenUpdating = False
            
            wb = excel_app.Workbooks.Open(str(self.output_path.absolute()))
            ws = wb.Worksheets(self.worksheet_name)
            
            logging.info("✓ Excel opened for screenshot insertion")
            
            # Step 3: Insert screenshots
            logging.info("Step 3: Inserting screenshots...")
            inserted_count = 0
            for position, image in sorted(self.screenshots.items()):
                cell = self.position_mapping.get(position)
                if not cell:
                    logging.warning(f"No cell mapping for Position {position}")
                    continue
                
                temp_file = self.temp_dir / f"position_{position}.png"
                image.save(temp_file, 'PNG')
                self.temp_files.append(temp_file)
                
                logging.info(f"Inserting Position {position} into cell {cell}")
                
                try:
                    cell_range = ws.Range(cell)
                    
                    if cell_range.MergeCells:
                        merge_area = cell_range.MergeArea
                        left = merge_area.Left
                        top = merge_area.Top
                        cell_width = merge_area.Width
                        cell_height = merge_area.Height
                    else:
                        left = cell_range.Left
                        top = cell_range.Top
                        cell_width = cell_range.Width
                        cell_height = cell_range.Height
                    
                    picture = ws.Shapes.AddPicture(
                        Filename=str(temp_file.absolute()),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Left=left,
                        Top=top,
                        Width=-1,
                        Height=-1
                    )
                    
                    margin = 8
                    target_width = cell_width - margin
                    target_height = cell_height - margin
                    
                    img_width = picture.Width
                    img_height = picture.Height
                    aspect_ratio = img_width / img_height
                    
                    if target_width / target_height > aspect_ratio:
                        new_height = target_height
                        new_width = new_height * aspect_ratio
                    else:
                        new_width = target_width
                        new_height = new_width / aspect_ratio
                    
                    picture.Width = new_width
                    picture.Height = new_height
                    
                    picture.Left = left + (cell_width - new_width) / 2
                    picture.Top = top + (cell_height - new_height) / 2
                    
                    logging.info(f"✓ Inserted Position {position} at {cell}")
                    inserted_count += 1
                    
                except Exception as e:
                    logging.error(f"Error inserting image at {cell}: {e}")
            
            logging.info(f"✓ Inserted {inserted_count} screenshot(s)")
            
            # Step 4: Save the workbook
            logging.info("Step 4: Saving workbook...")
            wb.Save()
            logging.info("✓ Workbook saved")
            
            # Step 5: Show Excel MAXIMIZED
            logging.info("Step 5: Showing Excel maximized...")
            import time
            import win32gui
            import win32con
            
            try:
                # Get window handle
                hwnd = excel_app.Hwnd
                
                # Maximize while hidden
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                time.sleep(0.1)
                
                # NOW make visible
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                
                # Activate
                wb.Activate()
                ws.Activate()
                
                # Bring to front
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetForegroundWindow(hwnd)
                
                logging.info("✓ Excel shown maximized")
                
            except Exception as e:
                logging.warning(f"Could not maximize Excel: {e}")
                # Fallback
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                try:
                    wb.Windows(1).WindowState = -4137
                except:
                    pass
            
            # Success!
            self.success = True
            logging.info("✓ Complete! Excel created with data and screenshots")
            
            # Cleanup and close GUI
            self.cleanup()
            self.root.destroy()
            
        except Exception as e:
            logging.error(f"Error in finish_and_insert: {e}")
            logging.exception("Full traceback:")
            messagebox.showerror("Error", f"Failed to create Excel:\n{e}")
            self.success = False
    
    def cancel(self):
        """Override cancel to set success flag."""
        if self.screenshots:
            result = messagebox.askyesno(
                "Cancel",
                "Discard all screenshots and exit?",
                icon="warning"
            )
            if not result:
                return
        
        self.success = False
        self.cleanup()
        self.root.destroy()


def capture_then_create_excel(csv_path, template_path, output_path, sheet_name):
    """
    New workflow: Capture screenshots FIRST, then create Excel with data + screenshots.
    
    Args:
        csv_path: Path to CSV file with coordinate data
        template_path: Path to Excel template file
        output_path: Path for output Excel file
        sheet_name: Name of worksheet to update
        
    Returns:
        True if successful, False otherwise
    """
    if not PIL_AVAILABLE:
        logging.error("Screenshot GUI requires Pillow")
        return False
    
    if not WIN32COM_AVAILABLE:
        logging.error("Screenshot GUI requires pywin32")
        return False
    
    try:
        # Create a special GUI that will create Excel when Done is clicked
        gui = ScreenshotGUIWithExcelCreation(csv_path, template_path, output_path, sheet_name)
        gui.show()
        return gui.success  # True if completed successfully
    except Exception as e:
        logging.error(f"Error in capture_then_create_excel: {e}")
        logging.exception("Full traceback:")
        return False


def open_screenshot_gui(excel_workbook, worksheet_name: str) -> bool:
    """Open the screenshot GUI for capturing and inserting images (OLD workflow)."""
    if not PIL_AVAILABLE:
        logging.error("Screenshot GUI requires Pillow")
        logging.error("Install with: pip install pillow")
        return False
    
    if not WIN32COM_AVAILABLE:
        logging.error("Screenshot GUI requires pywin32")
        return False
    
    try:
        gui = ScreenshotGUI(excel_workbook, worksheet_name)
        gui.show()
        return True
    except Exception as e:
        logging.error(f"Error opening screenshot GUI: {e}")
        return False


def open_screenshot_gui(excel_workbook, worksheet_name: str) -> bool:
    """Open the screenshot GUI for capturing and inserting images."""
    if not PIL_AVAILABLE:
        logging.error("Screenshot GUI requires Pillow")
        logging.error("Install with: pip install pillow")
        return False
    
    if not WIN32COM_AVAILABLE:
        logging.error("Screenshot GUI requires pywin32")
        return False
    
    try:
        gui = ScreenshotGUI(excel_workbook, worksheet_name)
        gui.show()
        return True
    except Exception as e:
        logging.error(f"Error opening screenshot GUI: {e}")
        return False