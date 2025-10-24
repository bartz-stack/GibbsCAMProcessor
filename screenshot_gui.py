"""
screenshot_gui.py
-----------------
Modern GUI for capturing screenshots and mapping them to Excel cells.

MODERN UI/UX VERSION - Modular Split
"""

import logging
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, List, Optional
import tempfile
import shutil

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    logging.error("PIL/Pillow not available")

try:
    import win32com.client as win32
    from pywintypes import com_error
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    logging.error("win32com not available")

# Import our modular components
try:
    from .screenshot_colors import ModernColors
    from .screenshot_capture import ScreenshotCapture
except ImportError:
    from screenshot_colors import ModernColors
    from screenshot_capture import ScreenshotCapture


# ============================================================================
# MAIN SCREENSHOT GUI
# ============================================================================
class ScreenshotGUI:
    """Modern professional GUI for managing multiple screenshots."""
    
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
        self._setup_window()
        self._setup_styles()
        self.setup_ui()

        # Align windows so GibbsCAM is behind the GUI at startup
        self.root.after(450, self._align_with_gibbscam_under_gui)
    
    def _setup_window(self):
        """Configure main window with modern dimensions."""
        self.root.title("GibbsCAM Screenshot Capture")
        
        # Modern window dimensions
        window_width = 920
        window_height = 760
        
        # Center on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.minsize(250, 250)
        self.root.configure(bg=ModernColors.BG_PRIMARY)
    
    def _setup_styles(self):
        """Configure ttk styles for modern look."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure scrollbar
        style.configure(
            "Modern.Vertical.TScrollbar",
            background=ModernColors.BG_SECONDARY,
            troughcolor=ModernColors.BG_PRIMARY,
            borderwidth=0,
            arrowsize=14
        )
        
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
                
                for key, value in config.CONFIG["SCREENSHOT_MAPPING"].items():
                    if key.upper().startswith("POSITION_"):
                        try:
                            pos_num = int(key.split("_")[1])
                            mapping[pos_num] = value
                            logging.info(f"✓ Loaded: Position {pos_num} -> {value}")
                        except (IndexError, ValueError) as e:
                            logging.warning(f"✗ Invalid key format: {key} - {e}")
            
            if mapping:
                logging.info(f"✓ Loaded {len(mapping)} screenshot position mappings from config")
                return mapping
            else:
                logging.warning("No screenshot mappings found in config, using defaults")
                
        except Exception as e:
            logging.error(f"Error loading screenshot mapping from config: {e}")
        
        logging.warning("Using hardcoded defaults (4 positions)")
        return {1: "G9", 2: "G31", 3: "A63", 4: "G63"}

    def _align_with_gibbscam_under_gui(self):
        """Bring GibbsCAM to foreground, then lift this Tk window above it."""
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
                win32gui.ShowWindow(hwnd, 9)
                win32gui.SetForegroundWindow(hwnd)
                logging.info("✓ GibbsCAM foregrounded (startup)")

            # Lift GUI above GibbsCAM
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(250, lambda: self.root.attributes("-topmost", False))
            logging.info("✓ GUI lifted above GibbsCAM")

        except Exception as e:
            logging.warning(f"Could not align GUI with GibbsCAM on startup: {e}")
        
    def setup_ui(self):
        """Setup the modern user interface."""
        # Header section
        self._create_header()
        
        # Main content area with custom scrollbar
        self._create_content_area()
        
        # Footer with action buttons
        self._create_footer()
    
    def _create_header(self):
        """Create modern header with instructions."""
        header_frame = tk.Frame(self.root, bg=ModernColors.BG_SECONDARY, height=65)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        # Title
        title_label = tk.Label(
            header_frame,
            text="📸 Screenshot Capture",
            font=("Segoe UI", 12, "bold"),
            bg=ModernColors.BG_SECONDARY,
            fg=ModernColors.TEXT_PRIMARY
        )
        title_label.pack(pady=(5, 5))
        
        # Instructions
        instructions = tk.Label(
            header_frame,
            text="Click 'Capture' → Position the red rectangle → Release to capture",
            font=("Segoe UI", 8),
            bg=ModernColors.BG_SECONDARY,
            fg=ModernColors.TEXT_SECONDARY
        )
        instructions.pack(pady=(0, 15))
    
    def _create_content_area(self):
        """Create scrollable content area for position cards."""
        content_container = tk.Frame(self.root, bg=ModernColors.BG_PRIMARY)
        content_container.pack(fill="both", expand=True, padx=15, pady=10)
        
        # Canvas with modern scrollbar
        self.canvas = tk.Canvas(
            content_container,
            bg=ModernColors.BG_PRIMARY,
            highlightthickness=0
        )
        
        scrollbar = ttk.Scrollbar(
            content_container,
            orient="vertical",
            command=self.canvas.yview,
            style="Modern.Vertical.TScrollbar"
        )
        
        self.content_frame = tk.Frame(self.canvas, bg=ModernColors.BG_PRIMARY)
        
        self.content_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Mouse wheel scrolling
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create position cards
        self.update_position_display()
    
    def _create_footer(self):
        """Create modern footer with action buttons."""
        # Remove existing if present
        if hasattr(self, 'footer_frame'):
            try:
                self.footer_frame.destroy()
            except:
                pass
        
        # Footer container
        self.footer_frame = tk.Frame(
            self.root,
            bg=ModernColors.BG_SECONDARY,
            height=80
        )
        self.footer_frame.pack(fill="x", side="bottom")
        self.footer_frame.pack_propagate(False)
        
        # Inner container for centering
        inner_frame = tk.Frame(self.footer_frame, bg=ModernColors.BG_SECONDARY)
        inner_frame.pack(expand=True, fill="both", padx=20, pady=15)
        
        # Left side - Add More button
        left_frame = tk.Frame(inner_frame, bg=ModernColors.BG_SECONDARY)
        left_frame.pack(side="left", fill="y")
        
        all_positions = sorted(self.position_mapping.keys())
        max_position = all_positions[-1] if all_positions else 0
        highest_visible = max(self.visible_positions) if self.visible_positions else 0
        
        if highest_visible < max_position:
            self.add_more_button = self._create_modern_button(
                left_frame,
                text="➕ Add More Positions",
                command=self.add_more_positions,
                bg=ModernColors.ACCENT_PRIMARY,
                fg=ModernColors.TEXT_PRIMARY
            )
            self.add_more_button.pack()
        
        # Center - Status indicator
        center_frame = tk.Frame(inner_frame, bg=ModernColors.BG_SECONDARY)
        center_frame.pack(side="left", expand=True, fill="both")
        
        self.status_label = tk.Label(
            center_frame,
            text=self._get_status_text(),
            font=("Segoe UI", 11, "bold"),
            bg=ModernColors.BG_SECONDARY,
            fg=ModernColors.TEXT_SECONDARY
        )
        self.status_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # Right side - Action buttons
        right_frame = tk.Frame(inner_frame, bg=ModernColors.BG_SECONDARY)
        right_frame.pack(side="right", fill="y")
        
        button_container = tk.Frame(right_frame, bg=ModernColors.BG_SECONDARY)
        button_container.pack()
        
        # Cancel button
        cancel_button = self._create_modern_button(
            button_container,
            text="✕ Cancel",
            command=self.cancel,
            bg=ModernColors.ACCENT_DANGER,
            fg=ModernColors.TEXT_PRIMARY
        )
        cancel_button.pack(side="left", padx=5)
        
        # Done button
        self.done_button = self._create_modern_button(
            button_container,
            text="✓ Done",
            command=self.finish_and_insert,
            bg=ModernColors.ACCENT_SUCCESS,
            fg=ModernColors.TEXT_PRIMARY,
            state="disabled" if len(self.screenshots) == 0 else "normal"
        )
        self.done_button.pack(side="left", padx=5)
    
    def _create_modern_button(self, parent, text, command, bg, fg, state="normal"):
        """Create a modern styled button with hover effects."""
        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Segoe UI", 10, "bold"),
            bg=bg,
            fg=fg,
            activebackground=self._lighten_color(bg),
            activeforeground=fg,
            padx=20,
            pady=10,
            cursor="hand2",
            relief="flat",
            borderwidth=0,
            state=state
        )
        
        # Hover effects
        def on_enter(e):
            if button['state'] == 'normal':
                button['bg'] = self._lighten_color(bg)
        
        def on_leave(e):
            button['bg'] = bg
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        
        return button
    
    def _lighten_color(self, hex_color, factor=1.2):
        """Lighten a hex color for hover effect."""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        rgb = tuple(min(255, int(c * factor)) for c in rgb)
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
    
    def _get_status_text(self):
        """Get current status text."""
        captured = len(self.screenshots)
        total = len(self.position_mapping)
        
        if captured == 0:
            return "No screenshots captured yet"
        elif captured == total:
            return f"✓ All {total} screenshots captured"
        else:
            return f"{captured} of {total} screenshots captured"
    
    def update_position_display(self):
        """Update the display of position cards."""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        self.position_widgets.clear()
        
        row = 0
        col = 0
        
        for position in self.visible_positions:
            if position not in self.position_mapping:
                continue
            
            widgets = self._create_position_card(position, row, col)
            self.position_widgets[position] = widgets
            
            col += 1
            if col >= 2:
                col = 0
                row += 1
        
        self.content_frame.update_idletasks()
    
    def _create_position_card(self, position, row, col):
        """Create a modern position card."""
        # Card container
        card_frame = tk.Frame(
            self.content_frame,
            bg=ModernColors.BG_SECONDARY,
            relief="flat",
            borderwidth=0
        )
        card_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # Configure grid weights
        self.content_frame.grid_rowconfigure(row, weight=1)
        self.content_frame.grid_columnconfigure(col, weight=1)
        
        # Inner content frame
        content = tk.Frame(card_frame, bg=ModernColors.BG_SECONDARY)
        content.pack(fill="both", expand=True, padx=2, pady=2)
        
        # Header with position number
        header = tk.Frame(content, bg=ModernColors.BG_TERTIARY, height=40)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        header_label = tk.Label(
            header,
            text=f"Position {position}",
            font=("Segoe UI", 11, "bold"),
            bg=ModernColors.BG_TERTIARY,
            fg=ModernColors.TEXT_PRIMARY
        )
        header_label.pack(side="left", padx=15, pady=10)
        
        # Status indicator dot
        status_dot = tk.Label(
            header,
            text="●",
            font=("Arial", 16),
            bg=ModernColors.BG_TERTIARY,
            fg=ModernColors.STATUS_EMPTY
        )
        status_dot.pack(side="right", padx=15)
        
        # Preview area
        preview_frame = tk.Frame(
            content,
            bg=ModernColors.BG_PRIMARY,
            width=420,
            height=180
        )
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)
        preview_frame.pack_propagate(False)
        
        preview_label = tk.Label(
            preview_frame,
            text="No screenshot",
            font=("Segoe UI", 10),
            bg=ModernColors.BG_PRIMARY,
            fg=ModernColors.TEXT_MUTED
        )
        preview_label.pack(expand=True)
        
        # Button area
        button_area = tk.Frame(content, bg=ModernColors.BG_SECONDARY, height=50)
        button_area.pack(fill="x", padx=10, pady=(0, 10))
        button_area.pack_propagate(False)
        
        button_container = tk.Frame(button_area, bg=ModernColors.BG_SECONDARY)
        button_container.pack(expand=True)
        
        # Capture button
        capture_button = self._create_modern_button(
            button_container,
            text="📷 Capture",
            command=lambda p=position: self.capture_screenshot(p),
            bg=ModernColors.ACCENT_SUCCESS,
            fg=ModernColors.TEXT_PRIMARY
        )
        capture_button.pack(side="left", padx=5)
        
        # Redo button
        redo_button = self._create_modern_button(
            button_container,
            text="🔄 Retake",
            command=lambda p=position: self.capture_screenshot(p),
            bg=ModernColors.ACCENT_WARNING,
            fg=ModernColors.TEXT_PRIMARY,
            state="disabled"
        )
        redo_button.pack(side="left", padx=5)
        
        return {
            'card_frame': card_frame,
            'preview_frame': preview_frame,
            'preview_label': preview_label,
            'capture_button': capture_button,
            'redo_button': redo_button,
            'status_dot': status_dot
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
        self.update_position_display()
        
        # Restore screenshots
        for position, image in self.screenshots.items():
            if position in self.position_widgets:
                self.update_preview(position, image)
                self.position_widgets[position]['redo_button'].config(state="normal")
                self.position_widgets[position]['status_dot'].config(
                    fg=ModernColors.STATUS_CAPTURED
                )
        
        # Recreate footer
        self._create_footer()
    
    def capture_screenshot(self, position):
        """Start screenshot capture for a specific position."""
        logging.info(f"Starting screenshot capture for Position {position}")
        
        self.root.withdraw()
        
        # Bring GibbsCAM to foreground
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
                self.root.after(300, delayed_start)
            else:
                logging.warning("No GibbsCAM (virtual.exe) window found")
                self._start_capture(position)
                
        except Exception as e:
            logging.warning(f"Could not bring GibbsCAM to foreground: {e}")
            self._start_capture(position)
    
    def _start_capture(self, position):
        """Internal method to start the actual capture."""
        # Load screenshot dimensions from config
        try:
            try:
                from . import config
            except ImportError:
                import config
            
            width_inches = float(config.get_value("SCREENSHOT", "WIDTH_INCHES", "3.0"))
            height_inches = float(config.get_value("SCREENSHOT", "HEIGHT_INCHES", "2.5"))
            dpi = config.get_int("SCREENSHOT", "DPI", 96)
            
            logging.info(f"Screenshot dimensions from config: {width_inches}\" x {height_inches}\" @ {dpi} DPI")
        except Exception as e:
            logging.warning(f"Could not load screenshot config, using defaults: {e}")
            width_inches = 3.0
            height_inches = 2.5
            dpi = 96
        
        capturer = ScreenshotCapture(width_inches, height_inches, dpi)
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
            self.position_widgets[position]['status_dot'].config(
                fg=ModernColors.STATUS_CAPTURED
            )
        
        self.update_status()
    
    def update_preview(self, position, image):
        """Update the preview for a position."""
        if position not in self.position_widgets:
            return
            
        widgets = self.position_widgets[position]
        preview_label = widgets['preview_label']
        
        preview_width = 400
        preview_height = 160
        
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
            self.status_label.config(text=self._get_status_text())
        
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
                
                hwnd = excel_app.Hwnd
                
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                time.sleep(0.1)
                
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                
                self.workbook.Activate()
                ws.Activate()
                
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetForegroundWindow(hwnd)
                
                logging.info("✓ Brought Excel to foreground and maximized")
                
            except Exception as e:
                logging.warning(f"Windows API maximize failed: {e}")
                try:
                    excel_app.Visible = True
                    excel_app.ScreenUpdating = True
                    self.workbook.Windows(1).WindowState = -4137
                    logging.info("✓ Excel maximized via COM fallback")
                except Exception as e2:
                    logging.warning(f"COM maximize fallback also failed: {e2}")
            
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
# EXCEL CREATION WORKFLOW
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
        self._setup_window()
        self._setup_styles()
        self.setup_ui()
        self.root.after(450, self._align_with_gibbscam_under_gui)
    
    def finish_and_insert(self):
        """
        Override: Create Excel with data, insert screenshots, then show maximized.
        This runs when user clicks Done button.
        """
        if not self.screenshots:
            messagebox.showwarning("No Screenshots", "Please capture at least one screenshot first.")
            return
        
        logging.info("Creating Excel file with data and screenshots...")
        
        try:
            # Check if file is already open with unsaved changes
            import win32com.client as win32
            import time
            import gc
            
            try:
                excel_app_check = win32.GetObject(Class="Excel.Application")
                
                for i in range(1, excel_app_check.Workbooks.Count + 1):
                    try:
                        wb_check = excel_app_check.Workbooks(i)
                        if Path(wb_check.FullName).resolve() == self.output_path.resolve():
                            logging.warning(f"File is already open: {self.output_path.name}")
                            
                            if not wb_check.Saved:
                                result = messagebox.askyesnocancel(
                                    "Unsaved Changes",
                                    f"The file '{self.output_path.name}' is already open with unsaved changes.\n\n"
                                    "Yes - Save and continue\n"
                                    "No - Discard changes and continue\n"
                                    "Cancel - Stop",
                                    icon="warning"
                                )
                                
                                if result is None:
                                    logging.info("User cancelled due to unsaved changes")
                                    return
                                elif result:
                                    logging.info("Saving existing workbook...")
                                    wb_check.Save()
                                    wb_check.Close(SaveChanges=False)
                                else:
                                    logging.info("Discarding changes...")
                                    wb_check.Close(SaveChanges=False)
                            else:
                                logging.info("Closing existing workbook (no unsaved changes)")
                                wb_check.Close(SaveChanges=False)
                            break
                    except:
                        pass
            except:
                pass
            
            # Import excel_mapper
            try:
                from . import excel_mapper
            except ImportError:
                import excel_mapper
            
            # Step 1: Create Excel with data
            logging.info("Step 1: Mapping CSV data to Excel...")
            result = excel_mapper.map_csv_to_excel(
                self.csv_path,
                self.template_path,
                self.output_path,
                self.worksheet_name,
                open_excel=False,
                enable_screenshots=False
            )
            
            if not result:
                logging.error("Failed to create Excel with data")
                messagebox.showerror("Error", "Failed to create Excel file with coordinate data")
                return
            
            logging.info("✓ Excel created with coordinate data")
            
            # Step 2: Open the Excel file
            logging.info("Step 2: Opening Excel to insert screenshots...")
            
            # Force garbage collection and wait
            gc.collect()
            time.sleep(0.3)
            
            # Get existing Excel instance or create new one
            try:
                excel_app = win32.GetObject(Class="Excel.Application")
                logging.info("Connected to existing Excel instance")
            except:
                excel_app = win32.Dispatch("Excel.Application")
                logging.info("Created new Excel instance")
            
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.ScreenUpdating = False
            
            # Open the workbook
            logging.info(f"Opening: {self.output_path}")
            wb = excel_app.Workbooks.Open(str(self.output_path.absolute()), ReadOnly=False)
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
            import win32gui
            import win32con
            
            try:
                hwnd = excel_app.Hwnd
                
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                time.sleep(0.1)
                
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                
                wb.Activate()
                ws.Activate()
                
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                     win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                win32gui.SetForegroundWindow(hwnd)
                
                logging.info("✓ Excel shown maximized")
                
            except Exception as e:
                logging.warning(f"Could not maximize Excel: {e}")
                excel_app.Visible = True
                excel_app.ScreenUpdating = True
                try:
                    wb.Windows(1).WindowState = -4137
                except:
                    pass
            
            self.success = True
            logging.info("✓ Complete! Excel created with data and screenshots")
            
            self.cleanup()
            self.root.destroy()
            
        except Exception as e:
            logging.error(f"Error in finish_and_insert: {e}")
            logging.exception("Full traceback:")
            messagebox.showerror("Error", f"Failed to create Excel:\n{e}")
            self.success = False
            self.cleanup()
            self.root.destroy()


# ============================================================================
# PUBLIC API FUNCTIONS
# ============================================================================
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
        gui = ScreenshotGUIWithExcelCreation(csv_path, template_path, output_path, sheet_name)
        gui.show()
        return gui.success
    except Exception as e:
        logging.error(f"Error in capture_then_create_excel: {e}")
        logging.exception("Full traceback:")
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