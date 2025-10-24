"""
screenshot_capture.py
---------------------
Handles screenshot capture with fixed-size selection rectangle.
"""

import logging
import tkinter as tk

try:
    from PIL import Image, ImageTk, ImageGrab
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    logging.error("PIL/Pillow not available - install with: pip install pillow")


class ScreenshotCapture:
    """Handles screenshot capture with a fixed size selection."""
    
    def __init__(self, width_inches=3.0, height_inches=2.5, dpi=96):
        """
        Initialize screenshot capture with specific dimensions.
        
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