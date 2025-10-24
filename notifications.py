"""
notifications.py
-----------------
Manages Windows toast notifications and error display GUI.
Uses windows-toasts for reliable Windows 10/11 notifications.
"""

import logging
from pathlib import Path
from typing import Optional

# Try to import windows-toasts (most reliable for Windows 10/11)
try:
    from windows_toasts import Toast, WindowsToaster, ToastDisplayImage
    WINDOWS_TOASTS_AVAILABLE = True
except ImportError:
    WINDOWS_TOASTS_AVAILABLE = False
    logging.debug("windows-toasts not available")

# Tkinter for GUI dialogs
try:
    import tkinter as tk
    from tkinter.scrolledtext import ScrolledText
    from tkinter import messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    tk = None
    ScrolledText = None
    messagebox = None
    TKINTER_AVAILABLE = False
    logging.debug("tkinter not available")


def show_toast(title, message, icon_path=None, duration=5, force_winotify=False):
    """
    Display a Windows toast notification.
    
    Args:
        title (str): Toast title
        message (str): Toast message body
        icon_path (Path or str): Optional icon file path
        duration (int): Duration in seconds (not used, kept for compatibility)
        force_winotify (bool): Ignored, kept for compatibility
    
    Returns:
        bool: True if toast was shown successfully
    """
    
    # Log icon info for debugging
    if icon_path:
        logging.info(f"Toast requested with icon: {icon_path}")
        icon_exists = Path(icon_path).exists() if icon_path else False
        logging.info(f"  Icon exists: {icon_exists}")
    else:
        logging.debug("Toast requested without icon")
    
    # Use windows-toasts (most reliable)
    if WINDOWS_TOASTS_AVAILABLE:
        try:
            toaster = WindowsToaster('GibbsCAM Processor')
            toast = Toast([title, message])
            
            # Add icon if provided and exists
            if icon_path:
                icon_path_obj = Path(icon_path)
                if icon_path_obj.exists():
                    try:
                        toast.AddImage(ToastDisplayImage.fromPath(str(icon_path_obj.absolute())))
                        logging.info(f"Icon added to toast: {icon_path_obj.absolute()}")
                    except Exception as e:
                        logging.warning(f"Could not add icon to toast: {e}")
            
            toaster.show_toast(toast)
            logging.info(f"✓ Toast sent: {title}")
            return True
            
        except Exception as e:
            logging.error(f"windows-toasts failed: {e}")
    else:
        logging.warning("windows-toasts not available - install with: pip install windows-toasts")
    
    # Fallback: console output
    _show_toast_console(title, message)
    return False


def _show_toast_console(title: str, message: str):
    """Fallback: Print formatted message to console."""
    separator = "=" * 70
    output = f"\n{separator}\n  {title}\n  {message}\n{separator}\n"
    print(output)
    logging.info(f"Toast displayed in console: {title}")


def show_error_gui(log_file: Path, title: str = "Error Log - GibbsCAM Processor"):
    """
    Display GUI window with error log contents.
    
    Args:
        log_file: Path to log file to display
        title: Window title
    """
    if not TKINTER_AVAILABLE:
        print(f"\nERROR LOG: See {log_file}")
        return

    log_file = Path(log_file)
    
    if not log_file.exists():
        logging.error(f"Log file not found: {log_file}")
        return

    try:
        root = tk.Tk()
        root.title(title)
        root.geometry("1000x650")

        # Create scrolled text widget with dark theme
        text = ScrolledText(
            root, 
            wrap="word", 
            width=120, 
            height=35,
            bg="#1e1e1e",
            fg="#d4d4d4",
            font=("Consolas", 9),
            insertbackground="white"
        )
        text.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure color tags
        text.tag_config("ERROR", foreground="#ff6b6b", font=("Consolas", 9, "bold"))
        text.tag_config("WARNING", foreground="#ffd93d")
        text.tag_config("INFO", foreground="#d4d4d4")

        # Load log file with syntax highlighting
        try:
            with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    if "[ERROR]" in line:
                        text.insert("end", line, "ERROR")
                    elif "[WARNING]" in line:
                        text.insert("end", line, "WARNING")
                    else:
                        text.insert("end", line, "INFO")
                        
            # Scroll to bottom
            text.see("end")
            
        except Exception as e:
            text.insert("end", f"Error loading log: {e}\n", "ERROR")

        # Make read-only
        text.config(state="disabled")

        # Button frame
        button_frame = tk.Frame(root, bg="#2d2d2d")
        button_frame.pack(fill="x", pady=10, padx=10)

        def open_log_file():
            """Open log in default text editor."""
            try:
                import os
                os.startfile(str(log_file))
            except Exception as e:
                logging.error(f"Failed to open log file: {e}")

        # Buttons
        open_btn = tk.Button(
            button_frame,
            text="Open Log File",
            command=open_log_file,
            padx=20,
            pady=8
        )
        open_btn.pack(side="right", padx=5)

        close_btn = tk.Button(
            button_frame,
            text="Close",
            command=root.destroy,
            padx=20,
            pady=8
        )
        close_btn.pack(side="right", padx=5)

        # Center window
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")

        root.mainloop()
        
    except Exception as e:
        logging.error(f"Error showing GUI: {e}")
        print(f"\nERROR LOG: See {log_file}")


def show_success_gui(log_file: Path, processed_count: int):
    """
    Display GUI window showing successful processing results.
    
    Args:
        log_file: Path to log file
        processed_count: Number of files processed
    """
    if not TKINTER_AVAILABLE:
        return

    try:
        root = tk.Tk()
        root.title("GibbsCAM Processor - Success")
        root.geometry("700x500")

        # Main frame with white background
        main_frame = tk.Frame(root, bg="white")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Success message
        success_msg = f"Processing Complete!\n\n{processed_count} file(s) processed successfully"
        success_label = tk.Label(
            main_frame,
            text=success_msg,
            font=("Arial", 14, "bold"),
            bg="white",
            fg="#28a745"
        )
        success_label.pack(pady=20)

        # Log preview section
        log_label = tk.Label(
            main_frame,
            text="Log Summary:",
            font=("Arial", 10, "bold"),
            bg="white"
        )
        log_label.pack(anchor="w", pady=(10, 5))

        # Log text area
        text = ScrolledText(
            main_frame,
            wrap="word",
            height=18,
            font=("Consolas", 9),
            bg="#f8f9fa"
        )
        text.pack(fill="both", expand=True)

        # Load log
        try:
            with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                text.insert("end", f.read())
            text.see("end")
        except Exception as e:
            text.insert("end", f"Could not load log: {e}")

        text.config(state="disabled")

        # Close button
        close_btn = tk.Button(
            main_frame,
            text="Close",
            command=root.destroy,
            padx=30,
            pady=10,
            font=("Arial", 10)
        )
        close_btn.pack(pady=15)

        # Center window
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")

        root.mainloop()
        
    except Exception as e:
        logging.error(f"Error showing success GUI: {e}")


def confirm_overwrite(filepath) -> bool:
    """
    Show confirmation dialog for overwriting existing file.
    
    Args:
        filepath: Path or filename that would be overwritten
        
    Returns:
        True if user confirms, False otherwise
    """
    if not TKINTER_AVAILABLE:
        # No GUI available, default to yes
        logging.warning("Cannot show overwrite prompt - tkinter not available")
        return True

    try:
        # Get just the filename for display
        filename = Path(filepath).name
        
        root = tk.Tk()
        root.withdraw()  # Hide main window
        root.attributes('-topmost', True)  # Bring to front
        root.update()
        
        result = messagebox.askyesno(
            "Overwrite File?",
            f"The file already exists:\n\n{filename}\n\nDo you want to overwrite it?",
            icon="warning",
            parent=root
        )
        
        root.destroy()
        logging.info(f"Overwrite prompt result: {'Yes' if result else 'No'}")
        return result
        
    except Exception as e:
        logging.warning(f"Error showing overwrite dialog: {e}")
        return True  # Default to yes on error
