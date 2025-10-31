# PdfXtract - PDF Asset Extraction Utility
#
# Copyright (c) 2025 Jonathan Schoenberger
#
# Description:
# PdfXtract is a desktop application with a modern graphical user interface (GUI)
# for extracting various assets from PDF documents. Built with CustomTkinter, it
# provides a user-friendly and theme-aware experience.
#
# Key Features:
#   - Image Extraction: Extracts all embedded images from a PDF and saves them
#     as individual files (e.g., PNG, JPG).
#   - Text Extraction: Collates all text content from the PDF into a single,
#     UTF-8 encoded .txt file.
#   - OCR Support: Can perform Optical Character Recognition (OCR) to extract
#     text from scanned documents or images within the PDF.
#   - HTML Conversion: Converts the entire PDF document into a single .html file,
#     preserving a basic layout structure.
#
# User Interface & Experience:
#   - Intuitive Controls: Simple buttons for selecting a PDF and an output folder.
#   - Drag & Drop: Supports dragging and dropping a PDF file directly onto the
#     application window to select it.
#   - Password Handling: Automatically prompts for a password if the selected
#     PDF is encrypted.
#   - Real-time Feedback: A status log and progress bar provide real-time updates
#     during extraction tasks.
#   - Non-Blocking Operations: Long-running extraction tasks are executed in
#     background threads to keep the UI responsive.
#   - Theming: Includes a toggle for light and dark appearance modes.
#   - Post-Extraction Prompt: Asks to open the output folder upon success.
#
# Dependencies/Requirements:
# To run this script, you need to install the following Python libraries:
# pip install customtkinter PyMuPDF Pillow tkinterdnd2 easyocr
#
# Note on PyMuPDF: The 'fitz' import comes from the PyMuPDF library. If you
# encounter an import error, it may be due to a conflicting package.
# Run 'pip uninstall fitz' and 'pip uninstall PyMuPDF', then reinstall with
# 'pip install PyMuPDF'.

# License:
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import os

# Force Python to use UTF-8 encoding for all text I/O, which can prevent
# UnicodeEncodeError on Windows when libraries interact with the system.
os.environ['PYTHONUTF8'] = '1'

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import Toplevel, font as tkfont
import fitz # PyMuPDF
from PIL import Image
from tkinterdnd2 import DND_FILES, TkinterDnD # For drag and drop
import io
from datetime import datetime

import sys
import subprocess
import threading
from queue import Queue

# --- PyInstaller Hook for tkinterdnd2 ---
# When running as a PyInstaller bundle, the path to the tkdnd library
# needs to be explicitly provided to the Tcl interpreter. This code checks
# if the app is running in a bundled environment and, if so, sets the
# appropriate environment variable before the TkinterDnD root is created.
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # This is the path to the bundled tkdnd library within the executable
    tkdnd_path = os.path.join(sys._MEIPASS, 'tkdnd')
    os.environ['TKDND_LIBRARY'] = tkdnd_path


APP_VERSION = "1.0"
COPYRIGHT_YEAR = datetime.now().year

class PasswordDialog(ctk.CTkToplevel):
    """A custom, centered dialog for password input."""
    def __init__(self, parent, title="Password Required", text="Enter password:"):
        super().__init__(parent)
        self.parent = parent
        self.title(f"🔒 {title}")
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()

        self._password = None

        ctk.CTkLabel(self, text=text, wraplength=250).pack(padx=20, pady=(20, 10))
        self._entry = ctk.CTkEntry(self, show="*", width=250)
        self._entry.pack(padx=20, pady=0)
        self._entry.bind("<Return>", self._on_ok)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(padx=20, pady=20)

        ctk.CTkButton(button_frame, text="OK", command=self._on_ok).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Cancel", command=self._on_cancel).pack(side="left", padx=10)

        self.parent.center_window(self)
        self._entry.focus_set()

    def _on_ok(self, event=None):
        self._password = self._entry.get()
        self.destroy()

    def _on_cancel(self, event=None):
        self._password = None
        self.destroy()

    def get_password(self):
        """Waits for the dialog to close and returns the password."""
        # This makes the dialog modal
        self.wait_window()
        return self._password



class PdfXtract(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        
        self.title("PdfXtract")
        self.geometry("550x470")
        self.minsize(550, 470) # Prevent the window from being resized smaller

        # Hide the window initially to prevent the "flash" on startup
        self.withdraw()

        self.pdf_path = ""
        self.output_folder = ""
        
        # OCR reader instance, initialized on first use to save memory/startup time
        self.ocr_reader = None

        # A mapping of task names to their respective task functions.
        # This is more robust than string checking.
        self.task_map = {
            "Image Extraction": self._extract_images_task,
            "Text Extraction": self._extract_text_task,
            "HTML Extraction": self._extract_html_task,
        }

        # Header frame for top-right buttons
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(pady=(5, 0), padx=10, fill="x", anchor="n")

        # Spacer to push buttons to the right
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=1)

        # --- Custom Drawn Title ---
        # Replace the CTkLabel with a standard tkinter Canvas for drawing
        self.title_canvas = tk.Canvas(self.header_frame, width=250, height=50, highlightthickness=0)
        self.title_canvas.grid(row=0, column=0, padx=10, pady=0, sticky="w")
        # The drawing itself is handled by _draw_title() which is called during theme updates


        # --- Icon Buttons Frame ---
        # Place icon buttons in their own frame for stable alignment
        self.icon_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.icon_frame.grid(row=0, column=2, sticky="e")

        icon_font = ctk.CTkFont(size=24)

        # The theme button is now the only icon in the header
        self.theme_button = ctk.CTkLabel(self.icon_frame, text="💡", font=icon_font, cursor="hand2")
        self.theme_button.pack(pady=5) # Use pack for a single item
        self.theme_button.bind("<Button-1>", lambda e: self.toggle_theme())

        # Bind mouse events for custom hover effect on icon buttons
        self.theme_button.bind("<Enter>", lambda e: self.on_icon_button_enter(self.theme_button))
        self.theme_button.bind("<Leave>", lambda e: self.on_icon_button_leave(self.theme_button))

        # Frame for file selection
        self.selection_frame = ctk.CTkFrame(self)
        self.selection_frame.pack(pady=20, padx=20, fill="x")

        self.select_pdf_button = ctk.CTkButton(self.selection_frame, text="Select PDF", command=self.select_pdf)
        self.select_pdf_button.pack(pady=10, padx=10, side="left")
        # Register selection_frame as a drop target for files
        self.selection_frame.drop_target_register(DND_FILES)
        # Also register the new canvas and its parent frame to ensure dropping on the title area works
        self.header_frame.drop_target_register(DND_FILES)
        self.title_canvas.drop_target_register(DND_FILES)
        self.selection_frame.dnd_bind('<<Drop>>', self.handle_pdf_drop)

        self.pdf_label = ctk.CTkLabel(self.selection_frame, text="No PDF selected", wraplength=350)
        self.pdf_label.pack(pady=10, padx=10, side="left", fill="x", expand=True)

        # Frame for output selection
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.pack(pady=0, padx=20, fill="x")

        self.select_output_button = ctk.CTkButton(self.output_frame, text="Output Folder", command=self.select_output_folder)
        self.select_output_button.pack(pady=10, padx=10, side="left")

        self.output_label = ctk.CTkLabel(self.output_frame, text="No output folder selected", wraplength=350)
        self.output_label.pack(pady=10, padx=10, side="left", fill="x", expand=True)

        # Extraction buttons frame
        self.extraction_buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.extraction_buttons_frame.pack(pady=20, padx=20, fill="x")
        # Configure columns to be of equal weight to center the buttons and labels
        self.extraction_buttons_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # --- Extraction Buttons ---
        self.extract_images_button = ctk.CTkButton(self.extraction_buttons_frame, text="Extract Images", command=self.extract_images, state="disabled")
        self.extract_images_button.grid(row=0, column=0, pady=10, padx=10)

        self.extract_text_button = ctk.CTkButton(self.extraction_buttons_frame, text="Extract Text", command=self.extract_text, state="disabled")
        self.extract_text_button.grid(row=0, column=1, pady=10, padx=10)

        self.extract_html_button = ctk.CTkButton(self.extraction_buttons_frame, text="Extract as HTML", command=self.extract_html, state="disabled")
        self.extract_html_button.grid(row=0, column=2, pady=10, padx=10)

        # --- OCR Checkbox ---
        self.ocr_var = ctk.StringVar(value="off")
        self.ocr_checkbox = ctk.CTkCheckBox(self.extraction_buttons_frame, text="Use OCR (extract image-based text)", variable=self.ocr_var, onvalue="on", offvalue="off")
        self.ocr_checkbox.grid(row=1, column=0, columnspan=3, pady=(10, 0), padx=10, sticky="w")

        # Status textbox
        self.status_textbox = ctk.CTkTextbox(self, height=100)
        self.status_textbox.pack(pady=10, padx=20, fill="both", expand=True)

        # --- Footer Frame for Progress Bar and About Icon ---
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(pady=(0, 10), padx=(20, 10), fill="x") # Left padding 20, Right padding 10
        footer_frame.grid_columnconfigure(0, weight=1) # Progress bar expands
        footer_frame.grid_columnconfigure(1, weight=0) # Icon does not expand

        self.progress_bar = ctk.CTkProgressBar(footer_frame, orientation="horizontal", mode="determinate")
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 8)) # Add padding to its right
        self.progress_bar.set(0) # Start at 0

        # Place the about button to the right of the progress bar
        self.about_button = ctk.CTkLabel(footer_frame, text="ⓘ", font=icon_font, cursor="hand2")
        self.about_button.grid(row=0, column=1, pady=(0, 4)) # Nudge down slightly
        self.about_button.bind("<Button-1>", lambda e: self.show_about_dialog())

        # Bind hover effects for the about button in its new location
        self.about_button.bind("<Enter>", lambda e: self.on_icon_button_enter(self.about_button))
        self.about_button.bind("<Leave>", lambda e: self.on_icon_button_leave(self.about_button))

        # Set initial theme button color after all widgets are created
        # and explicitly set the background color to prevent a "flash" on startup.
        self._update_theme_and_backgrounds()

        # Defer UI updates until after the mainloop has started
        self.after(20, self.center_window) # Increased delay slightly for smoother init
        self.after(20, lambda: self.log("Welcome to PdfXtract!"))

    def select_pdf(self):
        """Opens a file dialog to select a PDF file."""
        path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if path:
            self.pdf_path = path
            self.pdf_label.configure(text=os.path.basename(path))
            self.log(f"Selected PDF: {os.path.basename(path)}")
            self.update_button_state()

    def select_output_folder(self):
        """Opens a directory dialog to select an output folder."""
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_folder = path
            self.output_label.configure(text=path)
            self.log(f"Selected output folder: {path}")
            self.update_button_state()

    def update_button_state(self):
        """Enables the extract button only if both paths are set."""
        new_state = "normal" if self.pdf_path and self.output_folder else "disabled"
        self._set_extraction_buttons_state(new_state)

    def log(self, message):
        """Adds a message to the status textbox."""
        self.status_textbox.insert("end", message + "\n")
        self.status_textbox.see("end") # Auto-scroll

    def toggle_theme(self):
        """Toggles between light and dark theme and updates icon color."""
        current_mode = ctk.get_appearance_mode()
        new_mode = "Light" if current_mode == "Dark" else "Dark"
        ctk.set_appearance_mode(new_mode)

        self._update_theme_and_backgrounds()

    def update_theme_button_color(self):
        """Sets the lightbulb icon color based on the current theme."""
        # Get the default text color for the current theme.
        text_color = self._get_default_icon_color()
        
        self.theme_button.configure(
            text_color=text_color
        )
        # Also update the about button to match
        self.about_button.configure(
            text_color=text_color
        )

    def _get_default_icon_color(self):
        """Returns the appropriate default icon color based on the theme."""
        current_mode = ctk.get_appearance_mode()
        # Use black in Light mode for visibility, otherwise use the theme's default.
        return "black" if current_mode == "Light" else ctk.ThemeManager.theme["CTkButton"]["text_color"]

    def on_icon_button_enter(self, widget):
        """Changes the icon color to the theme's accent color on hover."""
        # Use the button's hover color for the text on hover for a better effect
        hover_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTkButton"]["fg_color"])
        widget.configure(text_color=hover_color)

    def on_icon_button_leave(self, widget):
        """Resets the icon color to its default when the mouse leaves."""
        default_color = self._get_default_icon_color()
        widget.configure(text_color=default_color)

    def _draw_title(self):
        """Draws the stylized title on the canvas and handles theme colors."""
        # Use the main window's background color for a "transparent" effect.
        bg_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTk"]["fg_color"])
        text_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTkLabel"]["text_color"])
        accent_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTkButton"]["fg_color"])
        red_color = "#E53935" # A nice, modern red

        # Configure canvas background and clear any previous drawing
        self.title_canvas.configure(bg=bg_color)
        self.title_canvas.delete("all")

        # Define fonts
        pdf_font = ("Impact", 30)
        xtract_font = ("Arial", 30, "bold")

        # Draw "Pdf" in the theme's accent color
        self.title_canvas.create_text(
            38, 28, text="Pdf", font=pdf_font, fill=accent_color
        )

        # Draw the "X" in red
        self.title_canvas.create_text(
            88, 28, text="X", font=xtract_font, fill=red_color
        )
        # Draw the rest of "tract" in the default text color
        self.title_canvas.create_text(
            142, 28, text="tract", font=xtract_font, fill=text_color
        )

    def _update_theme_and_backgrounds(self):
        """Centralized method to update theme-dependent colors and backgrounds."""
        self.configure(fg_color=ctk.ThemeManager.theme["CTk"]["fg_color"])
        self._draw_title() # Redraw the title with new theme colors
        self.update_theme_button_color()
        self.header_frame.dnd_bind('<<Drop>>', self.handle_pdf_drop) # Re-bind drop after theme change
    
    def center_window(self, window=None): # Consolidated and moved to avoid duplicate definition
        """Centers the given window on the screen or over the parent."""
        target_window = window if window is not None else self
        target_window.update_idletasks()

        if window is None or window == self:
            # Center on screen
            # Use winfo_reqwidth/height for more reliable dimensions at startup
            win_width = self.winfo_reqwidth()
            win_height = self.winfo_reqheight()
            x = (self.winfo_screenwidth() - win_width) // 2
            y = (self.winfo_screenheight() - win_height) // 2
            # Make the window visible now that it's centered
            self.geometry(f"+{x}+{y}")
            self.deiconify()
        else:
            # Center over parent window
            parent_x, parent_y = self.winfo_x(), self.winfo_y()
            parent_w, parent_h = self.winfo_reqwidth(), self.winfo_reqheight() # Use req_width/height for reliability
            child_w, child_h = window.winfo_reqwidth(), window.winfo_reqheight() # Use req_width/height for reliability
            x = parent_x + (parent_w - child_w) // 2
            y = parent_y + (parent_h - child_h) // 2
            window.geometry(f"+{x}+{y}")

    def _draw_about_title(self, canvas, parent_window):
        """Draws a smaller, stylized title on the canvas for the About dialog."""
        # Get theme-appropriate colors from the parent window
        bg_color = parent_window._apply_appearance_mode(ctk.ThemeManager.theme["CTk"]["fg_color"])
        text_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTkLabel"]["text_color"])
        accent_color = self._apply_appearance_mode(ctk.ThemeManager.theme["CTkButton"]["fg_color"])
        red_color = "#E53935"

        # Configure canvas background for a "transparent" look
        canvas.configure(bg=bg_color)
        canvas.delete("all")

        # Define smaller fonts for the dialog
        pdf_font = ("Impact", 24)
        xtract_font = ("Arial", 24, "bold")

        # Draw "Pdf" - Coordinates adjusted for centering
        canvas.create_text(55, 22, text="Pdf", font=pdf_font, fill=accent_color)
        # Draw "X" - Coordinates adjusted for centering
        canvas.create_text(93, 22, text="X", font=xtract_font, fill=red_color)
        # Draw "tract" - Coordinates adjusted for centering
        canvas.create_text(137, 22, text="tract", font=xtract_font, fill=text_color)

    def show_about_dialog(self):
        """Displays the about dialog window."""
        about_win = ctk.CTkToplevel(self)
        about_win.title("About PdfXtract")
        about_win.resizable(False, False)
        about_win.transient(self) # Keep on top of the main window

        # Use a CTkFrame for consistent theming
        about_frame = ctk.CTkFrame(about_win, fg_color="transparent")
        about_frame.pack(expand=True, fill="both", pady=(15, 0), padx=20)

        # Create and draw the custom title
        about_title_canvas = tk.Canvas(about_frame, width=200, height=40, highlightthickness=0)
        about_title_canvas.pack(pady=(0, 5))
        self._draw_about_title(about_title_canvas, about_win)

        ctk.CTkLabel(about_frame, text=f"Version {APP_VERSION}").pack(pady=0)
        
        description = "A PDF Asset Extraction Utility with OCR Support."
        ctk.CTkLabel(about_frame, text=description, wraplength=380, justify="center").pack(pady=10)
        
        ctk.CTkLabel(about_frame, text=f"© {COPYRIGHT_YEAR} Jonathan Schoenberger", font=ctk.CTkFont(size=12)).pack(pady=(10, 0))
        ctk.CTkButton(about_frame, text="OK", command=about_win.destroy, width=100).pack(pady=(15,10))

        about_win.grab_set() # Modal
        about_win.update_idletasks() # Force update to get correct width after canvas drawing
        self.center_window(about_win)

    def handle_pdf_drop(self, event):
        """Handles PDF file drops onto the application window."""
        # event.data contains the dropped file path(s)
        # It can be a single path or multiple paths enclosed in {} and separated by spaces
        file_paths = self.tk.splitlist(event.data)

        if not file_paths:
            self.log("No file dropped or invalid drop data.")
            return

        # We'll only process the first dropped file for simplicity
        dropped_file_path = file_paths[0]

        if not os.path.isfile(dropped_file_path):
            self.log(f"Dropped item is not a file: {dropped_file_path}")
            messagebox.showerror("Error", "Dropped item is not a valid file.")
            return

        if dropped_file_path.lower().endswith('.pdf'):
            self.pdf_path = dropped_file_path
            self.pdf_label.configure(text=os.path.basename(dropped_file_path))
            self.log(f"Dropped PDF: {os.path.basename(dropped_file_path)}")
            self.update_button_state()
        else:
            self.log(f"Dropped file is not a PDF: {os.path.basename(dropped_file_path)}")
            messagebox.showerror("Error", "Please drop a PDF file.")

    def _set_extraction_buttons_state(self, state):
        """Helper to set the state of all extraction buttons."""
        self.extract_images_button.configure(state=state)
        self.extract_text_button.configure(state=state)
        self.extract_html_button.configure(state=state)

    def _execute_task(self, task_function, task_name, **kwargs):
        """
        A wrapper to execute an extraction task with consistent UI feedback and error handling.
        
        Args:
            task_function: The actual extraction logic to run.
            task_name: A user-friendly name for the task (e.g., "Image Extraction").
        """
        self.after(0, self.progress_bar.set, 0)
        self.log(f"\nStarting {task_name.lower()}...")

        # Check if this is an OCR task and set the progress bar to indeterminate
        is_ocr_task = (task_function == self._extract_text_task and self.ocr_var.get() == "on")
        if is_ocr_task:
            self.after(0, self.progress_bar.configure, {"mode": "indeterminate"})
            self.after(0, self.progress_bar.start)

        try:
            result_message = task_function(**kwargs) # Pass kwargs to task_function
            if result_message: # Only log if there's a message (task wasn't aborted for password)
                self.log(result_message)
                # Ask the user if they want to open the output folder
                prompt = f"{result_message}\n\nWould you like to open the output folder?"
                if messagebox.askyesno("Success", prompt):
                    self._open_folder(self.output_folder)

        # Catch specific, known errors from the extraction process first.
        except RuntimeError as e:
            error_message = f"An error occurred during {task_name.lower()}: {e}"
            self.log(error_message)
            messagebox.showerror("Error", error_message)
        except Exception as e: # Catch any other unexpected errors.
            error_message = f"An error occurred during {task_name.lower()}: {e}"
            self.log(error_message)
            messagebox.showerror("Error", error_message)
        finally:
            # Reset progress bar after a short delay to allow user to see it's finished
            # Ensure progress bar is back in determinate mode
            if is_ocr_task:
                # Schedule UI updates on the main thread
                self.after(0, self.progress_bar.stop)
                self.after(0, self.progress_bar.configure, {"mode": "determinate"})
            # Reset progress bar value after a short delay
            self.after(1000, self.progress_bar.set, 0)

    def _run_extraction_in_thread(self, target_function, *args, **kwargs): # Added **kwargs
        """
        Runs a given extraction function in a separate thread to prevent GUI freezing.
        
        Args:
            target_function: The wrapper method to run (e.g., self._execute_task).
            *args: Arguments to pass to the target function.
            **kwargs: Keyword arguments to pass to the target function.
        """
        self._set_extraction_buttons_state("disabled")
        def task_wrapper():
            target_function(*args, **kwargs)
            self.after(0, self._set_extraction_buttons_state, "normal")

        threading.Thread(target=task_wrapper, daemon=True).start()

    def _open_folder(self, path):
        """Opens the specified folder in the system's file explorer."""
        try:
            if os.name == 'nt':  # For Windows
                os.startfile(path)
            elif sys.platform == 'darwin':  # For macOS
                subprocess.Popen(['open', path])
            else:  # For Linux and other Unix-like OS
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")

    def _prompt_for_password_and_retry(self, task_function, task_name):
        """
        Prompts the user for a password and, if one is provided, re-runs the
        extraction task with the password.
        """
        dialog = PasswordDialog(self, text="This PDF is password protected.\nEnter the password to continue:")
        if (password := dialog.get_password()) is not None:
            self.log("Retrying with provided password...")
            # Re-run the original task, but this time passing the password
            self._run_extraction_in_thread(self._execute_task, task_function, task_name, password=password) # This call is now correct
        else:
            self.log("Password entry cancelled. Extraction aborted.")
            # Re-enable buttons since the task is aborted
            self._set_extraction_buttons_state("normal")

    def extract_images(self):
        """Core logic to extract images from the selected PDF."""
        if not self.pdf_path or not self.output_folder:
            messagebox.showerror("Error", "Please select a PDF and an output folder first.")
            return

        self._run_extraction_in_thread(self._execute_task, self._extract_images_task, "Image Extraction")

    def _extract_images_task(self, **kwargs): # Added **kwargs
        """Wrapper to handle password for image extraction."""
        return self._perform_extraction(self._do_extract_images, "Image Extraction", **kwargs) # Pass kwargs

    def _do_extract_images(self, doc):
        """
        The actual image extraction logic that runs in a background thread.
        
        Args:
            doc: An opened PyMuPDF document object.
        """
        image_count = 0
        for page_num in range(len(doc)):
            # Update progress bar every 5 pages or on the last page to reduce overhead
            if page_num % 5 == 0 or page_num == len(doc) - 1:
                progress_value = (page_num + 1) / len(doc)
                self.after(0, self.progress_bar.set, progress_value)

            page = doc.load_page(page_num)
            image_list = page.get_images(full=True)

            if not image_list:
                continue

            for img_index, img_info in enumerate(image_list):
                xref = img_info[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                ext = base_image["ext"]

                filename = f"page{page_num+1}_img{img_index+1}.{ext}"
                save_path = os.path.join(self.output_folder, filename)

                with open(save_path, "wb") as img_file:
                    img_file.write(image_bytes)
                image_count += 1

        return f"Extraction complete! Found and saved {image_count} images." if image_count > 0 else "Extraction finished, but no images were found in the PDF."

    def extract_text(self):
        """Core logic to extract text from the selected PDF."""
        if not self.pdf_path or not self.output_folder:
            messagebox.showerror("Error", "Please select a PDF and an output folder first.")
            return

        self._run_extraction_in_thread(self._execute_task, self._extract_text_task, "Text Extraction")

    def _extract_text_task(self, **kwargs): # Added **kwargs
        """Wrapper to handle password for text extraction."""
        # Decide whether to use standard extraction or OCR based on the checkbox.
        if self.ocr_var.get() == "on":
            self.log("OCR mode enabled. This may take a while...")
            return self._perform_extraction(self._do_extract_text_ocr, "Text Extraction", **kwargs)
        else:
            return self._perform_extraction(self._do_extract_text, "Text Extraction", **kwargs)

    def _do_extract_text(self, doc):
        """
        The actual text extraction logic that runs in a background thread.
        
        Args:
            doc: An opened PyMuPDF document object.
        """
        full_text = []
        for page_num in range(len(doc)):
            # Update progress bar every 5 pages or on the last page
            if page_num % 5 == 0 or page_num == len(doc) - 1:
                progress_value = (page_num + 1) / len(doc)
                self.after(0, self.progress_bar.set, progress_value)

            page = doc.load_page(page_num)
            text = page.get_text()
            if text:
                full_text.append(f"--- Page {page_num + 1} ---\n{text}\n\n")

        return self._save_extracted_text(full_text)

    def _initialize_ocr(self):
        """Initializes the EasyOCR reader if it hasn't been already."""
        if self.ocr_reader is None:
            try:
                import easyocr # Defer import until it's actually needed
                import contextlib # For redirecting stdout/stderr
                self.log("Initializing OCR engine (this may download model files on first run)...")

                # Suppress the specific UserWarning from torch about 'pin_memory'
                import warnings
                with warnings.catch_warnings():
                    warnings.filterwarnings("ignore", category=UserWarning, message=".*pin_memory.*")
                    # Also suppress stdout/stderr from easyocr/torch during model download
                    with open(os.devnull, 'w') as f, contextlib.redirect_stdout(f), contextlib.redirect_stderr(f):
                        # This will download the required language model on the first run
                        # Explicitly tell easyocr to use the GPU. It will fall back to CPU if not available.
                        self.ocr_reader = easyocr.Reader(['en'], gpu=True, verbose=False)

                self.log("OCR engine ready.")
                return True # Indicate success
            except ModuleNotFoundError:
                error_msg = "The 'easyocr' library is required for OCR functionality but is not installed.\n\nPlease install it by running:\npip install easyocr"
                self.log(f"Error: {error_msg}")
                messagebox.showerror("Dependency Missing", error_msg)
                return False # Indicate failure
        return True # Already initialized

    def _save_extracted_text(self, full_text):
        """Saves the collected text parts to a .txt file."""
        if full_text:
            pdf_base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            filename = f"{pdf_base_name}_extracted_text.txt"
            save_path = os.path.join(self.output_folder, filename)

            with open(save_path, "w", encoding="utf-8") as f:
                f.write("".join(full_text))

            return f"Text extraction complete! Saved to {filename}."
        else:
            return "Text extraction finished, but no text was found in the PDF."

    def _do_extract_text_ocr(self, doc):
        """
        Performs OCR to extract text from each page, which is treated as an image.
        
        Args:
            doc: An opened PyMuPDF document object.
        """

        self._initialize_ocr() # Ensure the OCR engine is ready

        # Ensure the OCR engine is ready. If not, abort the extraction.
        if not self._initialize_ocr():
            raise RuntimeError("OCR initialization failed. Please install the 'easyocr' library.")
        full_text = []
        for page_num in range(len(doc)):
            # For determinate mode, you would update the bar like this:
            # progress_value = (page_num + 1) / len(doc)
            # self.after(0, self.progress_bar.set, progress_value)
            # But for indeterminate, we just log progress.

            self.after(0, self.log, f"Processing page {page_num + 1}/{len(doc)} with OCR... (this can be slow)")

            page = doc.load_page(page_num) # Load the page
            pix = page.get_pixmap(dpi=300) # Render page to an image at 300 DPI for better accuracy
            img_bytes = pix.tobytes("png") # Get image data as bytes

            try:
                # easyocr can read image data directly from bytes
                ocr_results = self.ocr_reader.readtext(img_bytes, detail=0, paragraph=True)
                text = "\n".join(ocr_results)
            finally:
                # Explicitly free memory from the large image objects
                del pix
                del img_bytes

            if text.strip():
                full_text.append(f"--- Page {page_num + 1} ---\n{text}\n\n")
        return self._save_extracted_text(full_text)

    def extract_html(self):
        """Kicks off the HTML extraction in a background thread."""
        if not self.pdf_path or not self.output_folder:
            messagebox.showerror("Error", "Please select a PDF and an output folder first.")
            return
        
        # Run the actual task in a background thread
        self._run_extraction_in_thread(self._execute_task, self._extract_html_task, "HTML Extraction")

    def _extract_html_task(self, **kwargs): # Added **kwargs
        """Wrapper to handle password for HTML extraction."""
        return self._perform_extraction(self._do_extract_html, "HTML Extraction", **kwargs) # Pass kwargs

    def _do_extract_html(self, doc):
        """
        The actual HTML extraction logic that runs in a background thread.
        
        Args:
            doc: An opened PyMuPDF document object.
        """
        html_parts = ["<html><head><title>Extracted Content</title></head><body>"]
        for page_num in range(len(doc)):
            # Update progress bar every 5 pages or on the last page
            if page_num % 5 == 0 or page_num == len(doc) - 1:
                progress_value = (page_num + 1) / len(doc)
                self.after(0, self.progress_bar.set, progress_value)

            page = doc.load_page(page_num)
            html_content = page.get_text("html")
            html_parts.append(f"<!-- Page {page_num + 1} -->\n{html_content}")

        html_parts.append("</body></html>")

        pdf_base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        filename = f"{pdf_base_name}_extracted_content.html"
        save_path = os.path.join(self.output_folder, filename)

        with open(save_path, "w", encoding="utf-8") as f:
            f.write("\n".join(html_parts))

        return f"HTML extraction complete! Saved to {filename}."

    def _perform_extraction(self, extraction_logic, task_name, **kwargs):
        """
        A generic method to open a PDF and run the specified extraction logic.
        Handles password authentication and errors.
        
        Args:
            extraction_logic: The function to call with the opened document (e.g., _do_extract_images).
            task_name: The user-facing name of the task.
            **kwargs: Keyword arguments, expected to contain 'password'.
        """
        try:
            password = kwargs.get('password') # Safely get the password from kwargs
            with fitz.open(self.pdf_path) as doc:
                if doc.is_encrypted and doc.needs_pass:
                    if not password:
                        # This is the first attempt. The PDF is encrypted, but we have no password.
                        # Schedule the password prompt in the main thread and abort this one.
                        self.after(0, self._prompt_for_password_and_retry, self._get_task_function_by_name(task_name), task_name)
                        return None # Abort this thread; a new one will be started if a password is provided.
                    
                    # We have a password, so try to authenticate.
                    if not doc.authenticate(password):
                        raise RuntimeError("Incorrect password or unable to decrypt PDF.")
                
                # If we reach here, the document is open and authenticated.
                # Call the extraction logic with only the doc object.
                return extraction_logic(doc)

        except RuntimeError as e:
            # Catch other PyMuPDF errors like corrupted files and re-raise them.
            raise e

    def _get_task_function_by_name(self, task_name):
        """
        Helper to get the correct task wrapper function based on its name.
        Uses the task_map dictionary for a robust lookup.
        """
        task_function = self.task_map.get(task_name)
        if task_function is None:
            self.log(f"Error: No task function found for '{task_name}'.")
        return task_function

if __name__ == "__main__":
    # Create the DnD-aware root window first, and then hide it.
    # The main app will be a Toplevel window on top of this hidden root.
    # This is the most stable way to integrate tkinterdnd2 with customtkinter.
    root = TkinterDnD.Tk()
    root.withdraw()

    ctk.set_appearance_mode("System")  # Set theme before creating any CTk windows
    ctk.set_default_color_theme("blue")

    app = PdfXtract(master=root)

    # Ensure closing the app window also quits the hidden root
    def on_closing():
        root.quit()
    app.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()
