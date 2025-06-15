#!/usr/bin/env python
import os
import re
import time
import random
import threading
import queue
import webbrowser
import requests
import pyautogui
import openpyxl
import gc
from tkinter import filedialog, messagebox, colorchooser, Label
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageOps
from io import BytesIO
import win32clipboard
import customtkinter as ctk
import tkinter as tk
from datetime import datetime, timedelta
import csv
import struct
import logging

# Set logging level for av module warnings
logging.getLogger('libav').setLevel(logging.ERROR)

# ----------------------- GLOBAL CONSTANTS & PATHS -----------------------
CURRENT_VERSION = "15"
GITHUB_API_URL = "https://api.github.com/repos/Parth-Sancheti-5/WabulkXpress/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/Parth-Sancheti-5/WabulkXpress/"
FLAG_FILE = "first_run.flag"
BIN_FOLDER = os.path.join(os.getcwd(), "bin")
TITLE_ICON_PATH = os.path.join(BIN_FOLDER, "loco.ico")
LOGO_PATH = os.path.join(BIN_FOLDER, "Logo.png")
WHATSAPP_BETA = os.path.join(BIN_FOLDER, "WhatsApp_Beta.lnk")
OUTPUT_IMG_FOLDER = os.path.join(os.getcwd(), "output_img")
DEFAULT_MIN_DELAY = 1
DEFAULT_MAX_DELAY = 10
VIDEO_PATH = os.path.join(BIN_FOLDER, "woi.mp4")
LOADING_GIF_PATH = os.path.join(BIN_FOLDER, "lod.gif")
if not os.path.exists(OUTPUT_IMG_FOLDER):
    os.makedirs(OUTPUT_IMG_FOLDER)

def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")

# ----------------------- Utility functions -----------------------
def copy_text_to_clipboard(text):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Clipboard error: {e}")

def copy_file_to_clipboard(file_path):
    """
    Copies a file to the clipboard as a file drop (CF_HDROP),
    ensuring that the actual file (not the path) is pasted.
    """
    DROPFILES_FORMAT = "IiiIII"
    DROPFILES_SIZE = struct.calcsize(DROPFILES_FORMAT)
    offset = DROPFILES_SIZE
    file_list = file_path + "\0\0"
    dropfiles_struct = struct.pack(DROPFILES_FORMAT, DROPFILES_SIZE, 0, 0, 0, offset, 1)
    data = dropfiles_struct + file_list.encode("utf-16-le")
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, data)
        win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Error copying file to clipboard: {e}")

# ----------------------- Custom Tooltip (HoverHint) Class -----------------------
class HoverHint(ctk.CTkToplevel):
    def __init__(self, widget, hint_text, image_path, *args, **kwargs):
        super().__init__(widget.master, *args, **kwargs)
        self.overrideredirect(True)
        self.geometry("300x120")  # Fixed size box (can tweak this)
        self.configure(bg="transparent")

        self.frame = ctk.CTkFrame(self, corner_radius=12)
        self.frame.pack(expand=True, fill="both", padx=5, pady=5)

        self.text_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        self.text_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)

        self.hint_label = ctk.CTkLabel(self.text_frame, text=hint_text, anchor="w", justify="left", wraplength=150)
        self.hint_label.pack(expand=True, fill="both")

        try:
            img = Image.open(image_path)
            box_width = 300  # Same as Toplevel width
            image_width = int(box_width * 0.4)  # 40% of box width
            img = img.resize((image_width, image_width), Image.Resampling.LANCZOS)
            self.hint_image = ctk.CTkImage(light_image=img, size=(image_width, image_width))
        except Exception as e:
            print(f"Image load error: {e}")
            self.hint_image = None

        self.image_label = ctk.CTkLabel(self.frame, image=self.hint_image, text="")
        self.image_label.pack(side="right", padx=10, pady=10)

        self.withdraw()
        self.widget = widget
        self.widget.bind("<Enter>", self.show_hint)
        self.widget.bind("<Leave>", self.hide_hint)
        self.widget.bind("<Motion>", self.move_hint)

    def show_hint(self, event=None):
        self.deiconify()
        self.lift()
        self.move_hint(event)

    def hide_hint(self, event=None):
        self.withdraw()

    def move_hint(self, event=None):
        if event:
            x = event.x_root + 10
            y = event.y_root + 10
            self.geometry(f"+{x}+{y}")

# ----------------------- Animated Button Class -----------------------
class AnimatedCTkButton(ctk.CTkButton):
    def __init__(self, *args, hover_fg_color="#0050a0", **kwargs):
        super().__init__(*args, **kwargs)
        self.hover_fg_color = hover_fg_color
        self.original_fg_color = self.cget("fg_color")
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
    def on_enter(self, event):
        self.configure(fg_color=self.hover_fg_color)
    def on_leave(self, event):
        self.configure(fg_color=self.original_fg_color)

# ----------------------- TkinterVideo Class -----------------------
class TkinterVideo(tk.Label):
    def __init__(self, master, path, scaled=True, keep_aspect=False, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.path = path
        self.scaled = scaled
        self.keep_aspect = keep_aspect
        self._stop = False
        self.frame_queue = queue.Queue()
        self.current_frame = None
        self._load_thread = threading.Thread(target=self._decode_video, daemon=True)
        self._load_thread.start()
        self.after(0, self._update_image)

    def _decode_video(self):
        try:
            import av
        except ImportError:
            print("av module not installed. Please run: pip install av")
            return
        try:
            container = av.open(self.path)
            stream = container.streams.video[0]
            stream.thread_type = "AUTO"
            delay = 1 / float(stream.average_rate)
            for frame in container.decode(stream):
                if self._stop:
                    break
                img = frame.to_image()
                self.frame_queue.put(img)
                time.sleep(delay)
            container.close()
        except Exception as e:
            print(f"Error in TkinterVideo: {e}")

    def _update_image(self):
        try:
            if not self.frame_queue.empty():
                img = self.frame_queue.get_nowait()
                if self.scaled:
                    w = self.winfo_width()
                    h = self.winfo_height()
                    if w and h:
                        if self.keep_aspect:
                            img = ImageOps.contain(img, (w, h))
                        else:
                            img = img.resize((w, h), Image.ANTIALIAS)
                self.current_frame = ImageTk.PhotoImage(img)
                self.configure(image=self.current_frame)
        except Exception as e:
            print(f"Error updating image: {e}")
        if not self._stop:
            self.after(30, self._update_image)

    def stop(self):
        self._stop = True

    def pause(self):
        self._stop = True

    def play(self):
        if self._stop:
            self._stop = False
            self._load_thread = threading.Thread(target=self._decode_video, daemon=True)
            self._load_thread.start()
            self.after(0, self._update_image)

# ----------------------- ProgressPopup Class -----------------------
class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, title, total):
        super().__init__(parent)
        self.transient(parent)  # This makes the window always on top of its parent
        self.title(title)
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("500x300")
        self.resizable(False, False)
        self.lift()  # Bring the window to the top
        self.attributes("-topmost", True)  # Ensure it stays on top
        self.configure(fg_color="white")
        self.total = total
        self.current = 0

        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(expand=True, fill="both", padx=20, pady=20)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        self.gif_label = ctk.CTkLabel(frame, text="")
        self.gif_label.grid(row=0, column=0, padx=10, pady=(30,10))
        self.load_gif(LOADING_GIF_PATH, size=(150, 150))
        
        self.progress_label = ctk.CTkLabel(frame, text=f"{self.current}/{self.total}", font=("Arial", 28, "bold"))
        self.progress_label.grid(row=1, column=0, padx=10, pady=(10,30))
        center_window(self)

    def load_gif(self, path, size=(150, 150)):
        try:
            image = Image.open(path)
            image = image.resize(size)
            self.gif_image = ImageTk.PhotoImage(image)
            self.gif_label.configure(image=self.gif_image)
        except Exception as e:
            print("Error loading GIF:", e)

    def update_progress(self, current):
        self.current = current
        if self.winfo_exists():
            self.after(0, lambda: self.progress_label.configure(text=f"{self.current}/{self.total}"))
            self.update_idletasks()

    def close(self):
        if self.winfo_exists():
            self.destroy()

# ----------------------- AnimatedGIF Class -----------------------
class AnimatedGIF(tk.Label):
    def __init__(self, master, filename, delay=100):
        self.master = master
        self.filename = filename
        self.delay = delay
        im = Image.open(filename)
        self.frames = []
        try:
            for i in range(1000):
                im.seek(i)
                frame = ImageTk.PhotoImage(im.copy())
                self.frames.append(frame)
        except EOFError:
            pass
        self.idx = 0
        super().__init__(master, image=self.frames[0])
        self.after(self.delay, self.play)
    def play(self):
        self.idx = (self.idx + 1) % len(self.frames)
        self.configure(image=self.frames[self.idx])
        self.after(self.delay, self.play)

# ----------------------- FIRST RUN POPUP -----------------------
class FirstRunPopup(ctk.CTkToplevel):
    def __init__(self, master, on_close_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Welcome!")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("700x400")
        self.resizable(False, False)
        self.on_close_callback = on_close_callback
        self.attributes("-topmost", True)
        
        self.container = ctk.CTkFrame(self, width=700, height=400)
        self.container.pack(fill="both", expand=True)
        
        if os.path.exists(VIDEO_PATH):
            self.video_player = TkinterVideo(self.container, VIDEO_PATH, scaled=True, keep_aspect=True)
            self.video_player.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            self.instruction_label = ctk.CTkLabel(self.container, text="Video file not found", font=("Arial", 16))
            self.instruction_label.place(relx=0.5, rely=0.5, anchor="center")
        
        self.bottom_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        self.bottom_frame.place(relx=0.5, rely=0.9, anchor="center")
        
        self.dont_show_var = ctk.BooleanVar(value=True)
        self.checkbox = ctk.CTkCheckBox(self.bottom_frame, text="Don't show this again", variable=self.dont_show_var)
        self.checkbox.grid(row=0, column=0, padx=10)
        self.ok_button = AnimatedCTkButton(self.bottom_frame, text="OK", fg_color="#0078D7", corner_radius=10, command=self.close_popup)
        self.ok_button.grid(row=0, column=1, padx=10)
        
        center_window(self)
        self.fade_in()

    def fade_in(self, alpha=0.0):
        if alpha < 1.0:
            self.attributes("-alpha", alpha)
            self.after(50, lambda: self.fade_in(alpha + 0.1))
        else:
            self.attributes("-alpha", 1.0)
        
    def close_popup(self):
        if self.dont_show_var.get():
            with open(FLAG_FILE, "w") as f:
                f.write("shown")
        self.on_close_callback()
        self.destroy()

# ----------------------- EXCEL TABLE -----------------------
class ExcelTable(ctk.CTkScrollableFrame):
    def __init__(self, master, main_app, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.main_app = main_app
        self.configure(corner_radius=15, scrollbar_fg_color="gray")
        self.rows = []
        self.add_header()
        self.prepopulate_rows(1)
    def add_header(self):
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=5, pady=3)
        ctk.CTkLabel(header_frame, text="S.No.", width=40, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Phone Number", width=200, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Name", width=200, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Status", width=60, anchor="center").pack(side="left", padx=5)
    def prepopulate_rows(self, count):
        for _ in range(count):
            self.add_row()
    def add_row(self):
        row_frame = ctk.CTkFrame(self, corner_radius=10)
        row_frame.pack(fill="x", padx=5, pady=3)
        sno_label = ctk.CTkLabel(row_frame, text=str(len(self.rows)+1), width=40, anchor="center")
        sno_label.pack(side="left", padx=5)
        phone_var = ctk.StringVar()
        phone_entry = ctk.CTkEntry(row_frame, textvariable=phone_var, placeholder_text="Enter number", width=200, corner_radius=10)
        phone_entry.pack(side="left", padx=5)
        phone_entry.bind("<Return>", lambda event, widget=phone_entry, var=phone_var: self.validate_phone(widget, var))
        phone_entry.bind("<KeyRelease>", self.check_add_row)
        name_var = ctk.StringVar()
        name_entry = ctk.CTkEntry(row_frame, textvariable=name_var, placeholder_text="Enter name", width=200, corner_radius=10)
        name_entry.pack(side="left", padx=5)
        name_entry.bind("<Return>", lambda event, var=name_var: self.validate_name(var))
        name_entry.bind("<KeyRelease>", self.check_add_row)
        bg_color = "#2B2B2B" if ctk.get_appearance_mode().lower() == "dark" else "#F0F0F0"
        indicator = tk.Canvas(row_frame, width=30, height=30, highlightthickness=0, bg=bg_color)
        indicator.create_oval(2, 2, 28, 28, fill="green", outline="")
        indicator.create_text(15, 15, text="✔", fill="white", font=("Arial", 12, "bold"))
        indicator.pack(side="left", padx=5)
        row_dict = {"sno": sno_label, "phone": phone_entry, "phone_var": phone_var,
                    "name": name_entry, "name_var": name_var, "indicator": indicator,
                    "indicator_state": 0, "row_frame": row_frame}
        indicator.bind("<Button-1>", lambda e, r=row_dict: self.toggle_indicator(r))
        self.rows.append(row_dict)
    def update_row_numbers(self):
        for idx, row in enumerate(self.rows, start=1):
            row["sno"].configure(text=str(idx))
    def toggle_indicator(self, row_dict):
        state = row_dict.get("indicator_state", 0)
        indicator = row_dict.get("indicator")
        if state == 0:
            indicator.delete("all")
            indicator.create_oval(2, 2, 28, 28, fill="red", outline="")
            indicator.create_text(15, 15, text="✖", fill="white", font=("Arial", 12, "bold"))
            row_dict["indicator_state"] = 1
            row_dict["skip"] = True
        elif state == 1:
            row_dict["row_frame"].destroy()
            self.rows.remove(row_dict)
            if not self.rows:
                self.add_row()
            self.update_row_numbers()
    def validate_phone(self, widget, var):
        import re

        text = var.get().strip()
        if not text:
            return

        original_text = text
        default_country = self.main_app.country_code_var.get()
        allowed_codes = ["+91", "+1", "+44", "+61", "+81", "+49", "+33", "+86", "+7"]

        text = text.replace(" ", "").replace("-", "")

        if text.startswith("+"):
            clean = re.sub(r"[^\d+]", "", text)
            digits_only = re.sub(r"\D", "", clean)
            if len(digits_only) < 10:
                self.main_app.log_live(f"⚠️ Invalid phone number detected: {original_text}")
            var.set(clean)
            widget.delete(0, "end")
            widget.insert(0, clean)
            return

        text = re.sub(r"^[^\d]+", "", text)

        matched = False
        for code in allowed_codes:
            code_digits = code.replace("+", "")
            if text.startswith(code_digits):
                text = "+" + text
                matched = True
                break

        if not matched:
            if default_country != "None":
                text = default_country + text

        final = re.sub(r"[^\d+]", "", text)
        digits_only = re.sub(r"\D", "", final)

        if len(digits_only) < 10 or not digits_only.isdigit():
            self.main_app.log_live(f"⚠️ Invalid phone number detected: {original_text}")

        var.set(final)
        widget.delete(0, "end")
        widget.insert(0, final)

    def validate_name(self, var):
        var.set(var.get().strip())
    def check_add_row(self, event):
        last_row = self.rows[-1]
        if last_row["phone_var"].get().strip() or last_row["name_var"].get().strip():
            self.add_row()
    def load_data(self, data):
        for child in self.winfo_children():
            child.destroy()
        self.rows = []
        self.add_header()
        for idx, entry in enumerate(data, start=1):
            row_frame = ctk.CTkFrame(self, corner_radius=10)
            row_frame.pack(fill="x", padx=5, pady=3)
            sno_label = ctk.CTkLabel(row_frame, text=str(idx), width=40, anchor="center")
            sno_label.pack(side="left", padx=5)
            phone_var = ctk.StringVar(value=entry.get("phone", ""))
            phone_entry = ctk.CTkEntry(row_frame, textvariable=phone_var, width=200, corner_radius=10)
            phone_entry.pack(side="left", padx=5)
            phone_entry.bind("<Return>", lambda event, widget=phone_entry, var=phone_var: self.validate_phone(widget, var))
            phone_entry.bind("<KeyRelease>", self.check_add_row)
            name_var = ctk.StringVar(value=entry.get("name", ""))
            name_entry = ctk.CTkEntry(row_frame, textvariable=name_var, width=200, corner_radius=10)
            name_entry.pack(side="left", padx=5)
            name_entry.bind("<Return>", lambda event, var=name_var: self.validate_name(var))
            name_entry.bind("<KeyRelease>", self.check_add_row)
            bg_color = "#2B2B2B" if ctk.get_appearance_mode().lower() == "dark" else "#F0F0F0"
            indicator = tk.Canvas(row_frame, width=30, height=30, highlightthickness=0, bg=bg_color)
            indicator.create_oval(2, 2, 28, 28, fill="green", outline="")
            indicator.create_text(15, 15, text="✔", fill="white", font=("Arial", 12, "bold"))
            indicator.pack(side="left", padx=5)
            row_dict = {"sno": sno_label, "phone": phone_entry, "phone_var": phone_var,
                        "name": name_entry, "name_var": name_var, "indicator": indicator,
                        "indicator_state": 0, "row_frame": row_frame}
            indicator.bind("<Button-1>", lambda e, r=row_dict: self.toggle_indicator(r))
            self.rows.append(row_dict)
        if not self.rows:
            self.add_row()
    def get_data(self):
        data = []
        for row in self.rows:
            phone = row["phone_var"].get().strip()
            name = row["name_var"].get().strip()
            if phone == "" and name == "":
                continue
            entry = {"phone": phone, "name": name}
            if "skip" in row:
                entry["skip"] = True
            data.append(entry)
        return data

# ----------------------- IMPORT DATABASE POPUP -----------------------
class ImportDatabasePopup(ctk.CTkToplevel):
    def __init__(self, master, import_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Import Excel/CSV Data")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("500x150")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.import_callback = import_callback
        ctk.CTkLabel(self, text="Import Excel/CSV Data", font=("Arial", 14, "bold")).pack(pady=5)
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(fill="x", padx=20, pady=5)
        left = ctk.CTkFrame(frame, corner_radius=10)
        left.pack(side="left", expand=True, fill="both", padx=5)
        ctk.CTkLabel(left, text="Phone Column (e.g., B):").pack(pady=2)
        self.phone_col_var = ctk.StringVar()
        self.phone_entry = ctk.CTkEntry(left, textvariable=self.phone_col_var, corner_radius=10)
        self.phone_entry.pack(pady=2, fill="x")
        right = ctk.CTkFrame(frame, corner_radius=10)
        right.pack(side="left", expand=True, fill="both", padx=5)
        ctk.CTkLabel(right, text="Name Column (e.g., C):").pack(pady=2)
        self.name_col_var = ctk.StringVar()
        self.name_entry = ctk.CTkEntry(right, textvariable=self.name_col_var, corner_radius=10)
        self.name_entry.pack(pady=2, fill="x")
        self.browse_btn = ctk.CTkButton(self, text="Browse Excel/CSV File", corner_radius=10, state="disabled", command=self.browse_file)
        self.browse_btn.pack(pady=10)
        center_window(self)
        self.phone_entry.bind("<KeyRelease>", self.check_fields)
        self.name_entry.bind("<KeyRelease>", self.check_fields)
    def check_fields(self, event):
        if self.phone_col_var.get().strip() and self.name_col_var.get().strip():
            self.browse_btn.configure(state="normal")
        else:
            self.browse_btn.configure(state="disabled")
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.csv *.xls")])
        if path:
            prog = ProgressPopup(self, "Loading Data", total=1)
            threading.Thread(target=lambda: self.import_callback(
                path,
                self.phone_col_var.get().upper(),
                self.name_col_var.get().upper()
            ), daemon=True).start()
            self.after(2000, prog.close)
            self.destroy()

# ----------------------- CUSTOM IMAGE GENERATOR WINDOW -----------------------
class CustomImageWindow(ctk.CTkToplevel):
    def __init__(self, master, excel_data, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Custom Image Generator")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("900x460")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.excel_data = excel_data
        self.configure(padx=10, pady=10)
        center_window(self)
        self.template_image_path = None
        self.font_file_path = None
        self.last_click = (50, 50)
        self.font_size_var = ctk.StringVar(value="50")
        self.text_color_var = ctk.StringVar(value="black")
        self.ratio_options = ["Original", "4:3", "16:9", "5:8", "1:1", "3:2", "21:9"]
        self.ratio_var = ctk.StringVar(value="Original")
        top_frame = ctk.CTkFrame(self, corner_radius=10)
        top_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(top_frame, text="Select Image Ratio:").pack(side="left", padx=5)
        self.ratio_menu = ctk.CTkOptionMenu(top_frame, values=self.ratio_options, variable=self.ratio_var, command=lambda x: self.update_preview())
        self.ratio_menu.pack(side="left", padx=5)
        self.control_frame = ctk.CTkFrame(self, corner_radius=10)
        self.control_frame.pack(side="left", fill="y", padx=10, pady=10)
        self.preview_frame = ctk.CTkFrame(self, corner_radius=10)
        self.preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        ctk.CTkLabel(self.control_frame, text="Custom Image Generator", font=("Arial", 16, "bold")).pack(pady=10)
        self.select_template_btn = ctk.CTkButton(self.control_frame, text="Select Template Image", corner_radius=10, command=self.select_template)
        self.select_template_btn.pack(pady=5, fill="x", padx=10)
        self.select_font_btn = ctk.CTkButton(self.control_frame, text="Select Font File", corner_radius=10, command=self.select_font)
        self.select_font_btn.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(self.control_frame, text="Font Size:").pack(pady=5)
        self.font_size_entry = ctk.CTkEntry(self.control_frame, textvariable=self.font_size_var, corner_radius=10)
        self.font_size_entry.pack(pady=5, fill="x", padx=10)
        self.font_size_entry.bind("<KeyRelease>", lambda e: self.update_preview())
        ctk.CTkLabel(self.control_frame, text="Text Color:").pack(pady=5)
        color_btn = ctk.CTkButton(self.control_frame, text="Choose Color", corner_radius=10, command=self.choose_color)
        color_btn.pack(pady=5, fill="x", padx=10)
        self.set_position_btn = ctk.CTkButton(self.control_frame, text="Set Text Position\n(Click Preview)", corner_radius=10, command=self.instruct_set_position)
        self.set_position_btn.pack(pady=5, fill="x", padx=10)
        self.generate_btn = ctk.CTkButton(self.control_frame, text="Generate Images", fg_color="#0078D7", corner_radius=10, command=self.generate_images_with_progress)
        self.generate_btn.pack(pady=20, fill="x", padx=10)
        self.canvas = ctk.CTkCanvas(self.preview_frame, bg="white", width=800, height=800)
        self.canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas.bind("<Button-1>", self.canvas_click)
        self.preview_image = None
    def choose_color(self):
        color = colorchooser.askcolor(title="Choose text color", parent=self)
        if color and color[1]:
            self.text_color_var.set(color[1])
            self.update_preview()
    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")], parent=self)
        if path:
            self.template_image_path = path
            self.update_preview()
    def select_font(self):
        path = filedialog.askopenfilename(filetypes=[("Font Files", "*.ttf")], parent=self)
        if path:
            self.font_file_path = path
    def canvas_click(self, event):
        self.last_click = (event.x, event.y)
        self.update_preview()
    def instruct_set_position(self):
        messagebox.showinfo("Set Position", "Click on the preview image to set the text position.", parent=self)
    def update_preview(self):
        if not self.template_image_path:
            return
        try:
            ratio = self.ratio_var.get()
            if ratio == "4:3":
                new_size = (800, 600)
            elif ratio == "16:9":
                new_size = (800, 450)
            elif ratio == "5:8":
                new_size = (800, 640)
            elif ratio == "1:1":
                new_size = (800, 800)
            elif ratio == "3:2":
                new_size = (800, int(800*2/3))
            elif ratio == "21:9":
                new_size = (800, int(800*9/21))
            elif ratio == "Original":
                img_temp = Image.open(self.template_image_path).convert("RGB")
                orig_size = img_temp.size
                ratio_val = min(800/orig_size[0], 800/orig_size[1])
                new_size = (int(orig_size[0]*ratio_val), int(orig_size[1]*ratio_val))
            img = Image.open(self.template_image_path).convert("RGB").resize(new_size, Image.Resampling.LANCZOS)
            draw = ImageDraw.Draw(img)
            font_size = int(self.font_size_var.get() or 50)
            font_path = self.font_file_path if self.font_file_path else "arial.ttf"
            font = ImageFont.truetype(font_path, font_size)
            draw.text(self.last_click, "{User_Name}", font=font, fill=self.text_color_var.get())
            self.preview_image = ImageTk.PhotoImage(img)
            self.canvas.config(width=new_size[0], height=new_size[1])
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, image=self.preview_image, anchor="nw")
        except Exception as e:
            messagebox.showerror("Preview Error", f"Error updating preview: {e}", parent=self)
    def generate_images_with_progress(self):
        total = len(self.excel_data)
        prog = ProgressPopup(self, "Generating Images", total)
        prog.geometry(
            f"500x300+{self.winfo_rootx() + (self.winfo_width()-500)//2}"
            f"+{self.winfo_rooty() + (self.winfo_height()-300)//2}"
        )
        self.after(100, lambda: threading.Thread(target=self.generate_images, args=(prog,), daemon=True).start())

    def generate_images(self, prog):
        if not self.template_image_path:
            messagebox.showerror("Error", "No template image selected.", parent=self)
            return
        try:
            font_size = int(self.font_size_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid font size.", parent=self)
            return
        font_path = self.font_file_path if self.font_file_path else "arial.ttf"
        text_color = self.text_color_var.get()
        text_pos = self.last_click
        ratio = self.ratio_var.get()
        if ratio == "4:3":
            new_size = (800, 600)
        elif ratio == "16:9":
            new_size = (800, 450)
        elif ratio == "5:8":
            new_size = (800, 640)
        elif ratio == "1:1":
            new_size = (800, 800)
        elif ratio == "3:2":
            new_size = (800, int(800*2/3))
        elif ratio == "21:9":
            new_size = (800, int(800*9/21))
        elif ratio == "Original":
            img_temp = Image.open(self.template_image_path).convert("RGB")
            orig_size = img_temp.size
            ratio_val = min(800/orig_size[0], 800/orig_size[1])
            new_size = (int(orig_size[0]*ratio_val), int(orig_size[1]*ratio_val))
        count = 0
        for idx, entry in enumerate(self.excel_data, start=1):
            try:
                prog.update_progress(idx)
            except Exception as e:
                print("Progress popup update error:", e)
            if entry.get("skip", False):
                continue
            phone = entry.get("phone", "").strip()
            if not phone:
                entry['image_path'] = None
                continue
            safe_phone = re.sub(r'[<>:"/\\|?*]', '_', phone)
            try:
                img = Image.open(self.template_image_path).convert("RGB").resize(new_size, Image.Resampling.LANCZOS)
                draw = ImageDraw.Draw(img)
                text_to_draw = entry.get("name", "").strip() or f"{idx}"
                font_obj = ImageFont.truetype(font_path, font_size)
                draw.text(text_pos, text_to_draw, font=font_obj, fill=text_color)
                output_path = os.path.join(OUTPUT_IMG_FOLDER, f"{safe_phone}.png")
                img.save(output_path)
                entry['image_path'] = output_path
                count += 1
            except Exception as ex:
                print(f"Error generating image for phone {phone}: {ex}")
        messagebox.showinfo("Generation Complete", f"Generated {count} images in {OUTPUT_IMG_FOLDER}.", parent=self)
        try:
            prog.close()
        except:
            pass
        self.destroy()

# ----------------------- SCHEDULE POPUP -----------------------
class SchedulePopup(ctk.CTkToplevel):
    def __init__(self, master, on_schedule_set, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Schedule Sending")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("300x220")
        self.resizable(False, False)
        self.on_schedule_set = on_schedule_set
        self.wm_attributes("-topmost", True)
        container = ctk.CTkFrame(self, corner_radius=10)
        container.pack(expand=True, fill="both", padx=20, pady=20)
        container.grid_columnconfigure(0, weight=1)
        
        time_frame = ctk.CTkFrame(container, corner_radius=10)
        time_frame.grid(row=0, column=0, pady=10)
        time_frame.grid_columnconfigure((0,1,2), weight=1)
        
        # Hour widget
        hour_frame = ctk.CTkFrame(time_frame, corner_radius=10)
        hour_frame.grid(row=0, column=0, padx=5)
        self.hour_var = tk.StringVar(value="7")
        self.hour_up = ctk.CTkButton(hour_frame, text="▲", width=30, command=self.increment_hour)
        self.hour_up.pack()
        self.hour_entry = ctk.CTkEntry(hour_frame, width=40, corner_radius=10, textvariable=self.hour_var, font=("Arial", 16))
        self.hour_entry.pack(pady=5)
        self.hour_down = ctk.CTkButton(hour_frame, text="▼", width=30, command=self.decrement_hour)
        self.hour_down.pack()
        
        # Minute widget
        min_frame = ctk.CTkFrame(time_frame, corner_radius=10)
        min_frame.grid(row=0, column=1, padx=5)
        self.min_var = tk.StringVar(value="0")
        self.min_up = ctk.CTkButton(min_frame, text="▲", width=30, command=self.increment_min)
        self.min_up.pack()
        self.min_entry = ctk.CTkEntry(min_frame, width=40, corner_radius=10, textvariable=self.min_var, font=("Arial", 16))
        self.min_entry.pack(pady=5)
        self.min_down = ctk.CTkButton(min_frame, text="▼", width=30, command=self.decrement_min)
        self.min_down.pack()
        
        # AM/PM widget
        ampm_frame = ctk.CTkFrame(time_frame, corner_radius=10)
        ampm_frame.grid(row=0, column=2, padx=5)
        self.ampm_var = tk.StringVar(value="AM")
        self.ampm_up = ctk.CTkButton(ampm_frame, text="▲", width=30, command=self.toggle_ampm)
        self.ampm_up.pack()
        self.ampm_entry = ctk.CTkEntry(ampm_frame, width=40, corner_radius=10, textvariable=self.ampm_var, font=("Arial", 16))
        self.ampm_entry.pack(pady=5)
        self.ampm_down = ctk.CTkButton(ampm_frame, text="▼", width=30, command=self.toggle_ampm)
        self.ampm_down.pack()
        
        bottom_frame = ctk.CTkFrame(self, corner_radius=10)
        bottom_frame.pack(side="bottom", fill="x", pady=10)
        ctk.CTkButton(bottom_frame, text="Set", corner_radius=10, font=("Arial", 16), command=self.set_schedule).pack(side="right", padx=10)
        ctk.CTkButton(bottom_frame, text="Cancel", corner_radius=10, font=("Arial", 16), command=self.destroy).pack(side="right", padx=10)
        center_window(self)
    def increment_hour(self):
        try:
            val = int(self.hour_var.get())
        except ValueError:
            val = 7
        val = 12 if val >= 12 else val + 1
        self.hour_var.set(str(val))
    def decrement_hour(self):
        try:
            val = int(self.hour_var.get())
        except ValueError:
            val = 7
        val = 1 if val <= 1 else val - 1
        self.hour_var.set(str(val))
    def increment_min(self):
        try:
            val = int(self.min_var.get())
        except ValueError:
            val = 0
        val = 0 if val >= 59 else val + 1
        self.min_var.set(str(val))
    def decrement_min(self):
        try:
            val = int(self.min_var.get())
        except ValueError:
            val = 0
        val = 59 if val <= 0 else val - 1
        self.min_var.set(str(val))
    def toggle_ampm(self):
        self.ampm_var.set("PM" if self.ampm_var.get().upper() == "AM" else "AM")
    def set_schedule(self):
        try:
            hour = int(self.hour_var.get())
            minute = int(self.min_var.get())
            if not (1 <= hour <= 12):
                raise ValueError("Hour must be 1-12")
            if not (0 <= minute < 60):
                raise ValueError("Minute must be between 0 and 59")
        except ValueError as ve:
            messagebox.showerror("Invalid Time", f"Please enter valid time values: {ve}", parent=self)
            return

        ampm = self.ampm_var.get().upper()
        if ampm == "PM" and hour != 12:
            hour_24 = hour + 12
        elif ampm == "AM" and hour == 12:
            hour_24 = 0
        else:
            hour_24 = hour

        now = datetime.now()
        schedule_time = now.replace(hour=hour_24, minute=minute, second=0, microsecond=0)
        if schedule_time <= now:
            schedule_time += timedelta(days=1)
        self.on_schedule_set(schedule_time)
        self.destroy()

# ----------------------- TRANSLATE POPUP -----------------------
class TranslatePopup(ctk.CTkToplevel):
    def __init__(self, master, process_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Translate Message")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("400x200")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.process_callback = process_callback
        ctk.CTkLabel(self, text="Select Target Language:", font=("Arial", 14, "bold")).pack(pady=10)
        languages = ["English", "Hindi", "Marathi", "Spanish", "French", "German", "Italian", "Portuguese", "Russian", "Chinese", "Japanese", "Korean", "Arabic"]
        self.language_var = tk.StringVar(value="English")
        ctk.CTkOptionMenu(self, values=languages, variable=self.language_var).pack(pady=10)
        ctk.CTkButton(self, text="OK", command=self.on_ok).pack(pady=10)
        center_window(self)
    def on_ok(self):
        lang = self.language_var.get()
        self.process_callback(lang)
        self.destroy()

# ----------------------- HTML REPORT GENERATION -----------------------
def generate_html_report(success, failure):
    import os, webbrowser
    total = success + failure
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>WabulkXpress Messaging Analytics</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
    * {{
        box-sizing: border-box;
        margin: 0;
        padding: 0;
    }}
    body {{
        font-family: 'Arial', sans-serif;
        color: #e0e0e0;
        position: relative;
        min-height: 100vh;
        overflow: hidden;
        padding: 20px;
    }}
    /* Background image with blur effect */
    body::before {{
        content: "";
        background: url("bin/bg.jpg") no-repeat center center fixed;
        background-size: cover;
        filter: blur(8px);
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        z-index: -2;
    }}
    /* Dark overlay for text clarity */
    body::after {{
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background-color: rgba(0, 0, 0, 0.6);
        z-index: -1;
    }}
    /* Container styling */
    .container {{
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        padding: 30px;
        gap: 20px;
        border-radius: 10px;
        background-color: rgba(30, 30, 30, 0.9);
        max-width: 1200px;
        margin: auto;
        box-shadow: 0 0 20px rgba(0,0,0,0.5);
        transition: transform 0.3s ease;
    }}
    .container:hover {{
        transform: scale(1.02);
    }}
    .info {{
        flex: 1;
        padding: 20px;
        background: rgba(0, 0, 0, 0.3);
        border-radius: 10px;
        margin-right: 20px;
    }}
    .info h2 {{
        margin-bottom: 15px;
    }}
    .info p {{
        margin-bottom: 10px;
        line-height: 1.5;
    }}
    .chart-container {{
        flex: 1;
        position: relative;
        max-width: 400px;
        background: rgba(0, 0, 0, 0.3);
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 0 15px rgba(0,0,0,0.4);
        transition: box-shadow 0.3s ease;
    }}
    .chart-container:hover {{
        box-shadow: 0 0 25px rgba(0,0,0,0.6);
    }}
    h1 {{
        margin-bottom: 20px;
        text-align: center;
        font-size: 2em;
    }}
    button {{
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        background-color: #0078D7;
        color: #fff;
        cursor: pointer;
        transition: background-color 0.3s ease, transform 0.3s ease;
        margin-top: 15px;
    }}
    button:hover {{
        background-color: #005fa3;
        transform: scale(1.05);
    }}
    /* Responsive design */
    @media (max-width: 768px) {{
        .container {{
            flex-direction: column;
        }}
        .chart-container {{
            margin-top: 20px;
            max-width: 100%;
        }}
    }}
</style>
</head>
<body onload="window.focus();">
<h1>WabulkXpress Messaging Analytics</h1>
<div class="container">
    <div class="info">
        <h2>Message Summary</h2>
        <p>Total Messages: <strong id="totalCount"></strong></p>
        <p>Success: <strong id="successCount"></strong></p>
        <p>Failure: <strong id="failureCount"></strong></p>
        <button onclick="window.close();">Close Report</button>
    </div>
    <div class="chart-container">
        <canvas id="pieChart"></canvas>
    </div>
</div>
<script>
    // Set the data values (you can replace these with dynamic values from your back-end)
    const total = {total};
    const success = {success};
    const failure = {failure};
    
    // Update the text content with dynamic data
    document.getElementById('totalCount').textContent = total;
    document.getElementById('successCount').textContent = success;
    document.getElementById('failureCount').textContent = failure;
    
    // Setup the Chart.js doughnut chart
    const ctx = document.getElementById('pieChart').getContext('2d');
    const data = {{
        labels: ['Success', 'Failure'],
        datasets: [{{
            data: [success, failure],
            backgroundColor: ['#4CAF50', '#F44336'],
            borderColor: ['#2E7D32', '#C62828'],
            borderWidth: 2,
        }}]
    }};
    const options = {{
        cutout: '70%',
        responsive: true,
        plugins: {{
            legend: {{
                position: 'bottom',
                labels: {{
                    color: '#e0e0e0'
                }}
            }}
        }}
    }};
    new Chart(ctx, {{
        type: 'doughnut',
        data: data,
        options: options
    }});
</script>
</body>
</html>"""
    report_path = os.path.join(os.getcwd(), "Report.html")
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        webbrowser.open("file:///" + report_path)
    except Exception as e:
        print(f"Error generating/opening HTML report: {e}")

# ----------------------- MAIN APPLICATION WINDOW -----------------------
class WabulkXpressApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.title("WabulkXpress")
        self.geometry("1400x900")
        if os.path.exists(TITLE_ICON_PATH):
            icon_img = Image.open(TITLE_ICON_PATH).resize((32,32), Image.Resampling.LANCZOS)
            icon_tk = ImageTk.PhotoImage(icon_img)
            self.wm_iconphoto(False, icon_tk)
        loco_icon = os.path.join(os.getcwd(), "bin", "loco.ico")
        if os.path.exists(loco_icon):
            self.iconbitmap(loco_icon)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.attachments = {"Picture": None, "Video": None, "Document": None}
        self.custom_image_enabled = False
        self.excel_data = []
        self.sending = False
        self.undo_stack = []
        self.redo_stack = []
        self.schedule_time = None
        self.first_cycle = True
        self.last_action = None
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=15)
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)
        self.header = ctk.CTkFrame(self, height=120, corner_radius=15)
        self.header.pack(side="top", fill="x", padx=10, pady=(10,0))
        self.main_area = ctk.CTkFrame(self, corner_radius=15)
        self.main_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.create_sidebar()
        self.create_header()
        self.create_main_area()
        self.gemini_api_key = "AIzaSyDmYy3CFKb0aoVRYZANAyp6X3jgKUe__6g"
        self.gemini_api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.gemini_api_key}"
        self.ai_prompts = {
            "Reframe": "Rephrase the following message in a single cohesive paragraph with no extra disclaimers or stars:",
            "Emoji": "Rewrite the following message with a few relevant emojis, keeping it concise:",
            "Professional": "Rewrite the following message in a polite, professional tone, no extra disclaimers or bullet points:",
            "Funny": "Rewrite the following message with a light, humorous style, no extra disclaimers or bullet points:",
            "Ask AI": "Ask AI: Please answer the following message without adding any extra formatting or stars:",
            "Translate": "Please translate the following message into {lang}, ensuring that you preserve the original formatting exactly no extra disclaimers or stars or Any othermessage from your side only the text translated:",
        }
        self.ai_menu = tk.Menu(self, tearoff=0)
        self.ai_menu.add_command(label="Reframe", command=lambda: self.process_ai("Reframe"))
        self.ai_menu.add_command(label="Emoji", command=lambda: self.process_ai("Emoji"))
        self.ai_menu.add_command(label="Professional", command=lambda: self.process_ai("Professional"))
        self.ai_menu.add_command(label="Funny", command=lambda: self.process_ai("Funny"))
        self.ai_menu.add_command(label="Ask AI", command=lambda: self.process_ai("Ask AI"))
        if ctk.get_appearance_mode().lower() == "dark":
            self.ai_menu.configure(bg="black", fg="white")
        else:
            self.ai_menu.configure(bg="white", fg="black")
        if not os.path.exists(FLAG_FILE):
            FirstRunPopup(self, self.first_run_closed).wait_window()
        else:
            self.first_run_closed()
        self.refresh_icons()
        # Removed UI-reload bindings for better stability

    def first_run_closed(self):
        self.log_live("Welcome to WabulkXpress!")
        
    def refresh_icons(self):
        self.github_button.configure(
            image=self.get_icon("github"),
            fg_color="transparent",
            hover_color="#333333",
            corner_radius=0
        )
        self.update_button.configure(
            image=self.get_icon("update"),
            fg_color="transparent",
            hover_color="#333333",
            corner_radius=0
        )
        self.theme_toggle_button.configure(
            image=self.get_icon("dark"),
            fg_color="transparent",
            hover_color="#333333",
            corner_radius=0
        )

    def create_sidebar(self):
        if os.path.exists(LOGO_PATH):
            img = Image.open(LOGO_PATH).resize((150,150), Image.Resampling.LANCZOS)
            self.sidebar_logo = ctk.CTkImage(img, size=(150,150))
            self.logo_label = ctk.CTkLabel(self.sidebar, image=self.sidebar_logo, text="")
            self.logo_label.pack(pady=(20,10))
        else:
            self.logo_label = ctk.CTkLabel(self.sidebar, text="Logo Missing")
            self.logo_label.pack(pady=(20,10))
        center_frame = ctk.CTkFrame(self.sidebar, corner_radius=0)
        center_frame.pack(expand=True, fill="both")
        center_frame.grid_columnconfigure(0, weight=1)
        center_frame.grid_columnconfigure(1, weight=1)
        self.start_stop_button = ctk.CTkButton(
            center_frame, text="Start", corner_radius=10, height=50, width=100, command=self.toggle_sending
        )
        self.start_stop_button.grid(row=0, column=0, padx=5, pady=10)
        arrow_path = os.path.join(BIN_FOLDER, "down_arrow.png")
        down_arrow_icon = None
        if os.path.exists(arrow_path):
            arrow_img = Image.open(arrow_path).resize((25,25), Image.Resampling.LANCZOS)
            down_arrow_icon = ctk.CTkImage(arrow_img, size=(25,25))
        self.schedule_button = ctk.CTkButton(
            center_frame,
            text="",
            image=down_arrow_icon,
            corner_radius=10,
            height=50,
            width=50,
            command=self.open_schedule_popup
        )
        self.schedule_button.grid(row=0, column=1, padx=5, pady=10)
        self.login_button = ctk.CTkButton(
            center_frame, text="Login", corner_radius=10, height=40, command=self.launch_whatsapp_beta
        )
        self.login_button.grid(row=1, column=0, columnspan=2, padx=5, pady=0)
        self.live_alerts = ctk.CTkTextbox(self.sidebar, height=250, corner_radius=10)
        self.live_alerts.pack(side="bottom", pady=10, padx=20)
        self.live_alerts.insert("0.0", "Live Alerts:\n")
        self.live_alerts.configure(state="disabled")

    def open_schedule_popup(self):
        SchedulePopup(self, self.set_schedule_time)

    def create_header(self):
        self.welcome_label = ctk.CTkLabel(self.header, text="Welcome!", font=("Arial", 24, "bold"))
        self.welcome_label.pack(side="left", padx=20)
        self.github_button = ctk.CTkButton(
            self.header,
            text="",
            width=40,
            command=lambda: webbrowser.open(GITHUB_RELEASES_URL),
        )
        self.github_button.pack(side="right", padx=10)
        # Added check_for_update method stub below
        self.update_button = ctk.CTkButton(
            self.header,
            text="",
            width=40,
            command=self.check_for_update,
        )
        self.update_button.pack(side="right", padx=10)
        # Added toggle_theme method stub below
        self.theme_toggle_button = ctk.CTkButton(
            self.header,
            text="",
            width=40,
            command=self.toggle_theme,
        )
        self.theme_toggle_button.pack(side="right", padx=10)

    def get_icon(self, icon_name):
        mode = ctk.get_appearance_mode().lower()
        if icon_name in ["github", "update"]:
            file_name = f"{icon_name}.png" if mode == "light" else f"{icon_name}_dark.png"
        elif icon_name == "dark":
            file_name = "dark.png" if mode == "light" else "light.png"
        else:
            file_name = f"{icon_name}.png"
        icon_path = os.path.join(os.getcwd(), "bin", file_name)
        size = (30,30)
        if os.path.exists(icon_path):
            img = Image.open(icon_path).resize(size, Image.Resampling.LANCZOS)
        else:
            img = Image.new("RGB", size, "#DBDBDB" if mode=="light" else "#2B2B2B")
        return ctk.CTkImage(img, size=size)

    def create_main_area(self):
        self.main_area.columnconfigure(0, weight=1)
        self.main_area.columnconfigure(1, weight=2)
        self.message_frame = ctk.CTkFrame(self.main_area, corner_radius=15)
        self.message_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.create_message_area(self.message_frame)
        self.excel_frame = ctk.CTkFrame(self.main_area, corner_radius=15)
        self.excel_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.create_excel_area(self.excel_frame)
        self.main_area.rowconfigure(0, weight=1)

    def create_message_area(self, parent):
        button_frame = ctk.CTkFrame(parent, corner_radius=10)
        button_frame.pack(fill="x", pady=10, padx=5)
        # Attachment button restricted to image file types only
        self.attachment_btn = ctk.CTkButton(button_frame, text="Select Attachment", corner_radius=10, height=40, width=200, command=self.handle_attachment)
        self.attachment_btn.pack(side="left", padx=10)
        self.custom_image_btn = ctk.CTkButton(button_frame, text="Custom Image Namer", corner_radius=10, command=self.open_custom_image_window, height=40, width=200)
        self.custom_image_btn.pack(side="left", padx=10)
        # Add hover popup hint for Custom Image button using woi_ci.png
        HoverHint(self.custom_image_btn, "Automatically places the receiver’s name onto your custom image template — perfect for personalized visual messages!", os.path.join(os.getcwd(), "bin", "woi_ci.png"))
        fmt_frame = ctk.CTkFrame(parent, corner_radius=10)
        fmt_frame.pack(fill="x", pady=10, padx=5)
        self.bold_btn = ctk.CTkButton(fmt_frame, text="B", width=30, command=lambda: self.apply_formatting("*"), corner_radius=5)
        self.bold_btn.pack(side="left", padx=2)
        self.italic_btn = ctk.CTkButton(fmt_frame, text="I", width=30, command=lambda: self.apply_formatting("_"), corner_radius=5)
        self.italic_btn.pack(side="left", padx=2)
        self.strike_btn = ctk.CTkButton(fmt_frame, text="S", width=30, command=lambda: self.apply_formatting("~"), corner_radius=5)
        self.strike_btn.pack(side="left", padx=2)
        self.mono_btn = ctk.CTkButton(fmt_frame, text="Code", width=40, command=lambda: self.apply_formatting("```"), corner_radius=5)
        self.mono_btn.pack(side="left", padx=2)
        self.username_btn = ctk.CTkButton(fmt_frame, text="UserName", width=70, command=self.insert_username_placeholder, corner_radius=5)
        self.username_btn.pack(side="left", padx=2)
        # Add hover popup hint for UserName button using woi_un.png
        HoverHint(self.username_btn, "Automatically inserts each contact’s name into your message for a personal touch!", os.path.join(os.getcwd(), "bin", "woi_un.png"))
        self.undo_btn = ctk.CTkButton(fmt_frame, text="Undo", width=50, command=self.undo, corner_radius=5)
        self.undo_btn.pack(side="right", padx=2)
        self.redo_btn = ctk.CTkButton(fmt_frame, text="Redo", width=50, command=self.redo, corner_radius=5)
        self.redo_btn.pack(side="right", padx=2)
        ctk.CTkLabel(parent, text="Message:", font=("Arial",14)).pack(pady=(10,0))
        self.text_area_frame = ctk.CTkFrame(parent, corner_radius=10)
        self.text_area_frame.pack(fill="both", expand=True, padx=5, pady=5)
        mode = ctk.get_appearance_mode().lower()
        bg_color = "gray" if mode == "dark" else "white"
        self.border_canvas = ctk.CTkCanvas(self.text_area_frame, highlightthickness=0, bd=0, bg=bg_color)
        self.border_canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.message_text = ctk.CTkTextbox(self.text_area_frame, height=100, corner_radius=10)
        self.message_text.pack(fill="both", expand=True, padx=0, pady=0)
        self.message_text.bind("<Control-z>", self.undo)
        self.message_text.bind("<Control-y>", self.redo)
        self.message_text.bind("<KeyRelease>", self.save_state)
        # Right-click context menu for formatting tools
        self.message_text.bind("<Button-3>", self.show_context_menu)
        self.context_menu = tk.Menu(self.message_text, tearoff=0)
        self.context_menu.add_command(label="Bold", command=lambda: self.apply_formatting("*"))
        self.context_menu.add_command(label="Italic", command=lambda: self.apply_formatting("_"))
        self.context_menu.add_command(label="Strike", command=lambda: self.apply_formatting("~"))
        self.context_menu.add_command(label="Code", command=lambda: self.apply_formatting("```"))
        self.context_menu.add_command(label="UserName", command=self.insert_username_placeholder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Undo", command=self.undo)
        self.context_menu.add_command(label="Redo", command=self.redo)
        self.translate_button = ctk.CTkButton(parent,
            text="",
            image=self.get_icon("trans") if os.path.exists(os.path.join(BIN_FOLDER, "trans.png")) else None,
            fg_color="white",
            hover_color="#f0f0f0",
            corner_radius=9999,
            width=50,
            height=50,
            command=self.open_translate_popup)
        self.translate_button.place(relx=1.0, rely=1.0, x=-120, y=-100, anchor="se")
        self.translate_button.bind("<ButtonPress>", lambda e: self.translate_button.configure(width=45, height=45))
        self.translate_button.bind("<ButtonRelease>", lambda e: self.translate_button.configure(width=50, height=50))
        ai_icon_path = os.path.join(os.getcwd(), "bin", "ai_icon.png")
        if os.path.exists(ai_icon_path):
            ai_img = Image.open(ai_icon_path).resize((25,25), Image.Resampling.LANCZOS)
            ai_icon = ctk.CTkImage(ai_img, size=(25,25))
        else:
            fallback_img = Image.new("RGB", (25,25), "blue")
            ai_icon = ctk.CTkImage(fallback_img, size=(25,25))
        self.ai_button = ctk.CTkButton(parent,
            text="",
            image=ai_icon,
            fg_color="white",
            hover_color="#f0f0f0",
            corner_radius=9999,
            width=50,
            height=50,
            command=self.show_ai_menu)
        self.ai_button.place(relx=1.0, rely=1.0, x=-40, y=-100, anchor="se")
        self.ai_button.bind("<ButtonPress>", lambda e: self.ai_button.configure(width=45, height=45))
        self.ai_button.bind("<ButtonRelease>", lambda e: self.ai_button.configure(width=50, height=50))
        delay_frame = ctk.CTkFrame(parent, corner_radius=10)
        delay_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(delay_frame, text="Min Delay (s)", font=("Arial",12)).pack(side="left", padx=10)
        self.min_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="Default 1", corner_radius=10, width=100)
        self.min_delay_entry.insert(0, str(DEFAULT_MIN_DELAY))
        self.min_delay_entry.pack(side="left", padx=10)
        ctk.CTkLabel(delay_frame, text="Max Delay (s)", font=("Arial",12)).pack(side="left", padx=10)
        self.max_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="Default 10", corner_radius=10, width=100)
        self.max_delay_entry.insert(0, str(DEFAULT_MAX_DELAY))
        self.max_delay_entry.pack(side="left", padx=10)

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def save_state(self, event=None):
        current_text = self.message_text.get("0.0", "end-1c")
        self.undo_stack.append(current_text)
        self.redo_stack.clear()

    def undo(self, event=None):
        if self.undo_stack:
            current_text = self.message_text.get("0.0", "end-1c")
            self.redo_stack.append(current_text)
            previous_text = self.undo_stack.pop()
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", previous_text)

    def redo(self, event=None):
        if self.redo_stack:
            current_text = self.message_text.get("0.0", "end-1c")
            self.undo_stack.append(current_text)
            next_text = self.redo_stack.pop()
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", next_text)

    def insert_username_placeholder(self):
        self.save_state()
        self.message_text.insert("insert", "()")

    def apply_formatting(self, symbol):
        self.save_state()
        try:
            sel_start = self.message_text.index("sel.first")
            sel_end = self.message_text.index("sel.last")
            text = self.message_text.get(sel_start, sel_end).strip()
            formatted = f"{symbol}{text}{symbol}"
            self.message_text.delete(sel_start, sel_end)
            self.message_text.insert(sel_start, formatted)
        except:
            text = self.message_text.get("0.0", "end-1c").strip()
            formatted = f"{symbol}{text}{symbol}"
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", formatted)

    def handle_attachment(self):
        # Open file chooser restricted to image files only
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if path:
            self.attachments["Picture"] = None
            self.attachments["Video"] = None
            self.attachments["Document"] = None
            self.attachments["Any"] = path
            self.log_live(f"Attachment selected: {os.path.basename(path)}")
            self.last_action = "attachment"

    def open_custom_image_window(self):
        if hasattr(self, "excel_table"):
            self.excel_data = self.excel_table.get_data()
        if not self.excel_data:
            messagebox.showerror("Error", "Please load or enter phone data first.")
            return
        self.custom_image_enabled = True
        self.last_action = "custom"
        CustomImageWindow(self, self.excel_data)

    def create_excel_area(self, parent):
        top_frame = ctk.CTkFrame(parent, corner_radius=10)
        top_frame.pack(fill="x", padx=5, pady=5)
        country_codes = ["+91", "None", "+1", "+44", "+61", "+81", "+49", "+33", "+86", "+7"]
        self.country_code_var = ctk.StringVar(value="+91")
        self.country_code_dropdown = ctk.CTkOptionMenu(top_frame, values=country_codes, variable=self.country_code_var)
        self.country_code_dropdown.pack(side="left", padx=5, pady=5)
        self.import_db_btn = ctk.CTkButton(top_frame, text="Import DataBase", corner_radius=10, command=self.open_import_popup)
        self.import_db_btn.pack(side="left", padx=5, pady=5)
        self.excel_table = ExcelTable(parent, main_app=self)
        self.excel_table.pack(fill="both", expand=True, padx=5, pady=5)

    def open_import_popup(self):
        ImportDatabasePopup(self, self.load_excel_data)

    def load_excel_data(self, path, phone_col, name_col):
        try:
            # Create progress popup
            prog = ProgressPopup(self, "Loading Data", total=1)
            prog.geometry(
                f"500x300+{self.winfo_rootx() + (self.winfo_width()-500)//2}"
                f"+{self.winfo_rooty() + (self.winfo_height()-300)//2}"
            )
            
            # Define the actual loading function
            def perform_loading():
                new_data = []
                if path.lower().endswith((".xlsx", ".xls")):
                    wb = openpyxl.load_workbook(path)
                    sheet = wb.active
                    total_rows = sheet.max_row - 1  # Subtract header row
                    prog.update_progress(0)
                    
                    count = 0
                    for row in range(2, sheet.max_row+1):
                        phone = sheet[f"{phone_col}{row}"].value
                        name = sheet[f"{name_col}{row}"].value
                        if phone is None or str(phone).strip() == "":
                            continue
                        if isinstance(phone, (int, float)):
                            phone = str(int(phone))
                        else:
                            phone = str(phone)
                        phone = phone.strip().replace(" ", "").replace("-", "")
                        if self.country_code_var.get() == "+1":
                            if phone.startswith("+"):
                                pass
                            else:
                                if len(phone) >= 10:
                                    extra = len(phone) - 10
                                    prefix = phone[:extra]
                                    if prefix == "1":
                                        phone = "+" + phone
                                    else:
                                        phone = "+1" + phone[1:]
                                else:
                                    phone = "+1" + phone
                        else:
                            country = self.country_code_var.get()
                            if country != "None" and not phone.startswith(country) and not phone.startswith("+"):
                                phone = country + phone
                        new_data.append({"phone": phone, "name": name.strip() if name else ""})
                        count += 1
                        if count % 10 == 0:  # Update progress every 10 rows
                            prog.update_progress(min(count, total_rows))
                    
                elif path.lower().endswith(".csv"):
                    with open(path, newline='', encoding='utf-8') as csvfile:
                        reader = csv.reader(csvfile)
                        rows = list(reader)
                        total_rows = len(rows) - 1  # Subtract header row
                        prog.update_progress(0)
                        
                        count = 0
                        for row in rows[1:]:
                            try:
                                phone = row[ord(phone_col.upper()) - ord("A")].strip()
                                name = row[ord(name_col.upper()) - ord("A")].strip() if len(row) > (ord(name_col.upper()) - ord("A")) else ""
                            except Exception as ex:
                                continue
                            if phone:
                                phone = phone.replace(" ", "").replace("-", "").strip()
                                if self.country_code_var.get() == "+1":
                                    if phone.startswith("+"):
                                        pass
                                    else:
                                        if len(phone) >= 10:
                                            extra = len(phone) - 10
                                            prefix = phone[:extra]
                                            if prefix == "1":
                                                phone = "+" + phone
                                            else:
                                                phone = "+1" + phone[1:]
                                        else:
                                            phone = "+1" + phone
                                else:
                                    country = self.country_code_var.get()
                                    if country != "None" and not phone.startswith(country) and not phone.startswith("+"):
                                        phone = country + phone
                                new_data.append({"phone": phone, "name": name})
                            count += 1
                            if count % 10 == 0:  # Update progress every 10 rows
                                prog.update_progress(min(count, total_rows))
                
                # Process the loaded data on the main thread
                self.after(0, lambda: self.process_loaded_data(new_data, prog))
            
            # Start loading in a separate thread
            self.after(100, lambda: threading.Thread(target=perform_loading, daemon=True).start())
            
        except Exception as e:
            if 'prog' in locals():
                prog.close()
            messagebox.showerror("Excel Error", f"Error loading Excel/CSV: {e}")

    def process_loaded_data(self, new_data, prog):
        try:
            if not new_data:
                messagebox.showerror("Excel Error", "No valid phone numbers found in Excel/CSV.")
                prog.close()
                return
                
            if self.excel_data:
                # Close progress popup before showing merge prompt
                prog.close()
                self.show_merge_prompt(new_data)
            else:
                self.excel_data = new_data
                self.excel_table.load_data(self.excel_data)
                self.log_live(f"Loaded {len(new_data)} entries from Excel/CSV.")
                prog.close()
        except Exception as e:
            prog.close()
            messagebox.showerror("Excel Error", f"Error processing Excel/CSV data: {e}")

    def show_merge_prompt(self, new_data):
        prompt = ctk.CTkToplevel(self)
        prompt.title("Choose Import Mode")
        prompt.iconbitmap(TITLE_ICON_PATH)
        prompt.geometry("300x150")
        prompt.resizable(False, False)
        prompt.wm_attributes("-topmost", True)
        center_window(prompt)
        ctk.CTkLabel(prompt, text="Import Excel Data", font=("Arial", 14, "bold")).pack(pady=10)
        ctk.CTkLabel(prompt, text="Do you want to merge with current data?").pack(pady=5)
        button_frame = ctk.CTkFrame(prompt)
        button_frame.pack(pady=10)
        def merge_data():
            self.excel_data.extend(new_data)
            self.excel_table.load_data(self.excel_data)
            self.log_live(f"Merged {len(new_data)} new entries. Total: {len(self.excel_data)}")
            prompt.destroy()
        def replace_data():
            self.excel_data = new_data
            self.excel_table.load_data(self.excel_data)
            self.log_live(f"Replaced with {len(new_data)} new entries.")
            prompt.destroy()
        merge_btn = ctk.CTkButton(button_frame, text="Merge", command=merge_data)
        merge_btn.grid(row=0, column=0, padx=5)
        replace_btn = ctk.CTkButton(button_frame, text="Add New", command=replace_data)
        replace_btn.grid(row=0, column=1, padx=5)

    def toggle_sending(self):
        if not self.sending:
            self.start_sending()
        else:
            self.stop_sending()

    def start_sending(self):
        data = self.excel_table.get_data()
        if not data:
            messagebox.showerror("Error", "No phone numbers loaded.")
            return
        msg = self.message_text.get("0.0", "end-1c").strip()
        if not msg and not self.custom_image_enabled and not ("Any" in self.attachments and self.attachments["Any"]):
            messagebox.showerror("Error", "No message text provided and no attachment selected.")
            return
        try:
            min_delay = float(self.min_delay_entry.get())
            max_delay = float(self.max_delay_entry.get())
            if min_delay < 0 or max_delay < 0 or min_delay > max_delay:
                raise ValueError
        except:
            min_delay, max_delay = DEFAULT_MIN_DELAY, DEFAULT_MAX_DELAY
        self.sending = True
        self.start_stop_button.configure(text="Stop")
        self.sent_count = 0
        self.first_cycle = True
        def schedule_and_send():
            if self.schedule_time:
                while datetime.now() < self.schedule_time:
                    remaining = (self.schedule_time - datetime.now()).total_seconds()
                    self.log_live(f"Time left until scheduled start: {int(remaining)} seconds")
                    time.sleep(1)
            self.sending_process(msg, data, min_delay, max_delay)
            generate_html_report(self.sent_count, len(data) - self.sent_count)
            self.schedule_time = None
        threading.Thread(target=schedule_and_send).start()

    def stop_sending(self):
        self.sending = False
        self.start_stop_button.configure(text="Start")
        self.log_live("Sending stopped.")

    def copy_image_to_clipboard(self, image_path):
        try:
            image = Image.open(image_path)
            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            self.log_live("Image copied to clipboard.")
        except Exception as e:
            self.log_live(f"Error copying image: {e}")

    def sending_process(self, msg, data, min_delay, max_delay):
        total = len(data)
        for i, entry in enumerate(data, start=1):
            if not self.sending:
                break
            if entry.get("skip", False):
                continue
            phone = entry["phone"].strip()
            if not phone:
                continue
            name = entry["name"]
            personalized_msg = msg.replace("()", name)
            self.log_live(f"Sending {i} of {total}...")
            attachment_path = None
            if self.last_action == "custom" and self.custom_image_enabled:
                file_name = f"{phone}.png"
                attachment_path = os.path.join(OUTPUT_IMG_FOLDER, file_name)
                if os.path.exists(attachment_path):
                    self.log_live(f"Custom image found for {phone}: {attachment_path}")
                else:
                    self.log_live(f"No custom image found for {phone}.")
                    attachment_path = None
            elif self.last_action == "attachment":
                if "Any" in self.attachments and self.attachments["Any"]:
                    attachment_path = self.attachments["Any"]
                    self.log_live("Using selected attachment.")
            elif self.custom_image_enabled:
                file_name = f"{phone}.png"
                attachment_path = os.path.join(OUTPUT_IMG_FOLDER, file_name)
                if os.path.exists(attachment_path):
                    self.log_live(f"Custom image found for {phone}: {attachment_path}")
                else:
                    self.log_live(f"No custom image found for {phone}.")
                    attachment_path = None
            elif "Any" in self.attachments and self.attachments["Any"]:
                attachment_path = self.attachments["Any"]
                self.log_live("Using selected attachment.")
            phone = re.sub(r"\s+", "", phone)
            country_code = self.country_code_var.get() if hasattr(self, "country_code_var") else ""
            if not phone.startswith("+") and country_code != "None" and not phone.startswith(country_code):
                phone = country_code + phone
            url = f"https://wa.me/{phone}"
            webbrowser.open(url)
            self.log_live(f"Opened chat for {phone}")
            time.sleep(10)
            if attachment_path and os.path.exists(attachment_path):
                ext = os.path.splitext(attachment_path)[1].lower()
                if ext in [".png", ".jpg", ".jpeg"]:
                    self.copy_image_to_clipboard(attachment_path)
                    pyautogui.hotkey('ctrl', 'v')
                    self.log_live("Attachment pasted (image).")
                    time.sleep(2)
                else:
                    copy_file_to_clipboard(attachment_path)
                    pyautogui.hotkey('ctrl', 'v')
                    self.log_live(f"Attachment pasted ({ext}).")
                    time.sleep(2)
            copy_text_to_clipboard(personalized_msg)
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.press('enter')
            self.log_live(f"Text message sent to {phone}: {personalized_msg}")
            self.sent_count += 1
            if self.sent_count % 10 == 0:
                if self.first_cycle:
                    webbrowser.open("https://www.google.com/")
                    webbrowser.open("https://www.google.com/")
                    self.first_cycle = False
                else:
                    webbrowser.open("https://www.google.com/")
                time.sleep(2)
                for _ in range(11):
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(0.3)
                self.log_live("Closed extra tabs.")
                time.sleep(5)
            delay = random.uniform(min_delay, max_delay)
            self.log_live(f"Waiting {delay:.1f} seconds before next...")
            time.sleep(delay)
        self.sending = False
        self.start_stop_button.configure(text="Start")
        self.log_live("All messages processed.")

    def set_schedule_time(self, scheduled_time):
        self.schedule_time = scheduled_time
        self.log_live(f"Scheduled sending at: {scheduled_time.strftime('%Y-%m-%d %I:%M %p')}")
        def schedule_and_trigger_start():
            while datetime.now() < self.schedule_time:
                remaining = (self.schedule_time - datetime.now()).total_seconds()
                self.log_live(f"Time left until scheduled start: {int(remaining)} seconds")
                time.sleep(1)
            # Automatically trigger the start_sending method when scheduled time is reached
            if not self.sending:  # Only start if not already sending
                self.after(0, self.start_sending)  # Use after() to call on main thread
        threading.Thread(target=schedule_and_trigger_start, daemon=True).start()

    def on_close(self):
        self.destroy()

    def log_live(self, text):
        self.live_alerts.configure(state="normal")
        self.live_alerts.insert("end", text + "\n")
        self.live_alerts.configure(state="disabled")
        self.live_alerts.see("end")

    def launch_whatsapp_beta(self):
        if os.path.exists(WHATSAPP_BETA):
            os.startfile(WHATSAPP_BETA)
        else:
            messagebox.showerror("Error", "WhatsApp Beta shortcut not found.")

    def process_ai(self, option):
        self.ai_button.configure(state="disabled")
        self.log_live(f"AI processing started with option: {option}")
        message = self.message_text.get("0.0", "end-1c").strip()
        if not message:
            self.ai_button.configure(state="normal")
            messagebox.showerror("Error", "No message to process.")
            return
        prompt = f"{self.ai_prompts[option]}\n\n{message}"
        threading.Thread(target=self.send_to_gemini, args=(prompt,), daemon=True).start()
    def send_to_gemini(self, prompt):
        try:
            response = requests.post(
                self.gemini_api_url,
                headers={"Content-Type": "application/json"},
                json={"contents": [{"parts": [{"text": prompt}]}]}
            )
            if response.status_code == 200:
                data = response.json()
                generated_text = data["candidates"][0]["content"]["parts"][0]["text"]
                generated_text = re.sub(r'^\*+|\*+$', '', generated_text).strip()
                self.after(0, lambda: self.update_message_text_typing(generated_text))
            else:
                self.after(0, lambda: [self.ai_button.configure(state="normal"), messagebox.showerror("API Error", f"Failed to get response: {response.status_code}")])
        except Exception as e:
            self.after(0, lambda: [self.ai_button.configure(state="normal"), messagebox.showerror("API Error", f"Error: {e}")])
    def update_message_text_typing(self, new_text):
        self.message_text.delete("0.0", "end")
        def type_character(index=0):
            if index < len(new_text):
                # Increase typing speed by reducing the delay from 50ms to 10ms
                # You can adjust this value - lower = faster
                typing_speed = 10  # milliseconds between characters
                self.message_text.insert("end", new_text[index])
                self.after(typing_speed, lambda: type_character(index + 1))
            else:
                self.ai_button.configure(state="normal")
                self.log_live("AI response complete.")
        type_character()

    def show_ai_menu(self):
        try:
            self.ai_menu.tk_popup(self.ai_button.winfo_rootx(), self.ai_button.winfo_rooty())
        finally:
            self.ai_menu.grab_release()

    def open_translate_popup(self):
        TranslatePopup(self, self.process_translate)
    def process_translate(self, language):
        self.translate_button.configure(state="disabled")
        self.log_live(f"Translation started for language: {language}")
        try:
            selected_text = self.message_text.get("sel.first", "sel.last")
        except:
            selected_text = self.message_text.get("0.0", "end-1c")
        prompt = f"Rephrase And Translate the following message into {language}, a single cohesive paragraph with no extra disclaimers or stars: \n\n{selected_text}"
        threading.Thread(target=self.send_translate_to_gemini, args=(prompt,), daemon=True).start()
    
    def send_translate_to_gemini(self, prompt):
        try:
            response = requests.post(
                self.gemini_api_url,
                headers={"Content-Type": "application/json"},
                json={"contents": [{"parts": [{"text": prompt}]}]}
            )
            if response.status_code == 200:
                data = response.json()
                generated_text = data["candidates"][0]["content"]["parts"][0]["text"]
                generated_text = re.sub(r'^\*+|\*+$', '', generated_text).strip()
                self.after(0, lambda: self.update_message_text_typing(generated_text))
                self.after(0, lambda: self.translate_button.configure(state="normal"))
            else:
                self.after(0, lambda: [self.translate_button.configure(state="normal"), 
                                    messagebox.showerror("API Error", f"Failed to get response: {response.status_code}")])
        except Exception as e:
            self.after(0, lambda: [self.translate_button.configure(state="normal"), 
                                messagebox.showerror("API Error", f"Error: {e}")])

    # ----- Stubs for missing functions -----
    def check_for_update(self):
        self.log_live("Checking for update...")
        try:
            response = requests.get(GITHUB_API_URL, timeout=5)
            if response.status_code == 200:
                release_info = response.json()
                latest_version = release_info.get("tag_name", "").lstrip("vV")
                if float(latest_version) > float(CURRENT_VERSION):
                    if messagebox.askyesno("Update Available", f"A new version ({latest_version}) is available.\nUpdate now?"):
                        webbrowser.open(GITHUB_RELEASES_URL)
                else:
                    messagebox.showinfo("No Update", f"You are running WabulkXpress-v{CURRENT_VERSION}")
            else:
                messagebox.showerror("Update Error", "Failed to fetch update info.")
        except Exception as e:
            messagebox.showerror("Update Error", f"Error: {e}")
        self.log_live("Update check completed.")

    def toggle_theme(self):
        # Toggle between light and dark mode.
        current = ctk.get_appearance_mode()
        if current.lower() == "light":
            ctk.set_appearance_mode("Dark")
            self.log_live("Switched to Dark mode.")
        else:
            ctk.set_appearance_mode("Light")
            self.log_live("Switched to Light mode.")
        self.refresh_icons()
    def on_close(self):
        self.sending = False
        for f in os.listdir(OUTPUT_IMG_FOLDER):
            try:
                os.remove(os.path.join(OUTPUT_IMG_FOLDER, f))
            except Exception as e:
                print(f"Error clearing file {f}: {e}")
        self.destroy()
    def show_ai_menu(self):
        x = self.ai_button.winfo_rootx()
        y = self.ai_button.winfo_rooty() - 100
        self.ai_menu.post(x, y)
if __name__ == "__main__":
    app = WabulkXpressApp()
    app.mainloop()

