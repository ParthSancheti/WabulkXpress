import os
import re
import time
import random
import threading
import webbrowser
import requests
import pyautogui
import openpyxl
from tkinter import filedialog, messagebox, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
from io import BytesIO
import win32clipboard
import customtkinter as ctk
import tkinter as tk
from datetime import datetime, timedelta

# ----------------------- GLOBAL CONSTANTS & PATHS ----------------------- #
CURRENT_VERSION = "12"  # version number
GITHUB_API_URL = "https://api.github.com/repos/Parth-Sancheti-5/WaBulkSender/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/Parth-Sancheti-5/WaBulkSender/releases"
FLAG_FILE = "first_run.flag"

BIN_FOLDER = os.path.join(os.getcwd(), "bin")
TITLE_ICON_PATH = os.path.join(BIN_FOLDER, "loco.ico")
LOGO_PATH = os.path.join(BIN_FOLDER, "Logo.png")
WHATSAPP_BETA = os.path.join(BIN_FOLDER, "WhatsApp_Beta.lnk")
OUTPUT_IMG_FOLDER = os.path.join(os.getcwd(), "output_img")
DEFAULT_MIN_DELAY = 1
DEFAULT_MAX_DELAY = 10
INSTRUCTION_IMAGE_PATH = os.path.join(BIN_FOLDER, "woi.png")

if not os.path.exists(OUTPUT_IMG_FOLDER):
    os.makedirs(OUTPUT_IMG_FOLDER)

# ----------------------- UTILITY: Center Window ----------------------- #
def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")

# ----------------------- FIRST RUN POPUP ----------------------- #
class FirstRunPopup(ctk.CTkToplevel):
    def __init__(self, master, on_close_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Welcome!")
        self.geometry("1200x800")
        self.iconbitmap(TITLE_ICON_PATH)
        self.resizable(False, False)
        self.on_close_callback = on_close_callback
        self.wm_attributes("-topmost", True)
        if os.path.exists(INSTRUCTION_IMAGE_PATH):
            img = Image.open(INSTRUCTION_IMAGE_PATH).resize((1280, 600), Image.Resampling.LANCZOS)
            self.instruction_photo = ImageTk.PhotoImage(img)
            self.instruction_label = ctk.CTkLabel(self, image=self.instruction_photo, text="")
            self.instruction_label.pack(pady=(20, 10), anchor="center")
        else:
            self.instruction_label = ctk.CTkLabel(self, text="Instruction Image Missing", font=("Arial", 16))
            self.instruction_label.pack(pady=(20, 10), anchor="center")
        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.pack(side="bottom", fill="x", padx=20, pady=10)
        self.dont_show_var = ctk.BooleanVar(value=True)
        self.checkbox = ctk.CTkCheckBox(bottom_frame, text="Don't show this again", variable=self.dont_show_var)
        self.checkbox.pack(side="left", padx=10)
        self.ok_button = ctk.CTkButton(bottom_frame, text="OK", fg_color="#0078D7", corner_radius=10, command=self.close_popup)
        self.ok_button.pack(side="right", padx=10)
        center_window(self)
    def close_popup(self):
        if self.dont_show_var.get():
            with open(FLAG_FILE, "w") as f:
                f.write("shown")
        self.on_close_callback()
        self.destroy()

# ----------------------- EXCEL-LIKE TABLE (Editable) ----------------------- #
class ExcelTable(ctk.CTkScrollableFrame):
    def __init__(self, master, main_app, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.main_app = main_app
        self.configure(corner_radius=15, scrollbar_fg_color="gray")
        self.rows = []
        self.add_header()
        self.prepopulate_rows(3)
    def add_header(self):
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=5, pady=3)
        ctk.CTkLabel(header_frame, text="S.No.", width=40, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Phone Number", width=200, anchor="center").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Name", width=200, anchor="center").pack(side="left", padx=5)
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
        self.rows.append({"sno": sno_label, "phone": phone_entry, "phone_var": phone_var,
                          "name": name_entry, "name_var": name_var})
    def validate_phone(self, widget, var):
        text = var.get().strip()
        text = re.sub(r"\s+", "", text)
        default_code = self.main_app.country_code_var.get()
        if default_code != "None" and not text.startswith(default_code) and not text.startswith("+"):
            text = default_code + text
        var.set(text)
        widget.delete(0, "end")
        widget.insert(0, text)
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
            self.rows.append({"sno": sno_label, "phone": phone_entry, "phone_var": phone_var,
                              "name": name_entry, "name_var": name_var})
    def get_data(self):
        data = []
        for row in self.rows:
            phone = row["phone_var"].get().strip()
            name = row["name_var"].get().strip()
            if phone:
                data.append({"phone": phone, "name": name})
        return data

# ----------------------- IMPORT DATABASE POPUP ----------------------- #
class ImportDatabasePopup(ctk.CTkToplevel):
    def __init__(self, master, import_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Import Excel Data")
        self.geometry("500x150")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.import_callback = import_callback
        ctk.CTkLabel(self, text="Import Excel Data", font=("Arial", 14, "bold")).pack(pady=5)
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
        self.browse_btn = ctk.CTkButton(self, text="Browse Excel File", corner_radius=10, state="disabled", command=self.browse_file)
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
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.import_callback(path, self.phone_col_var.get().upper(), self.name_col_var.get().upper())
            self.destroy()

# ----------------------- CUSTOM IMAGE GENERATOR WINDOW ----------------------- #
class CustomImageWindow(ctk.CTkToplevel):
    def __init__(self, master, excel_data, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Custom Image Generator")
        self.geometry("1200x700")
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
        self.ratio_var = ctk.StringVar(value="Original")
        top_frame = ctk.CTkFrame(self, corner_radius=10)
        top_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(top_frame, text="Select Image Ratio:").pack(side="left", padx=5)
        self.ratio_menu = ctk.CTkOptionMenu(top_frame, values=["Original", "4:3", "16:9"], variable=self.ratio_var)
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
        ctk.CTkLabel(self.control_frame, text="Text Color:").pack(pady=5)
        color_btn = ctk.CTkButton(self.control_frame, text="Choose Color", corner_radius=10, command=self.choose_color)
        color_btn.pack(pady=5, fill="x", padx=10)
        self.set_position_btn = ctk.CTkButton(self.control_frame, text="Set Text Position\n(Click Preview)", corner_radius=10, command=self.instruct_set_position)
        self.set_position_btn.pack(pady=5, fill="x", padx=10)
        self.generate_btn = ctk.CTkButton(self.control_frame, text="Generate Images", fg_color="#0078D7", corner_radius=10, command=self.generate_images)
        self.generate_btn.pack(pady=20, fill="x", padx=10)
        self.canvas = ctk.CTkCanvas(self.preview_frame, bg="white", width=800, height=800)
        self.canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas.bind("<Button-1>", self.canvas_click)
        self.preview_image = None
    def choose_color(self):
        color = colorchooser.askcolor(title="Choose text color", parent=self)
        if color and color[1]:
            self.text_color_var.set(color[1])
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
            else:
                new_size = (800, 800)
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
    def generate_images(self):
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
        else:
            new_size = (800, 800)
        count = 0
        for idx, entry in enumerate(self.excel_data, start=1):
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
        self.destroy()

# ----------------------- SCHEDULE POPUP ----------------------- #
class SchedulePopup(ctk.CTkToplevel):
    def __init__(self, master, on_schedule_set, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Schedule Sending")
        self.geometry("300x220")
        self.resizable(False, False)
        self.on_schedule_set = on_schedule_set
        self.wm_attributes("-topmost", True)
        container = ctk.CTkFrame(self, corner_radius=10)
        container.pack(expand=True, fill="both", padx=20, pady=20)
        self.hour_var = tk.IntVar(value=7)
        hour_frame = ctk.CTkFrame(container, corner_radius=10)
        hour_frame.grid(row=0, column=0, padx=5)
        self.hour_up = ctk.CTkButton(hour_frame, text="▲", width=30, command=self.increment_hour)
        self.hour_up.pack()
        self.hour_entry = ctk.CTkEntry(hour_frame, width=40, corner_radius=10, textvariable=self.hour_var, font=("Arial", 14))
        self.hour_entry.pack(pady=5)
        self.hour_down = ctk.CTkButton(hour_frame, text="▼", width=30, command=self.decrement_hour)
        self.hour_down.pack()
        self.min_var = tk.IntVar(value=0)
        min_frame = ctk.CTkFrame(container, corner_radius=10)
        min_frame.grid(row=0, column=1, padx=5)
        self.min_up = ctk.CTkButton(min_frame, text="▲", width=30, command=self.increment_min)
        self.min_up.pack()
        self.min_entry = ctk.CTkEntry(min_frame, width=40, corner_radius=10, textvariable=self.min_var, font=("Arial", 14))
        self.min_entry.pack(pady=5)
        self.min_down = ctk.CTkButton(min_frame, text="▼", width=30, command=self.decrement_min)
        self.min_down.pack()
        self.ampm_var = tk.StringVar(value="AM")
        ampm_frame = ctk.CTkFrame(container, corner_radius=10)
        ampm_frame.grid(row=0, column=2, padx=5)
        self.ampm_up = ctk.CTkButton(ampm_frame, text="▲", width=30, command=self.toggle_ampm)
        self.ampm_up.pack()
        self.ampm_entry = ctk.CTkEntry(ampm_frame, width=40, corner_radius=10, textvariable=self.ampm_var, font=("Arial", 14))
        self.ampm_entry.pack(pady=5)
        self.ampm_down = ctk.CTkButton(ampm_frame, text="▼", width=30, command=self.toggle_ampm)
        self.ampm_down.pack()
        bottom_frame = ctk.CTkFrame(self, corner_radius=10)
        bottom_frame.pack(side="bottom", fill="x", pady=10)
        ctk.CTkButton(bottom_frame, text="Set", corner_radius=10, font=("Arial", 14), command=self.set_schedule).pack(side="right", padx=10)
        ctk.CTkButton(bottom_frame, text="Cancel", corner_radius=10, font=("Arial", 14), command=self.destroy).pack(side="right", padx=10)
        center_window(self)
    def increment_hour(self):
        val = self.hour_var.get()
        val = 12 if val == 12 else val + 1
        self.hour_var.set(val)
    def decrement_hour(self):
        val = self.hour_var.get()
        val = 1 if val == 1 else val - 1
        self.hour_var.set(val)
    def increment_min(self):
        val = self.min_var.get()
        val = 0 if val == 59 else val + 1
        self.min_var.set(val)
    def decrement_min(self):
        val = self.min_var.get()
        val = 59 if val == 0 else val - 1
        self.min_var.set(val)
    def toggle_ampm(self):
        self.ampm_var.set("PM" if self.ampm_var.get() == "AM" else "AM")
    def set_schedule(self):
        hour = self.hour_var.get()
        minute = self.min_var.get()
        ampm = self.ampm_var.get()
        if ampm == "PM" and hour != 12:
            hour_24 = hour + 12
        elif ampm == "AM" and hour == 12:
            hour_24 = 0
        else:
            hour_24 = hour
        now = datetime.now()
        schedule_today = now.replace(hour=hour_24, minute=minute, second=0, microsecond=0)
        if schedule_today <= now:
            schedule_today += timedelta(days=1)
        self.on_schedule_set(schedule_today)
        self.destroy()

# ----------------------- MAIN APPLICATION WINDOW ----------------------- #
class WaBulkSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.title("WaBulkSender")
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
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=15)
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)
        self.header = ctk.CTkFrame(self, height=120, corner_radius=15)
        self.header.pack(side="top", fill="x", padx=10, pady=(10,0))
        self.main_area = ctk.CTkFrame(self, corner_radius=15)
        self.main_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.create_sidebar()
        self.create_header()
        self.create_main_area()
        self.gemini_api_key = "AIzaSyDmYy3CFKb0aoVRYZANAyp6X3jgKUe__6g"  # Replace with your actual key
        self.gemini_api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.gemini_api_key}"
        self.ai_prompts = {
            "Reframe": "Rephrase the following message in a single cohesive paragraph with no extra disclaimers or stars:",
            "Emoji": "Rewrite the following message with a few relevant emojis, keeping it concise:",
            "Professional": "Rewrite the following message in a polite, professional tone, no extra disclaimers or bullet points:",
            "Funny": "Rewrite the following message with a light, humorous style, no extra disclaimers or bullet points:"
        }
        self.ai_menu = tk.Menu(self, tearoff=0)
        self.ai_menu.add_command(label="Reframe", command=lambda: self.process_ai("Reframe"))
        self.ai_menu.add_command(label="Emoji", command=lambda: self.process_ai("Emoji"))
        self.ai_menu.add_command(label="Professional", command=lambda: self.process_ai("Professional"))
        self.ai_menu.add_command(label="Funny", command=lambda: self.process_ai("Funny"))
        if ctk.get_appearance_mode().lower() == "dark":
            self.ai_menu.configure(bg="black", fg="white")
        else:
            self.ai_menu.configure(bg="white", fg="black")
        if not os.path.exists(FLAG_FILE):
            FirstRunPopup(self, self.first_run_closed).wait_window()
        else:
            self.first_run_closed()
        self.refresh_icons()
        # For border animation on the message text area
        self.border_animating = False
        self.border_delay_ms = 80
        self.border_step = 0
        self.border_colors = [
            "#FF0000", "#FF7F00", "#FFFF00", "#00FF00",
            "#0000FF", "#4B0082", "#8F00FF"
        ]
    def first_run_closed(self):
        self.log_live("Welcome to WaBulkSender!")
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
        self.welcome_label = ctk.CTkLabel(self.header, text="Welcome!", font=("Arial",24,"bold"))
        self.welcome_label.pack(side="left", padx=20)
        self.github_button = ctk.CTkButton(
            self.header,
            text="",
            width=40,
            command=lambda: webbrowser.open(GITHUB_RELEASES_URL),
        )
        self.github_button.pack(side="right", padx=10)
        self.update_button = ctk.CTkButton(
            self.header,
            text="",
            width=40,
            command=self.check_for_update,
        )
        self.update_button.pack(side="right", padx=10)
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
        self.main_area.columnconfigure(1, weight=1)
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
        self.attachment_var = ctk.StringVar(value="Select Attachment")
        self.attachment_menu = ctk.CTkOptionMenu(
            button_frame,
            values=["Picture", "Video", "Document"],
            variable=self.attachment_var,
            command=self.handle_attachment,
            height=40,
            width=200
        )
        self.attachment_menu.pack(side="left", padx=10)
        self.custom_image_btn = ctk.CTkButton(
            button_frame, text="Custom Image Namer", corner_radius=10, command=self.open_custom_image_window, height=40, width=200
        )
        self.custom_image_btn.pack(side="left", padx=10)
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
        ai_icon_path = os.path.join(os.getcwd(), "bin", "ai_icon.png")
        if os.path.exists(ai_icon_path):
            ai_img = Image.open(ai_icon_path).resize((25,25), Image.Resampling.LANCZOS)
            ai_icon_ctk = ctk.CTkImage(ai_img, size=(25,25))
        else:
            fallback_img = Image.new("RGB", (25,25), "blue")
            ai_icon_ctk = ctk.CTkImage(fallback_img, size=(25,25))
        self.ai_button = ctk.CTkButton(
            parent,
            text="",
            image=ai_icon_ctk,
            fg_color="white",
            hover_color="#f0f0f0",
            corner_radius=9999,
            width=50,
            height=50,
            command=self.show_ai_menu
        )
        self.ai_button.place(relx=1.0, rely=1.0, x=-60, y=-60, anchor="se")
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
    def start_border_animation(self):
        self.border_animating = True
        self.border_step = 0
        self.animate_border()
    def stop_border_animation(self):
        self.border_animating = False
        self.border_canvas.delete("all")
    def animate_border(self):
        if not self.border_animating:
            return
        self.border_canvas.delete("all")
        w = self.text_area_frame.winfo_width()
        h = self.text_area_frame.winfo_height()
        top_color    = self.border_colors[(self.border_step + 0) % len(self.border_colors)]
        right_color  = self.border_colors[(self.border_step + 1) % len(self.border_colors)]
        bottom_color = self.border_colors[(self.border_step + 2) % len(self.border_colors)]
        left_color   = self.border_colors[(self.border_step + 3) % len(self.border_colors)]
        thickness = 5
        self.border_canvas.create_line(thickness//2, thickness//2, w-thickness//2, thickness//2, fill=top_color, width=thickness)
        self.border_canvas.create_line(w-thickness//2, thickness//2, w-thickness//2, h-thickness//2, fill=right_color, width=thickness)
        self.border_canvas.create_line(w-thickness//2, h-thickness//2, thickness//2, h-thickness//2, fill=bottom_color, width=thickness)
        self.border_canvas.create_line(thickness//2, h-thickness//2, thickness//2, thickness//2, fill=left_color, width=thickness)
        self.border_step = (self.border_step + 1) % len(self.border_colors)
        self.after(self.border_delay_ms, self.animate_border)
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
    def handle_attachment(self, selection):
        if selection in ["Picture", "Video", "Document"]:
            filetypes = []
            if selection == "Picture":
                filetypes = [("Image Files", "*.png;*.jpg;*.jpeg")]
            elif selection == "Video":
                filetypes = [("Video Files", "*.mp4;*.avi;*.mov")]
            elif selection == "Document":
                filetypes = [("Document Files", "*.pdf;*.docx;*.txt")]
            path = filedialog.askopenfilename(filetypes=filetypes)
            if path:
                self.attachments[selection] = path
                self.log_live(f"{selection} attached: {os.path.basename(path)}")
        self.attachment_var.set("Select Attachment")
    def open_custom_image_window(self):
        if hasattr(self, "excel_table"):
            self.excel_data = self.excel_table.get_data()
        if not self.excel_data:
            messagebox.showerror("Error", "Please load or enter phone data first.")
            return
        self.custom_image_enabled = True
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
            wb = openpyxl.load_workbook(path)
            sheet = wb.active
            new_data = []
            for row in range(2, sheet.max_row+1):
                phone = sheet[f"{phone_col}{row}"].value
                name = sheet[f"{name_col}{row}"].value if name_col else ""
                if phone:
                    phone = str(phone).strip()
                    phone = re.sub(r"\s+", "", phone)
                    phone = phone.replace("-", "")
                    phone = phone.lstrip("0")
                    country = self.country_code_var.get()
                    if country != "None":
                        country_digits = country.lstrip("+")
                        if phone.startswith(country_digits) and not phone.startswith("+"):
                            phone = "+" + phone
                        elif not phone.startswith("+"):
                            phone = country + phone
                    new_data.append({"phone": phone, "name": name.strip() if name else ""})
            if not new_data:
                messagebox.showerror("Excel Error", "No valid phone numbers found in Excel.")
                return
            if self.excel_data:
                self.show_merge_prompt(new_data)
            else:
                self.excel_data = new_data
                self.excel_table.load_data(self.excel_data)
                self.log_live(f"Loaded {len(new_data)} entries from Excel.")
        except Exception as e:
            messagebox.showerror("Excel Error", f"Error loading Excel: {e}")
    def show_merge_prompt(self, new_data):
        prompt = ctk.CTkToplevel(self)
        prompt.title("Choose Import Mode")
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
            self.log_live(f"Merged {len(new_data)} new entries. Total entries: {len(self.excel_data)}")
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
        if not msg:
            messagebox.showerror("Error", "No message text provided.")
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
        def schedule_and_send():
            if self.schedule_time:
                now = datetime.now()
                if now < self.schedule_time:
                    wait_seconds = (self.schedule_time - now).total_seconds()
                    self.log_live(
                        f"Waiting until scheduled time: {self.schedule_time.strftime('%Y-%m-%d %I:%M %p')} "
                        f"(about {int(wait_seconds)} seconds)"
                    )
                    while datetime.now() < self.schedule_time:
                        time.sleep(1)
                        if not self.sending:
                            return
            self.sending_process(msg, data, min_delay, max_delay)
            self.schedule_time = None
        threading.Thread(target=schedule_and_send, daemon=True).start()
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
        sent_count = 0
        total = len(data)
        for i, entry in enumerate(data, start=1):
            if not self.sending:
                break
            phone = entry["phone"]
            name = entry["name"]
            personalized_msg = msg.replace("()", name)
            self.log_live(f"Sending {i} of {total}...")
            if self.custom_image_enabled:
                file_name = f"{phone}.png"
                image_path = os.path.join(OUTPUT_IMG_FOLDER, file_name)
                if os.path.exists(image_path):
                    self.log_live(f"Custom image found for {phone}: {image_path}")
                else:
                    self.log_live(f"No custom image found for {phone}.")
                    image_path = None
            elif self.attachments["Picture"]:
                image_path = self.attachments["Picture"]
                self.log_live("Using attached picture.")
            else:
                image_path = None
            phone = re.sub(r"\s+", "", phone)
            country_code = self.country_code_var.get() if hasattr(self, "country_code_var") else ""
            if not phone.startswith("+") and country_code != "None" and not phone.startswith(country_code):
                phone = country_code + phone
            url = f"https://wa.me/{phone}"
            webbrowser.open(url)
            self.log_live(f"Opened chat for {phone}")
            time.sleep(10)
            if image_path and os.path.exists(image_path):
                self.copy_image_to_clipboard(image_path)
                pyautogui.hotkey('ctrl', 'v')
                self.log_live(f"Image pasted for {phone}.")
                time.sleep(2)
            for line in personalized_msg.split("\n"):
                pyautogui.write(line, interval=0.05)
                pyautogui.hotkey('shift', 'enter')
            time.sleep(0.5)
            pyautogui.press('enter')
            self.log_live(f"Text message sent to {phone}: {personalized_msg}")
            sent_count += 1
            if sent_count % 10 == 0:
                webbrowser.open("https://www.google.com/")
                time.sleep(2)
                for _ in range(10):
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(0.3)
                self.log_live("Closed 10 tabs.")
                time.sleep(5)
            delay = random.uniform(min_delay, max_delay)
            self.log_live(f"Waiting {delay:.1f} seconds before next...")
            time.sleep(delay)
        self.sending = False
        self.start_stop_button.configure(text="Start")
        self.log_live("All messages processed.")
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
                    messagebox.showinfo("No Update", f"You are running WaBulkSender-v{CURRENT_VERSION}")
            else:
                messagebox.showerror("Update Error", "Failed to fetch update info.")
        except Exception as e:
            messagebox.showerror("Update Error", f"Error: {e}")
        self.log_live("Update check completed.")
    def launch_whatsapp_beta(self):
        if os.path.exists(WHATSAPP_BETA):
            os.startfile(WHATSAPP_BETA)
            self.log_live("Launching WhatsApp Beta...")
        else:
            messagebox.showerror("Error", f"{WHATSAPP_BETA} not found.")
    def log_live(self, message):
        self.live_alerts.configure(state="normal")
        self.live_alerts.insert("end", f"{message}\n")
        self.live_alerts.see("end")
        self.live_alerts.configure(state="disabled")
    def toggle_theme(self):
        current = ctk.get_appearance_mode().lower()
        new_mode = "light" if current=="dark" else "dark"
        ctk.set_appearance_mode(new_mode)
        self.refresh_icons()
        if new_mode == "dark":
            self.ai_menu.configure(bg="black", fg="white")
        else:
            self.ai_menu.configure(bg="white", fg="black")
        self.log_live(f"Theme switched to {new_mode.capitalize()} Mode.")
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
    def process_ai(self, option):
        self.start_border_animation()
        message = self.message_text.get("0.0", "end-1c").strip()
        if not message:
            self.stop_border_animation()
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
                self.after(0, lambda: self.update_message_text(generated_text))
            else:
                self.after(0, lambda: [self.stop_border_animation(), messagebox.showerror("API Error", f"Failed to get response: {response.status_code}")])
        except Exception as e:
            self.after(0, lambda: [self.stop_border_animation(), messagebox.showerror("API Error", f"Error: {e}")])
    def update_message_text(self, new_text):
        self.message_text.delete("0.0", "end")
        self.message_text.insert("0.0", new_text)
        self.stop_border_animation()
    def start_border_animation(self):
        self.border_animating = True
        self.border_step = 0
        self.animate_border()
    def stop_border_animation(self):
        self.border_animating = False
        self.border_canvas.delete("all")
    def animate_border(self):
        if not self.border_animating:
            return
        self.border_canvas.delete("all")
        w = self.text_area_frame.winfo_width()
        h = self.text_area_frame.winfo_height()
        top_color = self.border_colors[(self.border_step + 0) % len(self.border_colors)]
        right_color = self.border_colors[(self.border_step + 1) % len(self.border_colors)]
        bottom_color = self.border_colors[(self.border_step + 2) % len(self.border_colors)]
        left_color = self.border_colors[(self.border_step + 3) % len(self.border_colors)]
        thickness = 5
        self.border_canvas.create_line(thickness//2, thickness//2, w-thickness//2, thickness//2, fill=top_color, width=thickness)
        self.border_canvas.create_line(w-thickness//2, thickness//2, w-thickness//2, h-thickness//2, fill=right_color, width=thickness)
        self.border_canvas.create_line(w-thickness//2, h-thickness//2, thickness//2, h-thickness//2, fill=bottom_color, width=thickness)
        self.border_canvas.create_line(thickness//2, h-thickness//2, thickness//2, thickness//2, fill=left_color, width=thickness)
        self.border_step = (self.border_step + 1) % len(self.border_colors)
        self.after(self.border_delay_ms, self.animate_border)

if __name__=="__main__":
    app = WaBulkSenderApp()
    app.mainloop()
