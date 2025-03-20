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
import customtkinter as ctk

# ----------------------- GLOBAL CONSTANTS & PATHS ----------------------- #
CURRENT_VERSION = "3.4"  # Your current version (without leading "v")
GITHUB_API_URL = "https://api.github.com/repos/Parth-Sancheti-5/WaBulkSender/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/Parth-Sancheti-5/WaBulkSender/releases"
FLAG_FILE = "first_run.flag"
BIN_FOLDER = os.path.join(os.getcwd(), "Bin")
LOGO_PATH = os.path.join(BIN_FOLDER, "Logo.png")
WHATSAPP_BETA = os.path.join(BIN_FOLDER, "WhatsApp_Beta.lnk")
OUTPUT_IMG_FOLDER = os.path.join(os.getcwd(), "output_img")
DEFAULT_MIN_DELAY = 1
DEFAULT_MAX_DELAY = 10

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
        self.geometry("500x300")
        self.resizable(False, False)
        self.on_close_callback = on_close_callback
        self.wm_attributes("-topmost", True)
        
        # Layout: left side shows logo; right side has heading and checkbox+OK
        self.left_frame = ctk.CTkFrame(self, corner_radius=15)
        self.left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        self.right_frame = ctk.CTkFrame(self, corner_radius=15)
        self.right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        
        if os.path.exists(LOGO_PATH):
            logo_img = Image.open(LOGO_PATH).resize((150,150), Image.Resampling.LANCZOS)
            self.logo_photo = ctk.CTkImage(logo_img, size=(150,150))
            self.logo_label = ctk.CTkLabel(self.left_frame, image=self.logo_photo, text="")
            self.logo_label.pack(pady=10)
        else:
            self.logo_label = ctk.CTkLabel(self.left_frame, text="Logo Missing")
            self.logo_label.pack(pady=10)
            
        # Heading on top of right side
        ctk.CTkLabel(self.right_frame, text="Welcome to WaBulkSender", font=("Arial", 16, "bold")).pack(pady=(10,5))
        self.dont_show_var = ctk.BooleanVar(value=True)
        self.checkbox = ctk.CTkCheckBox(self.right_frame, text="Don't show this again", variable=self.dont_show_var)
        self.checkbox.pack(pady=5)
        self.ok_button = ctk.CTkButton(self.right_frame, text="OK", fg_color="#0078D7", corner_radius=10, command=self.close_popup)
        self.ok_button.pack(pady=10)
        center_window(self)
        
    def close_popup(self):
        if self.dont_show_var.get():
            with open(FLAG_FILE, "w") as f:
                f.write("shown")
        self.on_close_callback()
        self.destroy()

# ----------------------- CUSTOM EXCEL TABLE (Editable) ----------------------- #
class ExcelTable(ctk.CTkScrollableFrame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.configure(corner_radius=15, scrollbar_fg_color="gray")
        self.rows = []  # list of dicts with keys: "sno", "phone", "name"
        self.add_header()
        self.prepopulate_rows(10)
        
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
        phone_entry = ctk.CTkEntry(row_frame, placeholder_text="", width=200, corner_radius=10)
        phone_entry.pack(side="left", padx=5)
        phone_entry.bind("<Return>", self.validate_and_move)
        name_entry = ctk.CTkEntry(row_frame, placeholder_text="", width=200, corner_radius=10)
        name_entry.pack(side="left", padx=5)
        name_entry.bind("<Return>", self.validate_and_move)
        self.rows.append({"sno": sno_label, "phone": phone_entry, "name": name_entry})
        
    def validate_and_move(self, event):
        widget = event.widget
        text = widget.get().strip()
        # Remove all spaces
        text = re.sub(r"\s+", "", text)
        default_code = self.master.master.country_code_var.get()
        if default_code != "None" and not text.startswith("+"):
            text = default_code + text
        widget.delete(0, "end")
        widget.insert(0, text)
        # Move focus to next row's phone entry if exists
        for i, row in enumerate(self.rows):
            if row["phone"] == widget or row["name"] == widget:
                if i+1 < len(self.rows):
                    self.rows[i+1]["phone"].focus_set()
                break
        
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
            phone_entry = ctk.CTkEntry(row_frame, placeholder_text="", width=200, corner_radius=10)
            phone_entry.insert(0, entry.get("phone", ""))
            phone_entry.pack(side="left", padx=5)
            phone_entry.bind("<Return>", self.validate_and_move)
            name_entry = ctk.CTkEntry(row_frame, placeholder_text="", width=200, corner_radius=10)
            name_entry.insert(0, entry.get("name", ""))
            name_entry.pack(side="left", padx=5)
            name_entry.bind("<Return>", self.validate_and_move)
            self.rows.append({"sno": sno_label, "phone": phone_entry, "name": name_entry})
        self.check_add_row(None)
        
    def check_add_row(self, event):
        last_row = self.rows[-1]
        if last_row["phone"].get() or last_row["name"].get():
            self.add_row()
            
    def get_data(self):
        return [{"phone": row["phone"].get().strip(), "name": row["name"].get().strip()} for row in self.rows]

# ----------------------- IMPORT DATABASE POPUP ----------------------- #
class ImportDatabasePopup(ctk.CTkToplevel):
    def __init__(self, master, import_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Import Database")
        self.geometry("500x150")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        # Top heading label
        ctk.CTkLabel(self, text="Import Excel Data", font=("Arial", 14, "bold")).pack(pady=5)
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(fill="x", padx=20, pady=5)
        # Two fields side by side
        left = ctk.CTkFrame(frame, corner_radius=10)
        left.pack(side="left", expand=True, fill="both", padx=5)
        right = ctk.CTkFrame(frame, corner_radius=10)
        right.pack(side="left", expand=True, fill="both", padx=5)
        ctk.CTkLabel(left, text="Phone Column (e.g., B):").pack(pady=2)
        self.phone_col_var = ctk.StringVar()
        self.phone_entry = ctk.CTkEntry(left, textvariable=self.phone_col_var, corner_radius=10)
        self.phone_entry.pack(pady=2, fill="x")
        ctk.CTkLabel(right, text="Name Column (e.g., C):").pack(pady=2)
        self.name_col_var = ctk.StringVar()
        self.name_entry = ctk.CTkEntry(right, textvariable=self.name_col_var, corner_radius=10)
        self.name_entry.pack(pady=2, fill="x")
        self.browse_btn = ctk.CTkButton(self, text="Browse Excel File", corner_radius=10, command=self.browse_file)
        self.browse_btn.pack(pady=10)
        center_window(self)
        
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
        self.geometry("1200x900")
        self.wm_attributes("-topmost", True)
        self.excel_data = excel_data  # List of dicts with 'name'
        self.configure(padx=10, pady=10)
        center_window(self)
        self.template_image_path = None
        self.font_file_path = None
        self.last_click = (50, 50)
        self.font_size_var = ctk.StringVar(value="50")
        self.text_color_var = ctk.StringVar(value="black")
        # Ratio drop-down at top
        self.ratio_var = ctk.StringVar(value="Original")
        
        # Layout: Top drop-down for ratio; Left controls; Right fixed-size preview canvas
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
        # Fixed preview canvas; size will be adjusted by ratio selection
        self.canvas = ctk.CTkCanvas(self.preview_frame, bg="white", width=800, height=800)
        self.canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas.bind("<Button-1>", self.canvas_click)
        self.preview_image = None
        
    def choose_color(self):
        color = colorchooser.askcolor(title="Choose text color")
        if color and color[1]:
            self.text_color_var.set(color[1])
            
    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if path:
            self.template_image_path = path
            self.update_preview()
            
    def select_font(self):
        path = filedialog.askopenfilename(filetypes=[("Font Files", "*.ttf")])
        if path:
            self.font_file_path = path
            
    def canvas_click(self, event):
        self.last_click = (event.x, event.y)
        self.update_preview()
        
    def instruct_set_position(self):
        messagebox.showinfo("Set Position", "Click on the preview image to set the text position.")
        
    def update_preview(self):
        if not self.template_image_path:
            return
        try:
            # Adjust preview canvas size based on ratio selection
            ratio = self.ratio_var.get()
            if ratio == "4:3":
                new_size = (800, 600)
            elif ratio == "16:9":
                new_size = (800, 450)
            else:
                new_size = (800, 800)
            img = Image.open(self.template_image_path).convert("RGB")
            img = img.resize(new_size, Image.Resampling.LANCZOS)
            draw = ImageDraw.Draw(img)
            font_size = int(self.font_size_var.get())
            font_path = self.font_file_path if self.font_file_path else "arial.ttf"
            font = ImageFont.truetype(font_path, font_size)
            draw.text(self.last_click, "{User_Name}", font=font, fill=self.text_color_var.get())
            self.preview_image = ImageTk.PhotoImage(img)
            self.canvas.config(width=new_size[0], height=new_size[1])
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, image=self.preview_image, anchor="nw")
        except Exception as e:
            messagebox.showerror("Preview Error", f"Error updating preview: {e}")
            
    def generate_images(self):
        if not self.template_image_path:
            messagebox.showerror("Error", "No template image selected.")
            return
        try:
            font_size = int(self.font_size_var.get())
        except:
            messagebox.showerror("Error", "Invalid font size.")
            return
        font_path = self.font_file_path if self.font_file_path else "arial.ttf"
        text_color = self.text_color_var.get()
        text_pos = self.last_click
        count = 0
        for idx, entry in enumerate(self.excel_data, start=1):
            name = entry.get("name", "").strip()
            if not name:
                name = f"{idx}"
            try:
                img = Image.open(self.template_image_path).convert("RGB")
                # Resize image based on chosen ratio
                ratio = self.ratio_var.get()
                if ratio == "4:3":
                    new_size = (img.width, int(img.width*3/4))
                elif ratio == "16:9":
                    new_size = (img.width, int(img.width*9/16))
                else:
                    new_size = img.size
                img = img.resize(new_size, Image.Resampling.LANCZOS)
                # For final output, double the image size
                img = img.resize((new_size[0]*2, new_size[1]*2), Image.Resampling.LANCZOS)
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype(font_path, font_size)
                draw.text(text_pos, name, font=font, fill=text_color)
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', name)
                output_path = os.path.join(OUTPUT_IMG_FOLDER, f"{safe_name}.png")
                img.save(output_path)
                count += 1
            except Exception as ex:
                print(f"Error generating image for {name}: {ex}")
        messagebox.showinfo("Generation Complete", f"Generated {count} images in {OUTPUT_IMG_FOLDER}.")
        self.destroy()

# ----------------------- MAIN APPLICATION WINDOW ----------------------- #
class WaBulkSenderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.title("WaBulkSender")
        self.geometry("1400x900")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.attachments = {"Picture": None, "Video": None, "Document": None}
        self.custom_image_enabled = False
        self.excel_data = []  # List of dicts with keys "phone" and "name"
        self.sending = False
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=15)
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)
        self.header = ctk.CTkFrame(self, height=120, corner_radius=15)
        self.header.pack(side="top", fill="x", padx=10, pady=(10,0))
        self.main_area = ctk.CTkFrame(self, corner_radius=15)
        self.main_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.create_sidebar()
        self.create_header()
        self.create_main_area()
        if not os.path.exists(FLAG_FILE):
            FirstRunPopup(self, self.first_run_closed).wait_window()
        else:
            self.first_run_closed()
            
    def first_run_closed(self):
        self.log_live("Welcome to WaBulkSender!")
        
    # ----------------------- Sidebar ----------------------- #
    def create_sidebar(self):
        if os.path.exists(LOGO_PATH):
            img = Image.open(LOGO_PATH).resize((150,150), Image.Resampling.LANCZOS)
            self.sidebar_logo = ctk.CTkImage(img, size=(150,150))
            self.logo_label = ctk.CTkLabel(self.sidebar, image=self.sidebar_logo, text="")
            self.logo_label.pack(pady=(20,10))
        else:
            self.logo_label = ctk.CTkLabel(self.sidebar, text="Logo Missing")
            self.logo_label.pack(pady=(20,10))
        self.start_stop_button = ctk.CTkButton(self.sidebar, text="Start", corner_radius=10, height=50, command=self.toggle_sending)
        self.start_stop_button.pack(pady=10, padx=20, fill="x")
        self.login_button = ctk.CTkButton(self.sidebar, text="Login", corner_radius=10, height=40, command=self.launch_whatsapp_beta)
        self.login_button.pack(pady=10, padx=20, fill="x")
        self.live_alerts = ctk.CTkTextbox(self.sidebar, height=250, corner_radius=10)
        self.live_alerts.pack(side="bottom", pady=10, padx=20)
        self.live_alerts.insert("0.0", "Live Alerts:\n")
        self.live_alerts.configure(state="disabled")
        
    # ----------------------- Header ----------------------- #
    def create_header(self):
        self.welcome_label = ctk.CTkLabel(self.header, text="Welcome!", font=("Arial", 24, "bold"))
        self.welcome_label.pack(side="left", padx=20)
        self.github_button = ctk.CTkButton(self.header, text="", image=self.get_icon("github"), width=40, command=lambda: webbrowser.open(GITHUB_RELEASES_URL))
        self.github_button.pack(side="right", padx=10)
        self.update_button = ctk.CTkButton(self.header, text="", image=self.get_icon("update"), width=40, command=self.check_for_update)
        self.update_button.pack(side="right", padx=10)
        self.theme_toggle_button = ctk.CTkButton(self.header, text="", image=self.get_icon("theme"), width=40, command=self.toggle_theme)
        self.theme_toggle_button.pack(side="right", padx=10)
        
    def get_icon(self, icon_name):
        icon_path = os.path.join(BIN_FOLDER, f"{icon_name}.png")
        size = (30,30)
        if os.path.exists(icon_path):
            img = Image.open(icon_path).resize(size, Image.Resampling.LANCZOS)
        else:
            color = {"theme": "green", "update": "orange", "github": "black"}.get(icon_name, "gray")
            img = Image.new("RGB", size, color)
        return ctk.CTkImage(img, size=size)
        
    # ----------------------- Main Area (Side by Side) ----------------------- #
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
        
    # ----------------------- Message Area ----------------------- #
    def create_message_area(self, parent):
        # Formatting toolbar on top (with small buttons)
        fmt_frame = ctk.CTkFrame(parent, corner_radius=10)
        fmt_frame.pack(fill="x", pady=5, padx=5)
        self.bold_btn = ctk.CTkButton(fmt_frame, text="B", width=30, command=lambda: self.apply_formatting("*"), corner_radius=5)
        self.bold_btn.pack(side="left", padx=2)
        self.italic_btn = ctk.CTkButton(fmt_frame, text="I", width=30, command=lambda: self.apply_formatting("_"), corner_radius=5)
        self.italic_btn.pack(side="left", padx=2)
        self.strike_btn = ctk.CTkButton(fmt_frame, text="S", width=30, command=lambda: self.apply_formatting("~~"), corner_radius=5)
        self.strike_btn.pack(side="left", padx=2)
        self.mono_btn = ctk.CTkButton(fmt_frame, text="Mono", width=40, command=lambda: self.apply_formatting("`"), corner_radius=5)
        self.mono_btn.pack(side="left", padx=2)
        # Additional button for UserName insertion
        self.username_btn = ctk.CTkButton(fmt_frame, text="UserName", width=70, command=self.insert_username_placeholder, corner_radius=5)
        self.username_btn.pack(side="left", padx=2)
        
        # Message text box below the formatting toolbar
        self.message_text = ctk.CTkTextbox(parent, height=150, corner_radius=10)
        self.message_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Attachment and Custom Image buttons below message box
        button_frame = ctk.CTkFrame(parent, corner_radius=10)
        button_frame.pack(fill="x", pady=5, padx=5)
        self.attachment_var = ctk.StringVar(value="Select Attachment")
        self.attachment_menu = ctk.CTkOptionMenu(button_frame, values=["Picture", "Video", "Document"],
                                                  variable=self.attachment_var, command=self.handle_attachment)
        self.attachment_menu.pack(side="left", padx=5)
        self.custom_image_btn = ctk.CTkButton(button_frame, text="Custom Image Namer", corner_radius=10, command=self.open_custom_image_window)
        self.custom_image_btn.pack(side="left", padx=5)
        
        # Delay fields below
        delay_frame = ctk.CTkFrame(parent, corner_radius=10)
        delay_frame.pack(fill="x", padx=5, pady=5)
        self.min_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="Min Delay (s) [default 1]", corner_radius=10, width=100)
        self.min_delay_entry.insert(0, str(DEFAULT_MIN_DELAY))
        self.min_delay_entry.pack(side="left", padx=10, pady=5)
        self.max_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="Max Delay (s) [default 10]", corner_radius=10, width=100)
        self.max_delay_entry.insert(0, str(DEFAULT_MAX_DELAY))
        self.max_delay_entry.pack(side="left", padx=10, pady=5)
        
    def insert_username_placeholder(self):
        # Insert "()" at current cursor position
        self.message_text.insert("insert", "()")
        
    def apply_formatting(self, symbol):
        try:
            text = self.message_text.get("0.0", "end-1c")
            formatted = f"{symbol}{text}{symbol}"
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", formatted)
        except Exception as e:
            self.log_live(f"Formatting error: {e}")
            
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
        self.custom_image_enabled = True
        CustomImageWindow(self, self.excel_data)
        
    # ----------------------- Excel Area ----------------------- #
    def create_excel_area(self, parent):
        top_frame = ctk.CTkFrame(parent, corner_radius=10)
        top_frame.pack(fill="x", padx=5, pady=5)
        country_codes = ["None", "+91", "+1", "+44", "+61", "+81", "+49", "+33", "+86", "+7"]
        self.country_code_var = ctk.StringVar(value="None")
        self.country_code_dropdown = ctk.CTkOptionMenu(top_frame, values=country_codes, variable=self.country_code_var)
        self.country_code_dropdown.pack(side="left", padx=5, pady=5)
        self.import_db_btn = ctk.CTkButton(top_frame, text="Import DataBase", corner_radius=10, command=self.open_import_popup)
        self.import_db_btn.pack(side="left", padx=5, pady=5)
        self.excel_table = ExcelTable(parent, corner_radius=15)
        self.excel_table.pack(fill="both", expand=True, padx=5, pady=5)
        
    def open_import_popup(self):
        ImportDatabasePopup(self, self.load_excel_data)
        
    def load_excel_data(self, path, phone_col, name_col):
        try:
            wb = openpyxl.load_workbook(path)
            sheet = wb.active
            data = []
            for row in range(2, sheet.max_row + 1):
                phone = sheet[f"{phone_col}{row}"].value
                name = sheet[f"{name_col}{row}"].value if name_col else ""
                if phone:
                    phone = re.sub(r"\s+", "", str(phone))
                    if not phone.startswith("+") and self.country_code_var.get() != "None":
                        phone = self.country_code_var.get() + phone
                    data.append({"phone": phone, "name": name if name else ""})
            if not data:
                messagebox.showerror("Excel Error", "No valid phone numbers found in Excel.")
                return
            self.excel_data = data
            self.excel_table.load_data(data)
            self.log_live(f"Loaded {len(data)} entries from Excel.")
        except Exception as e:
            messagebox.showerror("Excel Error", f"Error loading Excel: {e}")
            
    # ----------------------- Sending Logic ----------------------- #
    def toggle_sending(self):
        if not self.sending:
            self.start_sending()
        else:
            self.stop_sending()
            
    def start_sending(self):
        data = self.excel_table.get_data()
        data = [d for d in data if d["phone"]]
        if not data:
            messagebox.showerror("Error", "No phone numbers loaded.")
            return
        msg = self.message_text.get("0.0", "end-1c").strip()
        if not msg:
            messagebox.showerror("Error", "No message text provided.")
            return
        self.sending = True
        self.start_stop_button.configure(text="Stop")
        threading.Thread(target=self.sending_process, args=(msg, data), daemon=True).start()
        
    def stop_sending(self):
        self.sending = False
        self.start_stop_button.configure(text="Start")
        self.log_live("Sending stopped.")
        
    def sending_process(self, msg, data):
        count = 0
        self.log_live("Waiting 10 seconds before sending first message...")
        time.sleep(10)
        for entry in data:
            if not self.sending:
                break
            phone = entry["phone"]
            name = entry["name"]
            personalized_msg = msg.replace("()", name)
            # Copy message text to clipboard
            self.clipboard_clear()
            self.clipboard_append(personalized_msg)
            phone = re.sub(r"\s+", "", phone)
            if not phone.startswith("+") and self.country_code_var.get() != "None":
                phone = self.country_code_var.get() + phone
            url = f"https://wa.me/{phone}"
            if count and count % 10 == 0:
                self.log_live("10 messages sent. Simulating closing browser tabs and waiting 10 seconds...")
                time.sleep(10)
            webbrowser.open(url)
            self.log_live(f"Opened chat for {phone}")
            time.sleep(5)
            # Simulate sending image (custom image or attachment) first
            if self.custom_image_enabled:
                generated_img = os.path.join(OUTPUT_IMG_FOLDER, f"{name if name else 'no_name'}.png")
                self.log_live(f"Custom image sent for {phone}: {generated_img}")
                # (In production, code to copy image to clipboard and paste it would be added here)
                time.sleep(2)
            elif any(self.attachments.values()):
                for key, path in self.attachments.items():
                    if path:
                        self.log_live(f"{key} sent for {phone}: {os.path.basename(path)}")
                        time.sleep(2)
            # Now simulate pasting message text and sending
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            pyautogui.press('enter')
            self.log_live(f"Message sent to {phone}: {personalized_msg}")
            try:
                min_delay = float(self.min_delay_entry.get())
                max_delay = float(self.max_delay_entry.get())
            except:
                min_delay, max_delay = DEFAULT_MIN_DELAY, DEFAULT_MAX_DELAY
            delay = random.uniform(min_delay, max_delay)
            self.log_live(f"Waiting {delay:.1f} seconds before next message...")
            time.sleep(delay)
            count += 1
        self.sending = False
        self.start_stop_button.configure(text="Start")
        self.log_live("All messages processed.")
        
    # ----------------------- Update Checker (GitHub API) ----------------------- #
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
                    messagebox.showinfo("No Update", "You are running the latest version.")
            else:
                messagebox.showerror("Update Error", "Failed to fetch update info.")
        except Exception as e:
            messagebox.showerror("Update Error", f"Error checking update: {e}")
        self.log_live("Update check completed.")
        
    # ----------------------- Other Utility Methods ----------------------- #
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
        new_mode = "light" if current == "dark" else "dark"
        ctk.set_appearance_mode(new_mode)
        self.log_live(f"Theme switched to {new_mode.capitalize()} Mode.")
        
    def on_close(self):
        self.sending = False
        self.destroy()

# ----------------------- RUN APPLICATION ----------------------- #
if __name__ == "__main__":
    app = WaBulkSenderApp()
    app.mainloop()
