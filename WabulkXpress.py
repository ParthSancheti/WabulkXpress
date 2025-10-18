#!/usr/bin/env python
import os
import re
import google.generativeai as genai
import shutil
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
import urllib.parse
import logging
import sys
logger = logging.getLogger(__name__)
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv
# Load environment variables
load_dotenv()
# Set logging level for av module warnings
logging.getLogger('libav').setLevel(logging.ERROR)

# ----------------------- THEME COLORS -----------------------
# Pair 1: App Background (Root Window)
DARK_BG = "#000000"
LIGHT_BG = "#FFFFFF"
# Pair 2: Compartment Backgrounds (Sidebar, Header, Frames)
DARK_FG = "#111111"
LIGHT_FG = "#F5F5F5"
# Pair 3: Inner Elements (Text boxes, Excel cells/entries)
DARK_INNER = "#222222"
LIGHT_INNER = "#EAEAEA"
# ------------------------------------------------------------

# ----------------------- GLOBAL CONSTANTS & PATHS -----------------------
CURRENT_VERSION = "21"
GITHUB_API_URL = "https://api.github.com/repos/Parth-Sancheti-5/WabulkXpress/releases/latest"
GITHUB_RELEASES_URL = "https://github.com/Parth-Sancheti-5/WabulkXpress/"
FLAG_FILE = "first_run.flag"
BIN_FOLDER = os.path.join(os.getcwd(), "bin")
TITLE_ICON_PATH = os.path.join(BIN_FOLDER, "loco.ico")
LOGO_PATH = os.path.join(BIN_FOLDER, "Logo.png")
WHATSAPP_BETA = os.path.join(BIN_FOLDER, "WhatsApp_Beta.lnk")
OUTPUT_IMG_FOLDER = os.path.join(os.getcwd(), "output_img")
SESSION_DIR = os.path.join(os.getcwd(), "selenium_session")
DEFAULT_MIN_DELAY = 1
DEFAULT_MAX_DELAY = 10
VIDEO_PATH = os.path.join(BIN_FOLDER, "woi.mp4")
LOADING_GIF_PATH = os.path.join(BIN_FOLDER, "lod.gif")
AI_ICON_PATH = os.path.join(BIN_FOLDER, "ai_icon.png") # Added for AI Popup
# ----------------------- (NEW PATHS - Add these near your other global paths) -----------------------
# Schedule Popup Icons
UP_ARROW_LIGHT_PATH = os.path.join(BIN_FOLDER, "up_arrow.png")
UP_ARROW_DARK_PATH = os.path.join(BIN_FOLDER, "up_arrow_dark.png")
DOWN_ARROW_LIGHT_PATH = os.path.join(BIN_FOLDER, "down_arrow.png")
# ... (other paths)
UPDATE_CHECK_FILE = os.path.join(os.getcwd(), "update_check.dat")
# ...
DOWN_ARROW_DARK_PATH = os.path.join(BIN_FOLDER, "down_arrow_dark.png")
# -------------------------------------------------------------------------------------------------

if not os.path.exists(OUTPUT_IMG_FOLDER):
    os.makedirs(OUTPUT_IMG_FOLDER)
if not os.path.exists(SESSION_DIR):
    os.makedirs(SESSION_DIR)
# ----------------------- SELENIUM HELPER FUNCTIONS -----------------------
class CustomRoundedButton(ctk.CTkButton):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.configure(corner_radius=20, fg_color="green", font=("Arial", 18, "bold"))
        self.bind("<Enter>", self.on_hover)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonPress>", self.on_press)
        self.bind("<ButtonRelease>", self.on_release)
    def on_hover(self, event):
        self.configure(fg_color="#45A049")  # Lighter green on hover
    def on_leave(self, event):
        self.configure(fg_color="green")  # Default green
    def on_press(self, event):
        self.configure(fg_color="red")  # Red when pressed
    def on_release(self, event):
        self.configure(fg_color="green")  # Reset to green after release
class GuiLogger:
    def __init__(self, gui):
        self.gui = gui
    def log(self, message):
        self.gui.after(0, lambda: self.gui.log_live(message))
def normalize_phone(phone):
    phone = re.sub(r"[^\d+]", "", phone.strip())
    if not phone.startswith("+"):
        phone = "+91" + phone  # Default to +91 if no country code
    return phone
def selenium_login(gui_logger):
    options = webdriver.ChromeOptions()
    options.add_argument(f'--user-data-dir={SESSION_DIR}')
    options.add_argument('--profile-directory=Default')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--lang=en')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    gui_logger.log("Opening WhatsApp Web for login...")
    driver.get("https://web.whatsapp.com")
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@role="grid"]'))
        )
        gui_logger.log("✅ WhatsApp Web logged in successfully!")
    except Exception:
        gui_logger.log("❗️ QR code scan required. Please scan the QR code.")
        try:
            WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.XPATH, '//div[@role="grid"]'))
            )
            gui_logger.log("✅ WhatsApp Web logged in successfully!")
        except Exception:
            gui_logger.log("❌ Failed to log in. Please try again.")
    driver.quit()

def selenium_send_bulk(numbers, messages, attachments, min_delay, max_delay, gui_logger):
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import TimeoutException # Import TimeoutException

    options = webdriver.ChromeOptions()
    options.add_argument(f'--user-data-dir={SESSION_DIR}')
    options.add_argument('--profile-directory=Default')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--lang=en')
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    success, failure = 0, 0

    # --- XPATH SELECTORS ---
    INPUT_BOX_XPATH = '//div[@contenteditable="true"][@aria-placeholder="Type a message"]'
    # Selectors for the checkmark icons in the *last* message sent by the user
    # This targets messages with the 'message-out' class and looks for the status icon span within them.
    # We use 'last()' to target the most recent outgoing message.
    LAST_MESSAGE_STATUS_XPATH = (
        '(//div[contains(@class, "message-out")]//span[contains(@data-icon, "msg-check") '
        'or contains(@data-icon, "msg-dblcheck") '
        'or contains(@data-icon, "msg-dblcheck-ack")])[last()]'
    )
    # --- END XPATH SELECTORS ---

    for idx, number in enumerate(numbers):
        # Check if the stop button was pressed
        if not gui_logger.gui.sending:
            gui_logger.log("🛑 Stop signal received. Halting process...")
            break

        msg = messages[idx] if isinstance(messages, list) else messages
        attachment = attachments[idx] if attachments and idx < len(attachments) else None
        gui_logger.log(f"💬 [{idx+1}/{len(numbers)}] Sending to {number}...")

        # Open chat
        url = f"https://web.whatsapp.com/send?phone={number}"
        driver.get(url)

        # Wait for chat to load (main panel)
        try:
            WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.XPATH, '//div[@data-testid="conversation-panel-body"] | //div[@role="grid"]')) # Added alternative selector
            )
            time.sleep(random.uniform(1.5, 2.5)) # Slightly longer sleep after load
        except TimeoutException:
            gui_logger.log(f"❗️ Chat panel did not load for {number}. Skipping.")
            failure += 1
            continue
        except Exception as e:
             gui_logger.log(f"❗️ Error loading chat for {number}: {e}. Skipping.")
             failure += 1
             continue


        sent_something = False # Track if we *attempted* to send anything

        try:
            # --- 1. SEND ATTACHMENT ---
            if attachment and os.path.exists(attachment):
                try:
                    attach_btn = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="plus-rounded"] | //span[@data-icon="attach-menu-plus"]')) # Added alternative selector
                    )
                    attach_btn.click()
                    time.sleep(0.5) # Small pause after click

                    # Use a more specific selector for file input related to documents/media
                    file_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"] | //input[@type="file"]'))
                    )
                    file_input.send_keys(os.path.abspath(attachment))
                    time.sleep(0.5) # Pause after sending keys

                    # Wait for the preview screen's send button
                    send_btn = WebDriverWait(driver, 30).until( # Increased wait for potentially large files
                        EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
                    )
                    send_btn.click()
                    gui_logger.log(f"📎 Attachment sending initiated to {number}")
                    sent_something = True

                    # Wait for preview modal to close
                    WebDriverWait(driver, 15).until(EC.staleness_of(send_btn))
                    time.sleep(0.5) # Pause after modal closes

                except TimeoutException:
                    gui_logger.log(f"❌ Timed out waiting for attachment elements for {number}.")
                    # No failure increment here yet, maybe message will succeed
                except Exception as attach_err:
                     gui_logger.log(f"❌ Error sending attachment to {number}: {attach_err}")
                     # No failure increment here yet

            # --- 2. SEND MESSAGE ---
            if msg:
                try:
                    input_box = WebDriverWait(driver, 15).until( # Reduced wait slightly
                        EC.element_to_be_clickable((By.XPATH, INPUT_BOX_XPATH))
                    )
                    # Use JavaScript click as a fallback if normal click is intercepted
                    try:
                        input_box.click()
                    except Exception:
                         driver.execute_script("arguments[0].click();", input_box)

                    time.sleep(0.5) # Increased sleep

                    # Send message text
                    def _has_non_bmp(s: str) -> bool:
                        return any(ord(c) > 0xFFFF for c in s)

                    if _has_non_bmp(msg):
                        copy_text_to_clipboard(msg)
                        time.sleep(0.3)
                        input_box.send_keys(Keys.CONTROL, 'v')
                    else:
                        # Send char by char (more human-like)
                        for char in msg:
                            input_box.send_keys(char)
                            time.sleep(random.uniform(0.02, 0.08)) # Small random delay between chars

                    time.sleep(0.5) # Pause before sending
                    input_box.send_keys(Keys.ENTER)
                    gui_logger.log(f"✅ Message sending initiated to {number}")
                    sent_something = True

                except TimeoutException:
                    gui_logger.log(f"❌ Timed out waiting for message input box for {number}.")
                    # Fallback might still work if msg exists
                except Exception as msg_err:
                     gui_logger.log(f"❌ Error typing/sending message to {number}: {msg_err}")


        except Exception as outer_err:
            # Catch unexpected errors during the main send block
            gui_logger.log(f"❌ Unexpected error during send process for {number}: {outer_err}")


        # --- 5. WAIT FOR SEND CONFIRMATION (NEW) ---
        if sent_something:
            try:
                # Wait up to 30 seconds for the last outgoing message to show a checkmark
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, LAST_MESSAGE_STATUS_XPATH))
                )
                gui_logger.log(f"✔️ Send confirmed for {number}")
                success += 1 # Increment success *only* after confirmation
            except TimeoutException:
                gui_logger.log(f"⚠️ Send confirmation checkmark not found for {number} within 30s. Assuming failure for tally.")
                failure += 1 # Assume failure if no checkmark appears
            except Exception as check_err:
                gui_logger.log(f"⚠️ Error checking send status for {number}: {check_err}. Assuming failure for tally.")
                failure += 1
        elif not sent_something:
             # If we didn't even attempt to send anything (e.g., chat load failed earlier or errors prevented send attempts)
             # This case might already be covered by the initial chat load check, but added for clarity.
             # We already incremented failure earlier if chat load failed.
             # If send attempts failed but chat loaded, we log errors but don't double-count failure here.
             # Only increment failure if no send was attempted AND chat *did* load.
             if 'failure' not in locals() or failure == locals().get('initial_failure_count', 0): # Check if failure wasn't already counted
                failure +=1


        # --- 6. DELAY LOGIC ---
        if not gui_logger.gui.sending:
            gui_logger.log("🛑 Stop signal received. Halting process...")
            break

        random_delay = random.uniform(min_delay, max_delay)
        gui_logger.log(f"⏳ Waiting for {random_delay:.2f} seconds before next contact...")

        # Sleep incrementally
        for _ in range(int(random_delay)):
            if not gui_logger.gui.sending: break
            time.sleep(1)
        # Sleep for remaining fraction
        if gui_logger.gui.sending:
             time.sleep(random_delay - int(random_delay))


        # --- 7. 10 Contact Pause ---
        if (idx + 1) % 10 == 0 and (idx + 1) < len(numbers) and gui_logger.gui.sending:
            gui_logger.log(f"⏸️ 10 contacts processed. Pausing for 30 seconds...")
            for _ in range(30):
                if not gui_logger.gui.sending: break
                time.sleep(1)

    driver.quit()
    gui_logger.log(f"📊 Done: Success={success}, Failure={failure}")
    return success, failure

# ----------------------- UTILITY FUNCTIONS -----------------------
def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")
def copy_text_to_clipboard(text):
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"Clipboard error: {e}")
def copy_file_to_clipboard(file_path):
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
# ----------------------- CUSTOM WIDGET CLASSES -----------------------
class HoverHint(ctk.CTkToplevel):
    def __init__(self, widget, hint_text, image_path, *args, **kwargs):
        super().__init__(widget.master, *args, **kwargs)
        self.overrideredirect(True)
        self.geometry("300x120")
        self.configure(bg="transparent")
        
        mode = ctk.get_appearance_mode().lower()
        fg_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        
        self.frame = ctk.CTkFrame(self, corner_radius=12, fg_color=fg_color) # Theme
        self.frame.pack(expand=True, fill="both", padx=5, pady=5)
        self.text_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        self.text_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)
        self.hint_label = ctk.CTkLabel(self.text_frame, text=hint_text, anchor="w", justify="left", wraplength=150)
        self.hint_label.pack(expand=True, fill="both")
        try:
            img = Image.open(image_path)
            box_width = 300
            image_width = int(box_width * 0.4)
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
                            img = img.resize((w, h), Image.Resampling.LANCZOS)
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
class ProgressPopup(ctk.CTkToplevel):
    def __init__(self, parent, title, total):
        super().__init__(parent)
        self.transient(parent)
        self.title(title)
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("500x300")
        self.resizable(False, False)
        self.lift()
        self.attributes("-topmost", True)
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        fg_color = DARK_FG if mode == "dark" else LIGHT_FG
        
        self.configure(fg_color=bg_color) # Theme
        self.total = total
        self.current = 0
        frame = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color) # Theme
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
        
import sys
import os
import requests
import webbrowser
from PIL import Image, ImageTk, ImageDraw
import customtkinter as ctk
from tkinter import messagebox
# Assuming globals are defined elsewhere in the codebase
BIN_FOLDER = os.path.join(os.getcwd(), "bin")
ERROR_IMAGE_PATH = os.path.join(BIN_FOLDER, "alert_image.jpg")
PRODUCT_IMAGE_PATH = os.path.join(BIN_FOLDER, "product_key_image.png")
WHATSAPP_NUMBER = "+918007579299"
class AnimatedCTkButton(ctk.CTkButton):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<ButtonRelease-1>", self.on_release)
    def on_enter(self, event):
        pass  # Removed scale; rely on built-in hover
    def on_leave(self, event):
        pass  # Removed scale
    def on_press(self, event):
        self.configure(fg_color="#388E3C")  # Brief pulse
    def on_release(self, event):
        self.configure(fg_color=self.cget("fg_color"))
def round_corners(image_path, size=(200, 200), corner_radius=20):
    try:
        img = Image.open(image_path).resize(size, Image.Resampling.LANCZOS).convert("RGBA")
        mask = Image.new("L", size, 0)
        draw = ImageDraw.Draw(mask)
        # Draw rounded rectangle instead of full circle
        draw.rounded_rectangle((0, 0) + size, radius=corner_radius, fill=255)
        img.putalpha(mask)
        # Fix: Add size to CTkImage for proper scaling and display
        return ctk.CTkImage(light_image=img, dark_image=img, size=size)
    except Exception:
        return None
def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")
class AlertPopup(ctk.CTkToplevel):
    def __init__(self, master, msg, is_used=False, key=None):
        super().__init__(master)
        self.title("Alert")
        self.geometry("400x250")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        
        self.configure(fg_color=bg_color) # Theme
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda e: self.destroy())
        center_window(self)
        self.fade_in()
        self.container = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0) # Theme
        self.container.pack(fill="both", expand=True)
        # Alert Image - Bigger
        alert_img = round_corners(ERROR_IMAGE_PATH, (100, 100), corner_radius=15)
        if alert_img:
            img_label = ctk.CTkLabel(self.container, image=alert_img, text="")
            img_label.place(relx=0.5, rely=0.25, anchor="center")
        # Message
        label = ctk.CTkLabel(self.container, text=msg, font=("Arial", 16, "bold"), fg_color="transparent")
        label.place(relx=0.5, rely=0.5, anchor="center")
        # Buttons Frame
        buttons_frame = ctk.CTkFrame(self.container, fg_color="transparent", corner_radius=0)
        buttons_frame.place(relx=0.5, rely=0.8, anchor="center")
        if is_used:
            help_btn = AnimatedCTkButton(buttons_frame, text="Get Help", width=100, height=40, font=("Arial", 16, "bold"),
                                         fg_color="#4CAF50", hover_color="#45A049", corner_radius=15,
                                         command=lambda: self.open_help(key))
            help_btn.pack(side="left", padx=10)
        ok_btn = ctk.CTkButton(buttons_frame, text="OK", width=100, height=40, font=("Arial", 16, "bold"),
                               fg_color="#333333", hover_color="#555555", corner_radius=15,
                               command=self.destroy)
        ok_btn.pack(side="left", padx=10)
    def open_help(self, key):
        if key:
            message = f"Hi! Please revoke my previous product key—I’ve changed devices. Key: {key}. Thanks! 🙏"
        else:
            message = "Hi! Please revoke my previous product key—I’ve changed devices. Thanks! 🙏"
        encoded = message.replace(" ", "%20")
        webbrowser.open(f"https://wa.me/{WHATSAPP_NUMBER}?text={encoded}")
    def fade_in(self, alpha=0.0):
        if alpha < 1.0:
            self.attributes("-alpha", alpha)
            self.after(50, lambda: self.fade_in(alpha + 0.1))
        else:
            self.attributes("-alpha", 1.0)
class FirstRunPopup(ctk.CTkToplevel):
    def __init__(self, master, on_close_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.validated = False  # Flag to track successful validation
        self.title("Enter the Product Key")
        self.geometry("600x300")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        inner_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        
        self.configure(fg_color=bg_color) # Theme
        self.on_close_callback = on_close_callback
        self.protocol("WM_DELETE_WINDOW", self.terminate)
        self.bind("<Escape>", lambda e: self.terminate())
        center_window(self)
        self.fade_in()
        self.container = ctk.CTkFrame(self, fg_color="transparent", width=600, height=300, corner_radius=0) # Theme
        self.container.pack(fill="both", expand=True)
        # Title
        title = ctk.CTkLabel(self.container, text="Enter the Product Key", font=("Arial", 24, "bold"),
                             fg_color="transparent")
        title.pack(anchor="w", pady=(40, 20), padx=40)
        # Entry
        self.product_key_entry = ctk.CTkEntry(self.container, width=320, height=50,
                                              placeholder_text="K39U-33W3",
                                              font=("Arial", 16), corner_radius=15,
                                              fg_color=inner_color, # Theme
                                              border_width=0) # Theme
        self.product_key_entry.pack(anchor="w", pady=10, padx=40)
        self.product_key_entry.bind("<FocusIn>", self.clear_placeholder)
        self.product_key_entry.bind("<Return>", lambda e: self.register_product_key())
        # Buttons Row - Left aligned
        buttons_frame = ctk.CTkFrame(self.container, fg_color="transparent", corner_radius=0)
        buttons_frame.pack(anchor="w", pady=30, padx=40)
        self.get_key_button = AnimatedCTkButton(buttons_frame, text="Get Product Key",
                                                width=140, height=50, font=("Arial", 20, "bold"),
                                                fg_color="#4CAF50", hover_color="#45A049", corner_radius=15,
                                                command=self.get_product_key)
        self.get_key_button.pack(side="left", padx=(0, 20))
        self.register_button = AnimatedCTkButton(buttons_frame, text="Register",
                                                 width=140, height=50, font=("Arial", 20, "bold"),
                                                 fg_color="#4CAF50", hover_color="#45A049", corner_radius=15,
                                                 command=self.register_product_key)
        self.register_button.pack(side="left")
        # Image - Right side
        self.product_img = round_corners(PRODUCT_IMAGE_PATH, (210, 210), corner_radius=15)
        if self.product_img:
            self.product_image_label = ctk.CTkLabel(self.container, image=self.product_img, text="")
            self.product_image_label.place(relx=0.85, rely=0.5, anchor="center")
            self.product_image_label.bind("<Enter>", self.animate_image_zoom_in)
            self.product_image_label.bind("<Leave>", self.animate_image_zoom_out)
        else:
            # Fallback label
            fallback_label = ctk.CTkLabel(self.container, text="Image Missing", text_color="#FFFFFF")
            fallback_label.place(relx=0.85, rely=0.5, anchor="center")
    def clear_placeholder(self, event):
        if self.product_key_entry.get() == "Sample Product Key":
            self.product_key_entry.delete(0, "end")
    def animate_image_zoom_in(self, event):
        if hasattr(self, 'product_image_label') and self.product_img:
            # Recreate slightly larger image for zoom
            zoomed_img = round_corners(PRODUCT_IMAGE_PATH, (240, 240), corner_radius=15)
            if zoomed_img:
                self.after(50, lambda: self.product_image_label.configure(image=zoomed_img))
    def animate_image_zoom_out(self, event):
        if hasattr(self, 'product_image_label') and self.product_img:
            self.after(50, lambda: self.product_image_label.configure(image=self.product_img))
    def fade_in(self, alpha=0.0):
        if alpha < 1.0:
            self.attributes("-alpha", alpha)
            self.after(50, lambda: self.fade_in(alpha + 0.1))
        else:
            self.attributes("-alpha", 1.0)
    def terminate(self):
        if not self.validated:  # If not validated, terminate the entire app
            self.master.quit()  # Stop the mainloop
            sys.exit(0)  # Fully terminate the script
        self.destroy()  # Otherwise, just close the popup
    def get_product_key(self):
        message = "Hi! I would like to purchase WabulkXpress. What's the cost and how do I get a product key? 😊💸🚀"
        encoded = message.replace(" ", "%20")
        webbrowser.open(f"https://wa.me/{WHATSAPP_NUMBER}?text={encoded}")
    def validate_product_key(self, key):
        # Basic format check (fail fast)
        if not key or not re.match(r'^[A-Z0-9]{4}-[A-Z0-9]{4}$', key):
            AlertPopup(self, "The product key is invalid.")
            return False
        
        # Encode key for URL
        encoded_key = urllib.parse.quote(key, safe='')
        
        try:
            # Use the correct SCRIPT_ID
            SCRIPT_ID = "AKfycby80FXj-xqDbxEpVcQue77HI0LwIHSiHdjM4bjTs_66_r7zRrnPlLwGZVQBZQTjmZWNVA"
            url = f"https://script.google.com/macros/s/{SCRIPT_ID}/exec?key={encoded_key}"
            logger.info(f"Validating key '{key}' at URL: {url}")  # Debug log
            
            response = requests.get(url, timeout=30)  # Increased timeout for GAS cold starts
            response.raise_for_status()
            logger.info(f"Response status: {response.status_code}, content-type: {response.headers.get('content-type')}")  # Debug
            
            data = response.json()
            status = data.get("status")
            used = data.get("used", 1)
            message = data.get("message", "")  # From remade code
            
            logger.info(f"API response: {data}")  # Full debug
            
            if status == "valid":
                if used == 0:
                    # Mark as used
                    use_url = f"{url}&action=use"
                    use_response = requests.get(use_url, timeout=30)
                    use_response.raise_for_status()
                    use_data = use_response.json()
                    if use_data.get("status") == "success":
                        logger.info(f"Key '{key}' marked as used.")
                        return True
                    else:
                        logger.warning(f"Failed to mark key '{key}' as used: {use_data}")
                        AlertPopup(self, f"Server error: {use_data.get('message', 'try later.')}")
                        return False
                else:
                    AlertPopup(self, "This key has already been used.", is_used=True, key=key)
                    return False
            else:
                AlertPopup(self, f"The product key is invalid: {message}")
                return False
        except requests.exceptions.Timeout:
            logger.error("Validation request timed out.")
            AlertPopup(self, "Request timed out. Check internet or try again.")
            return False
        except requests.exceptions.RequestException as e:
            logger.error(f"Network/HTTP error: {e}")
            AlertPopup(self, "Check connection and try again.")
            return False
        except ValueError as e:  # JSON decode error
            error_snippet = response.text[:200] if 'response' in locals() else "No response"
            logger.error(f"Invalid JSON: {e}. Snippet: {error_snippet}")
            AlertPopup(self, "Server returned invalid data. Contact support.")
            return False
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            AlertPopup(self, "Server error, try later.")
            return False
    def register_product_key(self):
        key = self.product_key_entry.get().strip()
        if not key or key == "Sample Product Key":
            AlertPopup(self, "Please enter a valid product key.")
            return
        
        # Disable button during processing to prevent multiple clicks
        self.register_button.configure(state="disabled")
        
        if self.validate_product_key(key):
            self.validated = True  # Set the flag on successful validation
            self.create_first_run_flag()  # Create the first run flag immediately when video starts
            self.show_introduction_video()
        
        # Re-enable button (in case of failure)
        self.register_button.configure(state="normal")
    def show_introduction_video(self):
        video_window = ctk.CTkToplevel(self)
        video_window.title("Welcome!")
        video_window.iconbitmap(TITLE_ICON_PATH)
        video_window.geometry("700x400")
        video_window.resizable(False, False)
        video_window.attributes("-topmost", True)
        video_window.transient(self)
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        
        video_window.configure(fg_color=bg_color) # Theme
        video_window.protocol("WM_DELETE_WINDOW", video_window.destroy)
        video_window.bind("<Escape>", lambda e: video_window.destroy())
        center_window(video_window)
        
        video_window.container = ctk.CTkFrame(video_window, width=700, height=400, fg_color="transparent") # Theme
        video_window.container.pack(fill="both", expand=True)
        # Ensure the video is placed in the right container
        if os.path.exists(VIDEO_PATH):
            video_window.video_player = TkinterVideo(video_window.container, VIDEO_PATH, scaled=True, keep_aspect=True)
            video_window.video_player.place(relx=0, rely=0, relwidth=1, relheight=1)
            video_window.video_player.play()
        else:
            video_window.instruction_label = ctk.CTkLabel(video_window.container, text="Video file not found", font=("Arial", 16))
            video_window.instruction_label.place(relx=0.5, rely=0.5, anchor="center")
        # Bottom frame for the checkbox and OK button
        video_window.bottom_frame = ctk.CTkFrame(video_window.container, fg_color="transparent")
        video_window.bottom_frame.place(relx=0.5, rely=0.95, anchor="center")
        # Add the checkbox and OK button
        video_window.dont_show_var = ctk.BooleanVar(value=False)
        video_window.checkbox = ctk.CTkCheckBox(video_window.bottom_frame, text="Don't show this again", 
                                                variable=video_window.dont_show_var, font=("Arial", 14, "bold"))
        video_window.checkbox.pack(side="left", padx=20, pady=10)
        video_window.ok_button = ctk.CTkButton(video_window.bottom_frame, text="OK", fg_color="#29AE07", 
                                                hover_color="#009e05", font=("Arial", 16, "bold"), width=100, height=50,
                                                corner_radius=15, command=lambda: self.close_video_popup(video_window))
        video_window.ok_button.pack(side="left", padx=20, pady=10)
        # Bind close event to handle checkbox
        video_window.protocol("WM_DELETE_WINDOW", lambda: self.close_video_popup(video_window))
    def close_video_popup(self, video_window):
        if hasattr(video_window, 'dont_show_var') and video_window.dont_show_var.get():
            # Optionally handle "don't show again" - perhaps write to a separate flag or modify FLAG_FILE
            pass
        if hasattr(video_window, 'video_player'):
            video_window.video_player.stop()
        self.on_video_complete(video_window)
    def on_video_complete(self, video_window=None):
        if video_window:
            video_window.destroy()
        self.on_close_callback()
        self.destroy()
    def create_first_run_flag(self):
        with open(FLAG_FILE, 'w') as f:
            f.write("activated")
class ExcelTable(ctk.CTkScrollableFrame):
    def __init__(self, master, main_app, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.main_app = main_app
        # Remove scrollbar_fg_color to enable auto-hiding
        # Set button colors for the modern, rounded knob
        self.configure(
            corner_radius=15, 
            scrollbar_button_color="#555555",       # Knob color for dark mode
            scrollbar_button_hover_color="#333333"  # Knob hover color for dark mode
        )
        
        self.rows = []
        self.custom_cols_active = [] # Tracks which custom columns are shown
        
        self.header_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=5, pady=3, ipady=5) # Added padding
        
        self.add_header() # Initial header
        self.prepopulate_rows(1) # Initial row

    def add_header(self):
        # Clear existing header
        for widget in self.header_frame.winfo_children():
            widget.destroy()

        # Define column widths
        col_widths = {
            "sno": 40,
            "phone": 150,
            "name": 150,
            "custom1": 100,
            "custom2": 100,
            "status": 60
        }
        
        # Calculate dynamic width for flexible columns (phone, name, custom)
        num_flex_cols = 2 + len(self.custom_cols_active)
        # Give status and sno fixed width, divide rest
        
        header_font = ("Arial", 16, "bold") # Bigger font

        ctk.CTkLabel(self.header_frame, text="S.No.", width=col_widths["sno"], anchor="center", font=header_font).pack(side="left", padx=5)
        
        phone_label = ctk.CTkLabel(self.header_frame, text="Phone Number", anchor="center", font=header_font)
        phone_label.pack(side="left", padx=5, fill="x", expand=True)
        
        name_label = ctk.CTkLabel(self.header_frame, text="Name", anchor="center", font=header_font)
        name_label.pack(side="left", padx=5, fill="x", expand=True)

        if "custom1" in self.custom_cols_active:
            c1_label = ctk.CTkLabel(self.header_frame, text="Custom 1", anchor="center", font=header_font)
            c1_label.pack(side="left", padx=5, fill="x", expand=True)

        if "custom2" in self.custom_cols_active:
            c2_label = ctk.CTkLabel(self.header_frame, text="Custom 2", anchor="center", font=header_font)
            c2_label.pack(side="left", padx=5, fill="x", expand=True)
            
        ctk.CTkLabel(self.header_frame, text="Status", width=col_widths["status"], anchor="center", font=header_font).pack(side="left", padx=5)

    def prepopulate_rows(self, count):
        for _ in range(count):
            self.add_row()

    def add_row(self, data=None):
        if data is None:
            data = {}
            
        row_frame = ctk.CTkFrame(self, corner_radius=10, fg_color="transparent")
        row_frame.pack(fill="x", padx=5, pady=3)
        
        sno_label = ctk.CTkLabel(row_frame, text=str(len(self.rows)+1), width=40, anchor="center")
        sno_label.pack(side="left", padx=5)
        
        mode = ctk.get_appearance_mode().lower()
        cell_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        
        row_dict = {
            "sno": sno_label,
            "indicator_state": 0, 
            "row_frame": row_frame
        }

        # --- Phone Entry ---
        phone_var = ctk.StringVar(value=data.get("phone", ""))
        phone_entry = ctk.CTkEntry(row_frame, textvariable=phone_var, placeholder_text="Enter number", corner_radius=10, fg_color=cell_color, border_width=0)
        phone_entry.pack(side="left", padx=5, fill="x", expand=True)
        phone_entry.bind("<Return>", lambda event, widget=phone_entry, var=phone_var: self.validate_phone(widget, var))
        phone_entry.bind("<KeyRelease>", self.check_add_row)
        row_dict["phone"] = phone_entry
        row_dict["phone_var"] = phone_var

        # --- Name Entry ---
        name_var = ctk.StringVar(value=data.get("name", ""))
        name_entry = ctk.CTkEntry(row_frame, textvariable=name_var, placeholder_text="Enter name", corner_radius=10, fg_color=cell_color, border_width=0)
        name_entry.pack(side="left", padx=5, fill="x", expand=True)
        name_entry.bind("<Return>", lambda event, var=name_var: self.validate_name(var))
        name_entry.bind("<KeyRelease>", self.check_add_row)
        row_dict["name"] = name_entry
        row_dict["name_var"] = name_var

        # --- Custom 1 Entry (Dynamic) ---
        if "custom1" in self.custom_cols_active:
            custom1_var = ctk.StringVar(value=data.get("custom1", ""))
            custom1_entry = ctk.CTkEntry(row_frame, textvariable=custom1_var, placeholder_text="Custom 1", corner_radius=10, fg_color=cell_color, border_width=0)
            custom1_entry.pack(side="left", padx=5, fill="x", expand=True)
            custom1_entry.bind("<KeyRelease>", self.check_add_row)
            row_dict["custom1"] = custom1_entry
            row_dict["custom1_var"] = custom1_var
            
        # --- Custom 2 Entry (Dynamic) ---
        if "custom2" in self.custom_cols_active:
            custom2_var = ctk.StringVar(value=data.get("custom2", ""))
            custom2_entry = ctk.CTkEntry(row_frame, textvariable=custom2_var, placeholder_text="Custom 2", corner_radius=10, fg_color=cell_color, border_width=0)
            custom2_entry.pack(side="left", padx=5, fill="x", expand=True)
            custom2_entry.bind("<KeyRelease>", self.check_add_row)
            row_dict["custom2"] = custom2_entry
            row_dict["custom2_var"] = custom2_var

        # --- Status Indicator ---
        indicator_bg = DARK_BG if mode == "dark" else LIGHT_BG # Match main BG for "grid" look
        indicator = tk.Canvas(row_frame, width=30, height=30, highlightthickness=0, bg=indicator_bg)
        indicator.create_oval(5, 5, 28, 28, fill="green", outline="")
        indicator.create_text(16, 17, text="✔", fill="white", font=("Arial", 12, "bold"))
        indicator.pack(side="left", padx=5)
        row_dict["indicator"] = indicator
        
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
            indicator.create_oval(5, 5, 28, 28, fill="red", outline="")
            indicator.create_text(16, 17, text="✖", fill="white", font=("Arial", 12, "bold"))
            row_dict["indicator_state"] = 1
            row_dict["skip"] = True
        elif state == 1:
            row_dict["row_frame"].destroy()
            self.rows.remove(row_dict)
            if not self.rows:
                self.add_row()
            self.update_row_numbers()

    def validate_phone(self, widget, var):
        # (This function is unchanged)
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
        
        # Check if any field in the last row has text
        is_filled = False
        if last_row["phone_var"].get().strip() or last_row["name_var"].get().strip():
            is_filled = True
        
        if "custom1_var" in last_row and last_row["custom1_var"].get().strip():
            is_filled = True
            
        if "custom2_var" in last_row and last_row["custom2_var"].get().strip():
            is_filled = True

        if is_filled:
            self.add_row()

    def load_data(self, data):
        # 1. Clear all existing widgets
        for child in self.winfo_children():
            if child != self.header_frame: # Keep header frame
                child.destroy()
        self.rows = []

        # 2. Determine active custom columns from the new data
        self.custom_cols_active = []
        if data: # Check if data is not empty
            if any(d.get("custom1") for d in data):
                self.custom_cols_active.append("custom1")
            if any(d.get("custom2") for d in data):
                self.custom_cols_active.append("custom2")
        
        # 3. Re-draw the header dynamically
        self.add_header()

        # 4. Load the new data into rows
        for entry in data:
            self.add_row(data=entry) # Pass data to add_row

        # 5. Add one final empty row
        self.add_row()
        
        # 6. Update S.No.
        self.update_row_numbers()

    def get_data(self):
        data = []
        for row in self.rows:
            phone = row["phone_var"].get().strip()
            name = row["name_var"].get().strip()
            
            # Check custom fields only if they exist in the row_dict
            custom1 = ""
            if "custom1_var" in row:
                custom1 = row["custom1_var"].get().strip()
                
            custom2 = ""
            if "custom2_var" in row:
                custom2 = row["custom2_var"].get().strip()

            # Skip if all fields are empty
            if not phone and not name and not custom1 and not custom2:
                continue
                
            entry = {"phone": phone, "name": name}
            if "custom1" in self.custom_cols_active:
                entry["custom1"] = custom1
            if "custom2" in self.custom_cols_active:
                entry["custom2"] = custom2
                
            if "skip" in row:
                entry["skip"] = True
            data.append(entry)
        return data

class ImportDatabasePopup(ctk.CTkToplevel):
    def __init__(self, master, import_callback, merge_mode=False, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        
        self.import_callback = import_callback
        self.merge_mode = merge_mode # 'merge' or 'replace'
        action_text = "Merge" if merge_mode else "Import"
        
        self.title(f"{action_text} Excel/CSV Data")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("600x300") # Made wider for new fields
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.transient(master)

        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        fg_color = DARK_FG if mode == "dark" else LIGHT_FG
        inner_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        
        self.configure(fg_color=bg_color)
        
        title_label = ctk.CTkLabel(self, text=f"{action_text} Excel/CSV Data", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        frame = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color)
        frame.pack(fill="x", padx=20, pady=5)
        
        # Configure grid
        # Configure grid
        frame.columnconfigure((0, 1), weight=1) # 2 columns
        frame.rowconfigure((0, 1), weight=1)    # 2 rows

        # --- Column Definitions ---
        label_font = ("Arial", 14)
        entry_font = ("Arial", 16)
        
        # Phone Column (Required)
        phone_frame = ctk.CTkFrame(frame, fg_color="transparent")
        phone_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(phone_frame, text="Phone Column* (e.g., A)", font=label_font).pack(pady=2, anchor="w")
        self.phone_col_var = ctk.StringVar()
        self.phone_entry = ctk.CTkEntry(phone_frame, textvariable=self.phone_col_var, font=entry_font, corner_radius=10, fg_color=inner_color, border_width=0, placeholder_text="A")
        self.phone_entry.pack(pady=2, fill="x")

        # Name Column (Optional)
        name_frame = ctk.CTkFrame(frame, fg_color="transparent")
        name_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(name_frame, text="Name Column (e.g., B)", font=label_font).pack(pady=2, anchor="w")
        self.name_col_var = ctk.StringVar()
        self.name_entry = ctk.CTkEntry(name_frame, textvariable=self.name_col_var, font=entry_font, corner_radius=10, fg_color=inner_color, border_width=0, placeholder_text="B")
        self.name_entry.pack(pady=2, fill="x")

        # Custom 1 Column (Optional)
        c1_frame = ctk.CTkFrame(frame, fg_color="transparent")
        c1_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(c1_frame, text="Custom 1 Col (Opt.)", font=label_font).pack(pady=2, anchor="w")
        self.custom1_col_var = ctk.StringVar()
        self.custom1_entry = ctk.CTkEntry(c1_frame, textvariable=self.custom1_col_var, font=entry_font, corner_radius=10, fg_color=inner_color, border_width=0, placeholder_text="C")
        self.custom1_entry.pack(pady=2, fill="x")

        # Custom 2 Column (Optional)
        c2_frame = ctk.CTkFrame(frame, fg_color="transparent")
        c2_frame.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(c2_frame, text="Custom 2 Col (Opt.)", font=label_font).pack(pady=2, anchor="w")
        self.custom2_col_var = ctk.StringVar()
        self.custom2_entry = ctk.CTkEntry(c2_frame, textvariable=self.custom2_col_var, font=entry_font, corner_radius=10, fg_color=inner_color, border_width=0, placeholder_text="D")
        self.custom2_entry.pack(pady=2, fill="x")
        
        # --- Browse Button ---
        self.browse_btn = ctk.CTkButton(
            self, 
            text=f"{action_text} & Browse File", 
            corner_radius=10, 
            state="disabled", 
            command=self.browse_file,
            font=("Arial", 16, "bold"),
            height=40,
            width=250,  # <-- Set a fixed smaller width
            fg_color="green",
            hover_color="#45A049"
        )
        self.browse_btn.pack(pady=20, padx=20) # <-- Removed fill="x"
        
        center_window(self)
        self.phone_entry.bind("<KeyRelease>", self.check_fields)
        self.check_fields() # Initial check
        
        # --- New Enter Key Bindings ---
        def browse_file_if_enabled(event=None):
            # Only run if the button is not disabled
            if self.browse_btn.cget("state") == "normal":
                self.browse_file()

        # Bind Enter key to all entry fields
        self.phone_entry.bind("<Return>", browse_file_if_enabled)
        self.name_entry.bind("<Return>", browse_file_if_enabled)
        self.custom1_entry.bind("<Return>", browse_file_if_enabled)
        self.custom2_entry.bind("<Return>", browse_file_if_enabled)
        
    def check_fields(self, event=None):
        # Only require Phone column to be filled
        if self.phone_col_var.get().strip():
            self.browse_btn.configure(state="normal")
        else:
            self.browse_btn.configure(state="disabled")
            
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.xlsx *.csv *.xls")])
        if path:
            prog = ProgressPopup(self, "Loading Data", total=1)
            # Pass all 4 column values + merge_mode flag
            threading.Thread(target=lambda: self.import_callback(
                path,
                self.phone_col_var.get().upper(),
                self.name_col_var.get().upper(),
                self.custom1_col_var.get().upper(),
                self.custom2_col_var.get().upper(),
                self.merge_mode 
            ), daemon=True).start()
            self.after(2000, prog.close)
            self.destroy()

class CustomImageWindow(ctk.CTkToplevel):
    def __init__(self, master, excel_data, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Custom Image Generator")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("900x460")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.excel_data = excel_data
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        fg_color = DARK_FG if mode == "dark" else LIGHT_FG
        inner_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        
        self.configure(padx=10, pady=10, fg_color=bg_color) # Theme
        center_window(self)
        self.template_image_path = None
        self.font_file_path = None
        self.last_click = (50, 50)
        self.font_size_var = ctk.StringVar(value="50")
        self.text_color_var = ctk.StringVar(value="black")
        self.ratio_options = ["Original", "4:3", "16:9", "5:8", "1:1", "3:2", "21:9"]
        self.ratio_var = ctk.StringVar(value="Original")
        top_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color) # Theme
        top_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(top_frame, text="Select Image Ratio:").pack(side="left", padx=5)
        self.ratio_menu = ctk.CTkOptionMenu(top_frame, values=self.ratio_options, variable=self.ratio_var, command=lambda x: self.update_preview())
        self.ratio_menu.pack(side="left", padx=5)
        self.control_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color) # Theme
        self.control_frame.pack(side="left", fill="y", padx=10, pady=10)
        self.preview_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color) # Theme
        self.preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        ctk.CTkLabel(self.control_frame, text="Custom Image Generator", font=("Arial", 16, "bold")).pack(pady=10)
        self.select_template_btn = ctk.CTkButton(self.control_frame, text="Select Template Image", corner_radius=10, command=self.select_template)
        self.select_template_btn.pack(pady=5, fill="x", padx=10)
        self.select_font_btn = ctk.CTkButton(self.control_frame, text="Select Font File", corner_radius=10, command=self.select_font)
        self.select_font_btn.pack(pady=5, fill="x", padx=10)
        ctk.CTkLabel(self.control_frame, text="Font Size:").pack(pady=5)
        self.font_size_entry = ctk.CTkEntry(self.control_frame, textvariable=self.font_size_var, corner_radius=10, fg_color=inner_color, border_width=0) # Theme
        self.font_size_entry.pack(pady=5, fill="x", padx=10)
        self.font_size_entry.bind("<KeyRelease>", lambda e: self.update_preview())
        self.font_size_entry.bind("<Return>", lambda e: self.generate_images_with_progress())
        ctk.CTkLabel(self.control_frame, text="Text Color:").pack(pady=5)
        color_btn = ctk.CTkButton(self.control_frame, text="Choose Color", corner_radius=10, command=self.choose_color)
        color_btn.pack(pady=5, fill="x", padx=10)
        self.set_position_btn = ctk.CTkButton(self.control_frame, text="Set Text Position\n(Click Preview)", corner_radius=10, command=self.instruct_set_position)
        self.set_position_btn.pack(pady=5, fill="x", padx=10)
        self.generate_btn = ctk.CTkButton(
            self.control_frame,
            text="Generate Images",
            fg_color="green",       # Set to green
            hover_color="#45A049",  # Add green hover color
            corner_radius=10,
            command=self.generate_images_with_progress
        )
        self.generate_btn.pack(pady=20, fill="x", padx=10)
        self.canvas = ctk.CTkCanvas(self.preview_frame, bg=inner_color, width=800, height=800, highlightthickness=0, bd=0) # Theme
        self.canvas.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas.bind("<Button-1>", self.canvas_click)
        # Bind Enter key to the whole window
        self.bind("<Return>", lambda e: self.generate_images_with_progress())
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
class SchedulePopup(ctk.CTkToplevel):
    def __init__(self, master, on_schedule_set, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Schedule Sending")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("450x355") # Increased size for bigger elements
        self.resizable(False, False)
        self.on_schedule_set = on_schedule_set
        self.wm_attributes("-topmost", True)
        self.transient(master) # Keep popup on top of main app

        # --- Theme Colors ---
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        fg_color = DARK_FG if mode == "dark" else LIGHT_FG
        inner_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        self.btn_hover_color = "#333333" if mode == "dark" else "#E0E0E0"

        # --- Load Icons (Theme Aware) ---
        self.ICON_SIZE = (50, 50)
        self.BTN_SIZE = (80, 80)
        self.ENTRY_FONT = ("Arial", 48, "bold")
        self.LABEL_FONT = ("Arial", 16, "bold")
        
        try:
            # Load both light and dark versions
            up_img_light = Image.open(UP_ARROW_LIGHT_PATH).resize(self.ICON_SIZE, Image.Resampling.LANCZOS)
            up_img_dark = Image.open(UP_ARROW_DARK_PATH).resize(self.ICON_SIZE, Image.Resampling.LANCZOS)
            down_img_light = Image.open(DOWN_ARROW_LIGHT_PATH).resize(self.ICON_SIZE, Image.Resampling.LANCZOS)
            down_img_dark = Image.open(DOWN_ARROW_DARK_PATH).resize(self.ICON_SIZE, Image.Resampling.LANCZOS)

            # CTkImage will handle switching automatically
            self.up_icon = ctk.CTkImage(light_image=up_img_light, dark_image=up_img_dark, size=self.ICON_SIZE)
            self.down_icon = ctk.CTkImage(light_image=down_img_light, dark_image=down_img_dark, size=self.ICON_SIZE)
            
        except Exception as e:
            print(f"Error loading schedule icons: {e}")
            self.up_icon = None # Fallback
            self.down_icon = None # Fallback

        # --- Data Variables ---
        now = datetime.now()
        # Set default values based on current time
        current_hour_12 = now.hour % 12 if now.hour % 12 != 0 else 12
        self.hour_var = tk.StringVar(value=f"{current_hour_12:02d}")
        self.min_var = tk.StringVar(value=f"{now.minute:02d}")
        self.ampm_var = tk.StringVar(value="AM" if now.hour < 12 else "PM")

        # --- Layout ---
        self.configure(fg_color=bg_color) 
        
        # Main container to center everything
        container = ctk.CTkFrame(self, corner_radius=10, fg_color=fg_color)
        container.pack(expand=True, fill="both", padx=20, pady=20)

        container.rowconfigure(0, weight=3) # Time pickers
        container.rowconfigure(1, weight=1) # Set button
        container.rowconfigure(2, weight=1) # Cancel button
        container.columnconfigure(0, weight=1) # Center column

        # Frame for the 3 time pickers
        time_frame = ctk.CTkFrame(container, corner_radius=10, fg_color="transparent")
        time_frame.grid(row=0, column=0, sticky="nsew", pady=10)
        time_frame.columnconfigure((0, 1, 2), weight=1) # 3 columns for H, M, A/P
        time_frame.rowconfigure(0, weight=0) # Label row
        time_frame.rowconfigure(1, weight=1) # Picker row

        # --- Hour Picker ---
        ctk.CTkLabel(time_frame, text="Hour", font=self.LABEL_FONT).grid(row=0, column=0)
        hour_frame = self.create_spinner_widget(time_frame, self.hour_var, self.increment_hour, self.decrement_hour, inner_color)
        hour_frame.grid(row=1, column=0, padx=5)
        
        # --- Minute Picker ---
        ctk.CTkLabel(time_frame, text="Minute", font=self.LABEL_FONT).grid(row=0, column=1)
        min_frame = self.create_spinner_widget(time_frame, self.min_var, self.increment_min, self.decrement_min, inner_color)
        min_frame.grid(row=1, column=1, padx=5)

        # --- AM/PM Picker ---
        ctk.CTkLabel(time_frame, text="Period", font=self.LABEL_FONT).grid(row=0, column=2)
        ampm_frame = self.create_spinner_widget(time_frame, self.ampm_var, self.toggle_ampm, self.toggle_ampm, inner_color)
        ampm_frame.grid(row=1, column=2, padx=5)

        # --- Big Action Buttons ---
        self.set_btn = ctk.CTkButton(
            container, 
            text="Set Schedule", 
            corner_radius=100, 
            font=("Arial", 24, "bold"), 
            command=self.set_schedule,
            height=60,
            width=250,
            fg_color="green",
            hover_color="#45A049"
        )
        self.set_btn.grid(row=1, column=0, pady=10)

        self.cancel_btn = ctk.CTkButton(
            container, 
            text="Cancel", 
            corner_radius=100, 
            font=("Arial", 18, "bold"), 
            command=self.destroy,
            height=50,
            width=200,
            fg_color="#555555", # Neutral gray
            hover_color="#777777"
        )
        self.cancel_btn.grid(row=2, column=0, pady=5)

        center_window(self)
        # Bind Enter key to the whole window
        self.bind("<Return>", lambda e: self.set_schedule())    

    def create_spinner_widget(self, parent, text_var, up_cmd, down_cmd, inner_color):
        """Helper to create the up/down button and entry widget."""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.rowconfigure((0, 2), weight=1) # Buttons
        frame.rowconfigure(1, weight=0) # Entry
        frame.columnconfigure(0, weight=1)

        # Up Button
        up_btn = ctk.CTkButton(
            frame, 
            text="", 
            image=self.up_icon, 
            width=self.BTN_SIZE[0], 
            height=self.BTN_SIZE[1], 
            fg_color="transparent", 
            hover_color=self.btn_hover_color,
            command=up_cmd
        )
        up_btn.grid(row=0, column=0, pady=5)

        # Entry Display
        entry = ctk.CTkEntry(
            frame, 
            width=120, 
            height=70, 
            corner_radius=10, 
            textvariable=text_var, 
            font=self.ENTRY_FONT, 
            fg_color=inner_color, 
            border_width=0,
            justify="center",
            state="readonly" # Use readonly state
        )
        entry.grid(row=1, column=0, pady=5)

        # Down Button
        down_btn = ctk.CTkButton(
            frame, 
            text="", 
            image=self.down_icon, 
            width=self.BTN_SIZE[0], 
            height=self.BTN_SIZE[1], 
            fg_color="transparent", 
            hover_color=self.btn_hover_color,
            command=down_cmd
        )
        down_btn.grid(row=2, column=0, pady=5)
        
        return frame

    def increment_hour(self):
        try:
            val = int(self.hour_var.get())
        except ValueError:
            val = 12
        val = 1 if val >= 12 else val + 1
        self.hour_var.set(f"{val:02d}") # Format to 2 digits

    def decrement_hour(self):
        try:
            val = int(self.hour_var.get())
        except ValueError:
            val = 1
        val = 12 if val <= 1 else val - 1
        self.hour_var.set(f"{val:02d}") # Format to 2 digits

    def increment_min(self):
        try:
            val = int(self.min_var.get())
        except ValueError:
            val = 0
        val = 0 if val >= 59 else val + 1
        self.min_var.set(f"{val:02d}") # Format to 2 digits

    def decrement_min(self):
        try:
            val = int(self.min_var.get())
        except ValueError:
            val = 0
        val = 59 if val <= 0 else val - 1
        self.min_var.set(f"{val:02d}") # Format to 2 digits

    def toggle_ampm(self):
        self.ampm_var.set("PM" if self.ampm_var.get().upper() == "AM" else "AM")

    def set_schedule(self):
        try:
            # Get 0-padded values
            hour_str = self.hour_var.get()
            min_str = self.min_var.get()
            
            # Validate format (should be 2 digits)
            if not (hour_str.isdigit() and (len(hour_str) == 2 or len(hour_str) == 1)): # Allow 1 or 2 digits
                 hour_str = f"{int(hour_str):02d}" # Re-format
            if not (min_str.isdigit() and (len(min_str) == 2 or len(min_str) == 1)):
                 min_str = f"{int(min_str):02d}" # Re-format
            
            # Final check
            if not (hour_str.isdigit() and len(hour_str) == 2):
                 raise ValueError("Hour format invalid.")
            if not (min_str.isdigit() and len(min_str) == 2):
                 raise ValueError("Minute format invalid.")

            hour = int(hour_str)
            minute = int(min_str)
            
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
            hour_24 = 0 # 12 AM is 00:00
        else:
            hour_24 = hour
            
        now = datetime.now()
        schedule_time = now.replace(hour=hour_24, minute=minute, second=0, microsecond=0)
        
        # If schedule time is in the past, set it for the next day
        if schedule_time <= now:
            schedule_time += timedelta(days=1)
            
        self.on_schedule_set(schedule_time)
        self.destroy()


class TranslatePopup(ctk.CTkToplevel):
    def __init__(self, master, process_callback, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.title("Translate Message")
        self.iconbitmap(TITLE_ICON_PATH)
        self.geometry("400x200")
        self.resizable(False, False)
        self.wm_attributes("-topmost", True)
        self.process_callback = process_callback
        
        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_BG if mode == "dark" else LIGHT_BG
        
        self.configure(fg_color=bg_color) # Theme
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

# --- (Make sure AI_ICON_PATH is defined globally near other paths) ---
AI_ICON_PATH = os.path.join(BIN_FOLDER, "ai_icon.png")

# --- (Replace the existing AIPopup class with this) ---
# --- (Make sure AI_ICON_PATH is defined globally near other paths) ---
AI_ICON_PATH = os.path.join(BIN_FOLDER, "ai_icon.png") 

# --- (Replace the existing AIPopup class with this) ---
class AIPopup(ctk.CTkToplevel):
    def __init__(self, master, button):
        super().__init__(master)
        self.master = master
        self.overrideredirect(True) # Frameless

        mode = ctk.get_appearance_mode().lower()
        bg_color = DARK_FG if mode == "dark" else LIGHT_FG
        cell_color = DARK_INNER if mode == "dark" else LIGHT_INNER
        cell_hover = "#333333" if mode == "dark" else "#E0E0E0"
        # Removed border_color as border_width is 0
        tk_cell_bg = cell_color # Get hex value before potentially changing cell_color on hover

        # Removed border_width=1
        self.frame = ctk.CTkFrame(self, corner_radius=15, fg_color=bg_color, border_width=0) 
        self.frame.pack(padx=5, pady=5, ipadx=10, ipady=10) # Internal padding for overall popup
        # Configure columns for button grid - uniform ensures equal width
        self.frame.columnconfigure((0, 1, 2), weight=1, uniform="cell_col") 
        # Configure rows containing cells - uniform ensures equal height
        self.frame.rowconfigure((1, 2), weight=1) # Rows for cells
        self.frame.rowconfigure(0, weight=0) # Row for header

        # --- Header: AI Icon (Left) + Text (Center/Right) ---
        header_frame = ctk.CTkFrame(self.frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=3, pady=(10, 15), sticky="ew") # Increased bottom padding
        header_frame.columnconfigure(0, weight=0) # Icon column
        header_frame.columnconfigure(1, weight=1) # Text column expands

        ai_popup_icon = None
        if os.path.exists(AI_ICON_PATH):
            try:
                # Icon size for header
                ai_img = Image.open(AI_ICON_PATH).resize((30,30), Image.Resampling.LANCZOS) 
                ai_popup_icon = ctk.CTkImage(ai_img, size=(30,30))
            except Exception as e:
                print(f"Error loading AI icon for popup: {e}")

        if ai_popup_icon:
            icon_label = ctk.CTkLabel(header_frame, image=ai_popup_icon, text="")
            icon_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="w") # Place icon left
        
        # Add "Writing Tools" text
        header_text = ctk.CTkLabel(header_frame, text="Writing Tools", font=("Arial", 18, "bold"), anchor="w")
        header_text.grid(row=0, column=1, padx=(5, 10), pady=5, sticky="ew") # Place text next to icon


        # --- Action Buttons (Cells - 2 rows, 3 columns) ---
        # Using CTkImage for consistency now
        # emoji_font_config = ("Arial", 32) # Backup if images fail
        text_font = ctk.CTkFont(size=14) # Text font size

        # Updated data with image paths
        buttons_data = [
            ("Reframe", os.path.join(BIN_FOLDER, "reframe_icon.png"), "Reframe"),
            ("Emoji", os.path.join(BIN_FOLDER, "emoji_icon.png"), "Emoji"),
            ("Professional", os.path.join(BIN_FOLDER, "professional_icon.png"), "Professional"),
            ("Funny", os.path.join(BIN_FOLDER, "funny_icon.png"), "Funny"),
            ("Ask AI", os.path.join(BIN_FOLDER, "ask_ai_icon.png"), "Ask AI"),
            ("Translate", os.path.join(BIN_FOLDER, "translate_icon.png"), "Translate")
        ]

        # Define fixed size for cells and image size
        cell_width = 110 
        cell_height = 110
        image_size = (40, 40) # Size for icons within cells

        self.cell_frames = [] # To manage hover state

        for i, (text, img_path, key) in enumerate(buttons_data):
            # Create a frame for each button to act as the cell with FIXED size
            cell_frame = ctk.CTkFrame(self.frame, fg_color=cell_color, corner_radius=10, width=cell_width, height=cell_height)
            cell_frame.grid(row=(i // 3) + 1, column=(i % 3), sticky="nsew", padx=5, pady=5) 
            cell_frame.configure(cursor="hand2") # Make frame look clickable
            
            cell_frame.grid_propagate(False) # Enforce fixed size

            self.cell_frames.append({"frame": cell_frame, "original_color": cell_color, "hover_color": cell_hover})

            # --- Content inside the cell frame ---
            content_container = ctk.CTkFrame(cell_frame, fg_color="transparent")
            content_container.place(relx=0.5, rely=0.5, anchor="center") 

            # Image Label (Top) - Load image using CTkImage
            cell_icon = None
            emoji_label = None # Define variable outside try block
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path).resize(image_size, Image.Resampling.LANCZOS)
                    cell_icon = ctk.CTkImage(img, size=image_size)
                    emoji_label = ctk.CTkLabel(content_container, image=cell_icon, text="", anchor="center")
                    emoji_label.pack(pady=(10, 5)) # Adjusted padding
                except Exception as e:
                    print(f"Error loading cell icon {img_path}: {e}")
                    # Fallback to text if image fails
                    emoji_label = ctk.CTkLabel(content_container, text="?", font=("Arial", 30), anchor="center")
                    emoji_label.pack(pady=(10, 5)) 
            else:
                 # Fallback to text if path doesn't exist
                 print(f"Cell icon not found: {img_path}")
                 emoji_label = ctk.CTkLabel(content_container, text="X", font=("Arial", 30), anchor="center")
                 emoji_label.pack(pady=(10, 5)) 

            # Text Label (Bottom)
            text_label = ctk.CTkLabel(content_container, text=text, font=text_font, anchor="center")
            text_label.pack(pady=(0, 10)) # Adjusted padding

            # --- Bindings ---
            widgets_to_bind = [cell_frame, content_container, emoji_label, text_label]
            for widget in widgets_to_bind:
                 # Ensure widget is not None before binding
                 if widget: 
                    widget.bind("<Button-1>", lambda e, k=key, popup=self: self.master.process_ai(k, popup=popup))
                    
                    # Apply hover effect using the stored colors and helper methods
                    widget.bind("<Enter>", lambda e, frm=cell_frame, hov=cell_hover: self.on_cell_enter(frm, hov))
                    widget.bind("<Leave>", lambda e, frm=cell_frame, orig=cell_color: self.on_cell_leave(frm, orig))

        # --- Position Popup ---
        self.update_idletasks() # Ensure widgets are drawn
        self.frame.update_idletasks() # Ensure frame size is calculated based on grid
        
        popup_width = self.frame.winfo_reqwidth() + 10 # Add border/padding estimate
        popup_height = self.frame.winfo_reqheight() + 10 # Add border/padding estimate

        btn_x = button.winfo_rootx()
        btn_y = button.winfo_rooty()
        btn_width = button.winfo_width()
        
        # Position popup DIRECTLY ABOVE the button, centered horizontally
        x = btn_x + (btn_width // 2) - (popup_width // 2) 
        y = btn_y - popup_height - 5 # Place 5px above the button

        self.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        
        self.lift()
        self.attributes("-topmost", True)
        self.bind("<FocusOut>", lambda e: self.destroy_popup()) # Close when focus lost
        self.bind("<Escape>", lambda e: self.destroy_popup()) # Also close if Escape key is pressed
        self.focus_set()

    # --- Helper methods for hover effects (Simplified) ---
    def on_cell_enter(self, frame_widget, hover_color):
        # Only change the frame background on hover
        frame_widget.configure(fg_color=hover_color)

    def on_cell_leave(self, frame_widget, original_color):
        # Restore the original frame background
        frame_widget.configure(fg_color=original_color)

    def destroy_popup(self):
        # Check if the window still exists before destroying
        if self.winfo_exists():
            self.destroy()
def generate_html_report(success, failure):
    total = success + failure
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>WabulkXpress Messaging Analytics</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: 'Arial', sans-serif; color: #e0e0e0; position: relative; min-height: 100vh; overflow: hidden; padding: 20px; }}
    body::before {{ content: ""; background: url("bin/bg.jpg") no-repeat center center fixed; background-size: cover; filter: blur(8px); position: absolute; top: 0; left: 0; right: 0; bottom: 0; z-index: -2; }}
    body::after {{ content: ""; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background-color: rgba(0, 0, 0, 0.6); z-index: -1; }}
    .container {{ display: flex; align-items: flex-start; justify-content: space-between; padding: 30px; gap: 20px; border-radius: 10px; background-color: rgba(30, 30, 30, 0.9); max-width: 1200px; margin: auto; box-shadow: 0 0 20px rgba(0,0,0,0.5); transition: transform 0.3s ease; }}
    .container:hover {{ transform: scale(1.02); }}
    .info {{ flex: 1; padding: 20px; background: rgba(0, 0, 0, 0.3); border-radius: 10px; margin-right: 20px; }}
    .info h2 {{ margin-bottom: 15px; }}
    .info p {{ margin-bottom: 10px; line-height: 1.5; }}
    .chart-container {{ flex: 1; position: relative; max-width: 400px; background: rgba(0, 0, 0, 0.3); padding: 15px; border-radius: 10px; box-shadow: 0 0 15px rgba(0,0,0,0.4); transition: box-shadow 0.3s ease; }}
    .chart-container:hover {{ box-shadow: 0 0 25px rgba(0,0,0,0.6); }}
    h1 {{ margin-bottom: 20px; text-align: center; font-size: 2em; }}
    button {{ padding: 10px 20px; border: none; border-radius: 5px; background-color: #0078D7; color: #fff; cursor: pointer; transition: background-color 0.3s ease, transform 0.3s ease; margin-top: 15px; }}
    button:hover {{ background-color: #005fa3; transform: scale(1.05); }}
    @media (max-width: 768px) {{ .container {{ flex-direction: column; }} .chart-container {{ margin-top: 20px; max-width: 100%; }} }}
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
    const total = {total};
    const success = {success};
    const failure = {failure};
    document.getElementById('totalCount').textContent = total;
    document.getElementById('successCount').textContent = success;
    document.getElementById('failureCount').textContent = failure;
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
                labels: {{ color: '#e0e0e0' }}
            }}
        }}
    }};
    new Chart(ctx, {{ type: 'doughnut', data: data, options: options }});
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
class WabulkXpressApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark") # Theme
        # ctk.set_default_color_theme("blue") # Removed per request
        self.configure(fg_color=DARK_BG) # Theme
        self.title("WabulkXpress")
        self.geometry("1400x900")
        if os.path.exists(TITLE_ICON_PATH):
            icon_img = Image.open(TITLE_ICON_PATH).resize((32,32), Image.Resampling.LANCZOS)
            icon_tk = ImageTk.PhotoImage(icon_img)
            self.wm_iconphoto(False, icon_tk)
        loco_icon = os.path.join(os.getcwd(), "bin", "loco.ico")
        if os.path.exists(loco_icon):
            self.iconbitmap(loco_icon)
        self.protocol("WM_DELETE_WINDOW", self.on_close) # Ensure on_close is referenced
        self.attachments = {"Picture": None, "Video": None, "Document": None, "Any": None}
        self.custom_image_enabled = False
        self.excel_data = []
        self.sending = False
        self.undo_stack = []
        self.redo_stack = []
        self.schedule_time = None
        self.first_cycle = True
        self.last_action = None
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=15, fg_color=DARK_FG) # Theme
        self.sidebar.pack(side="left", fill="y", padx=10, pady=10)
        self.header = ctk.CTkFrame(self, height=120, corner_radius=15, fg_color=DARK_FG) # Theme
        self.header.pack(side="top", fill="x", padx=10, pady=(10,0))
        self.main_area = ctk.CTkFrame(self, corner_radius=15, fg_color="transparent") # Theme
        self.main_area.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.create_sidebar()
        self.create_header()
        self.create_main_area()
        self.gemini_api_key = "AIzaSyDmYy3CFKb0aoVRYZANAyp6X3jgKUe__6g"  # Fallback (replace with actual key or handle gracefully)
        self.gemini_api_url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.gemini_api_key}"
        self.ai_prompts = {
            "Reframe": "Rephrase the following message in a single cohesive paragraph with no extra disclaimers or stars:",
            "Emoji": "Rewrite the following message with a few relevant emojis, keeping it concise:",
            "Professional": "Rewrite the following message in a polite, professional tone, no extra disclaimers or bullet points:",
            "Funny": "Rewrite the following message with a light, humorous style, no extra disclaimers or bullet points:",
            "Ask AI": "Ask AI: Please answer the following message without adding any extra formatting or stars:",
            "Translate": "Please translate the following message into {lang}, ensuring that you preserve the original formatting exactly no extra disclaimers or stars or Any othermessage from your side only the text translated:",
        }
        self.schedule_background_update_check()
        
        if not os.path.exists(FLAG_FILE):
            FirstRunPopup(self, self.first_run_closed).wait_window()
        else:
            self.first_run_closed()
        self.refresh_icons()
        
    def first_run_closed(self):
        self.log_live("Welcome to WabulkXpress!")

    def schedule_background_update_check(self):
        """Checks if 7 days have passed and schedules a background update check."""
        try:
            last_check_time = 0.0
            if os.path.exists(UPDATE_CHECK_FILE):
                with open(UPDATE_CHECK_FILE, 'r') as f:
                    content = f.read().strip()
                    if content:
                        last_check_time = float(content)

            # Check if 7 days (in seconds) have passed
            seven_days_ago = time.time() - (7 * 24 * 60 * 60)

            if last_check_time < seven_days_ago:
                self.log_live("Performing automatic update check (last check > 7 days ago)...")
                # Run the check in a separate thread to avoid blocking startup
                threading.Thread(target=self.check_for_update_background, daemon=True).start()
            else:
                last_check_date = datetime.fromtimestamp(last_check_time).strftime('%Y-%m-%d %H:%M')
                print(f"DEBUG: Skipping automatic update check. Last check was at {last_check_date}") # Console log

        except Exception as e:
            self.log_live(f"⚠️ Error scheduling background update check: {e}")
            print(f"DEBUG: Error reading update check file: {e}") # Console log

    def check_for_update_background(self):
        """Performs the update check in the background and updates the timestamp."""
        update_available = False
        latest_version_tag = "0"
        try:
            response = requests.get(GITHUB_API_URL, timeout=15)
            response.raise_for_status()
            latest_version_tag = response.json().get("tag_name", "0")

            if float(latest_version_tag) > float(CURRENT_VERSION):
                update_available = True
            else:
                self.log_live(f"Automatic check: You are using the latest version ({CURRENT_VERSION}).")

            # Update the last check time *after* a successful check
            with open(UPDATE_CHECK_FILE, 'w') as f:
                f.write(str(time.time()))
            print(f"DEBUG: Updated last update check timestamp to {datetime.now()}") # Console log

        except Exception as e:
            self.log_live(f"⚠️ Automatic update check failed: {e}")
            print(f"DEBUG: Background update check failed: {e}") # Console log
            return # Don't proceed if check failed

        # If an update is available, schedule the prompt on the main thread
        if update_available:
            self.log_live(f"Automatic check: Update available: {latest_version_tag} (Current: {CURRENT_VERSION})")
            # Use self.after to run the messagebox on the main GUI thread
            self.after(0, self.prompt_for_update, latest_version_tag)

    def prompt_for_update(self, latest_version_tag):
        """Shows the update prompt messagebox on the main thread."""
        answer = messagebox.askyesno(
            "Update Available",
            f"A new version ({latest_version_tag}) is available!\n"
            f"Current version: {CURRENT_VERSION}\n\n"
            "Would you like to visit the releases page to download it?"
        )
        if answer:
            webbrowser.open(GITHUB_RELEASES_URL)
            
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
            img = Image.open(LOGO_PATH).resize((200,200), Image.Resampling.LANCZOS)
            self.sidebar_logo = ctk.CTkImage(img, size=(200,200))
            self.logo_label = ctk.CTkLabel(self.sidebar, image=self.sidebar_logo, text="")
            self.logo_label.pack(pady=(20,10))
        else:
            self.logo_label = ctk.CTkLabel(self.sidebar, text="Logo Missing")
            self.logo_label.pack(pady=(20,20))
        
        center_frame = ctk.CTkFrame(self.sidebar, corner_radius=0, fg_color="transparent") # Theme
        center_frame.pack(expand=True, fill="both")
        center_frame.grid_columnconfigure(0, weight=1)
        center_frame.grid_columnconfigure(1, weight=1)
        
        # Start button
        self.start_stop_button = CustomRoundedButton(
            center_frame, text="Start", corner_radius=100, height=50, width=100, command=self.toggle_sending
        )
        self.start_stop_button.grid(row=1, column=0, padx=0, pady=20)
        # Schedule button with an arrow icon and integrated into one button
        arrow_path = os.path.join(BIN_FOLDER, "down_arrow_dark.png")
        down_arrow_icon = None
        if os.path.exists(arrow_path):
            arrow_img = Image.open(arrow_path).resize((25, 25), Image.Resampling.LANCZOS)
            down_arrow_icon = ctk.CTkImage(arrow_img, size=(25, 25))
        self.schedule_button = CustomRoundedButton(
            center_frame,
            text="",
            image=down_arrow_icon,
            corner_radius=20,
            height=50,
            width=50,
            command=self.open_schedule_popup
        )
        self.schedule_button.grid(row=1, column=1, padx=0, pady=2)  # Reduced padx to 0 for smaller gap between start and schedule
        # Login button with full-rounded corners, bigger font, and green default color
        self.login_button = CustomRoundedButton(
            center_frame, text="Login", corner_radius=100, height=50, width=120, command=self.launch_whatsapp_beta
        )
        self.login_button.grid(row=2, column=0, columnspan=2, padx=5, pady=(10, 5))  # Reduced pady for smaller gap to start/schedule buttons
        self.live_alerts = ctk.CTkTextbox(
            self.sidebar, 
            height=250, 
            corner_radius=10, 
            font=ctk.CTkFont(size=16, weight="normal"),  # Bigger font for live alerts and all text
            padx=10,  # Added padding
            pady=10,   # Added padding
            fg_color=DARK_INNER, # Theme
            border_width=0 # Theme
        )
        self.live_alerts.pack(side="bottom", pady=20, padx=10) # Added horizontal padding
        self.live_alerts.insert("0.0", "Live Alerts:\n")
        self.live_alerts.configure(state="disabled")
        
    def open_schedule_popup(self):
        SchedulePopup(self, self.set_schedule_time)
        
    def create_header(self):
        # Double the height of the header box (assuming default ~40, set to 80)
        self.header.configure(height=70)
        
        # Use grid for better vertical centering control
        self.header.grid_columnconfigure(0, weight=1)
        self.header.grid_columnconfigure(1, weight=0)
        self.header.grid_rowconfigure(0, weight=1)
        
        # Left side: Vertical stack of labels
        left_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nw", padx=20, pady=10)
        left_frame.grid_rowconfigure(0, weight=0)
        left_frame.grid_rowconfigure(1, weight=0)
        
        self.welcome_label = ctk.CTkLabel(
            left_frame, 
            text="Welcome!", 
            font=("Arial", 30, "bold")  # Bigger font
        )
        self.welcome_label.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.subtitle_label = ctk.CTkLabel(
            left_frame, 
            text="To WaBulkExpress", 
            font=("Arial", 16)  # Smaller font
        )
        self.subtitle_label.grid(row=1, column=0, sticky="w")
        
        # Right side: Buttons vertically centered and bigger
        right_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="ne", padx=10, pady=10)
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)
        right_frame.grid_rowconfigure(1, weight=0)
        right_frame.grid_columnconfigure(0, weight=0)
        right_frame.grid_columnconfigure(1, weight=0)
        right_frame.grid_columnconfigure(2, weight=0)
        
        self.theme_toggle_button = ctk.CTkButton(
            right_frame,
            text="",
            width=70,  # Bigger
            height=70,  # Bigger and taller for bigger icons and vertical centering
            command=self.toggle_theme,
        )
        self.theme_toggle_button.grid(row=1, column=0, padx=(0, 5), sticky="e")
        
        self.update_button = ctk.CTkButton(
            right_frame,
            text="",
            width=70,  # Bigger
            height=70,  # Bigger and taller
            command=self.check_for_update,
        )
        self.update_button.grid(row=1, column=1, padx=5, sticky="e")
        
        self.github_button = ctk.CTkButton(
            right_frame,
            text="",
            width=70,  # Bigger
            height=70,  # Bigger and taller
            command=lambda: webbrowser.open(GITHUB_RELEASES_URL),
        )
        self.github_button.grid(row=1, column=2, padx=(5, 0), sticky="e")
        
    def get_icon(self, icon_name):
        mode = ctk.get_appearance_mode().lower()
        if icon_name in ["github", "update"]:
            file_name = f"{icon_name}.png" if mode == "light" else f"{icon_name}_dark.png"
        elif icon_name == "dark":
            file_name = "dark.png" if mode == "light" else "light.png"
        else:
            file_name = f"{icon_name}.png"
        icon_path = os.path.join(os.getcwd(), "bin", file_name)
        size = (40,40)
        if os.path.exists(icon_path):
            img = Image.open(icon_path).resize(size, Image.Resampling.LANCZOS)
        else:
            img = Image.new("RGB", size, "#DBDBDB" if mode=="light" else "#2B2B2B")
        return ctk.CTkImage(img, size=size)
        
    def create_main_area(self):
        # Changed weight to 1:3 to give Excel area more space
        self.main_area.columnconfigure(0, weight=1) 
        self.main_area.columnconfigure(1, weight=3)
        
        self.message_frame = ctk.CTkFrame(self.main_area, corner_radius=15, fg_color=DARK_FG) # Theme
        self.message_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.create_message_area(self.message_frame)
        
        self.excel_frame = ctk.CTkFrame(self.main_area, corner_radius=15, fg_color=DARK_FG) # Theme
        self.excel_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.create_excel_area(self.excel_frame)
        
        self.main_area.rowconfigure(0, weight=1)

    def create_message_area(self, parent):
            # --- Top Buttons (Side-by-side) ---
            top_button_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="transparent")
            top_button_frame.pack(fill="x", pady=(10, 5), padx=5)
            top_button_frame.columnconfigure((0, 1), weight=1) # Make columns expand equally

            self.attachment_btn = ctk.CTkButton(
                top_button_frame,
                text="Select Attachment",
                corner_radius=100,
                height=40,
                command=self.handle_attachment,
                font=("Arial", 16, "bold"),
                fg_color="green",
                hover_color="#45A049"
            )
            self.attachment_btn.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="ew") # Use grid

            self.custom_image_btn = ctk.CTkButton(
                top_button_frame,
                text="Custom Image Namer",
                corner_radius=100,
                command=self.open_custom_image_window,
                height=40,
                font=("Arial", 16, "bold"),
                fg_color="green",
                hover_color="#45A049"
            )
            self.custom_image_btn.grid(row=0, column=1, padx=(5, 10), pady=5, sticky="ew") # Use grid
            HoverHint(self.custom_image_btn, "Automatically places the receiver’s name onto your custom image template — perfect for personalized visual messages!", os.path.join(os.getcwd(), "bin", "woi_ci.png"))

            # --- Formatting Buttons (Two Rows) ---
            fmt_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="transparent") # Theme
            fmt_frame.pack(fill="x", pady=5, padx=5)
            # Configure columns for spacing in the first row
            fmt_frame.columnconfigure((0, 1, 2, 3), weight=0) # Left buttons don't expand
            fmt_frame.columnconfigure(4, weight=1) # Spacer expands
            fmt_frame.columnconfigure((5, 6), weight=0) # Right buttons don't expand

            # Style dictionary for small formatting buttons (width removed)
            btn_style_small = {
                "fg_color": "green",
                "hover_color": "#45A049",
                "corner_radius": 100,
                "font": ("Arial", 12, "bold"), # Smaller font
                # "width": 35, # REMOVED THIS LINE
                "height": 35 # Smaller height
            }

            # Row 0: Formatting Left, Undo/Redo Right
            self.bold_btn = ctk.CTkButton(fmt_frame, text="B", command=lambda: self.apply_formatting("*"), **btn_style_small, width=35) # Specify width here
            self.bold_btn.grid(row=0, column=0, padx=2, pady=2)

            self.italic_btn = ctk.CTkButton(fmt_frame, text="I", command=lambda: self.apply_formatting("_"), **btn_style_small, width=35) # Specify width here
            self.italic_btn.grid(row=0, column=1, padx=2, pady=2)

            self.strike_btn = ctk.CTkButton(fmt_frame, text="S", command=lambda: self.apply_formatting("~"), **btn_style_small, width=35) # Specify width here
            self.strike_btn.grid(row=0, column=2, padx=2, pady=2)

            self.mono_btn = ctk.CTkButton(fmt_frame, text="Code", command=lambda: self.apply_formatting("```"), **btn_style_small, width=40) # Specify width here
            self.mono_btn.grid(row=0, column=3, padx=2, pady=2)

            self.undo_btn = ctk.CTkButton(fmt_frame, text="Undo", command=self.undo, **btn_style_small, width=50) # Specify width here
            self.undo_btn.grid(row=0, column=5, padx=2, pady=2)

            self.redo_btn = ctk.CTkButton(fmt_frame, text="Redo", command=self.redo, **btn_style_small, width=50) # Specify width here
            self.redo_btn.grid(row=0, column=6, padx=(2,10), pady=2) # Add right padding

            # Row 1: Placeholders, evenly spaced using grid directly on fmt_frame
            placeholder_btn_style = {
                "fg_color": "green",
                "hover_color": "#45A049",
                "corner_radius": 100,
                "font": ("Arial", 12, "bold"), # Smaller font
                "height": 35 # Smaller height
            }
            # Create a sub-frame for row 1 buttons to manage spacing better
            row1_fmt_frame = ctk.CTkFrame(fmt_frame, fg_color="transparent")
            row1_fmt_frame.grid(row=1, column=0, columnspan=7, sticky="ew", pady=(5,0))
            row1_fmt_frame.columnconfigure((0,1,2), weight=1) # Even spacing within this sub-frame

            self.username_btn = ctk.CTkButton(row1_fmt_frame, text="User", command=lambda: self.insert_placeholder("{User_Name}"), **placeholder_btn_style)
            self.username_btn.grid(row=0, column=0, padx=5, pady=0, sticky="ew")
            HoverHint(self.username_btn, "Inserts {User_Name} placeholder", os.path.join(os.getcwd(), "bin", "woi_un.png"))

            self.custom1_btn = ctk.CTkButton(row1_fmt_frame, text="Custom 1", command=lambda: self.insert_placeholder("{Custom1}"), **placeholder_btn_style)
            self.custom1_btn.grid(row=0, column=1, padx=5, pady=0, sticky="ew")
            HoverHint(self.custom1_btn, "Inserts {Custom1} placeholder (from optional import column)", os.path.join(os.getcwd(), "bin", "woi_un.png"))

            self.custom2_btn = ctk.CTkButton(row1_fmt_frame, text="Custom 2", command=lambda: self.insert_placeholder("{Custom2}"), **placeholder_btn_style)
            self.custom2_btn.grid(row=0, column=2, padx=5, pady=0, sticky="ew")
            HoverHint(self.custom2_btn, "Inserts {Custom2} placeholder (from optional import column)", os.path.join(os.getcwd(), "bin", "woi_un.png"))

            # --- Message Area ---
            ctk.CTkLabel(parent, text="Message:", font=("Arial", 18, "bold")).pack(anchor="w", padx=10, pady=(10, 5)) # Increased top padding

            self.text_area_frame = ctk.CTkFrame(
                parent, 
                corner_radius=20,  # This will give the frame rounded corners
                fg_color=DARK_INNER,  # Ensure the color is set to something suitable
                border_width=0  # No border around the frame
            )
            self.text_area_frame.pack(fill="both", expand=True, padx=5, pady=(0,5))  # Adding padding if needed

            self.message_text = ctk.CTkTextbox(
                self.text_area_frame,
                height=100,
                corner_radius=10, # Rounded corners inherent
                fg_color=DARK_INNER,
                border_width=0,
                padx=10, # Increased padding inside textbox
                pady=10, # Increased padding inside textbox
                font=("Arial", 16) # Bigger font
            ) # Theme
            self.message_text.pack(fill="both", expand=True, padx=0, pady=0)
            self.message_text.bind("<Control-z>", self.undo)
            self.message_text.bind("<Control-y>", self.redo)
            self.message_text.bind("<KeyRelease>", self.save_state)
            self.message_text.bind("<Button-3>", self.show_context_menu)

            self.context_menu = tk.Menu(self.message_text, tearoff=0)
            self.context_menu.add_command(label="Bold", command=lambda: self.apply_formatting("*"))
            self.context_menu.add_command(label="Italic", command=lambda: self.apply_formatting("_"))
            self.context_menu.add_command(label="Strike", command=lambda: self.apply_formatting("~"))
            self.context_menu.add_command(label="Code", command=lambda: self.apply_formatting("```"))
            self.context_menu.add_separator()
            self.context_menu.add_command(label="Insert User", command=lambda: self.insert_placeholder("{User_Name}"))
            self.context_menu.add_command(label="Insert Custom 1", command=lambda: self.insert_placeholder("{Custom1}"))
            self.context_menu.add_command(label="Insert Custom 2", command=lambda: self.insert_placeholder("{Custom2}"))
            self.context_menu.add_separator()
            self.context_menu.add_command(label="Undo", command=self.undo)
            self.context_menu.add_command(label="Redo", command=self.redo)

            # --- AI Button ---
            ai_icon = None
            if os.path.exists(AI_ICON_PATH):
                try:
                    ai_img = Image.open(AI_ICON_PATH).resize((25,25), Image.Resampling.LANCZOS)
                    ai_icon = ctk.CTkImage(ai_img, size=(25,25))
                except Exception as e:
                    print(f"Error loading AI icon for button: {e}")

            mode = ctk.get_appearance_mode().lower()
            # The button's background (fg_color) should match its parent (text_area_frame)
            ai_bg_color = DARK_INNER if mode == "dark" else LIGHT_INNER
            # The button's hover color should match the main frames
            ai_hover_color = DARK_FG if mode == "dark" else LIGHT_FG

            self.ai_button = ctk.CTkButton(
                self.text_area_frame,  # This makes sure it's inside the text area frame
                text="",
                image=ai_icon,
                fg_color=ai_bg_color,  # Match parent frame
                hover_color=ai_hover_color, # Set initial hover color
                corner_radius=20,  # Rounded corners for the button itself
                width=50,
                height=50,
                command=self.show_ai_popup
            )
            self.ai_button.place(in_=self.text_area_frame, relx=1.0, rely=1.0, x=-15, y=-15, anchor="se")

            # Position relative to the text_area_frame for better anchoring
            self.ai_button.place(in_=self.text_area_frame, relx=1.0, rely=1.0, x=-15, y=-15, anchor="se")
            self.ai_button.bind("<ButtonPress>", lambda e: self.ai_button.configure(width=45, height=45))
            self.ai_button.bind("<ButtonRelease>", lambda e: self.ai_button.configure(width=50, height=50))

            # --- Delay Frame ---
            delay_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="transparent") # Theme
            delay_frame.pack(fill="x", padx=5, pady=(10, 5)) # Increased top padding

            delay_font = ("Arial", 14) # Increased font size

            ctk.CTkLabel(delay_frame, text="Min Delay (s)", font=delay_font).pack(side="left", padx=(10, 5))
            self.min_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="1", corner_radius=10, width=60, fg_color=DARK_INNER, border_width=0, font=delay_font) # Theme & Font
            self.min_delay_entry.insert(0, str(DEFAULT_MIN_DELAY))
            self.min_delay_entry.pack(side="left", padx=5)

            ctk.CTkLabel(delay_frame, text="Max Delay (s)", font=delay_font).pack(side="left", padx=(15, 5))
            self.max_delay_entry = ctk.CTkEntry(delay_frame, placeholder_text="10", corner_radius=10, width=60, fg_color=DARK_INNER, border_width=0, font=delay_font) # Theme & Font
            self.max_delay_entry.insert(0, str(DEFAULT_MAX_DELAY))
            self.max_delay_entry.pack(side="left", padx=5)

        
    def show_context_menu(self, event):
        try:
            # Update menu theme before showing
            mode = ctk.get_appearance_mode().lower()
            if mode == "dark":
                self.context_menu.configure(bg=DARK_FG, fg="white")
            else:
                self.context_menu.configure(bg=LIGHT_FG, fg="black")
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
            
    def save_state(self, event=None):
        current_text = self.message_text.get("0.0", "end-1c")
        if not self.undo_stack or current_text != self.undo_stack[-1]: # Avoid duplicate states
            self.undo_stack.append(current_text)
            self.redo_stack.clear()
        
    def undo(self, event=None):
        if len(self.undo_stack) > 1: # Need at least one state to revert to
            current_text = self.undo_stack.pop()
            self.redo_stack.append(current_text)
            previous_text = self.undo_stack[-1]
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", previous_text)
            
    def redo(self, event=None):
        if self.redo_stack:
            next_text = self.redo_stack.pop()
            self.undo_stack.append(next_text)
            self.message_text.delete("0.0", "end")
            self.message_text.insert("0.0", next_text)
            
    # New generic method for placeholders
    def insert_placeholder(self, placeholder_text):
        self.save_state()
        self.message_text.insert("insert", placeholder_text)
        self.save_state() # Save state *after* inserting
        
    def apply_formatting(self, symbol):
        self.save_state()
        try:
            sel_start_str = self.message_text.index("sel.first")
            sel_end_str = self.message_text.index("sel.last")
            
            # Get text and delete selection BEFORE inserting to avoid index issues
            text = self.message_text.get(sel_start_str, sel_end_str).strip()
            self.message_text.delete(sel_start_str, sel_end_str) 
            
            formatted = f"{symbol}{text}{symbol}"
            self.message_text.insert(sel_start_str, formatted)

            # Re-select the inserted formatted text (optional, good UX)
            end_index = f"{sel_start_str}+{len(formatted)}c"
            self.message_text.tag_add("sel", sel_start_str, end_index)
            self.message_text.mark_set("insert", end_index)
            
        except tk.TclError: # Handle case where no text is selected
            # Insert the symbols at the cursor position
            cursor_index = self.message_text.index("insert")
            formatted = f"{symbol}{symbol}"
            self.message_text.insert(cursor_index, formatted)
            # Place cursor between the symbols
            self.message_text.mark_set("insert", f"{cursor_index}+{len(symbol)}c")
            
        self.save_state() # Save state *after* formatting

    def handle_attachment(self):
        # Check if an attachment is already selected
        if self.attachments.get("Any"):
            # --- Deselect Logic ---
            self.attachments["Any"] = None
            self.log_live("ℹ️ Attachment deselected.")
            
            # Reset button to "Select" state (green)
            self.attachment_btn.configure(
                text="Select Attachment",
                fg_color="green",
                hover_color="#45A049"
            )
            
            # If the last action was 'attachment', clear it
            if self.last_action == "attachment":
                self.last_action = None
                
        else:
            # --- Select Logic (No attachment) ---
            filetypes = [
                ("All Files", "*.*"),
                ("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                ("Document Files", "*.pdf;*.docx;*.doc;*.txt;*.csv;*.xlsx;*.xls"),
                ("Audio/Video Files", "*.mp3;*.wav;*.m4a;*.mp4;*.mkv;*.avi")
            ]
            path = filedialog.askopenfilename(filetypes=filetypes)
            
            if path:
                # Clear other types just in case
                self.attachments["Picture"] = None
                self.attachments["Video"] = None
                self.attachments["Document"] = None
                # Set the new attachment
                self.attachments["Any"] = path
                
                self.log_live(f"📎 Attachment selected: {os.path.basename(path)}")
                self.last_action = "attachment"
                
                # Update button to "Deselect" state (red)
                self.attachment_btn.configure(
                    text="Deselect Attachment",
                    fg_color="#FF4040", # Red color (like stop button)
                    hover_color="#D32F2F"  # Darker red hover
                )
            
    def open_custom_image_window(self):
        if hasattr(self, "excel_table"):
            self.excel_data = self.excel_table.get_data()
        if not self.excel_data:
            messagebox.showerror("Error", "Please load or enter phone data first.")
            return
            
        self.custom_image_enabled = True
        self.last_action = "custom"
        
        # --- New Logic ---
        # 1. Update this button to "Deselect" state (red)
        self.custom_image_btn.configure(
            text="Deselect Custom",
            command=self.deselect_custom_images,  # Point to the new deselect method
            fg_color="#FF4040", # Red color
            hover_color="#D32F2F"
        )
        
        # 2. Deselect the single attachment if one was active
        if self.attachments.get("Any"):
            self.attachments["Any"] = None
            self.log_live("ℹ️ Single attachment deselected (Custom Images enabled).")
            # Reset the *other* button
            self.attachment_btn.configure(
                text="Select Attachment",
                fg_color="green",
                hover_color="#45A049"
            )
        # --- End New Logic ---

        CustomImageWindow(self, self.excel_data)
        
        # --- New Logic ---
        # 1. Update this button to "Deselect" state (red)
        self.custom_image_btn.configure(
            text="Deselect Custom",
            command=self.deselect_custom_images,  # Point to the new deselect method
            fg_color="#FF4040", # Red color
            hover_color="#D32F2F"
        )
        
        # 2. Deselect the single attachment if one was active
        if self.attachments.get("Any"):
            self.attachments["Any"] = None
            self.log_live("ℹ️ Single attachment deselected (Custom Images enabled).")
            # Reset the *other* button
            self.attachment_btn.configure(
                text="Select Attachment",
                fg_color="green",
                hover_color="#45A049"
            )
        # --- End New Logic ---

        CustomImageWindow(self, self.excel_data)


    def deselect_custom_images(self):
        """Deselects custom images, cleans the output folder, and resets the button."""
        if not self.custom_image_enabled:
            return  # Do nothing if already deselected

        try:
            if os.path.exists(OUTPUT_IMG_FOLDER):
                shutil.rmtree(OUTPUT_IMG_FOLDER)
                self.log_live(f"ℹ️ Cleared custom images from {OUTPUT_IMG_FOLDER}.")
            else:
                self.log_live("ℹ️ Custom images deselected. (Folder not found)")
            
            # Re-create the folder so it's ready for next time
            os.makedirs(OUTPUT_IMG_FOLDER, exist_ok=True) 

        except PermissionError:
            self.log_live(f"❌ Error: Could not delete {OUTPUT_IMG_FOLDER}. Files may be in use.")
            messagebox.showerror("Error", f"Could not clear custom images. Close any programs using files in {OUTPUT_IMG_FOLDER} and try again.")
            return  # Don't proceed if we failed to delete
        except Exception as e:
            self.log_live(f"❌ Error clearing images: {e}")
            # Still proceed to reset UI
        
        self.custom_image_enabled = False
        if self.last_action == "custom":
            self.last_action = None
        
        # Reset button to "Select" state
        self.custom_image_btn.configure(
            text="Custom Image Namer",
            command=self.open_custom_image_window,  # Point command back to the original
            fg_color="green",
            hover_color="#45A049"
        )
    
        
    def create_excel_area(self, parent):
        top_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        country_codes = ["+91", "None", "+1", "+44", "+61", "+81", "+49", "+33", "+86", "+7"]
        self.country_code_var = ctk.StringVar(value="+91")
        self.country_code_dropdown = ctk.CTkOptionMenu(
            top_frame, 
            values=country_codes, 
            variable=self.country_code_var,
            font=("Arial", 16), # Bigger font
            dropdown_font=("Arial", 16), # Bigger font
            height=40,
            corner_radius=10,
            fg_color="green",           # Main button background
            button_color="#45A049",     # Color for the arrow button
            button_hover_color="#3E8E41"  # Darker green for arrow hover
        )
        self.country_code_dropdown.pack(side="left", padx=5, pady=5)
        
        self.import_db_btn = ctk.CTkButton(
            top_frame, 
            text="Import DataBase", 
            corner_radius=10, 
            command=self.open_import_popup,
            font=("Arial", 16, "bold"), # Bigger font
            height=40,
            fg_color="green",
            hover_color="#45A049"
        )
        self.import_db_btn.pack(side="left", padx=5, pady=5)
        
        mode = ctk.get_appearance_mode().lower()
        excel_bg = DARK_FG if mode == "dark" else LIGHT_FG
        
        self.excel_table = ExcelTable(parent, main_app=self, fg_color=excel_bg)
        self.excel_table.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
    def open_import_popup(self):
        # Check if data already exists
        if self.excel_data:
            # Create a small popup to ask for Merge or Replace
            popup = ctk.CTkToplevel(self)
            popup.title("Import Mode")
            popup.geometry("350x150")
            popup.transient(self)
            popup.attributes("-topmost", True)
            center_window(popup)
            
            mode = ctk.get_appearance_mode().lower()
            bg_color = DARK_BG if mode == "dark" else LIGHT_BG
            fg_color = DARK_FG if mode == "dark" else LIGHT_FG
            
            popup.configure(fg_color=bg_color)
            
            ctk.CTkLabel(popup, text="Data already exists.", font=("Arial", 16, "bold")).pack(pady=10)
            ctk.CTkLabel(popup, text="How do you want to import?").pack(pady=5)
            
            btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
            btn_frame.pack(pady=10, fill="x", expand=True)
            btn_frame.columnconfigure((0,1), weight=1)

            def launch_import(merge_mode):
                popup.destroy()
                ImportDatabasePopup(self, self.load_excel_data, merge_mode=merge_mode)

            merge_btn = ctk.CTkButton(
                btn_frame, 
                text="Merge\n(Add to existing)", 
                command=lambda: launch_import(merge_mode=True),
                fg_color="green",
                hover_color="#45A049"
            )
            merge_btn.grid(row=0, column=0, padx=10, pady=5)

            replace_btn = ctk.CTkButton(
                btn_frame, 
                text="Replace\n(Clear old data)", 
                command=lambda: launch_import(merge_mode=False),
                fg_color="#D32F2F", # Red color
                hover_color="#B71C1C"
            )
            replace_btn.grid(row=0, column=1, padx=10, pady=5)
            
        else:
            # No data exists, just open the import window in "replace" mode (merge_mode=False)
            ImportDatabasePopup(self, self.load_excel_data, merge_mode=False)
        
    def load_excel_data(self, path, phone_col, name_col, custom1_col, custom2_col, merge_mode):
        try:
            prog = ProgressPopup(self, "Loading Data", total=1)
            prog.geometry(
                f"500x300+{self.winfo_rootx() + (self.winfo_width()-500)//2}"
                f"+{self.winfo_rooty() + (self.winfo_height()-300)//2}"
            )
            
            def perform_loading():
                new_data = []
                
                # Helper to convert col letter to index (A=0, B=1, ...)
                def col_to_index(col_letter):
                    if not col_letter or not col_letter.isalpha():
                        return -1
                    return ord(col_letter.upper()) - 65

                phone_idx = col_to_index(phone_col)
                name_idx = col_to_index(name_col)
                custom1_idx = col_to_index(custom1_col)
                custom2_idx = col_to_index(custom2_col)

                if path.endswith('.xlsx') or path.endswith('.xls'):
                    wb = openpyxl.load_workbook(path, data_only=True) # data_only=True to get values, not formulas
                    sheet = wb.active
                    
                    max_col_index = sheet.max_column - 1 # 0-based index
                    
                    # Validate required phone index
                    if phone_idx > max_col_index or phone_idx == -1:
                        self.log_live(f"❌ Error: Phone column '{phone_col}' not found in Excel.")
                        self.after(0, lambda: prog.close())
                        messagebox.showerror("Import Error", f"Phone column '{phone_col}' (Index {phone_idx}) is out of range. Max column index is {max_col_index}.")
                        return
                        
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        entry = {}
                        try:
                            phone_val = row[phone_idx] if len(row) > phone_idx else ""
                            phone = str(phone_val).strip() if phone_val is not None else ""
                            
                            if not phone: # Skip row if phone is empty
                                continue
                                
                            entry["phone"] = phone
                            
                            if name_idx != -1 and len(row) > name_idx:
                                name_val = row[name_idx]
                                entry["name"] = str(name_val).strip() if name_val is not None else ""
                            
                            if custom1_idx != -1 and len(row) > custom1_idx:
                                c1_val = row[custom1_idx]
                                entry["custom1"] = str(c1_val).strip() if c1_val is not None else ""
                                
                            if custom2_idx != -1 and len(row) > custom2_idx:
                                c2_val = row[custom2_idx]
                                entry["custom2"] = str(c2_val).strip() if c2_val is not None else ""
                            
                            new_data.append(entry)
                            
                        except Exception as e:
                            self.log_live(f"Error reading Excel row: {e}")
                            
                elif path.endswith('.csv'):
                    try:
                        with open(path, newline='', encoding='utf-8-sig') as csvfile: # Use utf-8-sig to handle BOM
                            # Sniff dialect
                            dialect = csv.Sniffer().sniff(csvfile.read(1024))
                            csvfile.seek(0) 
                            reader = csv.reader(csvfile, dialect)
                            
                            header = next(reader) # Read header row
                            
                            # --- CSV Column Name Matching ---
                            # Re-find indices based on header names (more robust for CSV)
                            try:
                                phone_idx_csv = header.index(phone_col)
                            except ValueError:
                                self.log_live(f"❌ Error: Phone column '{phone_col}' not found in CSV header.")
                                self.after(0, lambda: prog.close())
                                messagebox.showerror("Import Error", f"Phone column '{phone_col}' not found in CSV header: {header}")
                                return

                            try: name_idx_csv = header.index(name_col)
                            except ValueError: name_idx_csv = -1
                            
                            try: custom1_idx_csv = header.index(custom1_col)
                            except ValueError: custom1_idx_csv = -1

                            try: custom2_idx_csv = header.index(custom2_col)
                            except ValueError: custom2_idx_csv = -1
                            # ---
                            
                            for row in reader:
                                entry = {}
                                try:
                                    phone = row[phone_idx_csv].strip() if len(row) > phone_idx_csv else ""
                                    
                                    if not phone:
                                        continue
                                        
                                    entry["phone"] = phone
                                    
                                    if name_idx_csv != -1 and len(row) > name_idx_csv:
                                        entry["name"] = row[name_idx_csv].strip()
                                    
                                    if custom1_idx_csv != -1 and len(row) > custom1_idx_csv:
                                        entry["custom1"] = row[custom1_idx_csv].strip()
                                        
                                    if custom2_idx_csv != -1 and len(row) > custom2_idx_csv:
                                        entry["custom2"] = row[custom2_idx_csv].strip()
                                    
                                    new_data.append(entry)
                                    
                                except Exception as e:
                                    self.log_live(f"Error reading CSV row: {e}")
                                    
                    except Exception as e:
                         self.log_live(f"Error opening/reading CSV file: {e}")
                         messagebox.showerror("File Error", f"Could not read CSV file: {e}")
                         self.after(0, lambda: prog.close())
                         return

                # --- Handle Merge/Replace ---
                if merge_mode:
                    self.excel_data = self.excel_data + new_data
                    self.log_live(f"✅ Merged {len(new_data)} new contacts. Total: {len(self.excel_data)}")
                else:
                    self.excel_data = new_data
                    self.log_live(f"✅ Replaced with {len(new_data)} contacts from {os.path.basename(path)}")

                # --- Load data into the dynamic table ---
                self.after(0, lambda: self.excel_table.load_data(self.excel_data))
                self.after(0, lambda: prog.close())
                
            threading.Thread(target=perform_loading, daemon=True).start()
            
        except Exception as e:
            self.log_live(f"Error loading file: {e}")
            messagebox.showerror("Error", f"Failed to load file: {e}")
            
    def launch_whatsapp_beta(self):
        # --- New Logic to Clear Session ---
        try:
            if os.path.exists(SESSION_DIR):
                self.log_live("Clearing previous session for fresh login...")
                # Use shutil.rmtree to delete a non-empty directory
                shutil.rmtree(SESSION_DIR)
                self.log_live("Previous session cleared.")
            
            # Re-create the directory immediately
            if not os.path.exists(SESSION_DIR):
                 os.makedirs(SESSION_DIR)
                 
        except PermissionError:
            self.log_live(f"❌ Error: Session files are in use. Please close Chrome.")
            messagebox.showerror("Session Error", "Could not clear the session. Please close any open Chrome windows (especially WhatsApp Web) and try again.")
            return # Stop if we can't clear the session
        except Exception as e:
            self.log_live(f"❌ Error clearing session: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred while clearing the session: {e}")
            return # Stop on other errors
        # --- End of New Logic ---
            
        def run_login():
            logger = GuiLogger(self)
            selenium_login(logger)
        
        threading.Thread(target=run_login, daemon=True).start()
        
    def toggle_sending(self):
        if not self.sending:
            self.start_sending()
        else:
            self.stop_sending()
            
    def start_sending(self):
        self.sending = True
        self.start_stop_button.configure(text="Stop", fg_color="#FF4040")
        data = self.excel_table.get_data()
        if not data:
            messagebox.showerror("Error", "No phone numbers loaded.")
            self.stop_sending()
            return
        msg = self.message_text.get("0.0", "end-1c").strip()
        if not msg and not self.custom_image_enabled and not ("Any" in self.attachments and self.attachments["Any"]):
            messagebox.showerror("Error", "No message text provided and no attachment selected.")
            self.stop_sending()
            return
        attachment_path = self.attachments["Any"] if "Any" in self.attachments else None
        
        # --- Get Delay Settings ---
        try:
            min_delay = int(self.min_delay_entry.get())
            if min_delay <= 0: min_delay = DEFAULT_MIN_DELAY # Ensure positive
        except ValueError:
            min_delay = DEFAULT_MIN_DELAY
            self.log_live(f"⚠️ Invalid Min Delay, using default: {DEFAULT_MIN_DELAY}s")
            
        try:
            max_delay = int(self.max_delay_entry.get())
            if max_delay <= 0: max_delay = DEFAULT_MAX_DELAY # Ensure positive
        except ValueError:
            max_delay = DEFAULT_MAX_DELAY
            self.log_live(f"⚠️ Invalid Max Delay, using default: {DEFAULT_MAX_DELAY}s")

        if min_delay > max_delay:
            self.log_live(f"⚠️ Min Delay ({min_delay}) > Max Delay ({max_delay}). Swapping values.")
            min_delay, max_delay = max_delay, min_delay # Swap them
        
        self.log_live(f"🚀 Using random delay range: {min_delay}s to {max_delay}s")
        # --- End Delay Settings ---

        # Filter out skipped entries
        valid_data = [entry for entry in data if not entry.get("skip", False)]
        if not valid_data:
             messagebox.showerror("Error", "All entries are marked to be skipped.")
             self.stop_sending()
             return

        numbers = [normalize_phone(entry["phone"]) for entry in valid_data if entry.get("phone")] # Ensure phone exists
        
        # --- *** UPDATED PERSONALIZATION *** ---
        personalized_msgs = []
        for entry in valid_data:
            if not entry.get("phone"): # Skip if no phone number for this entry
                continue
                
            temp_msg = msg.replace("{User_Name}", entry.get("name", ""))
            
            # Add logic to replace {Custom1} and {Custom2}
            if "{Custom1}" in temp_msg:
                temp_msg = temp_msg.replace("{Custom1}", entry.get("custom1", ""))
            if "{Custom2}" in temp_msg:
                temp_msg = temp_msg.replace("{Custom2}", entry.get("custom2", ""))
                
            personalized_msgs.append(temp_msg)
        # --- *** END OF UPDATE *** ---

        # Check if numbers and messages lists are aligned
        if len(numbers) != len(personalized_msgs):
            self.log_live(f"❌ Error: Mismatch between numbers ({len(numbers)}) and messages ({len(personalized_msgs)}).")
            messagebox.showerror("Error", "Mismatch in contact data. Check for entries without phone numbers.")
            self.stop_sending()
            return

        def schedule_and_send():
            # Added Debug Logs
            print(f"DEBUG: Schedule thread started. Target time: {self.schedule_time}, Sending flag: {self.sending}")

            if self.schedule_time:
                print(f"DEBUG: Entering wait loop. Now: {datetime.now()}, Target: {self.schedule_time}")
                # Loop while current time is less than target AND sending flag is True
                while self.sending:
                    now = datetime.now()
                    if now >= self.schedule_time:
                        print(f"DEBUG: Target time reached or passed. Now: {now}, Target: {self.schedule_time}")
                        break # Exit the loop, time has come

                    remaining = (self.schedule_time - now).total_seconds()
                    # Log more frequently when close
                    log_interval = 10 if remaining > 60 else 5
                    if int(remaining) % log_interval == 0:
                        self.log_live(f"⏳ Time left until scheduled start: {int(remaining)} seconds")

                    # Check every 0.5 seconds for responsiveness
                    time.sleep(0.5)

                print(f"DEBUG: Exited wait loop. Now: {datetime.now()}, Sending flag: {self.sending}")

                # Check *why* the loop exited
                if not self.sending: # If user clicked "Stop" during the wait
                    self.log_live("🛑 Schedule cancelled by user.")
                    self.after(0, self.stop_sending) # Ensure UI updates correctly
                    print("DEBUG: Sending flag was False after loop.")
                    return # Stop the function here

                # If the loop exited because time was reached AND sending is still True
                print("DEBUG: Proceeding to send after schedule wait.")
                self.log_live("🚀 Scheduled time reached! Starting sending...")

            else: # No schedule was set, start immediately
                 print("DEBUG: No schedule time set, starting immediately.")
                 self.log_live("🚀 Starting sending now...")

            # --- Rest of the function (preparing attachments, calling selenium_send_bulk) ---
            logger = GuiLogger(self)
            send_attachments = []

            # Custom Image Attachment Logic... (Keep this section as is)
            if self.custom_image_enabled:
                self.log_live("📎 Attaching custom images...")
                generated_image_paths = {}
                if hasattr(self, 'excel_data') and self.excel_data:
                    for item in self.excel_data:
                        phone_norm = normalize_phone(item.get('phone', ''))
                        if 'image_path' in item and item['image_path']:
                            generated_image_paths[phone_norm] = item['image_path']
                for num in numbers:
                    path = generated_image_paths.get(num)
                    if path and os.path.exists(path): send_attachments.append(path)
                    else: send_attachments.append(None)

            # Single Attachment Logic... (Keep this section as is)
            elif attachment_path:
                self.log_live(f"📎 Attaching single file: {os.path.basename(attachment_path)}")
                send_attachments = [attachment_path] * len(numbers)
            else:
                send_attachments = None

            # --- Start Selenium Send ---
            if not self.sending: # Final check before potentially long operation
                print("DEBUG: Sending flag became False just before calling selenium_send_bulk.")
                self.after(0, self.stop_sending) # Ensure UI reset if somehow missed
                return

            print(f"DEBUG: Calling selenium_send_bulk with {len(numbers)} numbers, min_delay={min_delay}, max_delay={max_delay}")

            # Make the call to the Selenium function
            success, failure = selenium_send_bulk(
                numbers,
                personalized_msgs,
                send_attachments,
                min_delay,
                max_delay,
                logger
            )

            # --- Report and Cleanup --- (Keep this section as is)
            if (success + failure) > 0:
                generate_html_report(success, failure)
            else:
                self.log_live("ℹ️ No messages were attempted.")

            # Ensure stop_sending is called on the main thread after completion
            self.after(0, self.stop_sending)
            self.schedule_time = None # Clear schedule time after execution
            print("DEBUG: Schedule thread finished.")

        # Start the thread (This line remains the same)
        threading.Thread(target=schedule_and_send, daemon=True).start()

        
    def stop_sending(self):
        self.sending = False
        self.start_stop_button.configure(text="Start", fg_color="green") # Changed to green
        self.log_live("🛑 Sending stopped.") # Changed log message
        
    def set_schedule_time(self, schedule_time):
        self.schedule_time = schedule_time
        self.log_live(f"⏰ Scheduled sending for {schedule_time.strftime('%I:%M %p, %b %d, %Y')}")
        
    # New method to show the AI popup
    def show_ai_popup(self):
        AIPopup(self, self.ai_button)

    # Modified method to handle popup closing
    def process_ai(self, option, popup=None):
        if popup:
            popup.destroy() # Close the popup when an option is clicked
            
        text = self.message_text.get("0.0", "end-1c").strip()
        if not text and option != "Ask AI": # Allow "Ask AI" to run on empty text
            messagebox.showerror("Error", "No message text to process.")
            return
            
        if option == "Translate":
            TranslatePopup(self, self.translate_message)
        else:
            threading.Thread(target=self.call_gemini_api, args=(option, text), daemon=True).start()
            
    def translate_message(self, lang):
        text = self.message_text.get("0.0", "end-1c").strip()
        if not text:
            messagebox.showerror("Error", "No message text to translate.")
            return
        threading.Thread(target=self.call_gemini_api, args=("Translate", text, lang), daemon=True).start()
        
    # Add this import at the top of the file with other imports


    # In the __init__ method of WabulkXpressApp, add:
    genai.configure(api_key="AIzaSyBR9Os8LKyNQAwkeLKXO1OyGQnRjnaMe8w")

    # Ensure this method is properly indented inside the WabulkXpressApp class
    # (4 spaces or consistent with class indentation)
    def call_gemini_api(self, option, text, lang=None):
        try:
            self.log_live(f"🧠 Calling AI ({option})...")
            prompt = self.ai_prompts[option]
            if option == "Translate":
                prompt = prompt.format(lang=lang)
            
            full_prompt = f"{prompt}\n\n{text}"
            
            model = genai.GenerativeModel('gemini-2.0-flash')
            response = model.generate_content(full_prompt)
            
            processed_text = response.text
            
            if processed_text:
                self.after(0, lambda: self.update_message_text(processed_text))
                self.log_live(f"✅ AI {option} applied successfully!")
            else:
                self.log_live(f"⚠️ AI processing returned empty result for {option}.")
                logger.warning(f"Empty AI result. Raw response: {response}")
        except Exception as e:
            self.log_live(f"❌ Error in AI processing: {e}")
            messagebox.showerror("AI Error", f"Failed to process with AI: {e}")
            logger.error(f"Unexpected AI error: {e}", exc_info=True)



            
    def update_message_text(self, text):
        self.save_state() # Save before changing
        self.message_text.delete("0.0", "end")
        self.message_text.insert("0.0", text)
        self.save_state() # Save after changing
        
    def log_live(self, message):
        if not self.winfo_exists():
            return
        # Simple check to prevent excessive logging of the same message rapidly
        try:
             last_log = self.live_alerts.get("end-2l linestart", "end-1l lineend") # Get second to last line
             if message == last_log:
                  return 
        except tk.TclError:
             pass # Ignore error if textbox is empty or has only one line
             
        self.live_alerts.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.live_alerts.insert("end", f"[{timestamp}] {message}\n")
        self.live_alerts.see("end")
        self.live_alerts.configure(state="disabled")
        
    def check_for_update(self):
        try:
            self.log_live("Checking for updates...")
            response = requests.get(GITHUB_API_URL, timeout=15)
            response.raise_for_status()
            latest_version = response.json().get("tag_name", "0")

            # --- Add this block ---
            # Update the last check time after a successful manual check
            try:
                with open(UPDATE_CHECK_FILE, 'w') as f:
                    f.write(str(time.time()))
                print(f"DEBUG: Updated last update check timestamp (manual check) to {datetime.now()}") # Console log
            except Exception as e:
                print(f"DEBUG: Failed to update timestamp file after manual check: {e}") # Console log
            # --- End added block ---

            if float(latest_version) > float(CURRENT_VERSION):
                self.log_live(f"Update available: {latest_version} (Current: {CURRENT_VERSION})")
                # Use the prompt helper function
                self.prompt_for_update(latest_version)
            else:
                self.log_live("You are using the latest version.")
                messagebox.showinfo("Up to Date", f"You are using the latest version ({CURRENT_VERSION}).")
        # ... (rest of the error handling remains the same) ...
        except requests.exceptions.Timeout:
            self.log_live(f"Update check failed: Request timed out.")
            messagebox.showerror("Update Error", "Failed to check for updates: Request timed out.")
        except requests.exceptions.RequestException as e:
            self.log_live(f"Update check failed: {e}")
            messagebox.showerror("Update Error", f"Failed to check for updates: {e}")
        except Exception as e:
            self.log_live(f"Unexpected error during update check: {e}")
            messagebox.showerror("Update Error", f"An unexpected error occurred: {e}")
            logger.error(f"Update check error: {e}", exc_info=True)
                
    def toggle_theme(self):
        current_mode = ctk.get_appearance_mode().lower()
        new_mode = "Light" if current_mode == "dark" else "Dark"
        ctk.set_appearance_mode(new_mode)

        if new_mode == "Light":
            # Apply Light Theme
            self.configure(fg_color=LIGHT_BG)
            self.sidebar.configure(fg_color=LIGHT_FG)
            self.header.configure(fg_color=LIGHT_FG)
            self.main_area.configure(fg_color="transparent")
            self.message_frame.configure(fg_color=LIGHT_FG)
            self.excel_frame.configure(fg_color=LIGHT_FG)
            self.live_alerts.configure(fg_color=LIGHT_INNER)
            self.text_area_frame.configure(fg_color=LIGHT_INNER)
            self.message_text.configure(fg_color=LIGHT_INNER)
            self.min_delay_entry.configure(fg_color=LIGHT_INNER)
            self.max_delay_entry.configure(fg_color=LIGHT_INNER)
            self.excel_table.configure(
                fg_color=LIGHT_FG,
                scrollbar_button_color="#AAAAAA",       # Lighter gray knob for light mode
                scrollbar_button_hover_color="#888888"  # Medium gray knob on hover
            )
            # Update ExcelTable rows
            for row in self.excel_table.rows:
                if row["phone"].winfo_exists(): # Check if widget exists before configuring
                    row["phone"].configure(fg_color=LIGHT_INNER)
                    row["name"].configure(fg_color=LIGHT_INNER)
                    row["indicator"].configure(bg=LIGHT_INNER)
            # Update AI button hover color
            # Update AI button colors
            self.ai_button.configure(fg_color=LIGHT_INNER, hover_color=LIGHT_FG)
        else:
            # Apply Dark Theme
            self.configure(fg_color=DARK_BG)
            self.sidebar.configure(fg_color=DARK_FG)
            self.header.configure(fg_color=DARK_FG)
            self.main_area.configure(fg_color="transparent")
            self.message_frame.configure(fg_color=DARK_FG)
            self.excel_frame.configure(fg_color=DARK_FG)
            self.live_alerts.configure(fg_color=DARK_INNER)
            self.text_area_frame.configure(fg_color=DARK_INNER)
            self.message_text.configure(fg_color=DARK_INNER)
            self.min_delay_entry.configure(fg_color=DARK_INNER)
            self.max_delay_entry.configure(fg_color=DARK_INNER)
            self.excel_table.configure(
                fg_color=DARK_FG,
                scrollbar_button_color="#555555",       # Darker gray knob for dark mode
                scrollbar_button_hover_color="#333333"  # Even darker on hover
            )
            # Update ExcelTable rows
            for row in self.excel_table.rows:
                 if row["phone"].winfo_exists(): # Check if widget exists
                    row["phone"].configure(fg_color=DARK_INNER)
                    row["name"].configure(fg_color=DARK_INNER)
                    row["indicator"].configure(bg=DARK_INNER)
            # Update AI button hover color
            self.ai_button.configure(hover_color=DARK_FG)

        self.refresh_icons()
        
    def on_close(self):
        # Ensure cleanup happens before destroying the window
        if self.sending:
             # Optionally ask the user if they want to stop sending
             if messagebox.askyesno("Confirm Exit", "Sending is in progress. Are you sure you want to exit?"):
                  self.stop_sending() # Attempt graceful stop if possible
             else:
                  return # Prevent closing if user cancels

        # Stop video if playing
        if hasattr(self, "video_player") and self.video_player:
            try:
                self.video_player.stop()
            except Exception as e:
                print(f"Error stopping video player: {e}")
        
        # Add any other necessary cleanup here (e.g., close Selenium driver if running)
        
        self.destroy() # Destroy the main window
        sys.exit(0) # Ensure the script fully terminates

if __name__ == "__main__":
    # Setup basic logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # Redirect stdout/stderr for better error handling if running as executable (optional)
    # try:
    #     sys.stdout = open('stdout.log', 'w')
    #     sys.stderr = open('stderr.log', 'w')
    # except Exception as e:
    #     print(f"Could not redirect stdout/stderr: {e}")

    logger.info("Starting WabulkXpress Application...")
    app = WabulkXpressApp()
    try:
        app.mainloop()
    except Exception as e:
         logger.error("An unhandled exception occurred in the main loop:", exc_info=True)
         messagebox.showerror("Fatal Error", f"A critical error occurred:\n{e}\n\nPlease check the logs for details.")
    finally:
         logger.info("WabulkXpress Application finished.")