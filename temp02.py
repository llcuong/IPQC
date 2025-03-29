import tkinter as tk
import ctypes
import threading
import winreg
from PIL import Image, ImageTk
from tkinter import messagebox
import serial
import serial.tools.list_ports
import time
import datetime
from tkinter import ttk
import pyodbc
import re
import socket
import barcode
from io import BytesIO
from barcode.writer import ImageWriter
import os
import sys

ase_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")

"""Set DPI awareness for better scaling on high-DPI screens"""
ctypes.windll.shcore.SetProcessDpiAwareness(10)
REG_PATH = r"SOFTWARE\IPQC\Config"

user32 = ctypes.windll.user32
monitor_width, monitor_height = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
# print(f"Screen Resolution: {monitor_width}x{monitor_height}")

"""Main window setup"""
root = tk.Tk()
root.title("IPQC v.25.03.29")

"""Define scaling factors"""
user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
dpi = user32.GetDpiForSystem()
scaling = dpi / 96
# print(f"scaling ==> {scaling}")
screen_width = min(1600, int(0.9*root.winfo_screenwidth()*scaling))
screen_height = min(900, int(0.9*root.winfo_screenheight()*scaling))
# print(f"Window size: {screen_width}x{screen_height}")

root.iconbitmap("theme/icons/logo.ico")
root.minsize(screen_width, screen_height)
root.geometry(f"{screen_width}x{screen_height}")

style = ttk.Style(root)
root.tk.call('source', 'theme/azure.tcl')
style.theme_use('azure')
style.configure("Togglebutton", foreground='white')



"""Define app parameters"""
font_name = 'Arial'
font_size_base_on_ratio = int(screen_height * 0.015)
button_width_base_on_ratio = int(screen_width * 0.012)

bg_app_class_color_layer_1 = "#f4f4fe"   #00B9FF    => #333333
bg_app_class_color_layer_2 = "#ffffff"   #          => #414141
fg_app_class_color_layer_1 = '#000000'


showing_settings = False
showing_runcards = False
showing_advance_setting = False
showing_thickness_frame = False
showing_weight_frame = False

current_thickness_entry = ""
current_weight_entry = ""
error_msg = tk.StringVar()
error_msg.set("")
error_fg_color = "red"

error_thread = None
error_event = threading.Event()

weight_record_id = 0
thickness_record_id = 0

weight_com_thread = None
thickness_com_thread = None
class CustomOptionMenu(tk.OptionMenu):
    def __init__(self, master, variable, *options, command=None, **kwargs):
        if options is None:
            options = ["No COM"]
        super().__init__(master, variable, *options, **kwargs)
        self.config(font=("Helvetica", 14, "bold"),
                    bg="white",
                    fg="#333333",
                    activebackground="#f0f0f0",
                    activeforeground="black",
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground="#d1d1d1",
                    bd=0,
                    padx=12, pady=6,
                    width=8)
        self["menu"].config(font=("Helvetica", 14),
                            bg="white", fg="black",
                            activebackground="#e0e0e0",
                            activeforeground="black",
                            bd=0)
        if command:
            variable.trace_add("write", lambda *args: command(variable.get()))
def show_error_message(msg, fg_color_code, time_show):
    global error_thread, error_event
    if error_thread and error_thread.is_alive():
        error_event.set()
        error_thread.join()
    error_event = threading.Event()
    def clear_error_message():
        error_msg.set("")
        error_display_entry.config(fg=bg_app_class_color_layer_1)
    def update_message():
        if fg_color_code == 1:
            fg_color, fg_icon = "green", "✔"
        elif fg_color_code == 0:
            fg_color, fg_icon = "red", "❌"
        elif fg_color_code == -1:
            fg_color, fg_icon = "#595959", "↻"
        else:
            fg_color, fg_icon = "#7F7F7F", "ⓘ"
        error_msg.set(f"{fg_icon} {msg}")
        root.after(0, lambda: error_display_entry.config(fg=fg_color))
        if not error_event.wait(time_show / 1000):
            root.after(0, clear_error_message)
    error_thread = threading.Thread(target=update_message, daemon=True)
    error_thread.start()
def convert_to_uppercase(entry_name, max_char, accept_char):
    try:
        input_text = entry_name.get()
        if accept_char == 0:
            input_text = input_text.replace(',', '.')
            filtered_text = ''.join(char for char in input_text if char.isdigit() or char == '.')
            if filtered_text != input_text:
                threading.Thread(target=show_error_message, args=("Only numbers and '.' are allowed!", 0, 2000), daemon=True).start()
            input_text = filtered_text
        upper_entry_name = input_text.upper()
        if len(upper_entry_name) > max_char:
            entry_name.set(upper_entry_name[:max_char])
            threading.Thread(target=show_error_message, args=("Exceeded character limit!", 0, 2000), daemon=True).start()
        else:
            entry_name.set(upper_entry_name)
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass
def get_registry_value(name, default=""):
    """Retrieve a value from the Windows Registry."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, name)[0]
    except FileNotFoundError as e:
        return default
def set_registry_value(name, value):
    """Set a value in the Windows Registry."""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    except Exception as e:
        print(e)
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

