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

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")
"""Set DPI awareness for better scaling on high-DPI screens"""
ctypes.windll.shcore.SetProcessDpiAwareness(10)
REG_PATH = r"SOFTWARE\IPQC\Config"

user32 = ctypes.windll.user32
monitor_width, monitor_height = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
print(f"Screen Resolution: {monitor_width}x{monitor_height}")

"""Main window setup"""
root = tk.Tk()
root.title("IPQC")

"""Define scaling factors"""
user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
dpi = user32.GetDpiForSystem()
scaling = dpi / 96
print(f"scaling ==> {scaling}")
screen_width = min(1480, int(0.8*root.winfo_screenwidth()*scaling))
screen_height = min(900, int(0.8*root.winfo_screenheight()*scaling))
print(f"Window size: {screen_width}x{screen_height}")

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

bg_app_color = "#f4f4fe"   #00B9FF    => #333333
bg_param_color = "white"   #          => #414141


showing_settings = False
showing_runcards = False
showing_advance_setting = False
showing_thickness_frame = False
showing_weight_frame = False


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
def update_dimensions():
    global screen_width, screen_height, showing_settings, showing_runcards
    while True:

        """Define height - width size parameters"""


        root.update_idletasks()
        screen_width = root.winfo_width()
        screen_height = root.winfo_height()

        top_frame.config(width=int(screen_width), height=50)
        bottom_frame.config(width=int(screen_width), height=35)

        middle_frame_height = screen_height - 50 - 50
        middle_frame.place(x=0, y=50, width=screen_width, height=middle_frame_height)





        if showing_settings:
            middle_right_frame_width = 332
            middle_center_frame_width = 10
            middle_left_frame_width = screen_width - middle_right_frame_width - middle_center_frame_width
            middle_left_col3_width = 210
            middle_left_col2_width = 10
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width
        elif showing_runcards:
            middle_right_frame_width = 520
            middle_center_frame_width = 10
            middle_left_frame_width = screen_width - middle_right_frame_width - middle_center_frame_width
            if showing_thickness_frame:
                middle_left_col3_width = 0
                middle_left_col2_width = 0
            elif showing_weight_frame:
                middle_left_col3_width = 210
                middle_left_col2_width = 10
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width
        elif showing_advance_setting:
            middle_right_frame_width = 552
            middle_center_frame_width = 10
            middle_left_frame_width = screen_width - middle_right_frame_width - middle_center_frame_width
            middle_left_col3_width = 0
            middle_left_col2_width = 0
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width
        else:
            middle_right_frame_width = 0
            middle_center_frame_width = 0
            middle_left_frame_width = screen_width - middle_right_frame_width - middle_center_frame_width
            middle_left_col3_width = 420
            middle_left_col2_width = 10
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width





        middle_right_frame.place(x=middle_left_frame_width + middle_center_frame_width, y=0, width=middle_right_frame_width, height=middle_frame_height)
        middle_center_frame.place(x=middle_left_frame_width, y=0, width=middle_center_frame_width, height=middle_frame_height)
        middle_left_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        top_left_frame.place(x=0, y=0, width=int(screen_width * 0.3), height=50)
        top_right_frame.place(x=int(screen_width * 0.3), y=0, width=int(screen_width * 0.7), height=50)

        bottom_left_frame.place(x=0, y=0, width=int(screen_width * 0.3), height=50)
        bottom_right_frame.place(x=int(screen_width * 0.3), y=0, width=int(screen_width * 0.7), height=50)

        middle_left_weight_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)
        middle_left_thickness_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        middle_left_weight_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_weight_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_weight_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_left_thickness_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_thickness_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_thickness_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_right_runcard_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_advance_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)

        middle_right_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=180)
        middle_right_setting_frame_row3.place(x=0, y=40 + 200, width=middle_right_frame_width, height=50)

        middle_right_advance_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_advance_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width, height=50)

        middle_right_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width/2), height=40)
        middle_right_setting_frame_row1_col2.place(x=int(middle_right_frame_width/2), y=0, width=int(middle_right_frame_width/2), height=40)

        middle_right_setting_frame_row3_col1.place(x=int((middle_right_frame_width/2-154)/2), y=0, width=154, height=40)
        middle_right_setting_frame_row3_col2.place(x=int((middle_right_frame_width/2-154)/2+middle_right_frame_width/2), y=0, width=154, height=40)

        middle_right_advance_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width / 2), height=40)
        middle_right_advance_setting_frame_row1_col2.place(x=int(middle_right_frame_width / 2), y=0, width=int(middle_right_frame_width / 2), height=40)

        middle_right_advance_setting_frame_row3_col1.place(x=int((middle_right_frame_width/2-154)/2), y=0, width=154, height=40)
        middle_right_advance_setting_frame_row3_col2.place(x=int((middle_right_frame_width/2-154)/2+middle_right_frame_width/2), y=0, width=154, height=40)

        middle_right_setting_frame_row2_row1.place(x=0, y=20, width=int(middle_right_frame_width), height=160)

        middle_right_advance_setting_frame_row2_col1.place(x=0, y=0, width=332, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row2_col2.place(x=332, y=0, width=10, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row2_col3.place(x=342, y=0, width=middle_right_frame_width - 10 - 332, height=middle_frame_height - 40 - 50)

        middle_right_advance_setting_frame_row2_col1_row1.place(x=0, y=0, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row2.place(x=0, y=20, width=332, height=160)
        middle_right_advance_setting_frame_row2_col1_row3.place(x=0, y=180, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row4.place(x=0, y=200, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row5.place(x=0, y=210, width=332, height=180)
        #Add one frame here for "test connection" button
        middle_right_advance_setting_frame_row2_col1_row6.place(x=0, y=390, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row7.place(x=0, y=410, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row8.place(x=0, y=420, width=332, height=40)

        middle_right_advance_setting_frame_row2_col1_row5_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row5_col2.place(x=170, y=0, width=162, height=180)

        middle_right_advance_setting_frame_row2_col1_row8_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row8_col2.place(x=170, y=0, width=162, height=180)

        middle_right_advance_setting_frame_row2_col1_row5_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row2.place(x=0, y=40, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row3.place(x=0, y=80, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row4.place(x=0, y=120, width=170, height=40)

        middle_right_advance_setting_frame_row2_col1_row5_col2_row1.place(x=0, y=0, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row2.place(x=0, y=40, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row3.place(x=0, y=80, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row4.place(x=0, y=120, width=162, height=40)

        middle_right_advance_setting_frame_row2_col1_row8_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row8_col2_row1.place(x=0, y=0, width=162, height=40)

        root.update_idletasks()
        root.update()


"""Frame"""
top_frame = tk.Frame(root, bg=bg_app_color)
top_frame.place(relx=0, rely=0, height=50, anchor="nw")

top_left_frame = tk.Frame(top_frame, bg=bg_app_color)
top_right_frame = tk.Frame(top_frame, height=50, bg=bg_app_color)
top_right_frame.grid_columnconfigure(0, weight=1)





middle_frame = tk.Frame(root, bg=bg_app_color)
middle_frame.place(relx=0, rely=0, anchor="nw")

middle_left_frame = tk.Frame(middle_frame, bg=bg_app_color)
middle_left_frame.place(x=0, y=0)



middle_left_weight_frame = tk.Frame(middle_left_frame, bg=bg_app_color)
middle_left_thickness_frame = tk.Frame(middle_left_frame, bg=bg_app_color)
middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)


middle_left_weight_frame_col1_frame = tk.Frame(middle_left_weight_frame, bg=bg_param_color)
middle_left_weight_frame_col2_frame = tk.Frame(middle_left_weight_frame, bg=bg_app_color)
middle_left_weight_frame_col3_frame = tk.Frame(middle_left_weight_frame, bg=bg_param_color)

middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg=bg_param_color)
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_color)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_param_color)

middle_center_frame = tk.Frame(middle_frame, bg=bg_app_color)
middle_center_frame.place(x=0, y=0)

middle_right_frame = tk.Frame(middle_frame, bg=bg_app_color)
middle_right_frame.place(x=0, y=0)



middle_right_runcard_frame = tk.Frame(middle_right_frame, bg=bg_app_color)
middle_right_runcard_frame.place(x=0, y=0)

middle_right_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_color)
middle_right_setting_frame.place(x=0, y=0)

middle_right_advance_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_color)
middle_right_advance_setting_frame.place(x=0, y=0)

middle_right_setting_frame_row1 = tk.Frame(middle_right_setting_frame, bg=bg_app_color)
middle_right_setting_frame_row2 = tk.Frame(middle_right_setting_frame, bg=bg_param_color)
middle_right_setting_frame_row3 = tk.Frame(middle_right_setting_frame, bg=bg_app_color)

middle_right_advance_setting_frame_row1 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_color)
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_color)
middle_right_advance_setting_frame_row3 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_color)

middle_right_setting_frame_row1_col1 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_color)
middle_right_setting_frame_row1_col2 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_color)
middle_right_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row3_col1 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_color)
middle_right_setting_frame_row3_col2 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_color)

middle_right_advance_setting_frame_row1_col1 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_color)
middle_right_advance_setting_frame_row1_col2 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_color)
middle_right_advance_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_advance_setting_frame_row3_col1 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_color)
middle_right_advance_setting_frame_row3_col2 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_color)

bottom_frame = tk.Frame(root, bg=bg_app_color)
bottom_frame.place(relx=0, rely=1.0, height=50, anchor="sw")

bottom_left_frame = tk.Frame(bottom_frame, bg=bg_app_color)
bottom_right_frame = tk.Frame(bottom_frame, bg=bg_app_color)
bottom_right_frame.grid_columnconfigure(0, weight=1)


middle_right_setting_frame_row2_row1 = tk.Frame(middle_right_setting_frame_row2, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col2 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row5_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row8_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row8, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row8_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row8, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)


middle_right_advance_setting_frame_row2_col1_row8_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row8_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row8_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row8_col2, bg=bg_param_color)


"""Function"""
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
        # threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

if get_registry_value("is_current_entry", "weight") == "weight":
    showing_thickness_frame, showing_weight_frame = False, True
else:
    showing_thickness_frame, showing_weight_frame = True, False

def open_weight_frame():
    global showing_thickness_frame, showing_weight_frame
    set_registry_value("is_current_entry", "weight")
    showing_thickness_frame = False
    showing_weight_frame = True
    middle_left_thickness_frame.pack_forget()
    middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)
    middle_left_weight_frame.lift()
def open_thickness_frame():
    global showing_thickness_frame, showing_weight_frame
    set_registry_value("is_current_entry", "thickness")
    showing_weight_frame = False
    showing_thickness_frame = True
    middle_left_weight_frame.pack_forget()
    middle_left_thickness_frame.pack(fill=tk.BOTH, expand=True)
    middle_left_thickness_frame.lift()
def open_runcard_frame():
    global showing_settings, showing_runcards, showing_advance_setting
    showing_settings = False
    showing_advance_setting = False
    showing_runcards = True
    middle_right_setting_frame.pack_forget()
    middle_right_advance_setting_frame.pack_forget()
    middle_right_runcard_frame.pack(fill=tk.BOTH, expand=True)
    middle_right_runcard_frame.lift()
    set_registry_value("is_runcard_open", "1")
def open_setting_frame():
    global showing_settings, showing_runcards, showing_advance_setting
    showing_settings = True
    showing_runcards = False
    showing_advance_setting = False
    middle_right_runcard_frame.pack_forget()
    middle_right_advance_setting_frame.pack_forget()
    middle_right_setting_frame.pack(fill=tk.BOTH, expand=True)
    middle_right_setting_frame.lift()
def open_advance_setting_frame():
    global showing_settings, showing_runcards, showing_advance_setting
    showing_settings = False
    showing_advance_setting = True
    showing_runcards = False
    middle_right_runcard_frame.pack_forget()
    middle_right_setting_frame.pack_forget()
    middle_right_advance_setting_frame.pack(fill=tk.BOTH, expand=True)
    middle_right_advance_setting_frame.lift()
def close_frame():
    global showing_settings, showing_runcards, showing_advance_setting
    showing_settings = False
    showing_advance_setting = False
    showing_runcards = False
    middle_right_setting_frame.pack_forget()
    middle_right_runcard_frame.pack_forget()
    middle_right_advance_setting_frame.pack_forget()
def exit():
    set_registry_value("is_runcard_open", "0")
    root.destroy()




"""Entry"""
middle_right_setting_frame_row1_col1_label = tk.Label(middle_right_setting_frame_row1_col1, text="Cài đặt", font=(font_name, 18, "bold"), bg=bg_app_color)
middle_right_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_advance_setting_frame_row1_col1_label = tk.Label(middle_right_advance_setting_frame_row1_col1, text="Cài đặt bổ sung", font=(font_name, 18, "bold"), bg=bg_app_color)
middle_right_advance_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")




"""Settings"""
def update_com_ports(*menus):
    try:
        def monitor_com_ports():
            global com_ports
            while True:
                new_com_ports = [port.device for port in serial.tools.list_ports.comports() if "Bluetooth" not in port.description]
                if set(new_com_ports) != set(com_ports):
                    com_ports = new_com_ports
                    root.after(0, lambda: populate_com_menus(menus))
                time.sleep(2)
        def populate_com_menus(menus):
            for menu in menus:
                menu['menu'].delete(0, 'end')
                menu['menu'].add_command(label="-------", command=lambda m=menu: m.setvar(m.cget("textvariable"), value="-------"))
                for port in com_ports:
                    menu['menu'].add_command(label=port, command=lambda p=port, m=menu: m.setvar(m.cget("textvariable"), value=p))
        global com_ports
        com_ports = [port.device for port in serial.tools.list_ports.comports() if "Bluetooth" not in port.description]
        populate_com_menus(menus)
        if not hasattr(update_com_ports, "thread_started"):
            update_com_ports.thread_started = True
            com_port_thread = threading.Thread(target=monitor_com_ports, daemon=True)
            com_port_thread.start()
    except Exception as e:
        # threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def get_selected_frame():
    """Retrieve the last selected frame from the Windows Registry."""
    return get_registry_value("SelectedFrame", "Trọng lượng")

def set_selected_frame(value):
    """Save the selected frame to the Windows Registry."""
    set_registry_value("SelectedFrame", value)

weight_com_thread = None
thickness_com_thread = None
def switch_middle_left_frame(*args):
    global weight_com_thread, thickness_com_thread
    try:
        print(f"Switching to: {selected_middle_left_frame.get()}")
        if selected_middle_left_frame.get() == "Trọng lượng":
            middle_left_thickness_frame.pack_forget()
            middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)
        else:
            middle_left_weight_frame.pack_forget()
            middle_left_thickness_frame.pack(fill=tk.BOTH, expand=True)
        if "COM" in str(get_registry_value("selected_weight_com", "")):
            if weight_com_thread is None or not weight_com_thread.is_alive():
                pass
                # weight_com_thread = threading.Thread(target=weight_frame_com_port_insert_data, daemon=True)
                # weight_com_thread.start()
        if "COM" in str(get_registry_value("selected_thickness_com", "")):
            if thickness_com_thread is None or not thickness_com_thread.is_alive():
                pass
                # thickness_com_thread = threading.Thread(target=thickness_frame_com_port_insert_data, daemon=True)
                # thickness_com_thread.start()
    except Exception as e:
        # threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def save_setting_frame():
    try:
        # threading.Thread(target=show_error_message, args=(f"Save setting", 1, 3000), daemon=True).start()
        set_registry_value("COM1", selected_weight_com.get())
        set_registry_value("COM2", selected_thickness_com.get())
        # set_registry_value("ServerIP", server_ip.get())
        # set_registry_value("is_check_runcard", setting_check_runcard_switch.get())
        # set_registry_value("is_plant_name", plant_name.get())
        # set_registry_value("is_show_runcard", setting_show_runcard_switch.get())
        switch_middle_left_frame()
        set_selected_frame(selected_middle_left_frame.get())
        # database_test_connection()
        # manage_machine_buttons(middle_right_runcard_frame_col1, get_registry_value('is_plant_name', ""))
        # manage_line_buttons()
        # messagebox.showinfo("Success", "Save setting success!\nApplication will close automatically\nRe-open the app")
        # root.destroy()
    except Exception as e:
        # threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


selected_middle_left_frame = tk.StringVar(value=get_selected_frame())
switch_middle_left_frame()

selected_weight_com = tk.StringVar(value=get_registry_value("COM1", ""))
selected_thickness_com = tk.StringVar(value=get_registry_value("COM2", ""))

weight_label = tk.Label(middle_right_setting_frame_row2_row1, text="Trọng lượng:   ", font=(font_name, 18), bg='white')
weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
weight_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_weight_com, "")
weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

thickness_label = tk.Label(middle_right_setting_frame_row2_row1, text="Độ dày:", font=(font_name, 18), bg='white')
thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
thickness_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_thickness_com, "")
thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

frame_select_label = tk.Label(middle_right_setting_frame_row2_row1, text="Mặc định:", font=(font_name, 18), bg='white')
frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
frame_select_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")


advance_weight_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Trọng lượng:   ", font=(font_name, 18), bg='white')
advance_weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
advance_weight_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_weight_com, "")
advance_weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

advance_thickness_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Độ dày:", font=(font_name, 18), bg='white')
advance_thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
advance_thickness_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_thickness_com, "")
advance_thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

advance_frame_select_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Mặc định:", font=(font_name, 18), bg='white')
advance_frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
advance_frame_select_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
advance_frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")

update_com_ports(weight_menu, thickness_menu)
update_com_ports(advance_weight_menu, advance_thickness_menu)



server_ip_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row1, text="Server IP:       ", font=(font_name, 18), bg='white')
server_ip_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
server_ip = tk.StringVar(value=get_registry_value("ServerIP", ""))
server_ip_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row1, textvariable=server_ip, font=(font_name, 16), width=12, bg='white')
server_ip_entry.grid(row=0, column=0, padx=5, pady=5, sticky="w")


db_name_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row2, text="Database:", font=(font_name, 18), bg='white')
db_name_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
db_name = tk.StringVar(value=get_registry_value("Database", "PMG_DEVICE"))
db_name_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row2, textvariable=db_name, font=(font_name, 16), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
db_name_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")


user_id_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row3, text="User ID:", font=(font_name, 18), bg='white')
user_id_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
user_id = tk.StringVar(value=get_registry_value("UserID", "scadauser"))
user_id_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row3, textvariable=user_id, font=(font_name, 16), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
user_id_entry.grid(row=2, column=0, padx=5, pady=5, sticky="w")




password_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row4, text="Password:", font=(font_name, 18), bg='white')
password_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
password = tk.StringVar(value=get_registry_value("Password", "pmgscada+123"))
password_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row4, textvariable=password, font=(font_name, 16), width=12, show='*', bg='white', fg='red', readonlybackground='white', state='readonly')
password_entry.grid(row=3, column=0, padx=5, pady=5, sticky="w")


plant_name_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row8_col1_row1, text="Plant name:    ", font=(font_name, 18), bg='white')
plant_name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
plant_name = tk.StringVar(value=get_registry_value("is_plant_name", ""))
plant_name_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row8_col2_row1, textvariable=plant_name, font=(font_name, 16), width=12, bg='white')
plant_name_entry.grid(row=0, column=0, padx=5, pady=5, sticky="w")











"""Button"""
def on_enter_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_hover_icon)
def on_leave_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_icon)
top_open_weight_frame_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "weight.png")).resize((156, 36)))
top_open_weight_frame_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "weight_hover.png")).resize((156, 36)))
top_open_weight_frame_button = tk.Button(top_right_frame, image=top_open_weight_frame_button_icon, command=open_weight_frame, bg='#f0f2f6', width=156, height=36, relief="flat", borderwidth=0)
top_open_weight_frame_button.grid(row=0, column=3, padx=10, pady=5, sticky="e")
top_open_weight_frame_button.bind("<Enter>", on_enter_top_open_weight_frame_button)
top_open_weight_frame_button.bind("<Leave>", on_leave_top_open_weight_frame_button)


def on_enter_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_hover_icon)
def on_leave_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_icon)
top_open_thickness_frame_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "thickness.png")).resize((156, 36)))
top_open_thickness_frame_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "thickness_hover.png")).resize((156, 36)))
top_open_thickness_frame_button = tk.Button(top_right_frame, image=top_open_thickness_frame_button_icon, command=open_thickness_frame, bg='#f0f2f6', width=156, height=36, relief="flat", borderwidth=0)
top_open_thickness_frame_button.grid(row=0, column=4, padx=5, pady=5, sticky="e")
top_open_thickness_frame_button.bind("<Enter>", on_enter_top_open_thickness_frame_button)
top_open_thickness_frame_button.bind("<Leave>", on_leave_top_open_thickness_frame_button)


def on_enter_middle_open_menu_frame_button(event):
    middle_open_menu_frame_button.config(image=menu_hover_icon)
def on_leave_middle_open_menu_frame_button(event):
    middle_open_menu_frame_button.config(image=menu_icon)
menu_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "menu.png")).resize((42, 42)))
menu_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "menu_hover.png")).resize((42, 42)))
middle_open_menu_frame_button = tk.Button(top_left_frame, image=menu_icon, bd=0, command=None, bg=bg_app_color)
middle_open_menu_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")
middle_open_menu_frame_button.bind("<Enter>", on_enter_middle_open_menu_frame_button)
middle_open_menu_frame_button.bind("<Leave>", on_leave_middle_open_menu_frame_button)


def on_enter_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_hover_icon)
def on_leave_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_icon)
setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "setting.png")).resize((42, 42)))
setting_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "setting_hover.png")).resize((42, 42)))
middle_open_setting_frame_button = tk.Button(top_left_frame, image=setting_icon, bd=0, command=open_setting_frame, bg=bg_app_color)
middle_open_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")
middle_open_setting_frame_button.bind("<Enter>", on_enter_middle_open_setting_frame_button)
middle_open_setting_frame_button.bind("<Leave>", on_leave_middle_open_setting_frame_button)


def on_enter_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_hover_icon)
def on_leave_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_icon)
advance_setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "advance_setting.png")).resize((21, 21)))
advance_setting_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "advance_setting_hover.png")).resize((21, 21)))
middle_open_advance_setting_frame_button = tk.Button(middle_right_setting_frame_row1_col2, image=advance_setting_icon, bd=0, command=open_advance_setting_frame, bg=bg_app_color)
middle_open_advance_setting_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_advance_setting_frame_button.bind("<Enter>", on_enter_middle_open_advance_setting_frame_button)
middle_open_advance_setting_frame_button.bind("<Leave>", on_leave_middle_open_advance_setting_frame_button)


def on_enter_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_hover_icon)
def on_leave_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_icon)
return_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "return.png")).resize((21, 21)))
return_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "return_hover.png")).resize((21, 21)))
middle_open_return_frame_button = tk.Button(middle_right_advance_setting_frame_row1_col2, image=return_icon, bd=0, command=open_setting_frame, bg=bg_app_color)
middle_open_return_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_return_frame_button.bind("<Enter>", on_enter_middle_open_return_frame_button)
middle_open_return_frame_button.bind("<Leave>", on_leave_middle_open_return_frame_button)




def on_enter_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_hover_icon)
def on_leave_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_icon)
barcode_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "barcode.png")).resize((42, 42)))
barcode_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "barcode_hover.png")).resize((42, 42)))
middle_open_runcard_frame_button = tk.Button(top_left_frame, image=barcode_icon, bd=0, command=open_runcard_frame, bg=bg_app_color)
middle_open_runcard_frame_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")
middle_open_runcard_frame_button.bind("<Enter>", on_enter_middle_open_runcard_frame_button)
middle_open_runcard_frame_button.bind("<Leave>", on_leave_middle_open_runcard_frame_button)


close_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "close.png")).resize((134, 34)))
close_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "close_hover.png")).resize((134, 34)))
def on_enter_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon_hover)
def on_leave_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon)
close_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col2, image=close_icon, command=close_frame, bg=bg_app_color, width=134, height=34, relief="flat", borderwidth=0)
close_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")
close_setting_frame_button.bind("<Enter>", on_enter_close_setting_frame_button)
close_setting_frame_button.bind("<Leave>", on_leave_close_setting_frame_button)

def on_enter_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon_hover)
def on_leave_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon)
close_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col2, image=close_icon, command=close_frame, bg=bg_app_color, width=134, height=34, relief="flat", borderwidth=0)
close_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
close_advance_setting_frame_button.bind("<Enter>", on_enter_close_advance_setting_frame_button)
close_advance_setting_frame_button.bind("<Leave>", on_leave_close_advance_setting_frame_button)

save_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "save.png")).resize((134, 34)))
save_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "save_hover.png")).resize((134, 34)))
def on_enter_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon_hover)
def on_leave_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon)
save_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col1, image=save_icon, command=save_setting_frame, bg=bg_app_color, width=134, height=34, relief="flat", borderwidth=0)
save_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")
save_setting_frame_button.bind("<Enter>", on_enter_save_setting_frame_button)
save_setting_frame_button.bind("<Leave>", on_leave_save_setting_frame_button)


def on_enter_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon_hover)
def on_leave_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon)
exit_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "exit.png")).resize((134, 31)))
exit_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "exit_hover.png")).resize((134, 31)))
bottom_exit_button = tk.Button(bottom_right_frame, image=exit_icon, command=exit, bg='#f4f4fe', width=134, height=31, relief="flat", borderwidth=0)
bottom_exit_button.grid(row=0, column=0, columnspan=5, padx=5, pady=5, sticky="e")
bottom_exit_button.bind("<Enter>", on_enter_bottom_exit_button)
bottom_exit_button.bind("<Leave>", on_leave_bottom_exit_button)



"""Update loop"""
update_thread = threading.Thread(target=update_dimensions, daemon=True)
update_thread.start()
root.mainloop()