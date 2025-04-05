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
from tkcalendar import DateEntry
from cefpython3 import cefpython as cef
from ctypes import wintypes
import platform
cef_initialized = False
browser = None
barcode.base.Barcode.default_writer_options['write_text'] = False

sys.excepthook = cef.ExceptHook
cef.Initialize()
base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")

"""Set DPI awareness for better scaling on high-DPI screens"""
ctypes.windll.shcore.SetProcessDpiAwareness(10)
REG_PATH = r"SOFTWARE\IPQC\Config"

user32 = ctypes.windll.user32
monitor_width, monitor_height = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

"""Main window setup"""
root = tk.Tk()
root.title("IPQC v.25.03.29")
start_time = time.time()
"""Define scaling factors"""
user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
dpi = user32.GetDpiForSystem()
scaling = dpi / 96
screen_width = min(1600, int(0.9*root.winfo_screenwidth()*scaling))
screen_height = min(900, int(0.9*root.winfo_screenheight()*scaling))

root.iconbitmap("theme/icons/logo.ico")
root.minsize(screen_width, screen_height)
root.geometry(f"{screen_width}x{screen_height}")
root.protocol("WM_DELETE_WINDOW", lambda: None)

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
fg_app_class_color_layer_2 = '#ffffff'


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
weight_record_log_id = 0

weight_com_thread = None
thickness_com_thread = None


current_date = (datetime.datetime.now() - datetime.timedelta(hours=5) + datetime.timedelta(minutes=22)).date()
current_time = str(int((datetime.datetime.now() + datetime.timedelta(minutes=39)).strftime('%H')))
period_times = ['6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17',
                '18', '19', '20', '21', '22', '23', '0', '1', '2', '3', '4', '5']

selected_date = current_date
runcard_selected = ''

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
def get_registry_value(name, default=""):
    """Retrieve a value from the Windows Registry."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, name)[0]
    except FileNotFoundError as e:
        threading.Thread(target=show_error_message, args=(f"def get_registry_value => {e}", 0, 3000), daemon=True).start()
        return default
def set_registry_value(name, value):
    """Set a value in the Windows Registry."""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def set_registry_value => {e}", 0, 3000), daemon=True).start()

def hook_titlebar_click_to_shutdown_cef(hwnd):
    GWL_WNDPROC = -4
    WM_NCLBUTTONDOWN = 0x00A1

    WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, ctypes.c_uint, ctypes.c_int, ctypes.c_int)
    user32 = ctypes.windll.user32
    original_wndproc = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, ctypes.c_uint, ctypes.c_int, ctypes.c_int)()

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == WM_NCLBUTTONDOWN:
            print("🛑 Title bar clicked! Shutting down CEF...")
            try:
                cef.Shutdown()
            except:
                pass
        return user32.CallWindowProcW(original_wndproc, hwnd, msg, wparam, lparam)

    hwnd = root.winfo_id()
    wnd_proc_pointer = WNDPROC(wnd_proc)
    original_wndproc = user32.SetWindowLongW(hwnd, GWL_WNDPROC, wnd_proc_pointer)
if platform.system() == "Windows":
    hook_titlebar_click_to_shutdown_cef(root.winfo_id())




class CustomOptionMenu(tk.OptionMenu):
    def __init__(self, master, variable, *options, command=None, **kwargs):
        if options is None:
            options = ["No COM"]
        super().__init__(master, variable, *options, **kwargs)
        self.config(font=("Helvetica", 14, "bold"),
                    bg=bg_app_class_color_layer_2,
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


count = 0
cef_loaded = False
def update_dimensions():
    global screen_width, screen_height, showing_settings, showing_runcards, count, cef_loaded
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

        bottom_left_frame.place(x=0, y=0, width=int(screen_width * 0.7), height=50)
        bottom_right_frame.place(x=int(screen_width * 0.7), y=0, width=int(screen_width * 0.3), height=50)

        middle_left_weight_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)
        middle_left_thickness_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        middle_left_weight_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_weight_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_weight_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_left_thickness_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_thickness_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_thickness_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_left_weight_frame_col3_frame_row1.place(x=0, y=0, width=210, height=140)
        middle_left_weight_frame_col3_frame_row2.place(x=0, y=140, width=210, height=middle_frame_height-140)

        middle_left_thickness_frame_col3_frame_row1.place(x=0, y=0, width=210, height=140)
        middle_left_thickness_frame_col3_frame_row2.place(x=0, y=140, width=210, height=middle_frame_height-140)

        middle_left_weight_frame_col3_frame_row1_row1.place(x=0, y=20, width=120, height=40)
        middle_left_weight_frame_col3_frame_row1_row2.place(x=0, y=80, width=40, height=40)

        middle_left_thickness_frame_col3_frame_row1_row1.place(x=0, y=20, width=120, height=40)
        middle_left_thickness_frame_col3_frame_row1_row2.place(x=0, y=80, width=40, height=40)

        middle_left_weight_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_weight_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)

        middle_left_thickness_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_thickness_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)

        middle_left_weight_frame_col1_frame_row2_canvas.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height - 80)
        middle_left_weight_frame_col1_frame_row2_scrollbar.place(x=middle_left_col1_width - 20, y=0, width=20, height=middle_frame_height - 80)
        middle_left_weight_frame_col1_frame_row2_scrollable_frame.place(x=0, y=0, width=middle_left_col1_width - 20, height=middle_frame_height - 80)

        middle_left_thickness_frame_col1_frame_row2_canvas.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height - 80)
        middle_left_thickness_frame_col1_frame_row2_scrollbar.place(x=middle_left_col1_width - 20, y=0, width=20, height=middle_frame_height - 80)
        middle_left_thickness_frame_col1_frame_row2_scrollable_frame.place(x=0, y=0, width=middle_left_col1_width - 20, height=middle_frame_height - 80)

        middle_left_weight_frame_col3_frame_row2_canvas.place(x=0, y=0, width=210, height=middle_frame_height-140)
        middle_left_weight_frame_col3_frame_row2_scrollbar.place(x=190, y=0, width=20, height=middle_frame_height - 140)
        middle_left_weight_frame_col3_frame_row2_scrollable_frame.place(x=0, y=0, width=190, height=middle_frame_height - 140)

        middle_right_runcard_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_advance_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)

        middle_right_runcard_frame_row1.place(x=0, y=0, width=middle_right_frame_width-10, height=40)
        middle_right_runcard_frame_row2.place(x=0, y=40, width=middle_right_frame_width-10, height=720)
        middle_right_runcard_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width-5, height=40)

        if not cef_loaded and middle_right_runcard_frame_row2.winfo_width() > 100:
            run_cef_in_frame(middle_right_runcard_frame_row2, "http://10.13.104.181:10000/")
            cef_loaded = True
        # middle_right_runcard_frame_row2_col1.place(x=0, y=0, width=36, height=middle_frame_height - 40 - 40)
        # middle_right_runcard_frame_row2_col2.place(x=36, y=0, width=middle_right_frame_width - 10 - 36 - 36, height=middle_frame_height - 40 - 40)
        # middle_right_runcard_frame_row2_col3.place(x=middle_right_frame_width- 10 - 36, y=0, width=36, height=middle_frame_height - 40 - 40)

        # middle_right_runcard_frame_row2_col2_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 36 - 36, height=250)
        # middle_right_runcard_frame_row2_col2_row2.place(x=0, y=250, width=middle_right_frame_width - 10 - 36 - 36, height=50)
        # middle_right_runcard_frame_row2_col2_row3.place(x=0, y=300, width=middle_right_frame_width - 10 - 36 - 36, height=50)
        # middle_right_runcard_frame_row2_col2_row4.place(x=0, y=350, width=middle_right_frame_width - 10 - 36 - 36, height=200)
        # middle_right_runcard_frame_row2_col2_row5.place(x=0, y=550, width=middle_right_frame_width - 10 - 36 - 36, height=200)
        #
        # middle_right_runcard_frame_row2_col2_row1_row1.place(x=10, y=0, width=middle_right_frame_width - 10 - 36 - 36 - 20, height=250)

        # middle_right_runcard_frame_row2_col2_row4_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 36 - 36, height=30)
        # middle_right_runcard_frame_row2_col2_row4_row2.place(x=0, y=30, width=middle_right_frame_width - 10 - 36 - 36, height=80)
        # middle_right_runcard_frame_row2_col2_row4_row3.place(x=0, y=110, width=middle_right_frame_width - 10 - 36 - 36, height=30)
        # print(f"-----> {middle_right_runcard_frame_row2.winfo_height()}")
        # print(f"-----> {middle_right_runcard_frame_row2.winfo_width()}")
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
        middle_right_advance_setting_frame_row2_col3.place(x=342, y=0, width=middle_right_frame_width - 10 - 332 - 5, height=middle_frame_height - 40 - 50)

        middle_right_advance_setting_frame_row2_col1_row1.place(x=0, y=0, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row2.place(x=0, y=20, width=332, height=160)
        middle_right_advance_setting_frame_row2_col1_row3.place(x=0, y=180, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row4.place(x=0, y=200, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row5.place(x=0, y=210, width=332, height=160)
        middle_right_advance_setting_frame_row2_col1_row6.place(x=0, y=370, width=332, height=60)
        middle_right_advance_setting_frame_row2_col1_row7.place(x=0, y=430, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row8.place(x=0, y=450, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row9.place(x=0, y=460, width=332, height=40)
        middle_right_advance_setting_frame_row2_col1_row10.place(x=0, y=500, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row11.place(x=0, y=510, width=332, height=20)
        middle_right_advance_setting_frame_row2_col1_row12.place(x=0, y=530, width=332, height=30)
        middle_right_advance_setting_frame_row2_col1_row13.place(x=0, y=560, width=332, height=10)
        middle_right_advance_setting_frame_row2_col1_row14.place(x=0, y=570, width=332, height=100)
        middle_right_advance_setting_frame_row2_col1_row15.place(x=0, y=670, width=332, height=10)


        middle_right_advance_setting_frame_row2_col1_row5_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row5_col2.place(x=170, y=0, width=162, height=180)

        middle_right_advance_setting_frame_row2_col1_row6_col1.place(x=0, y=0, width=170, height=60)
        middle_right_advance_setting_frame_row2_col1_row6_col2.place(x=170, y=0, width=162, height=60)

        middle_right_advance_setting_frame_row2_col1_row9_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row9_col2.place(x=170, y=0, width=162, height=180)

        root.after(0, lambda: middle_right_advance_setting_frame_row2_col1_row5_col1_row1.place(x=0, y=0, width=170, height=40))
        root.after(0, lambda: middle_right_advance_setting_frame_row2_col1_row5_col1_row2.place(x=0, y=40, width=170, height=40))
        root.after(0, lambda: middle_right_advance_setting_frame_row2_col1_row5_col1_row3.place(x=0, y=80, width=170, height=40))
        root.after(0, lambda: middle_right_advance_setting_frame_row2_col1_row5_col1_row4.place(x=0, y=120, width=170, height=40))


        middle_right_advance_setting_frame_row2_col1_row5_col2_row1.place(x=0, y=0, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row2.place(x=0, y=40, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row3.place(x=0, y=80, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row4.place(x=0, y=120, width=162, height=40)

        middle_right_advance_setting_frame_row2_col1_row9_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row9_col2_row1.place(x=0, y=0, width=162, height=40)


        middle_right_advance_setting_frame_row2_col3_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 332 - 10, height=30)
        middle_right_advance_setting_frame_row2_col3_row2.place(x=0, y=30, width=middle_right_frame_width - 10 - 332 - 10, height=10)
        middle_right_advance_setting_frame_row2_col3_row3.place(x=0, y=40, width=middle_right_frame_width - 10 - 332 - 10, height=200)

        middle_right_advance_setting_frame_row2_col3_row4.place(x=0, y=240, width=middle_right_frame_width - 10 - 332 - 10, height=30)

        middle_right_advance_setting_frame_row2_col3_row5.place(x=0, y=270, width=middle_right_frame_width - 10 - 332 - 10, height=30)
        middle_right_advance_setting_frame_row2_col3_row6.place(x=0, y=300, width=middle_right_frame_width - 10 - 332 - 10, height=10)
        middle_right_advance_setting_frame_row2_col3_row7.place(x=0, y=310, width=middle_right_frame_width - 10 - 332 - 10, height=200)

        middle_right_advance_setting_frame_row2_col3_row8.place(x=0, y=510, width=middle_right_frame_width - 10 - 332 - 10, height=30)

        middle_right_advance_setting_frame_row2_col3_row9.place(x=0, y=540, width=middle_right_frame_width - 10 - 332 - 10, height=30)
        middle_right_advance_setting_frame_row2_col3_row10.place(x=0, y=570, width=middle_right_frame_width - 10 - 332 - 10, height=10)
        middle_right_advance_setting_frame_row2_col3_row11.place(x=0, y=580, width=middle_right_frame_width - 10 - 332 - 10, height=200)


        middle_right_advance_setting_frame_row2_col3_row3_11.place(x=0, y=0, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row3_12.place(x=110, y=0, width=80, height=40)
        middle_right_advance_setting_frame_row2_col3_row3_21.place(x=0, y=40, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row3_22.place(x=110, y=40, width=70, height=40)
        middle_right_advance_setting_frame_row2_col3_row3_31.place(x=0, y=80, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row3_32.place(x=110, y=80, width=70, height=40)

        middle_right_advance_setting_frame_row2_col3_row7_11.place(x=0, y=0, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row7_12.place(x=110, y=0, width=80, height=40)
        middle_right_advance_setting_frame_row2_col3_row7_21.place(x=0, y=40, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row7_22.place(x=110, y=40, width=70, height=40)
        middle_right_advance_setting_frame_row2_col3_row7_31.place(x=0, y=80, width=110, height=40)
        middle_right_advance_setting_frame_row2_col3_row7_32.place(x=110, y=80, width=70, height=40)

        middle_right_advance_setting_frame_row2_col3_row11_11.place(x=0, y=0, width=145, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_12.place(x=145, y=0, width=50, height=40)

        middle_right_advance_setting_frame_row2_col3_row11_21.place(x=0, y=40, width=145, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_22.place(x=145, y=40, width=50, height=40)

        middle_right_advance_setting_frame_row2_col3_row11_31.place(x=0, y=80, width=145, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_32.place(x=145, y=80, width=50, height=40)


        error_display_entry.place(x=5, y=10, width=int(screen_width * 0.7), height=25)

        if hasattr(weight_frame_write_insert_value, "tree"):
            treeview_width = middle_left_weight_frame_col1_frame_row2_canvas.winfo_width()
            treeview_height = middle_left_weight_frame_col1_frame_row2_canvas.winfo_height()
            weight_frame_write_insert_value.tree.config(height=int(treeview_height / 20))
            weight_frame_write_insert_value.tree.pack(fill="both", expand=True)
            for index, col in enumerate(weight_frame_write_insert_value.tree["columns"]):
                if index == 0:
                    record_width = int(treeview_width * 0.1)
                elif index == 5:
                    record_width = int(treeview_width * 0.3)
                else:
                    record_width = int((treeview_width * 0.6) / 4)
                weight_frame_write_insert_value.tree.column(col, width=record_width, anchor="center")

        if hasattr(thickness_frame_write_insert_value, "tree"):
            treeview_width = middle_left_thickness_frame_col1_frame_row2_canvas.winfo_width()
            treeview_height = middle_left_thickness_frame_col1_frame_row2_canvas.winfo_height()
            thickness_frame_write_insert_value.tree.config(height=int(treeview_height / 20))
            thickness_frame_write_insert_value.tree.pack(fill="both", expand=True)
            for index, col in enumerate(thickness_frame_write_insert_value.tree["columns"]):
                if index == 0:
                    record_width = int(treeview_width * 0.1)
                elif index == 1:
                    record_width = int(treeview_width * 0.25)
                else:
                    record_width = int((treeview_width * 0.65) / 5)
                thickness_frame_write_insert_value.tree.column(col, width=record_width, anchor="center")

        root.update_idletasks()
        root.update()
        if count == 0:
            print(f"-->==>===>{time.time() - start_time}")
            count = 1

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
        threading.Thread(target=show_error_message, args=(f"def convert_to_uppercase => {e}", 0, 3000), daemon=True).start()
        pass

conn_str = (
    f'DRIVER={{SQL Server}};'
    f'SERVER={get_registry_value("is_server_ip", "10.13.102.22")};'
    f'DATABASE={get_registry_value("is_db_name", "PMG_DEVICE")};'
    f'UID={get_registry_value("is_user_id", "scadauser")};'
    f'PWD={get_registry_value("is_password", "pmgscada+123")};'
)

"""Frame"""
top_frame = tk.Frame(root, bg=bg_app_class_color_layer_1)
top_frame.place(relx=0, rely=0, height=50, anchor="nw")

top_left_frame = tk.Frame(top_frame, bg=bg_app_class_color_layer_1)
top_right_frame = tk.Frame(top_frame, height=50, bg=bg_app_class_color_layer_1)
top_right_frame.grid_columnconfigure(0, weight=1)





middle_frame = tk.Frame(root, bg=bg_app_class_color_layer_1)
middle_frame.place(relx=0, rely=0, anchor="nw")

middle_left_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_left_frame.place(x=0, y=0)



middle_left_weight_frame = tk.Frame(middle_left_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame = tk.Frame(middle_left_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)


middle_left_weight_frame_col1_frame = tk.Frame(middle_left_weight_frame, bg=bg_app_class_color_layer_2)
middle_left_weight_frame_col2_frame = tk.Frame(middle_left_weight_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame = tk.Frame(middle_left_weight_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_2 )
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)


middle_left_weight_frame_col3_frame_row1 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row2 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col3_frame_row1 = tk.Frame(middle_left_thickness_frame_col3_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row2 = tk.Frame(middle_left_thickness_frame_col3_frame, bg=bg_app_class_color_layer_1)



middle_left_weight_frame_col3_frame_row1_row1 = tk.Frame(middle_left_weight_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row1_row2 = tk.Frame(middle_left_weight_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)


middle_left_thickness_frame_col3_frame_row1_row1 = tk.Frame(middle_left_thickness_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row1_row2 = tk.Frame(middle_left_thickness_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)


middle_left_weight_frame_col1_frame_row1 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col1_frame_row2 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)



middle_left_thickness_frame_col1_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_class_color_layer_1 )
middle_left_thickness_frame_col1_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_class_color_layer_1 )






middle_center_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_center_frame.place(x=0, y=0)

middle_right_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_right_frame.place(x=0, y=0)



middle_right_runcard_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame.place(x=0, y=0)

middle_right_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_setting_frame.place(x=0, y=0)

middle_right_advance_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame.place(x=0, y=0)

middle_right_runcard_frame_row1 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1 )
middle_right_runcard_frame_row2 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_2 )
middle_right_runcard_frame_row3 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1 )
middle_right_runcard_frame_row3.grid_columnconfigure(0, weight=1)


middle_right_runcard_frame_row2_col1 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row2_col3 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)

#
# middle_right_runcard_frame_row2_col2_row1 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row2 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row3 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row4 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row5 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
#
# middle_right_runcard_frame_row2_col2_row1_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row1, bg=bg_app_class_color_layer_2)

# middle_right_runcard_frame_row2_col2_row4_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row4_row2 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)
# middle_right_runcard_frame_row2_col2_row4_row3 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)


# bg_image = Image.open("theme/images/button-disabled.png")
# bg_image = bg_image.resize((400, 300), Image.LANCZOS)
# bg_photo = ImageTk.PhotoImage(bg_image)
# bg_label = tk.Label(middle_right_runcard_frame_row2_col2_row5, image=bg_photo)
# bg_label.place(x=0, y=0, relwidth=1, relheight=1)


middle_right_setting_frame_row1 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1 )
middle_right_setting_frame_row2 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_2 )
middle_right_setting_frame_row3 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1 )

middle_right_advance_setting_frame_row1 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row3 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1 )

middle_right_setting_frame_row1_col1 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1 )
middle_right_setting_frame_row1_col2 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1 )
middle_right_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row3_col1 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1 )
middle_right_setting_frame_row3_col2 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1 )

middle_right_advance_setting_frame_row1_col1 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row1_col2 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_advance_setting_frame_row3_col1 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row3_col2 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_class_color_layer_1 )

bottom_frame = tk.Frame(root, bg=bg_app_class_color_layer_1 )
bottom_frame.place(relx=0, rely=1.0, height=50, anchor="sw")

bottom_left_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1 )
bottom_right_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1 )
bottom_right_frame.grid_columnconfigure(0, weight=1)


middle_right_setting_frame_row2_row1 = tk.Frame(middle_right_setting_frame_row2, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col2 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )

middle_right_advance_setting_frame_row2_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row12 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row13 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row14 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col1_row15 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_class_color_layer_1 )




middle_right_advance_setting_frame_row2_col1_row5_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row6_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row6_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row9_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row9_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row9_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row9_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col2, bg=bg_app_class_color_layer_2 )





middle_right_advance_setting_frame_row2_col3_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_class_color_layer_2 )


middle_right_advance_setting_frame_row2_col3_row3_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row3_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_app_class_color_layer_2 )


middle_right_advance_setting_frame_row2_col3_row7_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row7_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_app_class_color_layer_2 )


middle_right_advance_setting_frame_row2_col3_row11_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row11_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col3_row11_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row11_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col3_row11_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col3_row11_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_app_class_color_layer_2 )












""""""
middle_right_setting_frame_row1_col1_label = tk.Label(middle_right_setting_frame_row1_col1, text="Cài đặt", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
middle_right_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_advance_setting_frame_row1_col1_label = tk.Label(middle_right_advance_setting_frame_row1_col1, text="Cài đặt nâng cao", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
middle_right_advance_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_runcard_frame_label = tk.Label(middle_right_runcard_frame_row1, text="Runcard", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
middle_right_runcard_frame_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")




if get_registry_value("is_current_entry", "weight") == "weight":
    showing_thickness_frame, showing_weight_frame = False, True
else:
    showing_thickness_frame, showing_weight_frame = True, False





error_display_entry = tk.Entry(bottom_left_frame, textvariable=error_msg, font=("Cambria", 12), bg=bg_app_class_color_layer_1 , fg=error_fg_color, bd=0, highlightthickness=0, readonlybackground=bg_app_class_color_layer_1 , state="readonly")


def weight_frame_write_insert_value(device_id, operator_id, runcard_id, weight_value):
    try:
        global weight_record_id
        weight_record_id += 1
        root.update_idletasks()
        root.update()
        frame_width = int(middle_left_weight_frame_col1_frame_row2_canvas.winfo_width())
        if not hasattr(weight_frame_write_insert_value, "tree"):
            columns = ("ID", "Device ID", "Operator ID", "Runcard ID", "Weight", "Timestamp")
            style = ttk.Style()
            style.theme_use("classic")
            style.configure("Treeview.Heading", font=("Arial", 12, "bold"), relief="flat", background="white",
                            foreground="black", borderwidth=1, highlightthickness=0)
            style.configure("Treeview", font=("Cambria", 13), borderwidth=0, relief="flat", background="white",
                            fieldbackground="white")
            style.map("Treeview", background=[("selected", "lightblue")])
            weight_frame_write_insert_value.tree = ttk.Treeview(middle_left_weight_frame_col1_frame_row2_scrollable_frame, columns=columns, show="headings", style="Treeview", height=int(middle_left_weight_frame_col1_frame_row2_scrollable_frame.winfo_height() - 80))
            for index, col in enumerate(columns):
                if index == 0:
                    record_width = int(frame_width * 0.1)
                elif index == 5:
                    record_width = int(frame_width * 0.3)
                else:
                    record_width = int(frame_width * 0.6 / 4)
                weight_frame_write_insert_value.tree.heading(col, text=col)
                weight_frame_write_insert_value.tree.column(col, width=record_width, anchor="center")
            weight_frame_write_insert_value.tree.pack(fill="both", expand=True)
        weight_frame_write_insert_value.tree.insert("", "0", values=( weight_record_id, device_id, operator_id, runcard_id, weight_value, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()


def weight_insert_data_to_db(device_name, runcard_id, weight_value, operator_id):
    global conn_str
    def insert_data():
        try:
            sql = f"""
                insert into [PMG_DEVICE].[dbo].[WeightDeviceData] (DeviceId, LotNo, Weight, UserId, CreationDate)
                values ('{device_name}', '{runcard_id}', {weight_value}, {operator_id}, '{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}')
            """
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute(sql)
                conn.commit()
                threading.Thread(target=show_error_message, args=(f"Inserted {entry_weight_runcard_id_entry.get()} - {entry_weight_weight_value_entry.get()}g to database!", 1, 2000), daemon=True).start()
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
    if all([device_name, runcard_id, weight_value, operator_id]):
        threading.Thread(target=insert_data, daemon=True).start()
    else:
        threading.Thread(target=show_error_message, args=("Empty field detected!", 0, 3000), daemon=True).start()


def weight_frame_mouser_pointer_in(event):
    global current_weight_entry
    if 'name_var' in event.widget.__dict__:
        current_weight_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")

def weight_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_weight_value_entry":
            weight_insert_data_to_db(entry_weight_device_name_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get(), entry_weight_operator_id_entry.get())
            weight_frame_write_insert_value(entry_weight_device_name_entry.get(), entry_weight_operator_id_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get())
            entry_weight_weight_value_entry.delete(0, tk.END)
            entry_weight_runcard_id_entry.delete(0, tk.END)
            entry_weight_runcard_id_entry.focus_set()
        else:
            event.widget.tk_focusNext().focus()
        return "break"
    except:
        threading.Thread(target=show_error_message, args=(f"def weight_frame_hit_enter_button => {e}", 0, 3000), daemon=True).start()



entry_weight_device_name_var = tk.StringVar()
entry_weight_device_name_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_device_name_var, 12, 1))
entry_weight_device_name_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Device ID", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 16, "bold"))
entry_weight_device_name_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_device_name_var, bg=bg_app_class_color_layer_2)
entry_weight_device_name_entry.name_var = "entry_weight_device_name_entry"
entry_weight_device_name_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_device_name_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(0, weight=1)



entry_weight_operator_id_var = tk.StringVar()
entry_weight_operator_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_operator_id_var, 5, 0))
entry_weight_operator_id_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Operator ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_operator_id_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_operator_id_var)
entry_weight_operator_id_entry.name_var = "entry_weight_operator_id_entry"
entry_weight_operator_id_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_operator_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(1, weight=1)



entry_weight_runcard_id_var = tk.StringVar()
entry_weight_runcard_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_runcard_id_var, 10, 1))
entry_weight_runcard_id_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Runcard ID", bg="#f4f4fe", bd=0, font=("Helvetica", 16, "bold"))
entry_weight_runcard_id_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_runcard_id_var)
entry_weight_runcard_id_entry.name_var = "entry_weight_runcard_id_entry"
entry_weight_runcard_id_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_runcard_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(2, weight=1)


entry_weight_weight_value_var = tk.StringVar()
entry_weight_weight_value_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_weight_value_var, 10, 0))
entry_weight_weight_value_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Trọng lượng", bg="#f4f4fe", bd=0, font=("Helvetica", 16, "bold"))
entry_weight_weight_value_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_weight_value_var)
entry_weight_weight_value_entry.name_var = "entry_weight_weight_value_entry"
entry_weight_weight_value_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_weight_value_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(3, weight=1)




middle_left_weight_frame_col1_frame_row2_canvas = tk.Canvas(middle_left_weight_frame_col1_frame_row2, bg=bg_app_class_color_layer_2 , highlightthickness=0)
middle_left_weight_frame_col1_frame_row2_scrollbar = tk.Scrollbar(middle_left_weight_frame_col1_frame_row2, orient="vertical", command=middle_left_weight_frame_col1_frame_row2_canvas.yview)
middle_left_weight_frame_col1_frame_row2_canvas.configure(yscrollcommand=middle_left_weight_frame_col1_frame_row2_scrollbar.set)
middle_left_weight_frame_col1_frame_row2_scrollable_frame = tk.Frame(middle_left_weight_frame_col1_frame_row2_canvas, bg=bg_app_class_color_layer_2)
middle_left_weight_frame_col1_frame_row2_canvas.create_window((0, 0), window=middle_left_weight_frame_col1_frame_row2_scrollable_frame, anchor="nw")
middle_left_weight_frame_col1_frame_row2_scrollable_frame.bind("<Configure>", lambda e: middle_left_weight_frame_col1_frame_row2_canvas.configure(scrollregion=middle_left_weight_frame_col1_frame_row2_canvas.bbox("all")))
middle_left_weight_frame_col1_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_col1_frame_row2_scrollbar.pack(side="right", fill="y")
all_entries = []















def thickness_frame_mouser_pointer_in(event):
    global current_thickness_entry
    if 'name_var' in event.widget.__dict__:
        current_thickness_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")

def thickness_frame_write_insert_value(runcard_id, cuon_bien, co_tay, ban_tay, ngon_tay, dau_ngon_tay):
    try:
        global thickness_record_id
        thickness_record_id += 1
        root.update_idletasks()
        root.update()
        frame_width = int(middle_left_thickness_frame_col1_frame_row2_canvas.winfo_width())
        if not hasattr(thickness_frame_write_insert_value, "tree"):
            columns = ("ID", "Runcard ID", "Cuốn biên", "Cổ tay", "Bàn tay", "Ngón tay", "Đầu ngón tay")
            style = ttk.Style()
            style.theme_use("classic")
            style.configure("Treeview.Heading", font=("Arial", 12, "bold"), relief="flat", background="white", foreground="black", borderwidth=1, highlightthickness=0)
            style.configure("Treeview", font=("Cambria", 13), borderwidth=0, relief="flat", background="white", fieldbackground="white")
            style.map("Treeview", background=[("selected", "lightblue")])
            thickness_frame_write_insert_value.tree = ttk.Treeview(middle_left_thickness_frame_col1_frame_row2_scrollable_frame, columns=columns, show="headings", style="Treeview", height=int(middle_left_thickness_frame_col1_frame_row2_canvas.winfo_height() - 80))
            for index, col in enumerate(columns):
                if index == 0:
                    record_width = int(frame_width*0.1)
                elif index == 1:
                    record_width = int(frame_width*0.2)
                else:
                    record_width = int(frame_width*0.7/5)
                thickness_frame_write_insert_value.tree.heading(col, text=col)
                thickness_frame_write_insert_value.tree.column(col, width=record_width, anchor="center")
            thickness_frame_write_insert_value.tree.pack(fill="both", expand=True)
        thickness_frame_write_insert_value.tree.insert("", "0", values=(thickness_record_id, runcard_id, cuon_bien, co_tay, ban_tay, ngon_tay, dau_ngon_tay))
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()


def thickness_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_thickness_dau_ngon_tay_entry":
            if all([entry_thickness_runcard_id_entry.get(), entry_thickness_cuon_bien_entry.get(),
                    entry_thickness_co_tay_entry.get(), entry_thickness_ban_tay_entry.get(),
                    entry_thickness_ngon_tay_entry.get(), current_widget.get()]):
                thickness_frame_write_insert_value(entry_thickness_runcard_id_entry.get(),
                                                   entry_thickness_cuon_bien_entry.get(),
                                                   entry_thickness_co_tay_entry.get(),
                                                   entry_thickness_ban_tay_entry.get(),
                                                   entry_thickness_ngon_tay_entry.get(),
                                                   current_widget.get())
                entry_thickness_dau_ngon_tay_entry.delete(0, tk.END)
                entry_thickness_ngon_tay_entry.delete(0, tk.END)
                entry_thickness_ban_tay_entry.delete(0, tk.END)
                entry_thickness_co_tay_entry.delete(0, tk.END)
                entry_thickness_cuon_bien_entry.delete(0, tk.END)
                entry_thickness_runcard_id_entry.delete(0, tk.END)
                entry_thickness_runcard_id_entry.focus_set()
        else:
            if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_thickness_runcard_id_entry":
                event.widget.tk_focusNext().focus()
            else:
                event.widget.tk_focusNext().focus()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()


entry_thickness_runcard_id_var = tk.StringVar()
entry_thickness_runcard_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_runcard_id_var, 12, 1))
entry_thickness_runcard_id_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Runcard ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_runcard_id_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_runcard_id_var)
entry_thickness_runcard_id_entry.name_var = "entry_thickness_runcard_id_entry"
entry_thickness_runcard_id_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_runcard_id_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(0, weight=1)



entry_thickness_cuon_bien_var = tk.StringVar()
entry_thickness_cuon_bien_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_cuon_bien_var, 6, 0))
entry_thickness_cuon_bien_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Cuốn biên", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_cuon_bien_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_cuon_bien_var)
entry_thickness_cuon_bien_entry.name_var = "entry_thickness_cuon_bien_entry"
entry_thickness_cuon_bien_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_cuon_bien_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(1, weight=1)



entry_thickness_co_tay_var = tk.StringVar()
entry_thickness_co_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_co_tay_var, 6, 0))
entry_thickness_co_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Cổ tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_co_tay_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_co_tay_var)
entry_thickness_co_tay_entry.name_var = "entry_thickness_co_tay_entry"
entry_thickness_co_tay_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_co_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(2, weight=1)



entry_thickness_ban_tay_var = tk.StringVar()
entry_thickness_ban_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ban_tay_var, 6, 0))
entry_thickness_ban_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Bàn tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_ban_tay_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_ban_tay_var)
entry_thickness_ban_tay_entry.name_var = "entry_thickness_ban_tay_entry"
entry_thickness_ban_tay_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ban_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(3, weight=1)



entry_thickness_ngon_tay_var = tk.StringVar()
entry_thickness_ngon_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ngon_tay_var, 6, 0))
entry_thickness_ngon_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Ngón tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_ngon_tay_label.grid(row=0, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_ngon_tay_var)
entry_thickness_ngon_tay_entry.name_var = "entry_thickness_ngon_tay_entry"
entry_thickness_ngon_tay_entry.grid(row=1, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(4, weight=1)



entry_thickness_dau_ngon_tay_var = tk.StringVar()
entry_thickness_dau_ngon_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_dau_ngon_tay_var, 6, 0))
entry_thickness_dau_ngon_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Đầu ngón tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_dau_ngon_tay_label.grid(row=0, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_thickness_dau_ngon_tay_var)
entry_thickness_dau_ngon_tay_entry.name_var = "entry_thickness_dau_ngon_tay_entry"
entry_thickness_dau_ngon_tay_entry.grid(row=1, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_dau_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(5, weight=1)





middle_left_thickness_frame_col1_frame_row2_canvas = tk.Canvas(middle_left_thickness_frame_col1_frame_row2, bg=bg_app_class_color_layer_2 , highlightthickness=0)
middle_left_thickness_frame_col1_frame_row2_scrollbar = tk.Scrollbar(middle_left_thickness_frame_col1_frame_row2, orient="vertical", command=middle_left_thickness_frame_col1_frame_row2_canvas.yview)
middle_left_thickness_frame_col1_frame_row2_canvas.configure(yscrollcommand=middle_left_thickness_frame_col1_frame_row2_scrollbar.set)
middle_left_thickness_frame_col1_frame_row2_scrollable_frame = tk.Frame(middle_left_thickness_frame_col1_frame_row2_canvas, bg=bg_app_class_color_layer_2)
middle_left_thickness_frame_col1_frame_row2_canvas.create_window((0, 0), window=middle_left_thickness_frame_col1_frame_row2_scrollable_frame, anchor="nw")
middle_left_thickness_frame_col1_frame_row2_scrollable_frame.bind("<Configure>", lambda e: middle_left_thickness_frame_col1_frame_row2_canvas.configure(scrollregion=middle_left_thickness_frame_col1_frame_row2_canvas.bbox("all")))
middle_left_thickness_frame_col1_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_thickness_frame_col1_frame_row2_scrollbar.pack(side="right", fill="y")
all_entries = []



middle_left_weight_frame_col3_frame_row2_canvas = tk.Canvas(middle_left_weight_frame_col3_frame_row2, bg=bg_app_class_color_layer_2, highlightthickness=0)
middle_left_weight_frame_col3_frame_row2_scrollbar = tk.Scrollbar(middle_left_weight_frame_col3_frame_row2, orient="vertical", command=middle_left_weight_frame_col3_frame_row2_canvas.yview)
middle_left_weight_frame_col3_frame_row2_scrollable_frame = tk.Frame(middle_left_weight_frame_col3_frame_row2_canvas, bg=bg_app_class_color_layer_2)
middle_left_weight_frame_col3_frame_row2_canvas.create_window((0, 0), window=middle_left_weight_frame_col3_frame_row2_scrollable_frame, anchor="nw")
middle_left_weight_frame_col3_frame_row2_canvas.configure(yscrollcommand=middle_left_weight_frame_col3_frame_row2_scrollbar.set)
middle_left_weight_frame_col3_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_col3_frame_row2_scrollbar.pack(side="right", fill="y")






























"""Settings"""

selected_middle_left_frame = tk.StringVar(value=get_registry_value("SelectedFrame", "Trọng lượng"))
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
        if not hasattr(update_com_ports, "thread_sq tarted"):
            update_com_ports.thread_started = True
            com_port_thread = threading.Thread(target=monitor_com_ports, daemon=True)
            com_port_thread.start()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def update_com_ports => {e}", 0, 3000), daemon=True).start()
        pass

current_thickness_entry = None
thickness_entry_list = [entry_thickness_cuon_bien_entry, entry_thickness_co_tay_entry, entry_thickness_ban_tay_entry,
                        entry_thickness_ngon_tay_entry, entry_thickness_dau_ngon_tay_entry]
def update_current_thickness_entry(event):
    global current_thickness_entry
    widget = event.widget
    if widget in thickness_entry_list:
        current_thickness_entry = widget
        print(f"Updated current entry: {current_thickness_entry}")
for entry in thickness_entry_list:
    entry.bind("<FocusIn>", update_current_thickness_entry)
def thickness_frame_com_port_insert_data():
    global current_thickness_entry
    if "COM" in get_registry_value("COM2", ""):
        print(f'Thickness COM: {get_registry_value("COM2", "")}')
        ser = serial.Serial(get_registry_value("COM2", ""), baudrate=9600, timeout=0.4)
        try:
            while ser.is_open:
                if ser.in_waiting > 0:
                    raw_value = ser.readline().decode('utf-8').strip()
                    if raw_value:
                        match = re.search(r"B\+([\d.]+)", raw_value)
                        if match:
                            thickness_value = float(match.group(1))
                            if entry_thickness_runcard_id_entry.get():
                                if current_thickness_entry and current_thickness_entry in thickness_entry_list:
                                    current_thickness_entry.delete(0, tk.END)
                                    current_thickness_entry.insert(0, thickness_value)
                                    current_index = thickness_entry_list.index(current_thickness_entry)
                                    if current_index < len(thickness_entry_list) - 1:
                                        thickness_entry_list[current_index + 1].focus_set()
                                    else:
                                        entry_thickness_dau_ngon_tay_entry.event_generate("<Return>")
                                else:
                                    entry_thickness_runcard_id_entry.event_generate("<Return>")
                                    thickness_entry_list[0].delete(0, tk.END)
                                    thickness_entry_list[0].insert(0, thickness_value)
                                    if len(thickness_entry_list) > 1:
                                        thickness_entry_list[1].focus_set()
                            else:
                                threading.Thread(target=show_error_message, args=("Chưa nhập giá trị Runcard!", 0, 3000),daemon=True).start()
                    else:
                        threading.Thread(target=show_error_message, args=("Kiểm tra lại kết nối với dụng cụ đo!", 0, 3000),daemon=True).start()
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"def thickness_frame_com_port_insert_data => {e}", 0, 3000),daemon=True).start()
            pass
        finally:
            if ser.is_open:
                ser.close()
    else:
        pass
def weight_frame_com_port_insert_data():
    global current_weight_entry, weight_record_log_id
    if "COM" in get_registry_value("COM1", ""):
        print(f'Weight COM: {get_registry_value("COM1", "")}')
        ser = serial.Serial(get_registry_value("COM1", ""), baudrate=9600, timeout=1)
        try:
            if ser.is_open:
                value = ser.readline().decode('utf-8').strip()
                if len(value):
                    weight_record_log_id += 1
                    if "g" not in value[-2:]:
                        threading.Thread(target=show_error_message, args=(f"Change your weight unit to gram!", 0, 3000), daemon=True).start()
                    weight_value = float(re.sub(r'[a-zA-Z]', '', ((value.replace(" ", "")).split(':')[-1])[:-1]))
                    if current_weight_entry == "entry_weight_runcard_id_entry":
                        entry_weight_runcard_id_entry.event_generate("<Return>")
                    entry_weight_weight_value_entry.insert(0, weight_value)
                    entry_weight_weight_value_entry.event_generate("<Return>")
                    entry_weight_weight_value_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.focus_set()
            else:
                ser.open()
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"def weight_frame_com_port_insert_data => {e}", 0, 3000), daemon=True).start()
            pass
        finally:
            if ser.is_open:
                ser.close()
    else:
        pass
def switch_middle_left_frame(*args):
    global weight_com_thread, thickness_com_thread
    try:
        print(f"Switching to: {selected_middle_left_frame.get()}")
        if selected_middle_left_frame.get() == "Trọng lượng":
            middle_left_thickness_frame.pack_forget()
            middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)
            middle_left_weight_frame.lift()
        else:
            middle_left_weight_frame.pack_forget()
            middle_left_thickness_frame.pack(fill=tk.BOTH, expand=True)
            middle_left_thickness_frame.lift()
        if "COM" in str(get_registry_value("COM1", "")):
            if weight_com_thread is None or not weight_com_thread.is_alive():
                weight_com_thread = threading.Thread(target=weight_frame_com_port_insert_data, daemon=True)
                weight_com_thread.start()
        if "COM" in str(get_registry_value("COM2", "")):
            if thickness_com_thread is None or not thickness_com_thread.is_alive():
                thickness_com_thread = threading.Thread(target=thickness_frame_com_port_insert_data, daemon=True)
                thickness_com_thread.start()
        set_registry_value("SelectedFrame", selected_middle_left_frame.get())
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def switch_middle_left_frame => {e}", 0, 3000), daemon=True).start()
        pass

switch_middle_left_frame()
selected_weight_com = tk.StringVar(value=get_registry_value("COM1", ""))
selected_thickness_com = tk.StringVar(value=get_registry_value("COM2", ""))

weight_label = tk.Label(middle_right_setting_frame_row2_row1, text="Trọng lượng:      ", font=(font_name, 14, "bold"), bg=bg_app_class_color_layer_2)
weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
weight_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_weight_com, "")
weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

thickness_label = tk.Label(middle_right_setting_frame_row2_row1, text="Độ dày:", font=(font_name, 14, "bold"), bg=bg_app_class_color_layer_2)
thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
thickness_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_thickness_com, "")
thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

frame_select_label = tk.Label(middle_right_setting_frame_row2_row1, text="Mặc định:", font=(font_name, 14, "bold"), bg=bg_app_class_color_layer_2)
frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
frame_select_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")
update_com_ports(weight_menu, thickness_menu)












"""Button"""
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
    # root.overrideredirect(True)

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
    # root.overrideredirect(False)
def save_setting():
    try:
        threading.Thread(target=show_error_message, args=(f"Save setting", 1, 3000), daemon=True).start()
        set_registry_value("COM1", selected_weight_com.get())
        set_registry_value("COM2", selected_thickness_com.get())
        switch_middle_left_frame()
        set_registry_value("SelectedFrame", selected_middle_left_frame.get())
        messagebox.showinfo("Success", "Save setting success!\nApplication will close automatically\nRe-open the app")
        root.destroy()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def save_setting() => {e}", 0, 3000), daemon=True).start()
        pass
def exit():
    try:
        set_registry_value("is_runcard_open", "0")
    except:
        pass
    os._exit(0)

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



def on_enter_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_hover_icon)
def on_leave_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_icon)
setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "setting.png")).resize((42, 42)))
setting_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "setting_hover.png")).resize((42, 42)))
middle_open_setting_frame_button = tk.Button(top_left_frame, image=setting_icon, bd=0, command=open_setting_frame, bg=bg_app_class_color_layer_1 )
middle_open_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")
middle_open_setting_frame_button.bind("<Enter>", on_enter_middle_open_setting_frame_button)
middle_open_setting_frame_button.bind("<Leave>", on_leave_middle_open_setting_frame_button)



def on_enter_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_hover_icon)
def on_leave_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_icon)
advance_setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "advance_setting.png")).resize((21, 21)))
advance_setting_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "advance_setting_hover.png")).resize((21, 21)))
middle_open_advance_setting_frame_button = tk.Button(middle_right_setting_frame_row1_col2, image=advance_setting_icon, bd=0, command=open_advance_setting_frame, bg=bg_app_class_color_layer_1 )
middle_open_advance_setting_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_advance_setting_frame_button.bind("<Enter>", on_enter_middle_open_advance_setting_frame_button)
middle_open_advance_setting_frame_button.bind("<Leave>", on_leave_middle_open_advance_setting_frame_button)



def on_enter_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_hover_icon)
def on_leave_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_icon)
return_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "return.png")).resize((21, 21)))
return_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "return_hover.png")).resize((21, 21)))
middle_open_return_frame_button = tk.Button(middle_right_advance_setting_frame_row1_col2, image=return_icon, bd=0, command=open_setting_frame, bg=bg_app_class_color_layer_1 )
middle_open_return_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_return_frame_button.bind("<Enter>", on_enter_middle_open_return_frame_button)
middle_open_return_frame_button.bind("<Leave>", on_leave_middle_open_return_frame_button)



def on_enter_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_hover_icon)
def on_leave_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_icon)
barcode_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "barcode.png")).resize((42, 42)))
barcode_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "barcode_hover.png")).resize((42, 42)))
middle_open_runcard_frame_button = tk.Button(top_left_frame, image=barcode_icon, bd=0, command=open_runcard_frame, bg=bg_app_class_color_layer_1 )
middle_open_runcard_frame_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")
middle_open_runcard_frame_button.bind("<Enter>", on_enter_middle_open_runcard_frame_button)
middle_open_runcard_frame_button.bind("<Leave>", on_leave_middle_open_runcard_frame_button)



close_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "close.png")).resize((134, 34)))
close_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "close_hover.png")).resize((134, 34)))
def on_enter_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon_hover)
def on_leave_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon)
close_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col2, image=close_icon, command=close_frame, bg=bg_app_class_color_layer_1 , width=134, height=34, relief="flat", borderwidth=0)
close_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")
close_setting_frame_button.bind("<Enter>", on_enter_close_setting_frame_button)
close_setting_frame_button.bind("<Leave>", on_leave_close_setting_frame_button)

def on_enter_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon_hover)
def on_leave_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon)
close_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col2, image=close_icon, command=close_frame, bg=bg_app_class_color_layer_1 , width=134, height=34, relief="flat", borderwidth=0)
close_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
close_advance_setting_frame_button.bind("<Enter>", on_enter_close_advance_setting_frame_button)
close_advance_setting_frame_button.bind("<Leave>", on_leave_close_advance_setting_frame_button)


def on_enter_close_runcard_frame_button(event):
    close_runcard_frame_button.config(image=close_icon_hover)
def on_leave_close_runcard_frame_button(event):
    close_runcard_frame_button.config(image=close_icon)
close_runcard_frame_button = tk.Button(middle_right_runcard_frame_row3, image=close_icon, command=close_frame, bg=bg_app_class_color_layer_1 , width=134, height=34, relief="flat", borderwidth=0)
close_runcard_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
close_runcard_frame_button.bind("<Enter>", on_enter_close_runcard_frame_button)
close_runcard_frame_button.bind("<Leave>", on_leave_close_runcard_frame_button)


save_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "save.png")).resize((134, 34)))
save_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "save_hover.png")).resize((134, 34)))
def on_enter_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon_hover)
def on_leave_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon)
save_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col1, image=save_icon, command=save_setting, bg=bg_app_class_color_layer_1 , width=134, height=34, relief="flat", borderwidth=0)
save_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_setting_frame_button.bind("<Enter>", on_enter_save_setting_frame_button)
save_setting_frame_button.bind("<Leave>", on_leave_save_setting_frame_button)

def on_enter_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon_hover)
def on_leave_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon)
save_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col1, image=save_icon, command=None, bg=bg_app_class_color_layer_1 , width=134, height=34, relief="flat", borderwidth=0)
save_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_advance_setting_frame_button.bind("<Enter>", on_enter_save_advance_setting_frame_button)
save_advance_setting_frame_button.bind("<Leave>", on_leave_save_advance_setting_frame_button)



def on_enter_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_hover_icon)
def on_leave_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_icon)
database_test_connection_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "connect.png")).resize((146, 38)))
database_test_connection_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "connect_hover.png")).resize((146, 38)))
database_test_connection_button = tk.Button(middle_right_advance_setting_frame_row2_col1_row6_col2, image=database_test_connection_button_icon, command=None, bg=bg_app_class_color_layer_1, width=146, height=38, relief="flat", borderwidth=0)
database_test_connection_button.grid(row=4, column=1, padx=5, pady=5, sticky="e")
database_test_connection_button.bind("<Enter>", on_enter_database_test_connection_button)
database_test_connection_button.bind("<Leave>", on_leave_database_test_connection_button)


delete_entry_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "delete.png")).resize((117, 36)))
delete_entry_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "delete_hover.png")).resize((117, 36)))
def on_enter_delete_weight_entry_button(event):
    delete_weight_entry_button.config(image=delete_entry_button_hover_icon)
def on_leave_delete_weight_entry_button(event):
    delete_weight_entry_button.config(image=delete_entry_button_icon)
delete_weight_entry_button = tk.Button(middle_left_weight_frame_col3_frame_row1_row1, image=delete_entry_button_icon, command=None, bg='#f0f2f6', width=120, height=40, relief="flat", borderwidth=0)
delete_weight_entry_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_weight_entry_button.bind("<Enter>", on_enter_delete_weight_entry_button)
delete_weight_entry_button.bind("<Leave>", on_leave_delete_weight_entry_button)

def on_enter_delete_thickness_entry_button(event):
    delete_thickness_entry_button.config(image=delete_entry_button_hover_icon)
def on_leave_delete_thickness_entry_button(event):
    delete_thickness_entry_button.config(image=delete_entry_button_icon)
delete_thickness_entry_button = tk.Button(middle_left_thickness_frame_col3_frame_row1_row1, image=delete_entry_button_icon, command=None, bg='#f0f2f6', width=120, height=40, relief="flat", borderwidth=0)
delete_thickness_entry_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_thickness_entry_button.bind("<Enter>", on_enter_delete_thickness_entry_button)
delete_thickness_entry_button.bind("<Leave>", on_leave_delete_thickness_entry_button)



delete_log_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "delete_all.png")).resize((36, 36)))
delete_log_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "delete_all_hover.png")).resize((36, 36)))
def on_enter_delete_weight_log_button(event):
    delete_weight_log_button.config(image=delete_log_button_hover_icon)
def on_leave_delete_weight_log_button(event):
    delete_weight_log_button.config(image=delete_log_button_icon)
delete_weight_log_button = tk.Button(middle_left_weight_frame_col3_frame_row1_row2, image=delete_log_button_icon, command=None, bg=bg_app_class_color_layer_1, width=40, height=40, relief="flat", borderwidth=0)
delete_weight_log_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_weight_log_button.bind("<Enter>", on_enter_delete_weight_log_button)
delete_weight_log_button.bind("<Leave>", on_leave_delete_weight_log_button)

def on_enter_delete_thickness_log_button(event):
    delete_thickness_log_button.config(image=delete_log_button_hover_icon)
def on_leave_delete_thickness_log_button(event):
    delete_thickness_log_button.config(image=delete_log_button_icon)
delete_thickness_log_button = tk.Button(middle_left_thickness_frame_col3_frame_row1_row2, image=delete_log_button_icon, command=None, bg=bg_app_class_color_layer_1, width=40, height=40, relief="flat", borderwidth=0)
delete_thickness_log_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_thickness_log_button.bind("<Enter>", on_enter_delete_thickness_log_button)
delete_thickness_log_button.bind("<Leave>", on_leave_delete_thickness_log_button)




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


def run_cef_in_frame(tk_frame, url):
    class FocusHandler:
        def OnSetFocus(self, browser, source):
            print("Blocked Chromium from stealing focus")
            return True
    class BrowserFrame(tk.Frame):
        def __init__(self, master, url):
            super().__init__(master)
            self.url = url
            self.browser = None
            self.pack(fill="both", expand=True)
            self._embedded = False

        def embed_browser(self):
            if self._embedded:
                return
            if not self.winfo_ismapped() or self.winfo_width() <= 0:
                self.after(0, self.embed_browser)
                return
            window_info = cef.WindowInfo()
            rect = [0, 0, self.winfo_width(), self.winfo_height()]
            window_info.SetAsChild(self.winfo_id(), rect)
            self.browser = cef.CreateBrowserSync(window_info, url=self.url)
            self.browser.SetFocusHandler(FocusHandler())
            self._embedded = True

        def resize_browser(self):
            if self.browser:
                try:
                    cef.WindowUtils.OnSize(self.winfo_id(), 0, 0, 0)
                except Exception as e:
                    pass
    browser_widget = BrowserFrame(tk_frame, url=url)
    def delayed_embed():
        browser_widget.embed_browser()
    tk_frame.after(0, delayed_embed)
    def cef_loop():
        if not root.winfo_exists():
            return
        try:
            cef.MessageLoopWork()
        except Exception as e:
            print("CEF error:", e)
        tk_frame.after(1, cef_loop)
    cef_loop()


def restore_entry_focus(entry_widget):
    def on_click(event):
        root.after(1, lambda: entry_widget.focus_force())
    entry_widget.bind("<Button-1>", on_click)

def on_close():
    os._exit(0)

restore_entry_focus(entry_weight_device_name_entry)
restore_entry_focus(entry_weight_operator_id_entry)
restore_entry_focus(entry_weight_runcard_id_entry)
restore_entry_focus(entry_weight_weight_value_entry)

restore_entry_focus(entry_thickness_runcard_id_entry)
restore_entry_focus(entry_thickness_cuon_bien_entry)
restore_entry_focus(entry_thickness_co_tay_entry)
restore_entry_focus(entry_thickness_ban_tay_entry)
restore_entry_focus(entry_thickness_ngon_tay_entry)
restore_entry_focus(entry_thickness_dau_ngon_tay_entry)
def on_close():
    os._exit(0)
root.protocol("WM_DELETE_WINDOW", on_close)





"""Update loop"""
update_thread = threading.Thread(target=update_dimensions, daemon=True)
update_thread.start()
root.mainloop()



"""
=> pyinstaller --windowed --onefile --name temp03 --add-data "theme/icons;theme/icons" --add-data "theme/assets;theme/assets" temp03.py --icon=theme/icons/logo.ico
"""