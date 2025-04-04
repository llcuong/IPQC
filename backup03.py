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

barcode.base.Barcode.default_writer_options['write_text'] = False

base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")

"""Set DPI awareness for better scaling on high-DPI screens"""
ctypes.windll.shcore.SetProcessDpiAwareness(10)
REG_PATH = r"SOFTWARE\IPQC\Config"

user32 = ctypes.windll.user32
monitor_width, monitor_height = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
# print(f"Screen Resolution: {monitor_width}x{monitor_height}")

"""Main window setup"""
root = tk.Tk()
root.title("IPQC v.25.03.29")
start_time = time.time()
"""Define scaling factors"""
user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
dpi = user32.GetDpiForSystem()
scaling = dpi / 96
# print(f"scaling ==> {scaling}")
screen_width = min(1600, int(0.9 * root.winfo_screenwidth() * scaling))
screen_height = min(900, int(0.9 * root.winfo_screenheight() * scaling))
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

bg_app_class_color_layer_1 = "#f4f4fe"  # 00B9FF    => #333333
bg_app_class_color_layer_2 = "#ffffff"  # => #414141
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


def update_dimensions():
    global screen_width, screen_height, showing_settings, showing_runcards, count
    while True:
        # print(f"App loaded in {time.time() - start_time:.4f} seconds")

        current_date = (datetime.datetime.now() - datetime.timedelta(hours=5) + datetime.timedelta(minutes=22)).date()

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

        middle_right_frame.place(x=middle_left_frame_width + middle_center_frame_width, y=0,
                                 width=middle_right_frame_width, height=middle_frame_height)
        middle_center_frame.place(x=middle_left_frame_width, y=0, width=middle_center_frame_width,
                                  height=middle_frame_height)
        middle_left_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        top_left_frame.place(x=0, y=0, width=int(screen_width * 0.3), height=50)
        top_right_frame.place(x=int(screen_width * 0.3), y=0, width=int(screen_width * 0.7), height=50)

        bottom_left_frame.place(x=0, y=0, width=int(screen_width * 0.7), height=50)
        bottom_right_frame.place(x=int(screen_width * 0.7), y=0, width=int(screen_width * 0.3), height=50)

        middle_left_weight_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)
        middle_left_thickness_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        middle_left_weight_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_weight_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width,
                                                  height=middle_frame_height)
        middle_left_weight_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0,
                                                  width=middle_left_col3_width, height=middle_frame_height)

        middle_left_thickness_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_thickness_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width,
                                                     height=middle_frame_height)
        middle_left_thickness_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0,
                                                     width=middle_left_col3_width, height=middle_frame_height)

        middle_left_weight_frame_col3_frame_row1.place(x=0, y=0, width=210, height=140)
        middle_left_weight_frame_col3_frame_row2.place(x=0, y=140, width=210, height=middle_frame_height - 140)

        middle_left_thickness_frame_col3_frame_row1.place(x=0, y=0, width=210, height=140)
        middle_left_thickness_frame_col3_frame_row2.place(x=0, y=140, width=210, height=middle_frame_height - 140)

        middle_left_weight_frame_col3_frame_row1_row1.place(x=0, y=20, width=120, height=40)
        middle_left_weight_frame_col3_frame_row1_row2.place(x=0, y=80, width=40, height=40)

        middle_left_thickness_frame_col3_frame_row1_row1.place(x=0, y=20, width=120, height=40)
        middle_left_thickness_frame_col3_frame_row1_row2.place(x=0, y=80, width=40, height=40)

        middle_left_weight_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_weight_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width,
                                                       height=middle_frame_height - 80)

        middle_left_thickness_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_thickness_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width,
                                                          height=middle_frame_height - 80)

        middle_left_weight_frame_col1_frame_row2_canvas.place(x=0, y=0, width=middle_left_col1_width,
                                                              height=middle_frame_height - 80)
        middle_left_weight_frame_col1_frame_row2_scrollbar.place(x=middle_left_col1_width - 20, y=0, width=20,
                                                                 height=middle_frame_height - 80)
        middle_left_weight_frame_col1_frame_row2_scrollable_frame.place(x=0, y=0, width=middle_left_col1_width - 20,
                                                                        height=middle_frame_height - 80)

        middle_left_thickness_frame_col1_frame_row2_canvas.place(x=0, y=0, width=middle_left_col1_width,
                                                                 height=middle_frame_height - 80)
        middle_left_thickness_frame_col1_frame_row2_scrollbar.place(x=middle_left_col1_width - 20, y=0, width=20,
                                                                    height=middle_frame_height - 80)
        middle_left_thickness_frame_col1_frame_row2_scrollable_frame.place(x=0, y=0, width=middle_left_col1_width - 20,
                                                                           height=middle_frame_height - 80)

        middle_left_weight_frame_col3_frame_row2_canvas.place(x=0, y=0, width=210, height=middle_frame_height - 140)
        middle_left_weight_frame_col3_frame_row2_scrollbar.place(x=190, y=0, width=20, height=middle_frame_height - 140)
        middle_left_weight_frame_col3_frame_row2_scrollable_frame.place(x=0, y=0, width=190,
                                                                        height=middle_frame_height - 140)

        middle_right_runcard_frame.place(x=0, y=0, width=middle_right_frame_width - 5, height=middle_frame_height)
        middle_right_setting_frame.place(x=0, y=0, width=middle_right_frame_width - 5, height=middle_frame_height)
        middle_right_advance_setting_frame.place(x=0, y=0, width=middle_right_frame_width - 5,
                                                 height=middle_frame_height)

        middle_right_runcard_frame_row1.place(x=0, y=0, width=middle_right_frame_width - 10, height=40)
        middle_right_runcard_frame_row2.place(x=0, y=40, width=middle_right_frame_width - 10, height=720)
        middle_right_runcard_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width - 5,
                                              height=40)

        middle_right_runcard_frame_row2_col1.place(x=0, y=0, width=36, height=middle_frame_height - 40 - 40)
        middle_right_runcard_frame_row2_col2.place(x=36, y=0, width=middle_right_frame_width - 10 - 36 - 36,
                                                   height=middle_frame_height - 40 - 40)
        middle_right_runcard_frame_row2_col3.place(x=middle_right_frame_width - 10 - 36, y=0, width=36,
                                                   height=middle_frame_height - 40 - 40)

        middle_right_runcard_frame_row2_col2_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 36 - 36,
                                                        height=250)
        middle_right_runcard_frame_row2_col2_row2.place(x=0, y=250, width=middle_right_frame_width - 10 - 36 - 36,
                                                        height=50)
        middle_right_runcard_frame_row2_col2_row3.place(x=0, y=300, width=middle_right_frame_width - 10 - 36 - 36,
                                                        height=50)
        middle_right_runcard_frame_row2_col2_row4.place(x=0, y=350, width=middle_right_frame_width - 10 - 36 - 36,
                                                        height=200)
        middle_right_runcard_frame_row2_col2_row5.place(x=0, y=550, width=middle_right_frame_width - 10 - 36 - 36,
                                                        height=200)

        middle_right_runcard_frame_row2_col2_row1_row1.place(x=10, y=0,
                                                             width=middle_right_frame_width - 10 - 36 - 36 - 20,
                                                             height=250)

        middle_right_runcard_frame_row2_col2_row4_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 36 - 36,
                                                             height=30)
        middle_right_runcard_frame_row2_col2_row4_row2.place(x=0, y=30, width=middle_right_frame_width - 10 - 36 - 36,
                                                             height=80)
        middle_right_runcard_frame_row2_col2_row4_row3.place(x=0, y=110, width=middle_right_frame_width - 10 - 36 - 36,
                                                             height=30)
        # print(f"-----> {middle_right_runcard_frame_row2_col2_row1_row1.winfo_height()}")
        # print(f"-----> {middle_right_runcard_frame_row2_col2_row1_row1.winfo_width()}")
        middle_right_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=180)
        middle_right_setting_frame_row3.place(x=0, y=40 + 200, width=middle_right_frame_width, height=50)

        middle_right_advance_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_advance_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width,
                                                      height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width,
                                                      height=50)

        middle_right_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width / 2), height=40)
        middle_right_setting_frame_row1_col2.place(x=int(middle_right_frame_width / 2), y=0,
                                                   width=int(middle_right_frame_width / 2), height=40)

        middle_right_setting_frame_row3_col1.place(x=int((middle_right_frame_width / 2 - 154) / 2), y=0, width=154,
                                                   height=40)
        middle_right_setting_frame_row3_col2.place(
            x=int((middle_right_frame_width / 2 - 154) / 2 + middle_right_frame_width / 2), y=0, width=154, height=40)

        middle_right_advance_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width / 2), height=40)
        middle_right_advance_setting_frame_row1_col2.place(x=int(middle_right_frame_width / 2), y=0,
                                                           width=int(middle_right_frame_width / 2), height=40)

        middle_right_advance_setting_frame_row3_col1.place(x=int((middle_right_frame_width / 2 - 154) / 2), y=0,
                                                           width=154, height=40)
        middle_right_advance_setting_frame_row3_col2.place(
            x=int((middle_right_frame_width / 2 - 154) / 2 + middle_right_frame_width / 2), y=0, width=154, height=40)

        middle_right_setting_frame_row2_row1.place(x=0, y=20, width=int(middle_right_frame_width), height=160)

        middle_right_advance_setting_frame_row2_col1.place(x=0, y=0, width=332, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row2_col2.place(x=332, y=0, width=10, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row2_col3.place(x=342, y=0, width=middle_right_frame_width - 10 - 332 - 5,
                                                           height=middle_frame_height - 40 - 50)

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

        middle_right_advance_setting_frame_row2_col1_row5_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row2.place(x=0, y=40, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row3.place(x=0, y=80, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row4.place(x=0, y=120, width=170, height=40)

        middle_right_advance_setting_frame_row2_col1_row5_col2_row1.place(x=0, y=0, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row2.place(x=0, y=40, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row3.place(x=0, y=80, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row4.place(x=0, y=120, width=162, height=40)

        middle_right_advance_setting_frame_row2_col1_row9_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row9_col2_row1.place(x=0, y=0, width=162, height=40)

        middle_right_advance_setting_frame_row2_col3_row1.place(x=0, y=0,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=30)
        middle_right_advance_setting_frame_row2_col3_row2.place(x=0, y=30,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=10)
        middle_right_advance_setting_frame_row2_col3_row3.place(x=0, y=40,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=200)

        middle_right_advance_setting_frame_row2_col3_row4.place(x=0, y=240,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=30)

        middle_right_advance_setting_frame_row2_col3_row5.place(x=0, y=270,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=30)
        middle_right_advance_setting_frame_row2_col3_row6.place(x=0, y=300,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=10)
        middle_right_advance_setting_frame_row2_col3_row7.place(x=0, y=310,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=200)

        middle_right_advance_setting_frame_row2_col3_row8.place(x=0, y=510,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=30)

        middle_right_advance_setting_frame_row2_col3_row9.place(x=0, y=540,
                                                                width=middle_right_frame_width - 10 - 332 - 10,
                                                                height=30)
        middle_right_advance_setting_frame_row2_col3_row10.place(x=0, y=570,
                                                                 width=middle_right_frame_width - 10 - 332 - 10,
                                                                 height=10)
        middle_right_advance_setting_frame_row2_col3_row11.place(x=0, y=580,
                                                                 width=middle_right_frame_width - 10 - 332 - 10,
                                                                 height=200)

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

        root.update_idletasks()
        root.update()
        if count == 0:
            print(f"-->==>===>{time.time() - start_time}")
            count = 1


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
                threading.Thread(target=show_error_message, args=("Only numbers and '.' are allowed!", 0, 2000),
                                 daemon=True).start()
            input_text = filtered_text
        upper_entry_name = input_text.upper()
        if len(upper_entry_name) > max_char:
            entry_name.set(upper_entry_name[:max_char])
            threading.Thread(target=show_error_message, args=("Exceeded character limit!", 0, 2000),
                             daemon=True).start()
        else:
            entry_name.set(upper_entry_name)
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def convert_to_uppercase => {e}", 0, 3000),
                         daemon=True).start()
        pass


def get_registry_value(name, default=""):
    """Retrieve a value from the Windows Registry."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, name)[0]
    except FileNotFoundError as e:
        threading.Thread(target=show_error_message, args=(f"def get_registry_value => {e}", 0, 3000),
                         daemon=True).start()
        return default


def set_registry_value(name, value):
    """Set a value in the Windows Registry."""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def set_registry_value => {e}", 0, 3000),
                         daemon=True).start()


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

middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_2)
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)

middle_left_weight_frame_col3_frame_row1 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row2 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col3_frame_row1 = tk.Frame(middle_left_thickness_frame_col3_frame,
                                                       bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row2 = tk.Frame(middle_left_thickness_frame_col3_frame,
                                                       bg=bg_app_class_color_layer_1)

middle_left_weight_frame_col3_frame_row1_row1 = tk.Frame(middle_left_weight_frame_col3_frame_row1,
                                                         bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row1_row2 = tk.Frame(middle_left_weight_frame_col3_frame_row1,
                                                         bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col3_frame_row1_row1 = tk.Frame(middle_left_thickness_frame_col3_frame_row1,
                                                            bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row1_row2 = tk.Frame(middle_left_thickness_frame_col3_frame_row1,
                                                            bg=bg_app_class_color_layer_1)

middle_left_weight_frame_col1_frame_row1 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col1_frame_row2 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col1_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame,
                                                       bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col1_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame,
                                                       bg=bg_app_class_color_layer_1)

middle_center_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_center_frame.place(x=0, y=0)

middle_right_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_right_frame.place(x=0, y=0)

middle_right_runcard_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame.place(x=0, y=0)

middle_right_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_setting_frame.place(x=0, y=0)

middle_right_advance_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame.place(x=0, y=0)

middle_right_runcard_frame_row1 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row2 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row3 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row3.grid_columnconfigure(0, weight=1)

middle_right_runcard_frame_row2_col1 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row2_col3 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row1 = tk.Frame(middle_right_runcard_frame_row2_col2,
                                                     bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row2 = tk.Frame(middle_right_runcard_frame_row2_col2,
                                                     bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row3 = tk.Frame(middle_right_runcard_frame_row2_col2,
                                                     bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4 = tk.Frame(middle_right_runcard_frame_row2_col2,
                                                     bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row5 = tk.Frame(middle_right_runcard_frame_row2_col2,
                                                     bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row1_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row1,
                                                          bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row4_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row4,
                                                          bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4_row2 = tk.Frame(middle_right_runcard_frame_row2_col2_row4,
                                                          bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4_row3 = tk.Frame(middle_right_runcard_frame_row2_col2_row4,
                                                          bg=bg_app_class_color_layer_2)

# bg_image = Image.open("theme/images/button-disabled.png")
# bg_image = bg_image.resize((400, 300), Image.LANCZOS)
# bg_photo = ImageTk.PhotoImage(bg_image)
# bg_label = tk.Label(middle_right_runcard_frame_row2_col2_row5, image=bg_photo)
# bg_label.place(x=0, y=0, relwidth=1, relheight=1)


middle_right_setting_frame_row1 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row2 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_2)
middle_right_setting_frame_row3 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1)

middle_right_advance_setting_frame_row1 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row3 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)

middle_right_setting_frame_row1_col1 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row1_col2 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row3_col1 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row3_col2 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1)

middle_right_advance_setting_frame_row1_col1 = tk.Frame(middle_right_advance_setting_frame_row1,
                                                        bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row1_col2 = tk.Frame(middle_right_advance_setting_frame_row1,
                                                        bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_advance_setting_frame_row3_col1 = tk.Frame(middle_right_advance_setting_frame_row3,
                                                        bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row3_col2 = tk.Frame(middle_right_advance_setting_frame_row3,
                                                        bg=bg_app_class_color_layer_1)

bottom_frame = tk.Frame(root, bg=bg_app_class_color_layer_1)
bottom_frame.place(relx=0, rely=1.0, height=50, anchor="sw")

bottom_left_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1)
bottom_right_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1)
bottom_right_frame.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row2_row1 = tk.Frame(middle_right_setting_frame_row2, bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1 = tk.Frame(middle_right_advance_setting_frame_row2,
                                                        bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col2 = tk.Frame(middle_right_advance_setting_frame_row2,
                                                        bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2,
                                                        bg=bg_app_class_color_layer_1)

middle_right_advance_setting_frame_row2_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row12 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row13 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row14 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col1_row15 = tk.Frame(middle_right_advance_setting_frame_row2_col1,
                                                              bg=bg_app_class_color_layer_1)

middle_right_advance_setting_frame_row2_col1_row5_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5,
                                                                  bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5,
                                                                  bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1_row6_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6,
                                                                  bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row6_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6,
                                                                  bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1_row9_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9,
                                                                  bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row9_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9,
                                                                  bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col1_row9_col1_row1 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row9_col1, bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col1_row9_col2_row1 = tk.Frame(
    middle_right_advance_setting_frame_row2_col1_row9_col2, bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                             bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2_col3_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                              bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col3,
                                                              bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row3_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row3_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3,
                                                                bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row7_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row7_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7,
                                                                bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row11_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row11_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row11_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row11_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)

middle_right_advance_setting_frame_row2_col3_row11_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)
middle_right_advance_setting_frame_row2_col3_row11_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11,
                                                                 bg=bg_app_class_color_layer_2)

""""""
middle_right_setting_frame_row1_col1_label = tk.Label(middle_right_setting_frame_row1_col1, text="Cài đặt",
                                                      bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
middle_right_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_advance_setting_frame_row1_col1_label = tk.Label(middle_right_advance_setting_frame_row1_col1,
                                                              text="Cài đặt nâng cao", bg=bg_app_class_color_layer_1,
                                                              bd=0, font=(font_name, 18, "bold"))
middle_right_advance_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_runcard_frame_label = tk.Label(middle_right_runcard_frame_row1, text="Runcard",
                                            bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
middle_right_runcard_frame_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")


def runcard_date():
    global selected_date
    calendar_style = ttk.Style()
    calendar_style.theme_use('clam')
    calendar_style.map('Custom.DateEntry.TEntry',
                       foreground=[('readonly', fg_app_class_color_layer_1)],
                       fieldbackground=[('readonly', bg_app_class_color_layer_2)],
                       background=[('readonly', bg_app_class_color_layer_2)],
                       bordercolor=[('focus', bg_app_class_color_layer_2), ('!focus', bg_app_class_color_layer_2)],
                       lightcolor=[('focus', bg_app_class_color_layer_2), ('!focus', bg_app_class_color_layer_2)],
                       highlightcolor=[('focus', bg_app_class_color_layer_2), ('!focus', bg_app_class_color_layer_2)],
                       )
    calendar_style.configure('Custom.DateEntry.TEntry', font=(font_name, 12), padding=4, borderwidth=1, relief='flat')
    calendar = DateEntry(middle_right_runcard_frame_row1,
                         width=12,
                         year=int(current_date.strftime("%Y")),
                         month=int(current_date.strftime("%m")),
                         day=int(current_date.strftime("%d")),
                         background=bg_app_class_color_layer_2,
                         foreground=fg_app_class_color_layer_1,
                         disabledbackground=bg_app_class_color_layer_2,
                         disabledforeground=fg_app_class_color_layer_2,
                         bordercolor=bg_app_class_color_layer_2,
                         headersbackground=bg_app_class_color_layer_2,
                         headersforeground=fg_app_class_color_layer_1,
                         normalbackground=bg_app_class_color_layer_2,
                         normalforeground=fg_app_class_color_layer_1,
                         weekendbackground=bg_app_class_color_layer_2,
                         weekendforeground=fg_app_class_color_layer_1,
                         othermonthforeground=bg_app_class_color_layer_2,
                         othermonthbackground=fg_app_class_color_layer_2,
                         othermonthweforeground=bg_app_class_color_layer_2,
                         othermonthwebackground=fg_app_class_color_layer_2,
                         disableddaybackground=bg_app_class_color_layer_2,
                         disableddayforeground="#aeaeae",
                         locale='vi_VN',
                         showweeknumbers=False,
                         mindate=current_date - datetime.timedelta(days=7),
                         maxdate=current_date,
                         borderwidth=2,
                         font=(font_name, 12),
                         date_pattern='dd-mm-yyyy',
                         style='Custom.DateEntry.TEntry',
                         state="readonly",
                         )
    calendar.grid(row=0, column=1, padx=5, pady=5, sticky="e")

    def on_date_change(event):
        global selected_date
        # print("Selected date:", calendar.get_date().strftime("%d-%m-%Y"))
        selected_date = calendar.get_date().strftime("%Y-%m-%d")
        root.after(0, runcard_period_button)
        root.after(0, runcard_machine_button)
        root.after(0, runcard_line_button)
        # root.after(0, runcard_wo_button)
        print(f"=> {selected_date}")

    calendar.bind("<<DateEntrySelected>>", on_date_change)
    root.after(0, runcard_period_button)
    root.after(0, runcard_machine_button)
    root.after(0, runcard_line_button)
    # root.after(0, runcard_wo_button)
    return calendar.get_date().strftime("%d-%m-%Y")


"""Function"""

if get_registry_value("is_current_entry", "weight") == "weight":
    showing_thickness_frame, showing_weight_frame = False, True
else:
    showing_thickness_frame, showing_weight_frame = True, False

error_display_entry = tk.Entry(bottom_left_frame, textvariable=error_msg, font=("Cambria", 12),
                               bg=bg_app_class_color_layer_1, fg=error_fg_color, bd=0, highlightthickness=0,
                               readonlybackground=bg_app_class_color_layer_1, state="readonly")


def weight_frame_mouser_pointer_in(event):
    global current_weight_entry
    if 'name_var' in event.widget.__dict__:
        current_weight_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")


def weight_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_weight_value_entry":
            messagebox.showinfo("Success",
                                "Fuck you!~")
        else:
            event.widget.tk_focusNext().focus()
        return "break"
    except:
        threading.Thread(target=show_error_message, args=(f"def weight_frame_hit_enter_button => {e}", 0, 3000),
                         daemon=True).start()


entry_weight_device_name_var = tk.StringVar()
entry_weight_device_name_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_device_name_var, 12, 1))
entry_weight_device_name_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Device ID",
                                          bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 16, "bold"))
entry_weight_device_name_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                          textvariable=entry_weight_device_name_var, bg=bg_app_class_color_layer_2)
entry_weight_device_name_entry.name_var = "entry_weight_device_name_entry"
entry_weight_device_name_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_device_name_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(0, weight=1)

entry_weight_operator_id_var = tk.StringVar()
entry_weight_operator_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_operator_id_var, 5, 0))
entry_weight_operator_id_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Operator ID", bg="#f4f4fe",
                                          bd=0, font=(font_name, 16, "bold"))
entry_weight_operator_id_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                          textvariable=entry_weight_operator_id_var)
entry_weight_operator_id_entry.name_var = "entry_weight_operator_id_entry"
entry_weight_operator_id_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_operator_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(1, weight=1)

entry_weight_runcard_id_var = tk.StringVar()
entry_weight_runcard_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_runcard_id_var, 10, 1))
entry_weight_runcard_id_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Runcard ID", bg="#f4f4fe",
                                         bd=0, font=("Helvetica", 16, "bold"))
entry_weight_runcard_id_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                         textvariable=entry_weight_runcard_id_var)
entry_weight_runcard_id_entry.name_var = "entry_weight_runcard_id_entry"
entry_weight_runcard_id_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_runcard_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(2, weight=1)

entry_weight_weight_value_var = tk.StringVar()
entry_weight_weight_value_var.trace_add("write",
                                        lambda *args: convert_to_uppercase(entry_weight_weight_value_var, 10, 0))
entry_weight_weight_value_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Trọng lượng", bg="#f4f4fe",
                                           bd=0, font=("Helvetica", 16, "bold"))
entry_weight_weight_value_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                           textvariable=entry_weight_weight_value_var)
entry_weight_weight_value_entry.name_var = "entry_weight_weight_value_entry"
entry_weight_weight_value_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_weight_value_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(3, weight=1)

middle_left_weight_frame_col1_frame_row2_canvas = tk.Canvas(middle_left_weight_frame_col1_frame_row2,
                                                            bg=bg_app_class_color_layer_2, highlightthickness=0)
middle_left_weight_frame_col1_frame_row2_scrollbar = tk.Scrollbar(middle_left_weight_frame_col1_frame_row2,
                                                                  orient="vertical",
                                                                  command=middle_left_weight_frame_col1_frame_row2_canvas.yview)
middle_left_weight_frame_col1_frame_row2_canvas.configure(
    yscrollcommand=middle_left_weight_frame_col1_frame_row2_scrollbar.set)
middle_left_weight_frame_col1_frame_row2_scrollable_frame = tk.Frame(middle_left_weight_frame_col1_frame_row2_canvas,
                                                                     bg=bg_app_class_color_layer_2)
middle_left_weight_frame_col1_frame_row2_canvas.create_window((0, 0),
                                                              window=middle_left_weight_frame_col1_frame_row2_scrollable_frame,
                                                              anchor="nw")
middle_left_weight_frame_col1_frame_row2_scrollable_frame.bind("<Configure>", lambda
    e: middle_left_weight_frame_col1_frame_row2_canvas.configure(
    scrollregion=middle_left_weight_frame_col1_frame_row2_canvas.bbox("all")))
middle_left_weight_frame_col1_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_col1_frame_row2_scrollbar.pack(side="right", fill="y")
all_entries = []


def thickness_frame_mouser_pointer_in(event):
    global current_thickness_entry
    if 'name_var' in event.widget.__dict__:
        current_thickness_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")


def thickness_frame_hit_enter_button(event):
    return None


entry_thickness_runcard_id_var = tk.StringVar()
entry_thickness_runcard_id_var.trace_add("write",
                                         lambda *args: convert_to_uppercase(entry_thickness_runcard_id_var, 12, 1))
entry_thickness_runcard_id_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Runcard ID",
                                            bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_runcard_id_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                            textvariable=entry_thickness_runcard_id_var)
entry_thickness_runcard_id_entry.name_var = "entry_thickness_runcard_id_entry"
entry_thickness_runcard_id_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_runcard_id_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(0, weight=1)

entry_thickness_cuon_bien_var = tk.StringVar()
entry_thickness_cuon_bien_var.trace_add("write",
                                        lambda *args: convert_to_uppercase(entry_thickness_cuon_bien_var, 6, 0))
entry_thickness_cuon_bien_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Cuốn biên", bg="#f4f4fe",
                                           bd=0, font=(font_name, 16, "bold"))
entry_thickness_cuon_bien_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                           textvariable=entry_thickness_cuon_bien_var)
entry_thickness_cuon_bien_entry.name_var = "entry_thickness_cuon_bien_entry"
entry_thickness_cuon_bien_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_cuon_bien_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(1, weight=1)

entry_thickness_co_tay_var = tk.StringVar()
entry_thickness_co_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_co_tay_var, 6, 0))
entry_thickness_co_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Cổ tay", bg="#f4f4fe", bd=0,
                                        font=(font_name, 16, "bold"))
entry_thickness_co_tay_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                        textvariable=entry_thickness_co_tay_var)
entry_thickness_co_tay_entry.name_var = "entry_thickness_co_tay_entry"
entry_thickness_co_tay_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_co_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(2, weight=1)

entry_thickness_ban_tay_var = tk.StringVar()
entry_thickness_ban_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ban_tay_var, 6, 0))
entry_thickness_ban_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Bàn tay", bg="#f4f4fe",
                                         bd=0, font=(font_name, 16, "bold"))
entry_thickness_ban_tay_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                         textvariable=entry_thickness_ban_tay_var)
entry_thickness_ban_tay_entry.name_var = "entry_thickness_ban_tay_entry"
entry_thickness_ban_tay_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ban_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(3, weight=1)

entry_thickness_ngon_tay_var = tk.StringVar()
entry_thickness_ngon_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ngon_tay_var, 6, 0))
entry_thickness_ngon_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Ngón tay", bg="#f4f4fe",
                                          bd=0, font=(font_name, 16, "bold"))
entry_thickness_ngon_tay_label.grid(row=0, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                          textvariable=entry_thickness_ngon_tay_var)
entry_thickness_ngon_tay_entry.name_var = "entry_thickness_ngon_tay_entry"
entry_thickness_ngon_tay_entry.grid(row=1, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(4, weight=1)

entry_thickness_dau_ngon_tay_var = tk.StringVar()
entry_thickness_dau_ngon_tay_var.trace_add("write",
                                           lambda *args: convert_to_uppercase(entry_thickness_dau_ngon_tay_var, 6, 0))
entry_thickness_dau_ngon_tay_label = tk.Label(middle_left_thickness_frame_col1_frame_row1, text="Đầu ngón tay",
                                              bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_dau_ngon_tay_label.grid(row=0, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_col1_frame_row1, font=(font_name, 18), bd=2,
                                              textvariable=entry_thickness_dau_ngon_tay_var)
entry_thickness_dau_ngon_tay_entry.name_var = "entry_thickness_dau_ngon_tay_entry"
entry_thickness_dau_ngon_tay_entry.grid(row=1, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_dau_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_col1_frame_row1.columnconfigure(5, weight=1)

middle_left_thickness_frame_col1_frame_row2_canvas = tk.Canvas(middle_left_thickness_frame_col1_frame_row2,
                                                               bg=bg_app_class_color_layer_2, highlightthickness=0)
middle_left_thickness_frame_col1_frame_row2_scrollbar = tk.Scrollbar(middle_left_thickness_frame_col1_frame_row2,
                                                                     orient="vertical",
                                                                     command=middle_left_thickness_frame_col1_frame_row2_canvas.yview)
middle_left_thickness_frame_col1_frame_row2_canvas.configure(
    yscrollcommand=middle_left_thickness_frame_col1_frame_row2_scrollbar.set)
middle_left_thickness_frame_col1_frame_row2_scrollable_frame = tk.Frame(
    middle_left_thickness_frame_col1_frame_row2_canvas, bg=bg_app_class_color_layer_2)
middle_left_thickness_frame_col1_frame_row2_canvas.create_window((0, 0),
                                                                 window=middle_left_thickness_frame_col1_frame_row2_scrollable_frame,
                                                                 anchor="nw")
middle_left_thickness_frame_col1_frame_row2_scrollable_frame.bind("<Configure>", lambda
    e: middle_left_thickness_frame_col1_frame_row2_canvas.configure(
    scrollregion=middle_left_thickness_frame_col1_frame_row2_canvas.bbox("all")))
middle_left_thickness_frame_col1_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_thickness_frame_col1_frame_row2_scrollbar.pack(side="right", fill="y")
all_entries = []

middle_left_weight_frame_col3_frame_row2_canvas = tk.Canvas(middle_left_weight_frame_col3_frame_row2,
                                                            bg=bg_app_class_color_layer_2, highlightthickness=0)
middle_left_weight_frame_col3_frame_row2_scrollbar = tk.Scrollbar(middle_left_weight_frame_col3_frame_row2,
                                                                  orient="vertical",
                                                                  command=middle_left_weight_frame_col3_frame_row2_canvas.yview)
middle_left_weight_frame_col3_frame_row2_scrollable_frame = tk.Frame(middle_left_weight_frame_col3_frame_row2_canvas,
                                                                     bg=bg_app_class_color_layer_2)
middle_left_weight_frame_col3_frame_row2_canvas.create_window((0, 0),
                                                              window=middle_left_weight_frame_col3_frame_row2_scrollable_frame,
                                                              anchor="nw")
middle_left_weight_frame_col3_frame_row2_canvas.configure(
    yscrollcommand=middle_left_weight_frame_col3_frame_row2_scrollbar.set)
middle_left_weight_frame_col3_frame_row2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_col3_frame_row2_scrollbar.pack(side="right", fill="y")

selected_period_icon = None
selected_period_button = None
selected_period_button_hover = None
middle_right_runcard_frame_row2_col3_frame_dict = {}
middle_right_runcard_frame_row2_col3_button_dict = {}
middle_right_runcard_frame_row2_col3_button_icon_dict = {}
middle_right_runcard_frame_row2_col3_button_icon_hover_dict = {}
preloaded_period_icons = {}
preloaded_period_hover_icons = {}


def preload_period_icons():
    """Preload and resize all period icons once at app start."""
    for period in period_times:
        period_icon_path = os.path.join(base_path, "theme", "icons", "period", f"time_{period}.png")
        period_hover_icon_path = os.path.join(base_path, "theme", "icons", "period", f"time_{period}_hover.png")

        if period not in preloaded_period_icons:
            preloaded_period_icons[period] = ImageTk.PhotoImage(Image.open(period_icon_path).resize((36, 26)))
            preloaded_period_hover_icons[period] = ImageTk.PhotoImage(
                Image.open(period_hover_icon_path).resize((36, 26)))


def runcard_manage_period_button(value, btn):
    global selected_period_button, selected_period_icon, selected_period_button_hover
    if selected_period_button:
        selected_period_button.config(
            image=middle_right_runcard_frame_row2_col3_button_icon_dict[selected_period_button_hover])
    btn.config(image=middle_right_runcard_frame_row2_col3_button_icon_hover_dict[value])
    selected_period_button = btn
    selected_period_icon = middle_right_runcard_frame_row2_col3_button_icon_hover_dict[value]
    selected_period_button_hover = value

    def wait_for_machine_then_load_wo():
        if selected_machine_button_hover is None:
            root.after(0, wait_for_machine_then_load_wo)
        else:
            runcard_wo_button()

    root.after(0, wait_for_machine_then_load_wo)


def runcard_period_button(index=0):
    if int(get_registry_value("is_connected", "1")) != 1:
        return
    if not preloaded_period_icons or not preloaded_period_hover_icons:
        preload_period_icons()
    if index < len(period_times):
        period = period_times[index]
        period_icon = preloaded_period_icons[period]
        period_hover_icon = preloaded_period_hover_icons[period]

        middle_right_runcard_frame_row2_col3_button_icon_dict[period] = period_icon
        middle_right_runcard_frame_row2_col3_button_icon_hover_dict[period] = period_hover_icon

        period_frame_name = f"middle_right_runcard_frame_row2_col3_row{index + 1}"
        frame = tk.Frame(middle_right_runcard_frame_row2_col3, bg=bg_app_class_color_layer_1)
        frame.place(x=0, y=index * 29, width=36, height=26)

        button = tk.Button(frame, image=period_icon, width=36, height=26, relief="flat")
        button.config(command=lambda p=period, b=button: runcard_manage_period_button(p, b))
        button.place(x=0, y=0, relwidth=1.0, relheight=1.0)

        middle_right_runcard_frame_row2_col3_frame_dict[period_frame_name] = frame
        middle_right_runcard_frame_row2_col3_button_dict[period] = button

        root.after(1, lambda: runcard_period_button(index + 1))
    else:
        if current_time in middle_right_runcard_frame_row2_col3_button_dict:
            root.after(0, lambda: runcard_manage_period_button(current_time,
                                                               middle_right_runcard_frame_row2_col3_button_dict[
                                                                   current_time]))


selected_machine_icon = None
selected_machine_button = None
selected_machine_button_hover = None
middle_right_machine_frame_row2_col1_frame_dict = {}
middle_right_machine_frame_row2_col1_button_dict = {}
middle_right_machine_frame_row2_col1_button_icon_dict = {}
middle_right_machine_frame_row2_col1_button_icon_hover_dict = {}
preloaded_machine_icons = {}
preloaded_machine_hover_icons = {}


def preload_machine_icons(machine_list):
    for index, machine in enumerate(machine_list):
        if machine not in preloaded_machine_icons:
            machine_icon_path = os.path.join(base_path, "theme", "icons", "machine", f"machine_{index+1}.png")
            machine_hover_icon_path = os.path.join(base_path, "theme", "icons", "machine",
                                                   f"machine_{index+1}_hover.png")
            preloaded_machine_icons[machine] = ImageTk.PhotoImage(Image.open(machine_icon_path).resize((36, 26)))
            preloaded_machine_hover_icons[machine] = ImageTk.PhotoImage(
                Image.open(machine_hover_icon_path).resize((36, 26)))


def runcard_get_machine_list(plant):
    global conn_str
    sql = f"""SELECT DISTINCT(Name) FROM [PMGMES].[dbo].[PMG_DML_DataModelList]
            WHERE DataModelTypeId = 'DMT000003' AND Name LIKE ?  ORDER BY Name"""
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, (f"%{plant}%",))
        result = cursor.fetchall()
        return [machine[0] for machine in result] if result else []


def runcard_manage_machine_button(value, btn):
    global selected_machine_button, selected_machine_icon, selected_machine_button_hover
    if selected_machine_button:
        selected_machine_button.config(
            image=middle_right_machine_frame_row2_col1_button_icon_dict[selected_machine_button_hover])
    btn.config(image=middle_right_machine_frame_row2_col1_button_icon_hover_dict[value])
    selected_machine_button = btn
    selected_machine_icon = middle_right_machine_frame_row2_col1_button_icon_hover_dict[value]
    selected_machine_button_hover = value

    def wait_for_machine_then_load_wo():
        if selected_machine_button_hover is None:
            root.after(0, wait_for_machine_then_load_wo)
        else:
            runcard_wo_button()

    root.after(0, wait_for_machine_then_load_wo)


def runcard_machine_button_lazy_loader(machine_list, index=0):
    if index < len(machine_list):
        machine = machine_list[index]
        machine_icon = preloaded_machine_icons[machine]
        machine_hover_icon = preloaded_machine_hover_icons[machine]

        middle_right_machine_frame_row2_col1_button_icon_dict[machine] = machine_icon
        middle_right_machine_frame_row2_col1_button_icon_hover_dict[machine] = machine_hover_icon

        machine_frame_name = f"middle_right_machine_frame_row2_col1_row{index + 1}"
        frame = tk.Frame(middle_right_runcard_frame_row2_col1, bg=bg_app_class_color_layer_1)
        frame.place(x=0, y=index * 29, width=36, height=26)

        button = tk.Button(frame, image=machine_icon, width=36, height=26, relief="flat")
        button.config(command=lambda p=machine, b=button: runcard_manage_machine_button(p, b))
        button.place(x=0, y=0, relwidth=1.0, relheight=1.0)

        middle_right_machine_frame_row2_col1_frame_dict[machine_frame_name] = frame
        middle_right_machine_frame_row2_col1_button_dict[machine] = button

        root.after(1, lambda: runcard_machine_button_lazy_loader(machine_list, index + 1))
    else:
        # Auto-select the first machine
        if machine_list:
            root.after(0, lambda: runcard_manage_machine_button(
                machine_list[0],
                middle_right_machine_frame_row2_col1_button_dict[machine_list[0]]
            ))


def runcard_machine_button():
    if int(get_registry_value("is_connected", "1")) != 1:
        return
    plant = get_registry_value("is_plant_name", "PVC1")
    machine_list = runcard_get_machine_list(plant)
    preload_machine_icons(machine_list)
    runcard_machine_button_lazy_loader(machine_list)


selected_line_icon = None
selected_line_button = None
selected_line_button_hover = None
middle_right_line_frame_row2_col1_row3_frame_dict = {}
middle_right_line_frame_row2_col1_row3_button_dict = {}
middle_right_line_frame_row2_col1_row3_button_icon_dict = {}
middle_right_line_frame_row2_col1_row3_button_icon_hover_dict = {}


def runcard_get_line_list(machine):
    global conn_str
    return ['A1', 'B1'] if 'PVC' in str(machine).upper() else ['A1', 'B1', 'A2', 'B2']


def runcard_manage_line_button(value, btn):
    global selected_line_button, selected_line_icon, selected_line_button_hover
    if selected_line_button:
        selected_line_button.config(
            image=middle_right_line_frame_row2_col1_row3_button_icon_dict[selected_line_button_hover])
    btn.config(image=middle_right_line_frame_row2_col1_row3_button_icon_hover_dict[value])
    selected_line_button = btn
    selected_line_icon = middle_right_line_frame_row2_col1_row3_button_icon_dict[value]
    selected_line_button_hover = value
    # root.after(0, runcard_wo_button)


def runcard_line_button():
    global selected_machine_button_hover
    line_list = runcard_get_line_list(selected_machine_button_hover)
    middle_width = int(len(line_list) * 48)
    side_width = int((432 - middle_width) / 2)
    middle_right_runcard_frame_row2_col2_row3_left_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3,
                                                                    bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_left_frame.place(x=0, y=0, width=side_width, height=40)
    middle_right_runcard_frame_row2_col2_row3_middle_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3,
                                                                      bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_middle_frame.place(x=side_width, y=0, width=middle_width, height=40)
    middle_right_runcard_frame_row2_col2_row3_right_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3,
                                                                     bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_right_frame.place(x=side_width + middle_width, y=0, width=side_width,
                                                                height=40)
    if int(get_registry_value("is_connected", "1")) == 1:
        for index, line in enumerate(line_list):
            line_icon_path = os.path.join(base_path, "theme", "icons", "line", f"line_{line}.png")
            line_hover_icon_path = os.path.join(base_path, "theme", "icons", "line", f"line_{line}_hover.png")
            line_icon = ImageTk.PhotoImage(Image.open(line_icon_path).resize((36, 26)))
            line_hover_icon = ImageTk.PhotoImage(Image.open(line_hover_icon_path).resize((36, 26)))

            middle_right_line_frame_row2_col1_row3_button_icon_dict[line] = line_icon
            middle_right_line_frame_row2_col1_row3_button_icon_hover_dict[line] = line_hover_icon
            line_frame_name = f"middle_right_runcard_frame_row2_col2_row3_col{index + 1}"
            middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name] = tk.Frame(
                middle_right_runcard_frame_row2_col2_row3_middle_frame, bg=bg_app_class_color_layer_1)
            middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name].place(x=index * 48, y=0, width=40,
                                                                                     height=29)

            button = tk.Button(middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name], image=line_icon,
                               width=36, height=26, relief="flat")
            button.config(command=lambda p=line, b=button: runcard_manage_line_button(p, b))

            button.place(x=0, y=0, relwidth=1.0, relheight=1.0)
            middle_right_line_frame_row2_col1_row3_button_dict[line] = button
        if line_list:
            runcard_manage_line_button(line_list[0], middle_right_line_frame_row2_col1_row3_button_dict[line_list[0]])


def runcard_get_runcard_id(date, machine, line, time):
    global conn_str
    sql = f"""SELECT rc.Id as runcard, wo.Id as workorder
            FROM [PMGMES].[dbo].[PMG_MES_RunCard] rc
            JOIN [PMGMES].[dbo].[PMG_MES_WorkOrder] wo
            on rc.WorkOrderId = wo.Id
            where rc.MachineName = ? and rc.LineName = ?
            and cast(Period as int) = ? and wo.StartDate is not NULL
            and ((rc.Period > 5 AND rc.InspectionDate = ?)
            OR (rc.Period <= 5 AND rc.InspectionDate = CAST(DATEADD(DAY, 1, CAST(? AS DATE)) AS DATE)))"""
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql, (machine, line, int(time), str(date), str(date),))
            result = cursor.fetchall()
            return [wo for wo in result] if result else []
    except:
        def format_sql(sql, params):
            for p in params:
                if isinstance(p, str):
                    p_str = f"'{p}'"
                else:
                    p_str = str(p)
                sql = sql.replace('?', p_str, 1)
            return sql

        print(format_sql(sql, (machine, line, int(time), str(date), str(date))))


selected_wo_button = None


def runcard_wo_button():
    global selected_wo_button, selected_wo_value, current_runcard_id
    global selected_date, selected_machine_button_hover, selected_line_button_hover
    for widget in middle_right_runcard_frame_row2_col2_row2.winfo_children():
        widget.destroy()

    wo_list = runcard_get_runcard_id(selected_date, selected_machine_button_hover, selected_line_button_hover,
                                     selected_period_button_hover)
    if not wo_list:
        global middle_right_runcard_frame_row2_col2_row4_row2
        global middle_right_runcard_frame_row2_col2_row4_row3
        # runcard_info_empty()
        root.after(0, lambda: runcard_info_show(
            [selected_date, '', selected_machine_button_hover[-8:-5], '', selected_machine_button_hover[-3:],
             selected_line_button_hover] + [""] * 13))
        for widget in middle_right_runcard_frame_row2_col2_row4_row2.winfo_children():
            widget.destroy()
        for widget in middle_right_runcard_frame_row2_col2_row4_row3.winfo_children():
            widget.destroy()
        barcode_label = tk.Label(middle_right_runcard_frame_row2_col2_row4_row2,
                                 text=f"沒有 Runcard !\nKhông có Runcard !", bg="white", bd=0, font=(font_name, 16))
        barcode_label.pack(expand=True)
        return
    button_width = 0 if len(wo_list) == 1 else 120
    button_height = 36
    padding = 8
    total_width = len(wo_list) * (button_width + padding)
    side_width = max((438 - total_width) // 2, 0)

    tk.Frame(middle_right_runcard_frame_row2_col2_row2, bg='white').place(x=0, y=0, width=side_width, height=40)
    middle_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row2, bg='white')
    middle_frame.place(x=side_width, y=0, width=total_width, height=40)

    tk.Frame(middle_right_runcard_frame_row2_col2_row2, bg='white').place(x=side_width + total_width, y=0,
                                                                          width=side_width, height=40)
    wo_button_dict = {}
    first_button = None

    def wo_button_clicked(wo_value, btn):
        global selected_wo_button, selected_wo_value, current_runcard_id
        if selected_wo_button:
            try:
                selected_wo_button.config(bg="#007bff" if len(wo_list) > 1 else "white", fg="white")
            except tk.TclError:
                selected_wo_button = None
        btn.config(bg="#E84A41" if len(wo_list) > 1 else "white", fg="white")
        selected_wo_button = btn
        selected_wo_value = wo_value
        current_runcard_id = wo_value
        print(f"==>--> {wo_value}")
        runcard_id_barcode(wo_value)
        root.after(0, lambda: runcard_info_show(get_runcard_info(wo_value)))
        # runcard_info_show()

    for index, (runcard, workorder) in enumerate(wo_list):
        frame = tk.Frame(middle_frame, bg='white')
        frame.place(x=index * (button_width + padding), y=2, width=button_width, height=button_height)
        btn = tk.Button(frame, text=workorder, bg="#007bff", fg="white", relief="flat", font=("Arial", 11),
                        command=lambda p=runcard, b=None: None)
        btn.place(x=0, y=0, relwidth=1.0, relheight=1.0)

        btn.configure(command=lambda p=runcard, b=btn: wo_button_clicked(p, b))
        wo_button_dict[runcard] = btn

        if first_button is None:
            first_button = (runcard, btn)
    if first_button:
        wo_button_clicked(first_button[0], first_button[1])


def runcard_info_show(info):
    for widget in middle_right_runcard_frame_row2_col2_row1_row1.winfo_children():
        widget.destroy()

    runcard_row1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                  highlightbackground="black", highlightthickness=1)
    runcard_row1_frame.place(x=0, y=0, width=418, height=25)
    runcard_row1_label = tk.Label(runcard_row1_frame, text=f"工站生產流程卡 Thẻ quy trình sản xuất", bg="white", bd=1,
                                  font=(font_name, 12, 'bold'))
    runcard_row1_label.pack(expand=True)

    runcard_row2_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row2_col1_frame.place(x=0, y=24, width=102, height=25)
    runcard_row2_col1_label = tk.Label(runcard_row2_col1_frame, text=f"Ngày", bg="white", bd=1, font=(font_name, 11))
    runcard_row2_col1_label.pack(expand=True)

    runcard_row2_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row2_col2_frame.place(x=101, y=24, width=108, height=25)
    runcard_row2_col2_label = tk.Label(runcard_row2_col2_frame, text=f"{info[0]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row2_col2_label.pack(expand=True)

    runcard_row2_col3_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row2_col3_frame.place(x=208, y=24, width=102, height=25)
    runcard_row2_col3_label = tk.Label(runcard_row2_col3_frame, text=f"Mã vật tư", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row2_col3_label.pack(expand=True)

    runcard_row2_col4_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row2_col4_frame.place(x=309, y=24, width=109, height=25)
    runcard_row2_col4_label = tk.Label(runcard_row2_col4_frame, text=f"{info[1]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row2_col4_label.pack(expand=True)

    runcard_row3_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row3_col1_frame.place(x=0, y=48, width=102, height=25)
    runcard_row3_col1_label = tk.Label(runcard_row3_col1_frame, text=f"Xưởng", bg="white", bd=1, font=(font_name, 11))
    runcard_row3_col1_label.pack(expand=True)

    runcard_row3_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row3_col2_frame.place(x=101, y=48, width=108, height=25)
    runcard_row3_col2_label = tk.Label(runcard_row3_col2_frame, text=f"{info[2]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row3_col2_label.pack(expand=True)

    runcard_row3_col3_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row3_col3_frame.place(x=208, y=48, width=102, height=25)
    runcard_row3_col3_label = tk.Label(runcard_row3_col3_frame, text=f"Mã khách hàng", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row3_col3_label.pack(expand=True)

    runcard_row3_col4_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row3_col4_frame.place(x=309, y=48, width=109, height=25)
    runcard_row3_col4_label = tk.Label(runcard_row3_col4_frame, text=f"{info[3]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row3_col4_label.pack(expand=True)

    runcard_row4_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row4_col1_frame.place(x=0, y=72, width=102, height=25)
    runcard_row4_col1_label = tk.Label(runcard_row4_col1_frame, text=f"Máy", bg="white", bd=1, font=(font_name, 11))
    runcard_row4_col1_label.pack(expand=True)

    runcard_row4_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row4_col2_frame.place(x=101, y=72, width=108, height=25)
    runcard_row4_col2_label = tk.Label(runcard_row4_col2_frame, text=f"{info[4]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row4_col2_label.pack(expand=True)

    runcard_row5_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row5_col1_frame.place(x=0, y=96, width=102, height=25)
    runcard_row5_col1_label = tk.Label(runcard_row5_col1_frame, text=f"Line", bg="white", bd=1, font=(font_name, 11))
    runcard_row5_col1_label.pack(expand=True)

    runcard_row5_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row5_col2_frame.place(x=101, y=96, width=108, height=25)
    runcard_row5_col2_label = tk.Label(runcard_row5_col2_frame, text=f"{info[5]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row5_col2_label.pack(expand=True)

    runcard_row45_col3_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                        highlightbackground="black", highlightthickness=1)
    runcard_row45_col3_frame.place(x=208, y=72, width=102, height=49)
    runcard_row45_col3_label = tk.Label(runcard_row45_col3_frame, text=f"Tên viết tắt\nkhách hàng", bg="white", bd=1,
                                        font=(font_name, 11))
    runcard_row45_col3_label.pack(expand=True)

    runcard_row45_col4_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                        highlightbackground="black", highlightthickness=1)
    runcard_row45_col4_frame.place(x=309, y=72, width=109, height=49)
    runcard_row45_col4_label = tk.Label(runcard_row45_col3_frame, text=f"", bg="white", bd=1, font=(font_name, 11))
    runcard_row45_col4_label.pack(expand=True)

    runcard_row6_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row6_col1_frame.place(x=0, y=120, width=102, height=25)
    runcard_row6_col1_label = tk.Label(runcard_row6_col1_frame, text=f"Công đơn", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row6_col1_label.pack(expand=True)

    runcard_row6_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row6_col2_frame.place(x=101, y=120, width=108, height=25)
    runcard_row6_col2_label = tk.Label(runcard_row6_col2_frame, text=f"{info[7]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row6_col2_label.pack(expand=True)

    runcard_row7_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row7_col1_frame.place(x=0, y=144, width=102, height=25)
    runcard_row7_col1_label = tk.Label(runcard_row7_col1_frame, text=f"AQL", bg="white", bd=1, font=(font_name, 11))
    runcard_row7_col1_label.pack(expand=True)

    runcard_row7_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row7_col2_frame.place(x=101, y=144, width=108, height=25)
    runcard_row7_col2_label = tk.Label(runcard_row7_col2_frame, text=f"{info[8]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row7_col2_label.pack(expand=True)

    runcard_row67_col3_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                        highlightbackground="black", highlightthickness=1)
    runcard_row67_col3_frame.place(x=208, y=120, width=102, height=49)
    runcard_row67_col3_label = tk.Label(runcard_row67_col3_frame, text=f"Loại", bg="white", bd=1, font=(font_name, 11))
    runcard_row67_col3_label.pack(expand=True)

    runcard_row67_col4_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                        highlightbackground="black", highlightthickness=1)
    runcard_row67_col4_frame.place(x=309, y=120, width=109, height=49)
    runcard_row67_col4_label = tk.Label(runcard_row67_col4_frame, text=f"{info[9]}", bg="white", bd=1,
                                        font=(font_name, 11))
    runcard_row67_col4_label.pack(expand=True)

    runcard_row8_col1_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row8_col1_frame.place(x=0, y=168, width=102, height=49)
    runcard_row8_col1_label = tk.Label(runcard_row8_col1_frame, text=f"Người kiểm tra", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row8_col1_label.pack(expand=True)

    runcard_row8_col2_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row8_col2_frame.place(x=101, y=168, width=108, height=49)
    runcard_row8_col2_label = tk.Label(runcard_row8_col2_frame, text=f"{info[10]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row8_col2_label.pack(expand=True)

    runcard_row8_col3_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row8_col3_frame.place(x=208, y=168, width=102, height=49)
    runcard_row8_col3_label = tk.Label(runcard_row8_col3_frame, text=f"Kích cỡ", bg="white", bd=1, font=(font_name, 11))
    runcard_row8_col3_label.pack(expand=True)

    runcard_row8_col4_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row1_row1, bg='white', bd=1,
                                       highlightbackground="black", highlightthickness=1)
    runcard_row8_col4_frame.place(x=309, y=168, width=109, height=49)
    runcard_row8_col4_label = tk.Label(runcard_row8_col4_frame, text=f"{info[11]}", bg="white", bd=1,
                                       font=(font_name, 11))
    runcard_row8_col4_label.pack(expand=True)


def get_runcard_info(runcard):
    global conn_str
    sql = """SELECT  rc.InspectionDate, wo.PartNo, rc.WorkCenterTypeName, wo.CustomerCode,
            rc.MachineName, rc.LineName, wo.CustomerName,  rc.WorkOrderId, wo.AQL,
            wo.ProductItem, au.Name, std.Size
            FROM [PMGMES].[dbo].[PMG_MES_RunCard] rc
            join [PMGMES].[dbo].[PMG_MES_WorkOrder] wo
            on wo.id = rc.WorkOrderId
            left join [PMGMES].[dbo].[PMG_MES_IPQCInspectingRecord] ir
            on ir.RunCardId = rc.id
            join [PMGMES].[dbo].[AbpUsers] au
            on rc.CreatorUserId = au.Id
            join [PMGMES].[dbo].[PMG_MES_PVC_SCADA_Std] std
            on std.PartNo = wo.PartNo
            WHERE rc.Id = ?
            Group by rc.InspectionDate, wo.PartNo, rc.WorkCenterTypeName, wo.CustomerCode,
            rc.MachineName, rc.LineName, wo.CustomerName,  rc.WorkOrderId, wo.AQL,
            wo.ProductItem, au.Name, std.Size"""
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, (runcard,))
        result = cursor.fetchall()
        return [info[-3:] if index == 4 else info for index, info in enumerate(result[0])] if result else []


def runcard_id_barcode(runcard):
    global middle_right_runcard_frame_row2_col2_row4_row2
    global middle_right_runcard_frame_row2_col2_row4_row3
    for widget in middle_right_runcard_frame_row2_col2_row4_row2.winfo_children():
        widget.destroy()
    for widget in middle_right_runcard_frame_row2_col2_row4_row3.winfo_children():
        widget.destroy()
    middle_right_runcard_frame_row2_col2_row4_row2 = tk.Frame(middle_right_runcard_frame_row2_col2_row4,
                                                              bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row4_row2.pack()
    barcode_class = barcode.get_barcode_class('code128')
    barcode_png = barcode_class(runcard, writer=ImageWriter())
    buffer = BytesIO()
    barcode_png.write(buffer)
    buffer.seek(0)
    pil_image = Image.open(buffer)
    pil_image = pil_image.resize((300, 100), Image.ANTIALIAS)
    barcode_image = ImageTk.PhotoImage(pil_image)

    barcode_runcard = tk.Label(middle_right_runcard_frame_row2_col2_row4_row2, image=barcode_image)
    barcode_runcard.image = barcode_image
    barcode_runcard.place(x=int((438 - 300) / 2), y=0, width=300, height=60)

    barcode_label = tk.Label(middle_right_runcard_frame_row2_col2_row4_row3, text=f"{runcard}", bg="white", bd=0,
                             font=(font_name, 16))
    barcode_label.pack(expand=True)


if int(get_registry_value("is_runcard_open", "0")) == 1:
    root.after(0, runcard_date)

"""Settings"""

selected_middle_left_frame = tk.StringVar(value=get_registry_value("SelectedFrame", "Trọng lượng"))


def update_com_ports(*menus):
    try:
        def monitor_com_ports():
            global com_ports
            while True:
                new_com_ports = [port.device for port in serial.tools.list_ports.comports() if
                                 "Bluetooth" not in port.description]
                if set(new_com_ports) != set(com_ports):
                    com_ports = new_com_ports
                    root.after(0, lambda: populate_com_menus(menus))
                time.sleep(2)

        def populate_com_menus(menus):
            for menu in menus:
                menu['menu'].delete(0, 'end')
                menu['menu'].add_command(label="-------",
                                         command=lambda m=menu: m.setvar(m.cget("textvariable"), value="-------"))
                for port in com_ports:
                    menu['menu'].add_command(label=port,
                                             command=lambda p=port, m=menu: m.setvar(m.cget("textvariable"), value=p))

        global com_ports
        com_ports = [port.device for port in serial.tools.list_ports.comports() if "Bluetooth" not in port.description]
        populate_com_menus(menus)
        if not hasattr(update_com_ports, "thread_started"):
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
                                threading.Thread(target=show_error_message,
                                                 args=("Chưa nhập giá trị Runcard!", 0, 3000), daemon=True).start()
                    else:
                        threading.Thread(target=show_error_message,
                                         args=("Kiểm tra lại kết nối với dụng cụ đo!", 0, 3000), daemon=True).start()
        except Exception as e:
            threading.Thread(target=show_error_message,
                             args=(f"def thickness_frame_com_port_insert_data => {e}", 0, 3000), daemon=True).start()
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
                        threading.Thread(target=show_error_message, args=(f"Change your weight unit to gram!", 0, 3000),
                                         daemon=True).start()
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
            threading.Thread(target=show_error_message, args=(f"def weight_frame_com_port_insert_data => {e}", 0, 3000),
                             daemon=True).start()
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
        threading.Thread(target=show_error_message, args=(f"def switch_middle_left_frame => {e}", 0, 3000),
                         daemon=True).start()
        pass


switch_middle_left_frame()
selected_weight_com = tk.StringVar(value=get_registry_value("COM1", ""))
selected_thickness_com = tk.StringVar(value=get_registry_value("COM2", ""))

weight_label = tk.Label(middle_right_setting_frame_row2_row1, text="Trọng lượng:      ", font=(font_name, 14, "bold"),
                        bg=bg_app_class_color_layer_2)
weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
weight_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_weight_com, "")
weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

thickness_label = tk.Label(middle_right_setting_frame_row2_row1, text="Độ dày:", font=(font_name, 14, "bold"),
                           bg=bg_app_class_color_layer_2)
thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
thickness_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_thickness_com, "")
thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

frame_select_label = tk.Label(middle_right_setting_frame_row2_row1, text="Mặc định:", font=(font_name, 14, "bold"),
                              bg=bg_app_class_color_layer_2)
frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
frame_select_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_middle_left_frame, "Trọng lượng",
                                     "Độ dày", command=switch_middle_left_frame)
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
        root.destroy()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"def exit() => {e}", 0, 3000), daemon=True).start()
        pass


def on_enter_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_hover_icon)


def on_leave_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_icon)


top_open_weight_frame_button_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "weight.png")).resize((156, 36)))
top_open_weight_frame_button_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "weight_hover.png")).resize((156, 36)))
top_open_weight_frame_button = tk.Button(top_right_frame, image=top_open_weight_frame_button_icon,
                                         command=open_weight_frame, bg='#f0f2f6', width=156, height=36, relief="flat",
                                         borderwidth=0)
top_open_weight_frame_button.grid(row=0, column=3, padx=10, pady=5, sticky="e")
top_open_weight_frame_button.bind("<Enter>", on_enter_top_open_weight_frame_button)
top_open_weight_frame_button.bind("<Leave>", on_leave_top_open_weight_frame_button)


def on_enter_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_hover_icon)


def on_leave_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_icon)


top_open_thickness_frame_button_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "thickness.png")).resize((156, 36)))
top_open_thickness_frame_button_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "thickness_hover.png")).resize((156, 36)))
top_open_thickness_frame_button = tk.Button(top_right_frame, image=top_open_thickness_frame_button_icon,
                                            command=open_thickness_frame, bg='#f0f2f6', width=156, height=36,
                                            relief="flat", borderwidth=0)
top_open_thickness_frame_button.grid(row=0, column=4, padx=5, pady=5, sticky="e")
top_open_thickness_frame_button.bind("<Enter>", on_enter_top_open_thickness_frame_button)
top_open_thickness_frame_button.bind("<Leave>", on_leave_top_open_thickness_frame_button)


def on_enter_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_hover_icon)


def on_leave_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_icon)


setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "setting.png")).resize((42, 42)))
setting_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "setting_hover.png")).resize((42, 42)))
middle_open_setting_frame_button = tk.Button(top_left_frame, image=setting_icon, bd=0, command=open_setting_frame,
                                             bg=bg_app_class_color_layer_1)
middle_open_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")
middle_open_setting_frame_button.bind("<Enter>", on_enter_middle_open_setting_frame_button)
middle_open_setting_frame_button.bind("<Leave>", on_leave_middle_open_setting_frame_button)


def on_enter_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_hover_icon)


def on_leave_middle_open_advance_setting_frame_button(event):
    middle_open_advance_setting_frame_button.config(image=advance_setting_icon)


advance_setting_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "advance_setting.png")).resize((21, 21)))
advance_setting_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "advance_setting_hover.png")).resize((21, 21)))
middle_open_advance_setting_frame_button = tk.Button(middle_right_setting_frame_row1_col2, image=advance_setting_icon,
                                                     bd=0, command=open_advance_setting_frame,
                                                     bg=bg_app_class_color_layer_1)
middle_open_advance_setting_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_advance_setting_frame_button.bind("<Enter>", on_enter_middle_open_advance_setting_frame_button)
middle_open_advance_setting_frame_button.bind("<Leave>", on_leave_middle_open_advance_setting_frame_button)


def on_enter_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_hover_icon)


def on_leave_middle_open_return_frame_button(event):
    middle_open_return_frame_button.config(image=return_icon)


return_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "return.png")).resize((21, 21)))
return_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "return_hover.png")).resize((21, 21)))
middle_open_return_frame_button = tk.Button(middle_right_advance_setting_frame_row1_col2, image=return_icon, bd=0,
                                            command=open_setting_frame, bg=bg_app_class_color_layer_1)
middle_open_return_frame_button.grid(row=0, column=0, padx=20, pady=5, sticky="e")
middle_open_return_frame_button.bind("<Enter>", on_enter_middle_open_return_frame_button)
middle_open_return_frame_button.bind("<Leave>", on_leave_middle_open_return_frame_button)


def on_enter_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_hover_icon)


def on_leave_middle_open_runcard_frame_button(event):
    middle_open_runcard_frame_button.config(image=barcode_icon)


barcode_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "barcode.png")).resize((42, 42)))
barcode_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "barcode_hover.png")).resize((42, 42)))
middle_open_runcard_frame_button = tk.Button(top_left_frame, image=barcode_icon, bd=0, command=open_runcard_frame,
                                             bg=bg_app_class_color_layer_1)
middle_open_runcard_frame_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")
middle_open_runcard_frame_button.bind("<Enter>", on_enter_middle_open_runcard_frame_button)
middle_open_runcard_frame_button.bind("<Leave>", on_leave_middle_open_runcard_frame_button)

close_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "close.png")).resize((134, 34)))
close_icon_hover = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "close_hover.png")).resize((134, 34)))


def on_enter_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon_hover)


def on_leave_close_setting_frame_button(event):
    close_setting_frame_button.config(image=close_icon)


close_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col2, image=close_icon, command=close_frame,
                                       bg=bg_app_class_color_layer_1, width=134, height=34, relief="flat",
                                       borderwidth=0)
close_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="e")
close_setting_frame_button.bind("<Enter>", on_enter_close_setting_frame_button)
close_setting_frame_button.bind("<Leave>", on_leave_close_setting_frame_button)


def on_enter_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon_hover)


def on_leave_close_advance_setting_frame_button(event):
    close_advance_setting_frame_button.config(image=close_icon)


close_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col2, image=close_icon,
                                               command=close_frame, bg=bg_app_class_color_layer_1, width=134, height=34,
                                               relief="flat", borderwidth=0)
close_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
close_advance_setting_frame_button.bind("<Enter>", on_enter_close_advance_setting_frame_button)
close_advance_setting_frame_button.bind("<Leave>", on_leave_close_advance_setting_frame_button)


def on_enter_close_runcard_frame_button(event):
    close_runcard_frame_button.config(image=close_icon_hover)


def on_leave_close_runcard_frame_button(event):
    close_runcard_frame_button.config(image=close_icon)


close_runcard_frame_button = tk.Button(middle_right_runcard_frame_row3, image=close_icon, command=close_frame,
                                       bg=bg_app_class_color_layer_1, width=134, height=34, relief="flat",
                                       borderwidth=0)
close_runcard_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
close_runcard_frame_button.bind("<Enter>", on_enter_close_runcard_frame_button)
close_runcard_frame_button.bind("<Leave>", on_leave_close_runcard_frame_button)

save_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "save.png")).resize((134, 34)))
save_icon_hover = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "save_hover.png")).resize((134, 34)))


def on_enter_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon_hover)


def on_leave_save_setting_frame_button(event):
    save_setting_frame_button.config(image=save_icon)


save_setting_frame_button = tk.Button(middle_right_setting_frame_row3_col1, image=save_icon, command=save_setting,
                                      bg=bg_app_class_color_layer_1, width=134, height=34, relief="flat", borderwidth=0)
save_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_setting_frame_button.bind("<Enter>", on_enter_save_setting_frame_button)
save_setting_frame_button.bind("<Leave>", on_leave_save_setting_frame_button)


def on_enter_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon_hover)


def on_leave_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon)


save_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col1, image=save_icon,
                                              command=None, bg=bg_app_class_color_layer_1, width=134, height=34,
                                              relief="flat", borderwidth=0)
save_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_advance_setting_frame_button.bind("<Enter>", on_enter_save_advance_setting_frame_button)
save_advance_setting_frame_button.bind("<Leave>", on_leave_save_advance_setting_frame_button)


def on_enter_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_hover_icon)


def on_leave_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_icon)


database_test_connection_button_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "connect.png")).resize((146, 38)))
database_test_connection_button_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "connect_hover.png")).resize((146, 38)))
database_test_connection_button = tk.Button(middle_right_advance_setting_frame_row2_col1_row6_col2,
                                            image=database_test_connection_button_icon, command=None,
                                            bg=bg_app_class_color_layer_1, width=146, height=38, relief="flat",
                                            borderwidth=0)
database_test_connection_button.grid(row=4, column=1, padx=5, pady=5, sticky="e")
database_test_connection_button.bind("<Enter>", on_enter_database_test_connection_button)
database_test_connection_button.bind("<Leave>", on_leave_database_test_connection_button)

delete_entry_button_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "delete.png")).resize((117, 36)))
delete_entry_button_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "delete_hover.png")).resize((117, 36)))


def on_enter_delete_weight_entry_button(event):
    delete_weight_entry_button.config(image=delete_entry_button_hover_icon)


def on_leave_delete_weight_entry_button(event):
    delete_weight_entry_button.config(image=delete_entry_button_icon)


delete_weight_entry_button = tk.Button(middle_left_weight_frame_col3_frame_row1_row1, image=delete_entry_button_icon,
                                       command=None, bg='#f0f2f6', width=120, height=40, relief="flat", borderwidth=0)
delete_weight_entry_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_weight_entry_button.bind("<Enter>", on_enter_delete_weight_entry_button)
delete_weight_entry_button.bind("<Leave>", on_leave_delete_weight_entry_button)


def on_enter_delete_thickness_entry_button(event):
    delete_thickness_entry_button.config(image=delete_entry_button_hover_icon)


def on_leave_delete_thickness_entry_button(event):
    delete_thickness_entry_button.config(image=delete_entry_button_icon)


delete_thickness_entry_button = tk.Button(middle_left_thickness_frame_col3_frame_row1_row1,
                                          image=delete_entry_button_icon, command=None, bg='#f0f2f6', width=120,
                                          height=40, relief="flat", borderwidth=0)
delete_thickness_entry_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_thickness_entry_button.bind("<Enter>", on_enter_delete_thickness_entry_button)
delete_thickness_entry_button.bind("<Leave>", on_leave_delete_thickness_entry_button)

delete_log_button_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "delete_all.png")).resize((36, 36)))
delete_log_button_hover_icon = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "delete_all_hover.png")).resize((36, 36)))


def on_enter_delete_weight_log_button(event):
    delete_weight_log_button.config(image=delete_log_button_hover_icon)


def on_leave_delete_weight_log_button(event):
    delete_weight_log_button.config(image=delete_log_button_icon)


delete_weight_log_button = tk.Button(middle_left_weight_frame_col3_frame_row1_row2, image=delete_log_button_icon,
                                     command=None, bg=bg_app_class_color_layer_1, width=40, height=40, relief="flat",
                                     borderwidth=0)
delete_weight_log_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_weight_log_button.bind("<Enter>", on_enter_delete_weight_log_button)
delete_weight_log_button.bind("<Leave>", on_leave_delete_weight_log_button)


def on_enter_delete_thickness_log_button(event):
    delete_thickness_log_button.config(image=delete_log_button_hover_icon)


def on_leave_delete_thickness_log_button(event):
    delete_thickness_log_button.config(image=delete_log_button_icon)


delete_thickness_log_button = tk.Button(middle_left_thickness_frame_col3_frame_row1_row2, image=delete_log_button_icon,
                                        command=None, bg=bg_app_class_color_layer_1, width=40, height=40, relief="flat",
                                        borderwidth=0)
delete_thickness_log_button.grid(row=0, column=0, padx=0, pady=0, sticky="e")
delete_thickness_log_button.bind("<Enter>", on_enter_delete_thickness_log_button)
delete_thickness_log_button.bind("<Leave>", on_leave_delete_thickness_log_button)


def on_enter_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon_hover)


def on_leave_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon)


exit_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "exit.png")).resize((134, 31)))
exit_icon_hover = ImageTk.PhotoImage(
    Image.open(os.path.join(base_path, "theme", "icons", "exit_hover.png")).resize((134, 31)))
bottom_exit_button = tk.Button(bottom_right_frame, image=exit_icon, command=exit, bg='#f4f4fe', width=134, height=31,
                               relief="flat", borderwidth=0)
bottom_exit_button.grid(row=0, column=0, columnspan=5, padx=5, pady=5, sticky="e")
bottom_exit_button.bind("<Enter>", on_enter_bottom_exit_button)
bottom_exit_button.bind("<Leave>", on_leave_bottom_exit_button)

"""Update loop"""
update_thread = threading.Thread(target=update_dimensions, daemon=True)
update_thread.start()
root.mainloop()

"""
=> pyinstaller --windowed --onefile --name temp03 --add-data "theme/icons;theme/icons" --add-data "theme/assets;theme/assets" temp03.py --icon=theme/icons/logo.ico
"""