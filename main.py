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
screen_width = min(1600, int(0.9*root.winfo_screenwidth()*scaling))
screen_height = min(900, int(0.9*root.winfo_screenheight()*scaling))
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

bg_app_color_layer1 = '#f4f4fe'
bg_app_color_layer2 = '#fefefe'

fg_app_color_layer1 = 'black'
gb_app_color_layer2 = '#333333'

showing_settings = False
showing_runcards = False
showing_advance_setting = False
showing_thickness_frame = False
showing_weight_frame = False

current_thickness_entry = ""
error_msg = tk.StringVar()
error_msg.set("")
error_fg_color = "red"

error_thread = None
error_event = threading.Event()

weight_record_id = 0
thickness_record_id = 0

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
                middle_left_col3_width = 210
                middle_left_col2_width = 10
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

        middle_left_weight_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_weight_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)

        middle_left_thickness_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_thickness_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)

        middle_left_weight_frame_left_2_canvas.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height - 80)
        middle_left_weight_frame_left_2_scrollbar.place(x=middle_left_col1_width - 20, y=0, width=20, height=middle_frame_height - 80)
        middle_left_weight_frame_left_2_scrollable_frame.place(x=0, y=0, width=middle_left_col1_width - 20, height=middle_frame_height - 80)







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

        middle_right_advance_setting_frame_row2_col1_row9_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row9_col2.place(x=170, y=0, width=162, height=180)

        middle_right_advance_setting_frame_row2_col1_row14_col1.place(x=0, y=0, width=210, height=100)
        middle_right_advance_setting_frame_row2_col1_row14_col2.place(x=210, y=0, width=122, height=100)

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

        middle_right_advance_setting_frame_row2_col1_row14_col1_row1.place(x=0, y=0, width=210, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col1_row2.place(x=0, y=30, width=210, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col1_row3.place(x=0, y=60, width=210, height=40)

        middle_right_advance_setting_frame_row2_col1_row14_col2_row1.place(x=0, y=6, width=122, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col2_row2.place(x=0, y=36, width=122, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col2_row3.place(x=0, y=66, width=122, height=30)


        middle_right_advance_setting_frame_row2_col3_row1.place(x=0, y=0, width=middle_right_frame_width - 10 - 332 - 10, height=30)

        middle_right_advance_setting_frame_row2_col3_row2.place(x=0, y=30, width=middle_right_frame_width - 10 - 332 - 10, height=10)
        middle_right_advance_setting_frame_row2_col3_row3.place(x=0, y=40, width=middle_right_frame_width - 10 - 332 - 10, height=200)
        middle_right_advance_setting_frame_row2_col3_row4.place(x=0, y=240, width=middle_right_frame_width - 10 - 332 - 10, height=30)

        middle_right_advance_setting_frame_row2_col3_row5.place(x=0, y=270, width=middle_right_frame_width - 10 - 332 - 10, height=30)
        middle_right_advance_setting_frame_row2_col3_row6.place(x=0, y=300, width=middle_right_frame_width - 10 - 332 - 10, height=10)
        middle_right_advance_setting_frame_row2_col3_row7.place(x=0, y=310, width=middle_right_frame_width - 10 - 332 - 10, height=200)

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

        error_display_entry.place(x=5, y=10, width=int(screen_width * 0.7), height=25)


        if hasattr(weight_frame_write_insert_value, "tree"):
            treeview_width = middle_left_col1_width
            treeview_height = middle_frame_height - 80
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



        root.update_idletasks()
        root.update()

def clear_error_message():
    error_msg.set("")
    error_display_entry.config(fg="red")

def show_error_message(msg, fg_color_code, time_show):
    global error_thread, error_event
    if error_thread and error_thread.is_alive():
        error_event.set()
        error_thread.join()
    error_event = threading.Event()
    def update_message():
        if fg_color_code == 1:
            fg_color = "green"
            fg_icon = "✔"
        elif fg_color_code == 0:
            fg_color = "red"
            fg_icon = "❌"
        elif fg_color_code == -1:
            fg_color = "#595959"
            fg_icon = "↻"
        else:
            fg_color = "#7F7F7F"
            fg_icon = "ⓘ"
        error_msg.set(f"{fg_icon} {msg}")
        root.after(0, lambda: error_display_entry.config(fg=fg_color))
        if not error_event.wait(time_show / 1000):
            root.after(0, clear_error_message)
    error_thread = threading.Thread(target=update_message, daemon=True)
    error_thread.start()

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

middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg='green')
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_color)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_param_color)



middle_left_weight_frame_col1_frame_row1 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_color)
middle_left_weight_frame_col1_frame_row2 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_color)



middle_left_thickness_frame_col1_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_color)
middle_left_thickness_frame_col1_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_color)






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
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_color)

middle_right_advance_setting_frame_row2_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row12 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col1_row13 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row15 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_param_color)



middle_right_advance_setting_frame_row2_col1_row5_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row9_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row9_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row9_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row9_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col2, bg=bg_param_color)


middle_right_advance_setting_frame_row2_col1_row14_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row14_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_param_color)

middle_right_advance_setting_frame_row2_col1_row14_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col1_row14_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_param_color)



middle_right_advance_setting_frame_row2_col3_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col3_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col3_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_app_color)
middle_right_advance_setting_frame_row2_col3_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_param_color)


middle_right_advance_setting_frame_row2_col3_row3_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row3_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_param_color)


middle_right_advance_setting_frame_row2_col3_row7_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)
middle_right_advance_setting_frame_row2_col3_row7_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_param_color)









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
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

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



error_display_entry = tk.Entry(bottom_left_frame, textvariable=error_msg, font=("Cambria", 12), bg=bg_app_color, fg=error_fg_color, bd=0, highlightthickness=0, readonlybackground=bg_app_color, state="readonly")






def check_runcard_correction(runcard):
    def insert_data():
        try:
            sql = "SELECT id FROM [PMGMES].[dbo].[PMG_MES_RunCard] WHERE id = ?"
            conn_str = (
                f'DRIVER={{SQL Server}};'
                f'SERVER={server_ip.get()};'
                f'DATABASE={db_name.get()};'
                f'UID={user_id.get()};'
                f'PWD={password.get()};'
            )
            # print(sql)
            # print(f"===>{server_ip.get()}")
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                cursor.execute(sql, (runcard,))
                result = cursor.fetchone()
                if result:
                    threading.Thread(target=show_error_message, args=(f"Runcard {runcard} có tồn tại", 1, 3000), daemon=True).start()
                    return True
                else:
                    threading.Thread(target=show_error_message, args=(f"Runcard {runcard} không tồn tại", 0, 3000), daemon=True).start()
                    return False
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
            return False
    if len(runcard) > 0:
        threading.Thread(target=insert_data, daemon=True).start()
    else:
        threading.Thread(target=show_error_message, args=("Chưa nhập Runcard", 0, 3000), daemon=True).start()

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

def weight_frame_mouser_pointer_in(event):
    if 'name_var' in event.widget.__dict__:
        print(f"Current Entry: {event.widget.name_var}")


def weight_insert_data_to_db(device_name, runcard_id, weight_value, operator_id):
    def insert_data():
        try:
            sql = f"""
                insert into [PMG_DEVICE].[dbo].[WeightDeviceData] (DeviceId, LotNo, Weight, UserId, CreationDate)
                values ('{device_name}', '{runcard_id}', {weight_value}, {operator_id}, '{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}')
            """
            conn_str = (
                f'DRIVER={{SQL Server}};'
                f'SERVER={server_ip.get()};'
                f'DATABASE={db_name.get()};'
                f'UID={user_id.get()};'
                f'PWD={password.get()};'
            )
            # print(sql)
            # print(f"===>{server_ip.get()}")
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

def weight_frame_write_insert_value(device_id, operator_id, runcard_id, weight_value):
    try:
        global weight_record_id
        weight_record_id += 1
        root.update_idletasks()
        root.update()
        frame_width = int(1000)
        if not hasattr(weight_frame_write_insert_value, "tree"):
            columns = ("ID", "Device ID", "Operator ID", "Runcard ID", "Weight", "Timestamp")
            style = ttk.Style()
            style.theme_use("classic")
            style.configure("Treeview.Heading", font=("Arial", 12, "bold"), relief="flat", background="white", foreground="black", borderwidth=1, highlightthickness=0)
            style.configure("Treeview", font=("Cambria", 13), borderwidth=0, relief="flat", background="white", fieldbackground="white")
            style.map("Treeview", background=[("selected", "lightblue")])
            weight_frame_write_insert_value.tree = ttk.Treeview(middle_left_weight_frame_left_2_scrollable_frame, columns=columns, show="headings", style="Treeview", height=int(middle_left_weight_frame_left_2_canvas.winfo_height() - 80))
            for index, col in enumerate(columns):
                if index == 0:
                    record_width = int(frame_width*0.1)
                elif index == 5:
                    record_width = int(frame_width*0.3)
                else:
                    record_width = int(frame_width*0.6/4)
                weight_frame_write_insert_value.tree.heading(col, text=col)
                weight_frame_write_insert_value.tree.column(col, width=record_width, anchor="center")
            weight_frame_write_insert_value.tree.pack(fill="both", expand=True)
        weight_frame_write_insert_value.tree.insert("", "0", values=(weight_record_id, device_id, operator_id, runcard_id, weight_value, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

def weight_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_weight_value_entry":
            is_connected = int(get_registry_value("is_connected", "0"))
            if is_connected == 1:
                if all([entry_weight_device_name_entry.get(), entry_weight_operator_id_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get()]):
                    if int(is_check_runcard_switch.get()) == 1:
                        if check_runcard_correction(entry_weight_runcard_id_entry.get()):
                            weight_insert_data_to_db(entry_weight_device_name_entry.get(),
                                                     entry_weight_runcard_id_entry.get(), current_widget.get(),
                                                     entry_weight_operator_id_entry.get())
                            weight_frame_write_insert_value(entry_weight_device_name_entry.get(),
                                                            entry_weight_operator_id_entry.get(),
                                                            entry_weight_runcard_id_entry.get(), current_widget.get())
                    else:
                        weight_insert_data_to_db(entry_weight_device_name_entry.get(),
                                                 entry_weight_runcard_id_entry.get(), current_widget.get(),
                                                 entry_weight_operator_id_entry.get())
                        weight_frame_write_insert_value(entry_weight_device_name_entry.get(),
                                                        entry_weight_operator_id_entry.get(),
                                                                        entry_weight_runcard_id_entry.get(), current_widget.get())
                    entry_weight_weight_value_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.focus_set()

                else:
                    threading.Thread(target=show_error_message, args=("Empty field detected!", 0, 3000), daemon=True).start()
            else:
                messagebox.showerror("Error", "Connect to database first!")
        else:
            if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_runcard_id_entry":
                if int(is_check_runcard_switch.get()) == 1:
                    if check_runcard_correction(entry_weight_runcard_id_entry.get()):
                        event.widget.tk_focusNext().focus()
                else:
                    event.widget.tk_focusNext().focus()
            else:
                event.widget.tk_focusNext().focus()
            return "break"
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


entry_weight_device_name_var = tk.StringVar()
entry_weight_device_name_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_device_name_var, 12, 1))
entry_weight_device_name_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Device ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_device_name_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_device_name_var)
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
entry_weight_runcard_id_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Runcard ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_runcard_id_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_runcard_id_var)
entry_weight_runcard_id_entry.name_var = "entry_weight_runcard_id_entry"
entry_weight_runcard_id_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_runcard_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(2, weight=1)


entry_weight_weight_value_var = tk.StringVar()
entry_weight_weight_value_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_weight_value_var, 10, 0))
entry_weight_weight_value_label = tk.Label(middle_left_weight_frame_col1_frame_row1, text="Trọng lượng", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_weight_value_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry = tk.Entry(middle_left_weight_frame_col1_frame_row1, font=(font_name, 18), bd=2, textvariable=entry_weight_weight_value_var)
entry_weight_weight_value_entry.name_var = "entry_weight_weight_value_entry"
entry_weight_weight_value_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_weight_value_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_col1_frame_row1.columnconfigure(3, weight=1)




middle_left_weight_frame_left_2_canvas = tk.Canvas(middle_left_weight_frame_col1_frame_row2, bg=bg_param_color, highlightthickness=0)
middle_left_weight_frame_left_2_scrollbar = tk.Scrollbar(middle_left_weight_frame_col1_frame_row2, orient="vertical", command=middle_left_weight_frame_left_2_canvas.yview)
middle_left_weight_frame_left_2_canvas.configure(yscrollcommand=middle_left_weight_frame_left_2_scrollbar.set)
middle_left_weight_frame_left_2_scrollable_frame = tk.Frame(middle_left_weight_frame_left_2_canvas, bg='white')
middle_left_weight_frame_left_2_canvas.create_window((0, 0), window=middle_left_weight_frame_left_2_scrollable_frame, anchor="nw")
middle_left_weight_frame_left_2_scrollable_frame.bind("<Configure>", lambda e: middle_left_weight_frame_left_2_canvas.configure(scrollregion=middle_left_weight_frame_left_2_canvas.bbox("all")))
middle_left_weight_frame_left_2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_left_2_scrollbar.pack(side="right", fill="y")
all_entries = []













def thickness_frame_mouser_pointer_in(event):
    global current_thickness_entry
    if 'name_var' in event.widget.__dict__:
        current_thickness_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")


def thickness_frame_hit_enter_button(event):
    return None

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
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def get_selected_frame():
    """Retrieve the last selected frame from the Windows Registry."""
    return get_registry_value("SelectedFrame", "Trọng lượng")

def set_selected_frame(value):
    """Save the selected frame to the Windows Registry."""
    set_registry_value("SelectedFrame", value)

def weight_frame_com_port_insert_data():
    global weight_record_log_id
    if "COM" in get_registry_value("COM1", ""):
        print(f"Weight COM: {selected_weight_com.get()}")
        ser = serial.Serial(selected_weight_com.get(), baudrate=9600, timeout=float(get_registry_value("is_weight_timeout", "0.3")))
        try:
            while True:
                if ser.is_open:
                    value = ser.readline().decode('utf-8').strip()
                    if len(value):
                        weight_record_log_id += 1
                        # threading.Thread(target=root.after, args=(
                        # 0, update_com_port_weight_log_display, f"{str(weight_record_log_id).zfill(4)}   {value}"),
                        #                  daemon=True).start()
                        if "g" not in value[-2:]:
                            threading.Thread(target=show_error_message,
                                             args=(f"Change your weight unit to gram!", 0, 3000), daemon=True).start()
                        if 'ST' in value:
                            weight_value = float(
                                re.sub(r'[a-zA-Z]', '', ((value.replace(" ", "")).split(':')[-1])[:-1]))
                            if weight_value >= 0:
                                entry_weight_runcard_id_entry.event_generate("<Return>")
                                entry_weight_weight_value_entry.insert(0, weight_value)
                                entry_weight_weight_value_entry.event_generate("<Return>")
                                entry_weight_weight_value_entry.delete(0, tk.END)
                                entry_weight_runcard_id_entry.delete(0, tk.END)
                                entry_weight_runcard_id_entry.focus_set()
                else:
                    ser.open()
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
            pass
        finally:
            if ser.is_open:
                ser.close()
    else:
        pass

com_data_labels = []
def update_com_port_weight_log_display(data):
    try:
        global com_data_labels
        def add_data():
            label = tk.Label(middle_left_weight_frame_right_frame_log_scrollable_frame, text=data, font=("Arial", 13), bg='white', anchor="w", justify="left")
            label.pack(fill="x", padx=5, pady=2)
            com_data_labels.append(label)
            if len(com_data_labels) > 20:
                com_data_labels[0].destroy()
                com_data_labels.pop(0)
            middle_left_weight_frame_right_frame_log_scrollable_frame.update_idletasks()
            middle_left_weight_frame_right_frame_log_canvas.configure(scrollregion=middle_left_weight_frame_right_frame_log_canvas.bbox("all"))
        root.after(0, add_data)
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

weight_com_thread = None
thickness_com_thread = None
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
                pass
                # thickness_com_thread = threading.Thread(target=thickness_frame_com_port_insert_data, daemon=True)
                # thickness_com_thread.start()
        set_selected_frame(selected_middle_left_frame.get())
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def save_setting_frame():
    try:
        threading.Thread(target=show_error_message, args=(f"Save setting", 1, 3000), daemon=True).start()
        set_registry_value("COM1", selected_weight_com.get())
        set_registry_value("COM2", selected_thickness_com.get())
        switch_middle_left_frame()
        set_selected_frame(selected_middle_left_frame.get())
        messagebox.showinfo("Success", "Save setting success!\nApplication will close automatically\nRe-open the app")
        root.destroy()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def save_advance_setting_frame():
    try:
        threading.Thread(target=show_error_message, args=(f"Save setting", 1, 3000), daemon=True).start()
        set_registry_value("COM1", selected_weight_com.get())
        set_registry_value("COM2", selected_thickness_com.get())
        set_registry_value("ServerIP", server_ip.get())
        switch_middle_left_frame()
        set_selected_frame(selected_middle_left_frame.get())
        set_registry_value("is_plant_name", plant_name.get())
        set_registry_value("is_show_runcard_switch", is_show_runcard_switch.get())
        set_registry_value("is_check_runcard_switch", is_check_runcard_switch.get())
        set_registry_value("is_input_runcard_switch", is_input_runcard_switch.get())
        set_registry_value("is_weight_keyboard", is_weight_keyboard.get())
        set_registry_value("is_weight_timeout", is_weight_timeout.get())
        set_registry_value("is_weight_delay", is_weight_delay.get())
        set_registry_value("is_thickness_keyboard", is_thickness_keyboard.get())
        set_registry_value("is_thickness_timeout", is_thickness_timeout.get())
        set_registry_value("is_thickness_delay", is_thickness_delay.get())
        messagebox.showinfo("Success", "Save setting success!\nApplication will close automatically\nRe-open the app")
        root.destroy()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def database_test_connection():
    try:
        def test_connection():
            conn_str = (
                f'DRIVER={{SQL Server}};'
                f'SERVER={server_ip.get()};'
                f'DATABASE={db_name.get()};'
                f'UID={user_id.get()};'
                f'PWD={password.get()};'
            )
            timeout_event = threading.Event()
            threading.Thread(target=show_error_message, args=(f"Connecting to database...", -1, 5000), daemon=True).start()
            def connect_to_db():
                try:
                    conn = pyodbc.connect(conn_str, timeout=5)
                    cursor = conn.cursor()
                    cursor.execute("SELECT 1")
                    conn.close()
                    if len(plant_name.get()):
                        # threading.Thread(target=runcard_machine_line_list, args=(plant_name.get(),), daemon=True).start()
                        pass
                    if len(str(get_registry_value("ServerIP", "0"))) > 0:
                        set_registry_value("is_connected", "1")
                    threading.Thread(target=show_error_message, args=("Connection successful!", 1, 2000), daemon=True).start()
                    timeout_event.set()
                except pyodbc.Error as e:
                    set_registry_value("is_connected", "0")
                    threading.Thread(target=show_error_message, args=(f"Connection failed: {str(e)}", 0, 2000), daemon=True).start()
                    timeout_event.set()
            connection_thread = threading.Thread(target=connect_to_db)
            connection_thread.start()
            connection_thread.join(timeout=5)
            if not timeout_event.is_set():
                set_registry_value("is_connected", "0")
                threading.Thread(target=show_error_message, args=("Connection timeout: Unable to connect within 5 seconds.", 0, 2000), daemon=True).start()
        threading.Thread(target=test_connection, daemon=True).start()
    except Exception as e:
        set_registry_value("is_connected", "0")
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass
selected_middle_left_frame = tk.StringVar(value=get_selected_frame())
switch_middle_left_frame()

selected_weight_com = tk.StringVar(value=get_registry_value("COM1", ""))
selected_thickness_com = tk.StringVar(value=get_registry_value("COM2", ""))

weight_label = tk.Label(middle_right_setting_frame_row2_row1, text="Trọng lượng:      ", font=(font_name, 14, "bold"), bg=bg_param_color)
weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
weight_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_weight_com, "")
weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

thickness_label = tk.Label(middle_right_setting_frame_row2_row1, text="Độ dày:", font=(font_name, 14, "bold"), bg=bg_param_color)
thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
thickness_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_thickness_com, "")
thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

frame_select_label = tk.Label(middle_right_setting_frame_row2_row1, text="Mặc định:", font=(font_name, 14, "bold"), bg=bg_param_color)
frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
frame_select_menu = CustomOptionMenu(middle_right_setting_frame_row2_row1, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")


advance_weight_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Trọng lượng:      ", font=(font_name, 14, "bold"), bg='white')
advance_weight_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
advance_weight_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_weight_com, "")
advance_weight_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

advance_thickness_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Độ dày:", font=(font_name, 14, "bold"), bg='white')
advance_thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
advance_thickness_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_thickness_com, "")
advance_thickness_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

advance_frame_select_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row2, text="Mặc định:", font=(font_name, 14, "bold"), bg='white')
advance_frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
advance_frame_select_menu = CustomOptionMenu(middle_right_advance_setting_frame_row2_col1_row2, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
advance_frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")

update_com_ports(weight_menu, thickness_menu)
update_com_ports(advance_weight_menu, advance_thickness_menu)



server_ip_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row1, text="Server IP:", font=(font_name, 14, "bold"), bg='white')
server_ip_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
server_ip = tk.StringVar(value=get_registry_value("ServerIP", ""))
server_ip_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row1, textvariable=server_ip, font=(font_name, 16), width=12, bg='white')
server_ip_entry.grid(row=0, column=0, padx=5, pady=5, sticky="w")


db_name_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row2, text="Database:", font=(font_name, 14, "bold"), bg='white')
db_name_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
db_name = tk.StringVar(value=get_registry_value("Database", "PMG_DEVICE"))
db_name_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row2, textvariable=db_name, font=(font_name, 16), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
db_name_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")


user_id_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row3, text="User ID:", font=(font_name, 14, "bold"), bg='white')
user_id_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
user_id = tk.StringVar(value=get_registry_value("UserID", "scadauser"))
user_id_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row3, textvariable=user_id, font=(font_name, 16), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
user_id_entry.grid(row=2, column=0, padx=5, pady=5, sticky="w")



password_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row5_col1_row4, text="Password:", font=(font_name, 14, "bold"), bg='white')
password_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
password = tk.StringVar(value=get_registry_value("Password", "pmgscada+123"))
password_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row5_col2_row4, textvariable=password, font=(font_name, 16), width=12, show='*', bg='white', fg='red', readonlybackground='white', state='readonly')
password_entry.grid(row=3, column=0, padx=5, pady=5, sticky="w")


plant_name_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row9_col1_row1, text="Plant name:", font=(font_name, 14, "bold"), bg='white')
plant_name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
plant_name = tk.StringVar(value=get_registry_value("is_plant_name", ""))
plant_name.trace_add("write", lambda *args: convert_to_uppercase(plant_name, 3, 1))
plant_name_entry = tk.Entry(middle_right_advance_setting_frame_row2_col1_row9_col2_row1, textvariable=plant_name, font=(font_name, 16), width=12, bg='white')
plant_name_entry.grid(row=0, column=0, padx=5, pady=5, sticky="w")


runcard_setting_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row12, text="Runcard", font=(font_name, 14, "bold"), bg=bg_app_color)
runcard_setting_label.grid(row=0, column=0, padx=0, pady=0)



is_show_runcard_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row14_col1_row1, text="Runcard show:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_show_runcard_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_show_runcard_switch = tk.StringVar(value=get_registry_value("is_show_runcard", "0"))
is_show_runcard_switch_button = ttk.Checkbutton(middle_right_advance_setting_frame_row2_col1_row14_col2_row1, variable=is_show_runcard_switch, onvalue='1', offvalue='0', style="Switch")
is_show_runcard_switch_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")


is_check_runcard_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row14_col1_row2, text="Runcard check:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_check_runcard_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_check_runcard_switch = tk.StringVar(value=get_registry_value("is_check_runcard_switch", "0"))
is_check_runcard_switch_button = ttk.Checkbutton(middle_right_advance_setting_frame_row2_col1_row14_col2_row2, variable=is_check_runcard_switch, onvalue='1', offvalue='0', style="Switch")
is_check_runcard_switch_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")


is_input_runcard_label = tk.Label(middle_right_advance_setting_frame_row2_col1_row14_col1_row3, text="Runcard input:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_input_runcard_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_input_runcard_switch = tk.StringVar(value=get_registry_value("is_input_runcard_switch", "0"))
is_input_runcard_switch_button = ttk.Checkbutton(middle_right_advance_setting_frame_row2_col1_row14_col2_row3, variable=is_input_runcard_switch, onvalue='1', offvalue='0', style="Switch")
is_input_runcard_switch_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")



weight_setting_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row1, text="Trọng lượng", font=(font_name, 14, "bold"), bg=bg_app_color)
weight_setting_label.grid(row=0, column=0, padx=0, pady=0)

is_weight_keyboard_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row3_11, text="Keyboard:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_weight_keyboard_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_weight_keyboard = tk.StringVar(value=get_registry_value("is_weight_keyboard", "0"))
is_weight_keyboard_button = ttk.Checkbutton(middle_right_advance_setting_frame_row2_col3_row3_12, variable=is_weight_keyboard, onvalue='1', offvalue='0', style="Switch")
is_weight_keyboard_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

is_weight_timeout_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row3_21, text="Timeout:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_weight_timeout_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_weight_timeout = tk.StringVar(value=get_registry_value("is_weight_timeout", "0.3"))
is_weight_timeout_entry = tk.Entry(middle_right_advance_setting_frame_row2_col3_row3_22, textvariable=is_weight_timeout, font=(font_name, 14), width=6, bg='white', fg='black')
is_weight_timeout_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")

is_weight_delay_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row3_31, text="Delay:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_weight_delay_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_weight_delay = tk.StringVar(value=get_registry_value("is_weight_delay", "0.0"))
is_weight_delay_entry = tk.Entry(middle_right_advance_setting_frame_row2_col3_row3_32, textvariable=is_weight_delay, font=(font_name, 14), width=6, bg='white', fg='black')
is_weight_delay_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")

thickness_setting_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row5, text="Độ dày", font=(font_name, 14, "bold"), bg=bg_app_color)
thickness_setting_label.grid(row=0, column=0, padx=0, pady=0)


is_thickness_keyboard_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row7_11, text="Keyboard:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_thickness_keyboard_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_thickness_keyboard = tk.StringVar(value=get_registry_value("is_thickness_keyboard", "0"))
is_thickness_keyboard_button = ttk.Checkbutton(middle_right_advance_setting_frame_row2_col3_row7_12, variable=is_thickness_keyboard, onvalue='1', offvalue='0', style="Switch")
is_thickness_keyboard_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

is_thickness_timeout_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row7_21, text="Timeout:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_thickness_timeout_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_thickness_timeout = tk.StringVar(value=get_registry_value("is_thickness_timeout", "0.3"))
is_thickness_timeout_entry = tk.Entry(middle_right_advance_setting_frame_row2_col3_row7_22, textvariable=is_thickness_timeout, font=(font_name, 14), width=6, bg='white', fg='black')
is_thickness_timeout_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")

is_thickness_delay_label = tk.Label(middle_right_advance_setting_frame_row2_col3_row7_31, text="Delay:", font=(font_name, 12, "bold"), bg=bg_param_color)
is_thickness_delay_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
is_thickness_delay = tk.StringVar(value=get_registry_value("is_thickness_delay", "0"))
is_thickness_delay_entry = tk.Entry(middle_right_advance_setting_frame_row2_col3_row7_32, textvariable=is_thickness_delay, font=(font_name, 14), width=6, bg='white', fg='black')
is_thickness_delay_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")








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
save_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_setting_frame_button.bind("<Enter>", on_enter_save_setting_frame_button)
save_setting_frame_button.bind("<Leave>", on_leave_save_setting_frame_button)

def on_enter_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon_hover)
def on_leave_save_advance_setting_frame_button(event):
    save_advance_setting_frame_button.config(image=save_icon)
save_advance_setting_frame_button = tk.Button(middle_right_advance_setting_frame_row3_col1, image=save_icon, command=save_advance_setting_frame, bg=bg_app_color, width=134, height=34, relief="flat", borderwidth=0)
save_advance_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="e")
save_advance_setting_frame_button.bind("<Enter>", on_enter_save_advance_setting_frame_button)
save_advance_setting_frame_button.bind("<Leave>", on_leave_save_advance_setting_frame_button)


def on_enter_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_hover_icon)
def on_leave_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_icon)
database_test_connection_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "connect.png")).resize((146, 38)))
database_test_connection_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "connect_hover.png")).resize((146, 38)))
database_test_connection_button = tk.Button(middle_right_advance_setting_frame_row2_col1_row6, image=database_test_connection_button_icon, command=database_test_connection, bg='#f0f2f6', width=146, height=38, relief="flat", borderwidth=0)
database_test_connection_button.grid(row=4, column=1, padx=5, pady=5, sticky="e")
database_test_connection_button.bind("<Enter>", on_enter_database_test_connection_button)
database_test_connection_button.bind("<Leave>", on_leave_database_test_connection_button)

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