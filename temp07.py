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
draw_runcard = False
runcard_selected_list = ['', '', '', '', ''] #=> date, machine, line, period, workorder
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
        threading.Thread(target=show_error_message, args=(f"def convert_to_uppercase => {e}", 0, 3000), daemon=True).start()
        pass
def update_dimensions():
    global screen_width, screen_height, showing_settings, showing_runcards, draw_runcard
    while True:
        current_date = (datetime.datetime.now() - datetime.timedelta(hours=5) + datetime.timedelta(minutes=22)).date()




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


conn_str = (
    f'DRIVER={{SQL Server}};'
    f'SERVER={get_registry_value("is_server_ip", "10.13.102.22")};'
    f'DATABASE={get_registry_value("is_db_name", "PMG_DEVICE")};'
    f'UID={get_registry_value("is_user_id", "scadauser")};'
    f'PWD={get_registry_value("is_password", "pmgscada+123")};'
)





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

middle_left_weight_frame_col1_frame_row1 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col1_frame_row2 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_app_class_color_layer_1)

middle_left_weight_frame_col3_frame_row1 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row2 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_app_class_color_layer_1)

middle_left_weight_frame_col3_frame_row1_row1 = tk.Frame(middle_left_weight_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)
middle_left_weight_frame_col3_frame_row1_row2 = tk.Frame(middle_left_weight_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_2)
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col3_frame_row1 = tk.Frame(middle_left_thickness_frame_col3_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row2 = tk.Frame(middle_left_thickness_frame_col3_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col1_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col1_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_app_class_color_layer_1)

middle_left_thickness_frame_col3_frame_row1_row1 = tk.Frame(middle_left_thickness_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)
middle_left_thickness_frame_col3_frame_row1_row2 = tk.Frame(middle_left_thickness_frame_col3_frame_row1, bg=bg_app_class_color_layer_1)




middle_center_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_center_frame.place(x=0, y=0)

middle_right_frame = tk.Frame(middle_frame, bg=bg_app_class_color_layer_1)
middle_right_frame.place(x=0, y=0)

middle_right_runcard_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame.place(x=0, y=0)

middle_right_runcard_frame_row1 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row2 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row3 = tk.Frame(middle_right_runcard_frame, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row3.grid_columnconfigure(0, weight=1)

middle_right_runcard_frame_row2_col1 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_1)
middle_right_runcard_frame_row2_col3 = tk.Frame(middle_right_runcard_frame_row2, bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row1 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row2 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row3 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row5 = tk.Frame(middle_right_runcard_frame_row2_col2, bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row1_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row1, bg=bg_app_class_color_layer_2)

middle_right_runcard_frame_row2_col2_row4_row1 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4_row2 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)
middle_right_runcard_frame_row2_col2_row4_row3 = tk.Frame(middle_right_runcard_frame_row2_col2_row4, bg=bg_app_class_color_layer_2)


middle_right_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_setting_frame.place(x=0, y=0)

middle_right_setting_frame_row1 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row2 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_2)
middle_right_setting_frame_row3 = tk.Frame(middle_right_setting_frame, bg=bg_app_class_color_layer_1)

middle_right_setting_frame_row1_col1 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row1_col2 = tk.Frame(middle_right_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row2_row1 = tk.Frame(middle_right_setting_frame_row2, bg=bg_app_class_color_layer_2)

middle_right_setting_frame_row3_col1 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1)
middle_right_setting_frame_row3_col2 = tk.Frame(middle_right_setting_frame_row3, bg=bg_app_class_color_layer_1)




middle_right_advance_setting_frame = tk.Frame(middle_right_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame.place(x=0, y=0)

middle_right_advance_setting_frame_row1 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row3 = tk.Frame(middle_right_advance_setting_frame, bg=bg_app_class_color_layer_1)

middle_right_advance_setting_frame_row1_col1 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row1_col2 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_advance_setting_frame_row2_col1 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col2 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_app_class_color_layer_1 )

middle_right_advance_setting_frame_row3_col1 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_class_color_layer_1)
middle_right_advance_setting_frame_row3_col2 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_app_class_color_layer_1)

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

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row6_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row6_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row6, bg=bg_app_class_color_layer_2 )

middle_right_advance_setting_frame_row2_col1_row9_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_app_class_color_layer_2 )
middle_right_advance_setting_frame_row2_col1_row9_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_app_class_color_layer_2 )

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


bottom_frame = tk.Frame(root, bg=bg_app_class_color_layer_1 )
bottom_frame.place(relx=0, rely=1.0, height=50, anchor="sw")

bottom_left_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1 )
bottom_right_frame = tk.Frame(bottom_frame, bg=bg_app_class_color_layer_1 )
bottom_right_frame.grid_columnconfigure(0, weight=1)




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
        root.after(0, runcard_wo_button)
        print(f"=> {selected_date}")
    calendar.bind("<<DateEntrySelected>>", on_date_change)
    root.after(0, runcard_period_button)
    root.after(0, runcard_machine_button)
    root.after(0, runcard_line_button)
    root.after(0, runcard_wo_button)
    return calendar.get_date().strftime("%d-%m-%Y")

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






# runcard_selected_list


selected_period_icon = None
selected_period_button = None
selected_period_button_hover = None
middle_right_runcard_frame_row2_col3_frame_dict = {}
middle_right_runcard_frame_row2_col3_button_dict = {}
middle_right_runcard_frame_row2_col3_button_icon_dict = {}
middle_right_runcard_frame_row2_col3_button_icon_hover_dict = {}

def runcard_manage_period_button(value, btn):
    global selected_period_button, selected_period_icon, selected_period_button_hover
    if selected_period_button:
        selected_period_button.config(image=middle_right_runcard_frame_row2_col3_button_icon_dict[selected_period_button_hover])
    btn.config(image=middle_right_runcard_frame_row2_col3_button_icon_hover_dict[value])
    selected_period_button = btn
    selected_period_icon = middle_right_runcard_frame_row2_col3_button_icon_hover_dict[value]
    selected_period_button_hover = value
    runcard_wo_button()
def runcard_period_button():
    if int(get_registry_value("is_connected", "1")) == 1:
        for index, period in enumerate(period_times):
            period_icon_path = os.path.join(base_path, "theme", "icons", "period", f"time_{period}.png")
            period_hover_icon_path = os.path.join(base_path, "theme", "icons", "period", f"time_{period}_hover.png")
            period_icon = ImageTk.PhotoImage(Image.open(period_icon_path).resize((36, 26)))
            period_hover_icon = ImageTk.PhotoImage(Image.open(period_hover_icon_path).resize((36, 26)))

            middle_right_runcard_frame_row2_col3_button_icon_dict[period] = period_icon
            middle_right_runcard_frame_row2_col3_button_icon_hover_dict[period] = period_hover_icon
            period_frame_name = f"middle_right_runcard_frame_row2_col3_row{index + 1}"
            middle_right_runcard_frame_row2_col3_frame_dict[period_frame_name] = tk.Frame(middle_right_runcard_frame_row2_col3, bg=bg_app_class_color_layer_1)
            middle_right_runcard_frame_row2_col3_frame_dict[period_frame_name].place(x=0, y=index * 29, width=36, height=26)

            button = tk.Button(middle_right_runcard_frame_row2_col3_frame_dict[period_frame_name], image=period_icon, width=36, height=26, relief="flat")
            button.config(command=lambda p=period, b=button: runcard_manage_period_button(p,b))

            button.place(x=0, y=0, relwidth=1.0, relheight=1.0)
            middle_right_runcard_frame_row2_col3_button_dict[period] = button
        if current_time in middle_right_runcard_frame_row2_col3_button_dict:
            runcard_manage_period_button(current_time, middle_right_runcard_frame_row2_col3_button_dict[current_time])




selected_machine_icon = None
selected_machine_button = None
selected_machine_button_hover = None
middle_right_machine_frame_row2_col1_frame_dict = {}
middle_right_machine_frame_row2_col1_button_dict = {}
middle_right_machine_frame_row2_col1_button_icon_dict = {}
middle_right_machine_frame_row2_col1_button_icon_hover_dict = {}
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
        selected_machine_button.config(image=middle_right_machine_frame_row2_col1_button_icon_dict[selected_machine_button_hover])
    btn.config(image=middle_right_machine_frame_row2_col1_button_icon_hover_dict[value])
    selected_machine_button = btn
    selected_machine_icon = middle_right_machine_frame_row2_col1_button_icon_hover_dict[value]
    selected_machine_button_hover = value
    runcard_wo_button()
def runcard_machine_button():
    if int(get_registry_value("is_connected", "1")) == 1:
        machine_list = runcard_get_machine_list(get_registry_value("is_plant_name", "PVC1"))
        for index, machine in enumerate(machine_list):
            machine_icon_path = os.path.join(base_path, "theme", "icons", "machine", f"machine_{index+1}.png")
            machine_hover_icon_path = os.path.join(base_path, "theme", "icons", "machine", f"machine_{index+1}_hover.png")
            machine_icon = ImageTk.PhotoImage(Image.open(machine_icon_path).resize((36, 26)))
            machine_hover_icon = ImageTk.PhotoImage(Image.open(machine_hover_icon_path).resize((36, 26)))

            middle_right_machine_frame_row2_col1_button_icon_dict[machine] = machine_icon
            middle_right_machine_frame_row2_col1_button_icon_hover_dict[machine] = machine_hover_icon
            machine_frame_name = f"middle_right_machine_frame_row2_col1_row{index + 1}"
            middle_right_machine_frame_row2_col1_frame_dict[machine_frame_name] = tk.Frame(middle_right_runcard_frame_row2_col1, bg=bg_app_class_color_layer_1)
            middle_right_machine_frame_row2_col1_frame_dict[machine_frame_name].place(x=0, y=index * 29, width=36, height=26)

            button = tk.Button(middle_right_machine_frame_row2_col1_frame_dict[machine_frame_name], image=machine_icon, width=36, height=26, relief="flat")
            button.config(command=lambda p=machine, b=button: runcard_manage_machine_button(p,b))

            button.place(x=0, y=0, relwidth=1.0, relheight=1.0)
            middle_right_machine_frame_row2_col1_button_dict[machine] = button
        if machine_list:
            runcard_manage_machine_button(machine_list[0], middle_right_machine_frame_row2_col1_button_dict[machine_list[0]])





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
        selected_line_button.config(image=middle_right_line_frame_row2_col1_row3_button_icon_dict[selected_line_button_hover])
    btn.config(image=middle_right_line_frame_row2_col1_row3_button_icon_hover_dict[value])
    selected_line_button = btn
    selected_line_icon = middle_right_line_frame_row2_col1_row3_button_icon_dict[value]
    selected_line_button_hover = value
    runcard_wo_button()

def runcard_line_button():
    global selected_machine_button_hover
    line_list = runcard_get_line_list(selected_machine_button_hover)
    middle_width = int(len(line_list) * 48)
    side_width = int((432 - middle_width) / 2)
    middle_right_runcard_frame_row2_col2_row3_left_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3, bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_left_frame.place(x=0, y=0, width=side_width, height=40)
    middle_right_runcard_frame_row2_col2_row3_middle_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3, bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_middle_frame.place(x=side_width, y=0, width=middle_width, height=40)
    middle_right_runcard_frame_row2_col2_row3_right_frame = tk.Frame(middle_right_runcard_frame_row2_col2_row3, bg=bg_app_class_color_layer_2)
    middle_right_runcard_frame_row2_col2_row3_right_frame.place(x=side_width + middle_width, y=0, width=side_width, height=40)
    if int(get_registry_value("is_connected", "1")) == 1:
        for index, line in enumerate(line_list):
            line_icon_path = os.path.join(base_path, "theme", "icons", "line", f"line_{line}.png")
            line_hover_icon_path = os.path.join(base_path, "theme", "icons", "line", f"line_{line}_hover.png")
            line_icon = ImageTk.PhotoImage(Image.open(line_icon_path).resize((36, 26)))
            line_hover_icon = ImageTk.PhotoImage(Image.open(line_hover_icon_path).resize((36, 26)))

            middle_right_line_frame_row2_col1_row3_button_icon_dict[line] = line_icon
            middle_right_line_frame_row2_col1_row3_button_icon_hover_dict[line] = line_hover_icon
            line_frame_name = f"middle_right_runcard_frame_row2_col2_row3_col{index + 1}"
            middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name] = tk.Frame(middle_right_runcard_frame_row2_col2_row3_middle_frame, bg=bg_app_class_color_layer_1)
            middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name].place(x=index * 48, y=0, width=40, height=29)

            button = tk.Button(middle_right_line_frame_row2_col1_row3_frame_dict[line_frame_name], image=line_icon, width=36, height=26, relief="flat")
            button.config(command=lambda p=line, b=button: runcard_manage_line_button(p,b))

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
            cursor.execute(sql, (machine, line, int(time), str(date), str(date), ))
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














def open_setting():
    middle_right_setting_frame_row1_col1_label = tk.Label(middle_right_setting_frame_row1_col1, text="Cài đặt", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
    middle_right_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")




def open_advance_setting():
    middle_right_advance_setting_frame_row1_col1_label = tk.Label(middle_right_advance_setting_frame_row1_col1, text="Cài đặt nâng cao", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
    middle_right_advance_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")




def open_runcard():
    middle_right_runcard_frame_label = tk.Label(middle_right_runcard_frame_row1, text="Runcard", bg=bg_app_class_color_layer_1, bd=0, font=(font_name, 18, "bold"))
    middle_right_runcard_frame_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
