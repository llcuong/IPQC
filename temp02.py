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


base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")

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

root.iconbitmap(os.path.join(base_path, "theme", "icons", "logo.ico"))
root.minsize(screen_width, screen_height)
root.geometry(f"{screen_width}x{screen_height}")

style = ttk.Style(root)
root.tk.call('source', "theme/azure.tcl")
style.theme_use('azure')
style.configure("Togglebutton", foreground='red')

"""Define app parameters"""
font_name = 'Arial'
font_size_base_on_ratio = int(screen_height * 0.015)
button_width_base_on_ratio = int(screen_width * 0.012)

theme_mode = 1
if theme_mode == 1:
    bg_color_layer1 = '#f4f4fe'
    bg_color_layer2 = '#ffffff'
    fg_color_layer1 = '#000000'
    fg_color_layer2 = '#e0e0e0'


showing_settings = False
showing_runcards = False
showing_advance_setting = False
showing_thickness_frame = False
showing_weight_frame = False

error_msg = tk.StringVar()
error_msg.set("")
error_fg_color = "red"
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
top_frame = tk.Frame(root, bg=bg_color_layer1)
top_frame.place(relx=0, rely=0, height=50, anchor="nw")

top_left_frame = tk.Frame(top_frame, bg=bg_color_layer1)
top_right_frame = tk.Frame(top_frame, height=50, bg=bg_color_layer1)
top_right_frame.grid_columnconfigure(0, weight=1)





middle_frame = tk.Frame(root, bg=bg_color_layer1)
middle_frame.place(relx=0, rely=0, anchor="nw")



middle_left_frame = tk.Frame(middle_frame, bg=bg_color_layer1)
middle_left_frame.place(x=0, y=0)

middle_left_weight_frame = tk.Frame(middle_left_frame, bg=bg_color_layer1)
middle_left_weight_frame_col1_frame = tk.Frame(middle_left_weight_frame, bg=bg_color_layer2)
middle_left_weight_frame_col2_frame = tk.Frame(middle_left_weight_frame, bg=bg_color_layer1)
middle_left_weight_frame_col3_frame = tk.Frame(middle_left_weight_frame, bg=bg_color_layer2)

middle_left_weight_frame_col1_frame_row1 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_color_layer1)
middle_left_weight_frame_col1_frame_row2 = tk.Frame(middle_left_weight_frame_col1_frame, bg=bg_color_layer1)

middle_left_weight_frame_col3_frame_row1 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_color_layer1)
middle_left_weight_frame_col3_frame_row2 = tk.Frame(middle_left_weight_frame_col3_frame, bg=bg_color_layer1)



middle_left_thickness_frame = tk.Frame(middle_left_frame, bg=bg_color_layer1)
middle_left_thickness_frame_col1_frame = tk.Frame(middle_left_thickness_frame, bg=bg_color_layer2)
middle_left_thickness_frame_col2_frame = tk.Frame(middle_left_thickness_frame, bg=bg_color_layer1)
middle_left_thickness_frame_col3_frame = tk.Frame(middle_left_thickness_frame, bg=bg_color_layer2)

middle_left_thickness_frame_col1_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_color_layer1)
middle_left_thickness_frame_col1_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_color_layer1)

middle_left_thickness_frame_col3_frame_row1 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_color_layer1)
middle_left_thickness_frame_col3_frame_row2 = tk.Frame(middle_left_thickness_frame_col1_frame, bg=bg_color_layer1)

middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)






middle_center_frame = tk.Frame(middle_frame, bg=bg_color_layer1)
middle_center_frame.place(x=0, y=0)




middle_right_frame = tk.Frame(middle_frame, bg=bg_color_layer1)
middle_right_frame.place(x=0, y=0)


middle_right_runcard_frame = tk.Frame(middle_right_frame, bg=bg_color_layer1)
middle_right_runcard_frame.place(x=0, y=0)

middle_right_setting_frame = tk.Frame(middle_right_frame, bg=bg_color_layer1)
middle_right_setting_frame.place(x=0, y=0)

middle_right_setting_frame_row1 = tk.Frame(middle_right_setting_frame, bg=bg_color_layer1)
middle_right_setting_frame_row2 = tk.Frame(middle_right_setting_frame, bg=bg_color_layer2)
middle_right_setting_frame_row3 = tk.Frame(middle_right_setting_frame, bg=bg_color_layer1)

middle_right_setting_frame_row1_col1 = tk.Frame(middle_right_setting_frame_row1, bg=bg_color_layer1)
middle_right_setting_frame_row1_col2 = tk.Frame(middle_right_setting_frame_row1, bg=bg_color_layer1)
middle_right_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_setting_frame_row3_col1 = tk.Frame(middle_right_setting_frame_row3, bg=bg_color_layer1)
middle_right_setting_frame_row3_col2 = tk.Frame(middle_right_setting_frame_row3, bg=bg_color_layer1)

middle_right_setting_frame_row2_row1 = tk.Frame(middle_right_setting_frame_row2, bg=bg_color_layer2)





middle_right_advance_setting_frame = tk.Frame(middle_right_frame, bg=bg_color_layer1)
middle_right_advance_setting_frame.place(x=0, y=0)

middle_right_advance_setting_frame_row1 = tk.Frame(middle_right_advance_setting_frame, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_color_layer1)
middle_right_advance_setting_frame_row3 = tk.Frame(middle_right_advance_setting_frame, bg=bg_color_layer1)

middle_right_advance_setting_frame_row1_col1 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row1_col2 = tk.Frame(middle_right_advance_setting_frame_row1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row1_col2.grid_columnconfigure(0, weight=1)

middle_right_advance_setting_frame_row2_col1 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col2 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3 = tk.Frame(middle_right_advance_setting_frame_row2, bg=bg_color_layer1)

middle_right_advance_setting_frame_row2_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col1_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col1_row12 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col1_row13 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row15 = tk.Frame(middle_right_advance_setting_frame_row2_col1, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row5_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row5_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col1_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col1, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row5_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row5_col2_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row5_col2, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row9_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row9_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row9_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row9_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row9_col2, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row14_col1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14_col2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row14_col1_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14_col1_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14_col1_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col1, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col1_row14_col2_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14_col2_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col1_row14_col2_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col1_row14_col2, bg=bg_color_layer2)

middle_right_advance_setting_frame_row3_col1 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row3_col2 = tk.Frame(middle_right_advance_setting_frame_row3, bg=bg_color_layer1)

middle_right_advance_setting_frame_row2_col3_row1 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3_row2 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row4 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3_row5 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3_row6 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row8 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3_row9 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer1)
middle_right_advance_setting_frame_row2_col3_row10 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11 = tk.Frame(middle_right_advance_setting_frame_row2_col3, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col3_row3_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row3_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row3, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col3_row7_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row7_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row7, bg=bg_color_layer2)

middle_right_advance_setting_frame_row2_col3_row11_11 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11_12 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11_21 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11_22 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11_31 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)
middle_right_advance_setting_frame_row2_col3_row11_32 = tk.Frame(middle_right_advance_setting_frame_row2_col3_row11, bg=bg_color_layer2)







bottom_frame = tk.Frame(root, bg=bg_color_layer1)
bottom_frame.place(relx=0, rely=1.0, height=50, anchor="sw")

bottom_left_frame = tk.Frame(bottom_frame, bg=bg_color_layer1)
bottom_right_frame = tk.Frame(bottom_frame, bg=bg_color_layer1)
bottom_right_frame.grid_columnconfigure(0, weight=1)




"""Entry"""
middle_right_setting_frame_row1_col1_label = tk.Label(middle_right_setting_frame_row1_col1, text="Cài đặt", font=(font_name, 18, "bold"), bg=bg_color_layer1)
middle_right_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

middle_right_advance_setting_frame_row1_col1_label = tk.Label(middle_right_advance_setting_frame_row1_col1, text="Cài đặt bổ sung", font=(font_name, 18, "bold"), bg=bg_color_layer1)
middle_right_advance_setting_frame_row1_col1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

error_display_entry = tk.Entry(bottom_left_frame, textvariable=error_msg, font=("Cambria", 12), bg=bg_color_layer1, fg=error_fg_color, bd=0, highlightthickness=0, readonlybackground=bg_color_layer1, state="readonly")












def on_enter_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon_hover)
def on_leave_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon)
exit_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "exit.png")).resize((134, 31)))
exit_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "theme", "icons", "exit_hover.png")).resize((134, 31)))
bottom_exit_button = tk.Button(bottom_right_frame, image=exit_icon, command=None, bg='#f4f4fe', width=134, height=31, relief="flat", borderwidth=0)
bottom_exit_button.grid(row=0, column=0, columnspan=5, padx=5, pady=5, sticky="e")
bottom_exit_button.bind("<Enter>", on_enter_bottom_exit_button)
bottom_exit_button.bind("<Leave>", on_leave_bottom_exit_button)



def update_dimensions():
    global screen_width, screen_height, showing_settings, showing_runcards
    while True:
        root.update_idletasks()
        screen_width = root.winfo_width()
        screen_height = root.winfo_height()

        top_frame_height = 50
        middle_frame_height = screen_height - 50 - 50
        bottom_frame_height = 50

        top_frame.place(x=0, y=0, width=screen_width, height=top_frame_height)
        middle_frame.place(x=0, y=top_frame_height, width=screen_width, height=middle_frame_height)
        bottom_frame.place(x=0, y=screen_height - 50, width=screen_width, height=bottom_frame_height)

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


        top_left_frame.place(x=0, y=0, width=int(screen_width * 0.3), height=50)
        top_right_frame.place(x=int(screen_width * 0.3), y=0, width=int(screen_width * 0.7), height=50)

        middle_right_frame.place(x=middle_left_frame_width + middle_center_frame_width, y=0, width=middle_right_frame_width, height=middle_frame_height)


        middle_center_frame.place(x=middle_left_frame_width, y=0, width=middle_center_frame_width, height=middle_frame_height)

        middle_left_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)
        middle_left_weight_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)
        middle_left_thickness_frame.place(x=0, y=0, width=middle_left_frame_width, height=middle_frame_height)

        middle_left_weight_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_weight_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_weight_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_left_weight_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_weight_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)


        middle_left_thickness_frame_col1_frame.place(x=0, y=0, width=middle_left_col1_width, height=middle_frame_height)
        middle_left_thickness_frame_col2_frame.place(x=middle_left_col1_width, y=0, width=middle_left_col2_width, height=middle_frame_height)
        middle_left_thickness_frame_col3_frame.place(x=middle_left_col1_width + middle_left_col2_width, y=0, width=middle_left_col3_width, height=middle_frame_height)

        middle_left_thickness_frame_col1_frame_row1.place(x=0, y=0, width=middle_left_col1_width, height=80)
        middle_left_thickness_frame_col1_frame_row2.place(x=0, y=80, width=middle_left_col1_width, height=middle_frame_height - 80)



        middle_right_runcard_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)



        middle_right_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=180)
        middle_right_setting_frame_row3.place(x=0, y=40 + 200, width=middle_right_frame_width, height=50)

        middle_right_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width/2), height=40)
        middle_right_setting_frame_row1_col2.place(x=int(middle_right_frame_width/2), y=0, width=int(middle_right_frame_width/2), height=40)

        middle_right_setting_frame_row2_row1.place(x=0, y=20, width=int(middle_right_frame_width), height=160)

        middle_right_setting_frame_row3_col1.place(x=int((middle_right_frame_width/2-154)/2), y=0, width=154, height=40)
        middle_right_setting_frame_row3_col2.place(x=int((middle_right_frame_width/2-154)/2+middle_right_frame_width/2), y=0, width=154, height=40)



        middle_right_advance_setting_frame.place(x=0, y=0, width=middle_right_frame_width-5, height=middle_frame_height)
        middle_right_advance_setting_frame_row1.place(x=0, y=0, width=middle_right_frame_width, height=40)
        middle_right_advance_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=middle_frame_height - 40 - 50)
        middle_right_advance_setting_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width, height=50)

        middle_right_advance_setting_frame_row1_col1.place(x=0, y=0, width=int(middle_right_frame_width / 2), height=40)
        middle_right_advance_setting_frame_row1_col2.place(x=int(middle_right_frame_width / 2), y=0, width=int(middle_right_frame_width / 2), height=40)

        middle_right_advance_setting_frame_row3_col1.place(x=int((middle_right_frame_width/2-154)/2), y=0, width=154, height=40)
        middle_right_advance_setting_frame_row3_col2.place(x=int((middle_right_frame_width/2-154)/2+middle_right_frame_width/2), y=0, width=154, height=40)

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
        middle_right_advance_setting_frame_row2_col1_row5_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row2.place(x=0, y=40, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row3.place(x=0, y=80, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col1_row4.place(x=0, y=120, width=170, height=40)

        middle_right_advance_setting_frame_row2_col1_row5_col2.place(x=170, y=0, width=162, height=180)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row1.place(x=0, y=0, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row2.place(x=0, y=40, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row3.place(x=0, y=80, width=162, height=40)
        middle_right_advance_setting_frame_row2_col1_row5_col2_row4.place(x=0, y=120, width=162, height=40)

        middle_right_advance_setting_frame_row2_col1_row9_col1.place(x=0, y=0, width=170, height=180)
        middle_right_advance_setting_frame_row2_col1_row9_col1_row1.place(x=0, y=0, width=170, height=40)
        middle_right_advance_setting_frame_row2_col1_row9_col2_row1.place(x=0, y=0, width=162, height=40)

        middle_right_advance_setting_frame_row2_col1_row9_col2.place(x=170, y=0, width=162, height=180)

        middle_right_advance_setting_frame_row2_col1_row14_col1.place(x=0, y=0, width=210, height=100)
        middle_right_advance_setting_frame_row2_col1_row14_col1_row1.place(x=0, y=0, width=210, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col1_row2.place(x=0, y=30, width=210, height=30)
        middle_right_advance_setting_frame_row2_col1_row14_col1_row3.place(x=0, y=60, width=210, height=40)

        middle_right_advance_setting_frame_row2_col1_row14_col2.place(x=210, y=0, width=122, height=100)
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

        middle_right_advance_setting_frame_row2_col3_row11_11.place(x=0, y=0, width=140, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_12.place(x=140, y=0, width=50, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_21.place(x=0, y=40, width=140, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_22.place(x=140, y=40, width=50, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_31.place(x=0, y=80, width=140, height=40)
        middle_right_advance_setting_frame_row2_col3_row11_32.place(x=140, y=80, width=50, height=40)

        middle_left_thickness_frame_col3_frame_row1.place(x=0, y=0, width=middle_left_col3_width, height=80)
        middle_left_thickness_frame_col3_frame_row2.place(x=0, y=80, width=middle_left_col3_width, height=80)

        bottom_left_frame.place(x=0, y=0, width=int(screen_width * 0.7), height=50)
        bottom_right_frame.place(x=int(screen_width * 0.7), y=0, width=int(screen_width * 0.3), height=50)

        root.update_idletasks()
        root.update()

update_thread = threading.Thread(target=update_dimensions, daemon=True)
update_thread.start()
root.mainloop()