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
            middle_left_col3_width = 0
            middle_left_col2_width = 0
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width
        elif showing_runcards:
            middle_right_frame_width = 520
            middle_center_frame_width = 10
            middle_left_frame_width = screen_width - middle_right_frame_width - middle_center_frame_width
            middle_left_col3_width = 0
            middle_left_col2_width = 0
            middle_left_col1_width = middle_left_frame_width - middle_left_col3_width - middle_left_col2_width
        elif showing_advance_setting:
            middle_right_frame_width = 520
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
        middle_right_setting_frame_row2.place(x=0, y=40, width=middle_right_frame_width, height=middle_frame_height - 40 - 50)
        middle_right_setting_frame_row3.place(x=0, y=middle_frame_height - 40, width=middle_right_frame_width, height=50)

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
middle_right_advance_setting_frame_row2 = tk.Frame(middle_right_advance_setting_frame, bg=bg_param_color)
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
def open_weight_frame():
    set_registry_value("is_current_entry", "weight")
    middle_left_thickness_frame.pack_forget()
    middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)
    middle_left_weight_frame.lift()
def open_thickness_frame():
    set_registry_value("is_current_entry", "thickness")
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