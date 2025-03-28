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
import os
import sys


if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")


"""Set DPI awareness for better scaling on high-DPI screens"""
ctypes.windll.shcore.SetProcessDpiAwareness(10)
REG_PATH = r"SOFTWARE\IPQC\Settings"

"""Main window setup"""
root = tk.Tk()
root.title("IPQC")
root.iconbitmap(os.path.join(base_path, "icon", "logo.ico"))

root.minsize(1280, 700)
root.geometry("1280x800")

style = ttk.Style(root)
root.tk.call('source', os.path.join(base_path, "assets", "azure.tcl"))
style.theme_use('azure')
style.configure("Togglebutton", foreground='white')


"""Get user scaling resolution"""
user32 = ctypes.windll.user32
user32.SetProcessDPIAware()
dpi = user32.GetDpiForSystem()

"""Define scaling factors"""
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

canvas = tk.Canvas(root, bg="#f4f4fe")
canvas.pack(fill=tk.BOTH, expand=True)
font_name = 'Arial'
font_size_ratio = int(screen_height * 0.015)
button_width_ratio = int(screen_width * 0.012)
showing_settings = False
showing_runcards = False
weight_record_id = 0
thickness_record_id = 0
weight_record_header_show = 0
weight_tree = None
weight_record_log_id = 0
error_msg = tk.StringVar()
error_msg.set("")
error_thread = None
error_event = threading.Event()
error_fg_color = "red"
current_thickness_entry = ""

current_date = datetime.datetime.now() + datetime.timedelta(minutes=39)
current_time = str(int(current_date.strftime('%H')))
selected_time_button = None
selected_time_icon = None
selected_time_button_period = None

selected_machine_button = None
selected_machine_icon = None
selected_machine_button_machine = None

selected_line_button = None
selected_line_icon = None
selected_line_button_line = None

current_com_thread = None

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
    """Dynamically update button size and label font based on window width"""
    global screen_width, screen_height, showing_settings, weight_record_header_show
    while True:

        screen_width = root.winfo_width()
        screen_height = root.winfo_height()

        font_size_ratio = min(int(screen_height * 0.013), 12)
        left_margin = min(int(screen_height * 0.012), 18)

        middle_frame_height = screen_height - top_frame.winfo_height() - bottom_frame.winfo_height() - left_margin - 2

        top_frame.place(relx=0, rely=0, anchor="nw", x=0, y=0)
        top_frame.config(width=screen_width)

        top_left_frame.config(width=int(screen_width * 0.3), height=50)
        top_right_frame.config(width=int(screen_width * 0.7), height=50)
        top_right_frame.place(x=screen_width * 0.3, y=0, width=screen_width * 0.7, height=50)

        middle_frame.place(relx=0, rely=0, anchor="nw", x=0, y=51)
        middle_frame.config(width=screen_width, height=middle_frame_height)

        if showing_settings:
            setting_width = 332
            border_width = 10
            new_setting_width = 0    #138
            new_border_width = 0    #10
        elif showing_runcards:
            setting_width = 520
            border_width = 10
            new_setting_width = 0
            new_border_width = 0
        else:
            setting_width = 0
            border_width = 0
            new_setting_width = 332
            new_border_width = 10

        middle_center_frame.config(width=border_width, height=middle_frame_height)
        middle_right_frame.config(width=int(setting_width), height=middle_frame_height)
        middle_left_frame.config(width=int(screen_width - setting_width - border_width), height=middle_frame_height)

        middle_left_frame.place(x=0, y=0, width=int(screen_width - setting_width - border_width), height=middle_frame_height)
        middle_center_frame.place(x=int(screen_width - setting_width - border_width), y=0, width=border_width, height=middle_frame_height)
        middle_right_frame.place(x=int(screen_width - setting_width), y=0, width=setting_width, height=middle_frame_height)





        middle_left_weight_frame_center_frame.config(width=border_width, height=middle_frame_height)
        middle_left_weight_frame_right_frame.config(width=int(setting_width), height=middle_frame_height)
        middle_left_weight_frame_left_frame.config(width=int(screen_width - setting_width - border_width), height=middle_frame_height)

        middle_left_weight_frame_left_frame.place(x=0, y=0, width=int(screen_width - new_setting_width - new_border_width), height=middle_frame_height)
        middle_left_weight_frame_center_frame.place(x=int(screen_width - new_setting_width - new_border_width), y=0, width=new_border_width, height=middle_frame_height)
        middle_left_weight_frame_right_frame.place(x=int(screen_width - new_setting_width), y=0, width=new_setting_width, height=middle_frame_height)







        middle_left_thickness_frame_center_frame.config(width=border_width, height=middle_frame_height)
        middle_left_thickness_frame_right_frame.config(width=int(setting_width), height=middle_frame_height)
        middle_left_thickness_frame_left_frame.config(width=int(screen_width - setting_width - border_width), height=middle_frame_height)

        middle_left_thickness_frame_left_frame.place(x=0, y=0, width=int(screen_width - new_setting_width - new_border_width), height=middle_frame_height)
        middle_left_thickness_frame_center_frame.place(x=int(screen_width - new_setting_width - new_border_width), y=0, width=new_border_width, height=middle_frame_height)
        middle_left_thickness_frame_right_frame.place(x=int(screen_width - new_setting_width), y=0, width=new_setting_width, height=middle_frame_height)






        middle_left_weight_frame_left_1_frame.config(width=border_width, height=middle_frame_height)
        middle_left_weight_frame_left_1_frame.place(x=0, y=0, width=int(screen_width - new_setting_width - new_border_width - setting_width - border_width), height=80)

        middle_left_weight_frame_left_2_frame.config(width=border_width, height=middle_frame_height)
        middle_left_weight_frame_left_2_frame.place(x=5, y=85, width=int(screen_width - new_setting_width - new_border_width - setting_width - border_width - 10), height=int(middle_frame_height-85))




        middle_left_thickness_frame_left_1_frame.config(width=border_width, height=middle_frame_height)
        middle_left_thickness_frame_left_1_frame.place(x=0, y=0, width=int(screen_width - new_setting_width - new_border_width - setting_width - border_width), height=80)

        middle_left_thickness_frame_left_2_frame.config(width=border_width, height=middle_frame_height)
        middle_left_thickness_frame_left_2_frame.place(x=5, y=85, width=int(screen_width - new_setting_width - new_border_width - setting_width - border_width - 10), height=int(middle_frame_height-85))






        middle_left_weight_frame_left_2_canvas.update_idletasks()
        middle_left_weight_frame_left_2_canvas.configure(scrollregion=middle_left_weight_frame_left_2_canvas.bbox("all"))


        if hasattr(weight_frame_write_insert_value, "tree"):
            treeview_width = middle_left_weight_frame_left_2_canvas.winfo_width()
            treeview_height = middle_left_weight_frame_left_2_canvas.winfo_height()
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
            treeview_width = middle_left_thickness_frame_left_2_canvas.winfo_width()
            treeview_height = middle_left_thickness_frame_left_2_canvas.winfo_height()
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




        bottom_frame.place(relx=0, rely=1.0, x=0, y=-left_margin, height=36, width=screen_width, anchor="sw")
        bottom_frame_left_frame.place(x=0, y=0, height=36, width=screen_width - 160)
        bottom_frame_right_frame.place(x=screen_width - 160, y=0, height=36, width=160)

        error_display_entry.place(x=5, y=5, width=screen_width - 210, height=25)
        bottom_exit_button.config(font=(font_name, font_size_ratio))
        bottom_exit_button.place(x=0, y=0, height=36, width=156)


        top_open_thickness_frame_button.config(width=156, height=36, font=(font_name, font_size_ratio))
        top_open_weight_frame_button.config(width=156, height=36, font=(font_name, font_size_ratio))

        root.update_idletasks()
        root.update()



def open_weight_frame():
    set_registry_value("is_current_entry", "weight")
    middle_left_thickness_frame.pack_forget()
    middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)


def open_thickness_frame():
    set_registry_value("is_current_entry", "thickness")
    middle_left_weight_frame.pack_forget()
    middle_left_thickness_frame.pack(fill=tk.BOTH, expand=True)


def open_setting_frame():
    global showing_settings, showing_runcards
    showing_settings = True
    showing_runcards = False
    middle_right_runcard_frame.pack_forget()
    middle_right_setting_frame.pack(fill=tk.BOTH, expand=True)

def open_runcard_frame():
    global showing_settings, showing_runcards
    showing_settings = False
    showing_runcards = True
    middle_right_setting_frame.pack_forget()
    middle_right_runcard_frame.pack(fill=tk.BOTH, expand=True)
    set_registry_value("is_runcard_open", 1)

def close_setting_frame():
    global showing_settings, showing_runcards
    showing_settings = False
    showing_runcards = False
    middle_right_setting_frame.pack_forget()
    middle_right_runcard_frame.pack_forget()

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

def get_registry_value(name, default="COM1"):
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
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

top_frame = tk.Frame(root, height=50, bg='#f4f4fe', width=screen_width)
top_frame.place(relx=0, rely=0, anchor="nw")

top_left_frame = tk.Frame(top_frame, height=50, bg='#f4f4fe', width=int(screen_width * 0.3))
top_left_frame.place(x=0, y=0, width=screen_width * 0.3, height=50)

top_right_frame = tk.Frame(top_frame, height=50, bg='#f4f4fe', width=int(screen_width * 0.7))
top_right_frame.place(x=screen_width * 0.3, y=0, width=screen_width * 0.7, height=50)
top_right_frame.columnconfigure(0, weight=1)






middle_frame = tk.Frame(root, bg='white')
middle_frame.place(relx=0, rely=0, anchor="nw")

middle_left_frame = tk.Frame(middle_frame, bg='white', width=int(screen_width))
middle_left_frame.place(x=0, y=51, width=screen_width)

middle_center_frame = tk.Frame(middle_frame, bg='#f4f4fe', width=int(screen_width * 0))
middle_center_frame.place(x=0, y=51, width=screen_width * 0)

middle_right_frame = tk.Frame(middle_frame, bg='white', width=int(screen_width * 0))
middle_right_frame.place(x=0, y=51, width=screen_width * 0)


middle_left_weight_frame = tk.Frame(middle_left_frame, bg='white', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_thickness_frame = tk.Frame(middle_left_frame, bg='white', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_weight_frame.pack(fill=tk.BOTH, expand=True)


middle_left_thickness_frame_left_frame =  tk.Frame(middle_left_thickness_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_thickness_frame_left_frame.grid(row=0, column=0, sticky="nw")
middle_left_thickness_frame_left_frame.grid_propagate(False)


middle_left_thickness_frame_left_1_frame = tk.Frame(middle_left_thickness_frame_left_frame, bg='#f4f4fe', width=middle_left_thickness_frame_left_frame.winfo_width(), height=80)
middle_left_thickness_frame_left_1_frame.grid(row=0, column=0, sticky="nw")
middle_left_thickness_frame_left_1_frame.grid_propagate(False)


middle_left_thickness_frame_left_2_frame = tk.Frame(middle_left_thickness_frame_left_frame, bg='white', width=middle_left_thickness_frame_left_frame.winfo_width(), height=int(middle_left_frame.winfo_height() - 80))
middle_left_thickness_frame_left_2_frame.grid(row=1, column=0, sticky="nw")
middle_left_thickness_frame_left_2_frame.grid_propagate(False)


middle_left_thickness_frame_left_2_canvas = tk.Canvas(middle_left_thickness_frame_left_2_frame, width=middle_left_thickness_frame_left_frame.winfo_width() - 20, bg='white', highlightthickness=0)
middle_left_thickness_frame_left_2_scrollbar = tk.Scrollbar(middle_left_thickness_frame_left_2_frame, width=20, orient="vertical", command=middle_left_thickness_frame_left_2_canvas.yview)
middle_left_thickness_frame_left_2_canvas.configure(yscrollcommand=middle_left_thickness_frame_left_2_scrollbar.set)

middle_left_thickness_frame_left_2_scrollable_frame = tk.Frame(middle_left_thickness_frame_left_2_canvas, bg='white', width=middle_left_thickness_frame_left_frame.winfo_width()-20)
middle_left_thickness_frame_left_2_canvas.create_window((0, 0), window=middle_left_thickness_frame_left_2_scrollable_frame, anchor="nw")
middle_left_thickness_frame_left_2_scrollable_frame.bind("<Configure>", lambda e: middle_left_weight_frame_left_2_canvas.configure(scrollregion=middle_left_weight_frame_left_2_canvas.bbox("all")))
middle_left_thickness_frame_left_2_canvas.pack(side="left", fill="both", expand=True)
middle_left_thickness_frame_left_2_scrollbar.pack(side="right", fill="y")
all_entries = []

















def thickness_frame_mouser_pointer_in(event):
    global current_thickness_entry
    if 'name_var' in event.widget.__dict__:
        current_thickness_entry = event.widget.name_var
        print(f"Current Entry: {event.widget.name_var}")



def thickness_insert_data_to_db(runcard_id, cuon_bien, co_tay, ban_tay, ngon_tay, dau_ngon_tay):
    def insert_data():
        try:
            sql = f"""
                       insert into [PMG_DEVICE].[dbo].[ThicknessDeviceData] (RunCardId, DeviceId, Roll, Cuff, Palm, Finger, FingerTip, Cdt)
                       values ('{runcard_id}', '{str(socket.gethostbyname(socket.gethostname()))}', {cuon_bien}, {co_tay}, {ban_tay}, {ngon_tay}, {dau_ngon_tay}, '{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}')
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
                threading.Thread(target=show_error_message, args=(f"Inserted {runcard_id} to database!", 1, 2000), daemon=True).start()
        except Exception as e:
            threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
    if all([runcard_id, cuon_bien, co_tay, ban_tay, ngon_tay, dau_ngon_tay]):
        threading.Thread(target=insert_data, daemon=True).start()
    else:
        threading.Thread(target=show_error_message, args=("Empty field detected!", 0, 3000), daemon=True).start()




def thickness_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_thickness_dau_ngon_tay_entry":
            is_connected = int(get_registry_value("is_connected", "0"))
            if is_connected == 1:
                if all([entry_thickness_runcard_id_entry.get(), entry_thickness_cuon_bien_entry.get(), entry_thickness_co_tay_entry.get(), entry_thickness_ban_tay_entry.get(), entry_thickness_ngon_tay_entry.get(), current_widget.get()]):
                    if int(setting_check_runcard_switch.get()) == 1:
                        if check_runcard_correction(entry_weight_runcard_id_entry.get()):
                            thickness_insert_data_to_db(entry_thickness_runcard_id_entry.get(),
                                                               entry_thickness_cuon_bien_entry.get(),
                                                               entry_thickness_co_tay_entry.get(),
                                                               entry_thickness_ban_tay_entry.get(),
                                                               entry_thickness_ngon_tay_entry.get(),
                                                               current_widget.get())
                            thickness_frame_write_insert_value(entry_thickness_runcard_id_entry.get(),
                                                               entry_thickness_cuon_bien_entry.get(),
                                                               entry_thickness_co_tay_entry.get(),
                                                               entry_thickness_ban_tay_entry.get(),
                                                               entry_thickness_ngon_tay_entry.get(),
                                                               current_widget.get())
                    else:
                        thickness_insert_data_to_db(entry_thickness_runcard_id_entry.get(),
                                                    entry_thickness_cuon_bien_entry.get(),
                                                    entry_thickness_co_tay_entry.get(),
                                                    entry_thickness_ban_tay_entry.get(),
                                                    entry_thickness_ngon_tay_entry.get(),
                                                    current_widget.get())
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
                    threading.Thread(target=show_error_message, args=("Empty field detected!", 0, 3000), daemon=True).start()
            else:
                messagebox.showerror("Error", "Connect to database first!")
        else:
            if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_thickness_runcard_id_entry":
                if int(setting_check_runcard_switch.get()) == 1:
                    if check_runcard_correction(entry_thickness_runcard_id_entry.get()):
                        event.widget.tk_focusNext().focus()
                else:
                    event.widget.tk_focusNext().focus()
            else:
                event.widget.tk_focusNext().focus()
        return "break"
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()



entry_thickness_runcard_id_var = tk.StringVar()
entry_thickness_runcard_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_runcard_id_var, 12, 1))
entry_thickness_runcard_id_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Runcard ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_runcard_id_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_runcard_id_var)
entry_thickness_runcard_id_entry.name_var = "entry_thickness_runcard_id_entry"
entry_thickness_runcard_id_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_thickness_runcard_id_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_runcard_id_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(0, weight=1)



entry_thickness_cuon_bien_var = tk.StringVar()
entry_thickness_cuon_bien_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_cuon_bien_var, 6, 0))
entry_thickness_cuon_bien_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Cuốn biên", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_cuon_bien_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_cuon_bien_var)
entry_thickness_cuon_bien_entry.name_var = "entry_thickness_cuon_bien_entry"
entry_thickness_cuon_bien_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_thickness_cuon_bien_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_cuon_bien_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(1, weight=1)



entry_thickness_co_tay_var = tk.StringVar()
entry_thickness_co_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_co_tay_var, 6, 0))
entry_thickness_co_tay_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Cổ tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_co_tay_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_co_tay_var)
entry_thickness_co_tay_entry.name_var = "entry_thickness_co_tay_entry"
entry_thickness_co_tay_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_thickness_co_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_co_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(2, weight=1)




entry_thickness_ban_tay_var = tk.StringVar()
entry_thickness_ban_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ban_tay_var, 6, 0))
entry_thickness_ban_tay_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Bàn tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_ban_tay_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_ban_tay_var)
entry_thickness_ban_tay_entry.name_var = "entry_thickness_ban_tay_entry"
entry_thickness_ban_tay_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_thickness_ban_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ban_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(3, weight=1)



entry_thickness_ngon_tay_var = tk.StringVar()
entry_thickness_ngon_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_ngon_tay_var, 6, 0))
entry_thickness_ngon_tay_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Ngón tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_ngon_tay_label.grid(row=0, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_ngon_tay_var)
entry_thickness_ngon_tay_entry.name_var = "entry_thickness_ngon_tay_entry"
entry_thickness_ngon_tay_entry.grid(row=1, column=4, padx=5, pady=0, sticky='ew')
entry_thickness_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(4, weight=1)



entry_thickness_dau_ngon_tay_var = tk.StringVar()
entry_thickness_dau_ngon_tay_var.trace_add("write", lambda *args: convert_to_uppercase(entry_thickness_dau_ngon_tay_var, 6, 0))
entry_thickness_dau_ngon_tay_label = tk.Label(middle_left_thickness_frame_left_1_frame, text="Đầu ngón tay", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_thickness_dau_ngon_tay_label.grid(row=0, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry = tk.Entry(middle_left_thickness_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_thickness_dau_ngon_tay_var)
entry_thickness_dau_ngon_tay_entry.name_var = "entry_thickness_dau_ngon_tay_entry"
entry_thickness_dau_ngon_tay_entry.grid(row=1, column=5, padx=5, pady=0, sticky='ew')
entry_thickness_dau_ngon_tay_entry.bind('<FocusIn>', thickness_frame_mouser_pointer_in)
entry_thickness_dau_ngon_tay_entry.bind('<Return>', thickness_frame_hit_enter_button)
middle_left_thickness_frame_left_1_frame.columnconfigure(5, weight=1)



def thickness_frame_write_insert_value(runcard_id, cuon_bien, co_tay, ban_tay, ngon_tay, dau_ngon_tay):
    try:
        global thickness_record_id
        thickness_record_id += 1
        root.update_idletasks()
        root.update()
        frame_width = int(middle_left_thickness_frame_left_2_canvas.winfo_width())
        if not hasattr(thickness_frame_write_insert_value, "tree"):
            columns = ("ID", "Runcard ID", "Cuốn biên", "Cổ tay", "Bàn tay", "Ngón tay", "Đầu ngón tay")
            style = ttk.Style()
            style.theme_use("classic")
            style.configure("Treeview.Heading", font=("Arial", 12, "bold"), relief="flat", background="white", foreground="black", borderwidth=1, highlightthickness=0)
            style.configure("Treeview", font=("Cambria", 13), borderwidth=0, relief="flat", background="white", fieldbackground="white")
            style.map("Treeview", background=[("selected", "lightblue")])
            thickness_frame_write_insert_value.tree = ttk.Treeview(middle_left_thickness_frame_left_2_scrollable_frame, columns=columns, show="headings", style="Treeview", height=int(middle_left_thickness_frame_left_2_canvas.winfo_height() - 80))
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





































middle_left_thickness_frame_center_frame = tk.Frame(middle_left_thickness_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_thickness_frame_center_frame.grid(row=0, column=1, sticky="nw")
middle_left_thickness_frame_center_frame.grid_propagate(False)

middle_left_thickness_frame_right_frame = tk.Frame(middle_left_thickness_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_thickness_frame_right_frame.grid(row=0, column=2, sticky="nw")
middle_left_thickness_frame_right_frame.grid_propagate(False)


























middle_left_weight_frame_left_frame = tk.Frame(middle_left_weight_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_weight_frame_left_frame.grid(row=0, column=0, sticky="nw")
middle_left_weight_frame_left_frame.grid_propagate(False)



middle_left_weight_frame_left_1_frame = tk.Frame(middle_left_weight_frame_left_frame, bg='#f4f4fe', width=middle_left_weight_frame_left_frame.winfo_width(), height=80)
middle_left_weight_frame_left_1_frame.grid(row=0, column=0, sticky="nw")
middle_left_weight_frame_left_1_frame.grid_propagate(False)



middle_left_weight_frame_left_2_frame = tk.Frame(middle_left_weight_frame_left_frame, bg='white', width=middle_left_weight_frame_left_frame.winfo_width(), height=int(middle_left_frame.winfo_height() - 80))
middle_left_weight_frame_left_2_frame.grid(row=1, column=0, sticky="nw")
middle_left_weight_frame_left_2_frame.grid_propagate(False)

middle_left_weight_frame_left_2_canvas = tk.Canvas(middle_left_weight_frame_left_2_frame, width=middle_left_weight_frame_left_frame.winfo_width() - 20, bg='white', highlightthickness=0)
middle_left_weight_frame_left_2_scrollbar = tk.Scrollbar(middle_left_weight_frame_left_2_frame, width=20, orient="vertical", command=middle_left_weight_frame_left_2_canvas.yview)
middle_left_weight_frame_left_2_canvas.configure(yscrollcommand=middle_left_weight_frame_left_2_scrollbar.set)

middle_left_weight_frame_left_2_scrollable_frame = tk.Frame(middle_left_weight_frame_left_2_canvas, bg='white', width=middle_left_weight_frame_left_frame.winfo_width()-20)
middle_left_weight_frame_left_2_canvas.create_window((0, 0), window=middle_left_weight_frame_left_2_scrollable_frame, anchor="nw")
middle_left_weight_frame_left_2_scrollable_frame.bind("<Configure>", lambda e: middle_left_weight_frame_left_2_canvas.configure(scrollregion=middle_left_weight_frame_left_2_canvas.bbox("all")))
middle_left_weight_frame_left_2_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_left_2_scrollbar.pack(side="right", fill="y")
all_entries = []



def weight_frame_write_insert_value(device_id, operator_id, runcard_id, weight_value):
    try:
        global weight_record_id
        weight_record_id += 1
        root.update_idletasks()
        root.update()
        frame_width = int(middle_left_weight_frame_left_2_canvas.winfo_width())
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





def weight_frame_hit_enter_button(event):
    try:
        current_widget = event.widget
        if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_weight_value_entry":
            is_connected = int(get_registry_value("is_connected", "0"))
            if is_connected == 1:
                if all([entry_weight_device_name_entry.get(), entry_weight_operator_id_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get()]):
                    if int(setting_check_runcard_switch.get()) == 1:
                        if check_runcard_correction(entry_weight_runcard_id_entry.get()):
                            weight_insert_data_to_db(entry_weight_device_name_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get(), entry_weight_operator_id_entry.get())
                            weight_frame_write_insert_value(entry_weight_device_name_entry.get(), entry_weight_operator_id_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get())
                    else:
                        weight_insert_data_to_db(entry_weight_device_name_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get(), entry_weight_operator_id_entry.get())
                        weight_frame_write_insert_value(entry_weight_device_name_entry.get(), entry_weight_operator_id_entry.get(), entry_weight_runcard_id_entry.get(), current_widget.get())

                    entry_weight_weight_value_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.delete(0, tk.END)
                    entry_weight_runcard_id_entry.focus_set()
                else:
                    threading.Thread(target=show_error_message, args=("Empty field detected!", 0, 3000), daemon=True).start()
            else:
                messagebox.showerror("Error", "Connect to database first!")
        else:
            if hasattr(current_widget, 'name_var') and current_widget.name_var == "entry_weight_runcard_id_entry":
                if int(setting_check_runcard_switch.get()) == 1:
                    if check_runcard_correction(entry_weight_runcard_id_entry.get()):
                        event.widget.tk_focusNext().focus()
                else:
                    event.widget.tk_focusNext().focus()
            else:
                event.widget.tk_focusNext().focus()
        return "break"
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()



def weight_frame_mouser_pointer_in(event):
    if 'name_var' in event.widget.__dict__:
        print(f"Current Entry: {event.widget.name_var}")


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


entry_weight_device_name_var = tk.StringVar(value=get_registry_value("is_plant_name", ""))
entry_weight_device_name_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_device_name_var, 12, 1))
entry_weight_device_name_label = tk.Label(middle_left_weight_frame_left_1_frame, text="Xưởng", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_device_name_label.grid(row=0, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry = tk.Entry(middle_left_weight_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_weight_device_name_var, fg='black', readonlybackground="#f4f4fe", state="readonly")
entry_weight_device_name_entry.name_var = "entry_weight_device_name_entry"
entry_weight_device_name_entry.grid(row=1, column=0, padx=5, pady=0, sticky='ew')
entry_weight_device_name_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_device_name_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_left_1_frame.columnconfigure(0, weight=1)


entry_weight_operator_id_var = tk.StringVar()
entry_weight_operator_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_operator_id_var, 5, 0))
entry_weight_operator_id_label = tk.Label(middle_left_weight_frame_left_1_frame, text="Số thẻ", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_operator_id_label.grid(row=0, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry = tk.Entry(middle_left_weight_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_weight_operator_id_var)
entry_weight_operator_id_entry.name_var = "entry_weight_operator_id_entry"
entry_weight_operator_id_entry.grid(row=1, column=1, padx=5, pady=0, sticky='ew')
entry_weight_operator_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_operator_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_left_1_frame.columnconfigure(1, weight=1)


entry_weight_runcard_id_var = tk.StringVar()
entry_weight_runcard_id_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_runcard_id_var, 10, 1))
entry_weight_runcard_id_label = tk.Label(middle_left_weight_frame_left_1_frame, text="Runcard ID", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_runcard_id_label.grid(row=0, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry = tk.Entry(middle_left_weight_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_weight_runcard_id_var)
entry_weight_runcard_id_entry.name_var = "entry_weight_runcard_id_entry"
entry_weight_runcard_id_entry.grid(row=1, column=2, padx=5, pady=0, sticky='ew')
entry_weight_runcard_id_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_runcard_id_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_left_1_frame.columnconfigure(2, weight=1)


entry_weight_weight_value_var = tk.StringVar()
entry_weight_weight_value_var.trace_add("write", lambda *args: convert_to_uppercase(entry_weight_weight_value_var, 10, 0))
entry_weight_weight_value_label = tk.Label(middle_left_weight_frame_left_1_frame, text="Trọng lượng", bg="#f4f4fe", bd=0, font=(font_name, 16, "bold"))
entry_weight_weight_value_label.grid(row=0, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry = tk.Entry(middle_left_weight_frame_left_1_frame, font=(font_name, 18), bd=2, textvariable=entry_weight_weight_value_var)
entry_weight_weight_value_entry.name_var = "entry_weight_weight_value_entry"
entry_weight_weight_value_entry.grid(row=1, column=3, padx=5, pady=0, sticky='ew')
entry_weight_weight_value_entry.bind('<FocusIn>', weight_frame_mouser_pointer_in)
entry_weight_weight_value_entry.bind('<Return>', weight_frame_hit_enter_button)
middle_left_weight_frame_left_1_frame.columnconfigure(3, weight=1)





middle_left_weight_frame_center_frame = tk.Frame(middle_left_weight_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_weight_frame_center_frame.grid(row=0, column=1, sticky="nw")
middle_left_weight_frame_center_frame.grid_propagate(False)

middle_left_weight_frame_right_frame = tk.Frame(middle_left_weight_frame, bg='#f4f4fe', width=middle_left_frame.winfo_width(), height=middle_left_frame.winfo_height())
middle_left_weight_frame_right_frame.grid(row=0, column=2, sticky="nw")
middle_left_weight_frame_right_frame.grid_propagate(False)




middle_left_weight_frame_right_left_frame = tk.Frame(middle_left_weight_frame_right_frame, bg='#f4f4fe', width=int(332/2) - 40, height=int(screen_height - 50 - 35 - 20))
middle_left_weight_frame_right_left_frame.grid(row=0, column=0, sticky="nw")
middle_left_weight_frame_right_left_frame.grid_propagate(False)


middle_left_weight_frame_right_right_frame = tk.Frame(middle_left_weight_frame_right_frame, bg='white', width=int(332/2) + 40, height=int(screen_height - 50 - 35 - 20))
middle_left_weight_frame_right_right_frame.grid(row=0, column=1, sticky="nw")
middle_left_weight_frame_right_right_frame.grid_propagate(False)




middle_left_weight_frame_right_frame_log_canvas = tk.Canvas(middle_left_weight_frame_right_right_frame, bg='white', highlightthickness=0, width=int(332/2) + 20, height=int(screen_height - 50 - 35 - 20))
middle_left_weight_frame_right_frame_log_scrollbar = tk.Scrollbar(middle_left_weight_frame_right_right_frame, orient="vertical", width=20, command=middle_left_weight_frame_right_frame_log_canvas.yview)
middle_left_weight_frame_right_frame_log_scrollable_frame = tk.Frame(middle_left_weight_frame_right_frame_log_canvas, bg='white')


middle_left_weight_frame_right_frame_log_canvas.create_window((0, 0), window=middle_left_weight_frame_right_frame_log_scrollable_frame, anchor="nw")
middle_left_weight_frame_right_frame_log_canvas.configure(yscrollcommand=middle_left_weight_frame_right_frame_log_scrollbar.set)

middle_left_weight_frame_right_frame_log_canvas.pack(side="left", fill="both", expand=True)
middle_left_weight_frame_right_frame_log_scrollbar.pack(side="right", fill="y")




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






middle_right_runcard_frame = tk.Frame(middle_right_frame, bg='white', width=middle_right_frame.winfo_width(), height=middle_right_frame.winfo_height())
middle_right_setting_frame = tk.Frame(middle_right_frame, bg='white', width=middle_right_frame.winfo_width(), height=middle_right_frame.winfo_height())


middle_right_setting_frame_top_frame = tk.Frame(middle_right_setting_frame, bg='#f4f4fe', width=middle_right_frame.winfo_width(), height=int(middle_right_frame.winfo_height()*0.9))
middle_right_setting_frame_bottom_frame = tk.Frame(middle_right_setting_frame, bg='#f4f4fe', width=middle_right_frame.winfo_width(), height=int(middle_right_frame.winfo_height()*0.1))
middle_right_setting_frame_top_frame.pack(fill=tk.BOTH, expand=True)
middle_right_setting_frame_bottom_frame.pack(fill=tk.BOTH, expand=False)


middle_right_setting_frame_top_0_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='#f4f4fe', width=332, height=48)
middle_right_setting_frame_top_0_frame.grid(row=0, column=0, sticky="nw")
middle_right_setting_frame_top_0_frame.grid_propagate(False)


middle_right_setting_frame_top_1_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='white', width=332, height=140)
middle_right_setting_frame_top_1_frame.grid(row=1, column=0, sticky="nw")
middle_right_setting_frame_top_1_frame.grid_propagate(False)


middle_right_setting_frame_top_2_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='#f4f4fe', width=332, height=8)
middle_right_setting_frame_top_2_frame.grid(row=2, column=0, sticky="nw")
middle_right_setting_frame_top_2_frame.grid_propagate(False)


middle_right_setting_frame_top_3_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='white', width=332, height=220)
middle_right_setting_frame_top_3_frame.grid(row=3, column=0, sticky="nw")
middle_right_setting_frame_top_3_frame.grid_propagate(False)

middle_right_setting_frame_top_4_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='#f4f4fe', width=332, height=8)
middle_right_setting_frame_top_4_frame.grid(row=4, column=0, sticky="nw")
middle_right_setting_frame_top_4_frame.grid_propagate(False)


middle_right_setting_frame_top_5_frame = tk.Frame(middle_right_setting_frame_top_frame, bg='white', width=332, height=220)
middle_right_setting_frame_top_5_frame.grid(row=5, column=0, sticky="nw")
middle_right_setting_frame_top_5_frame.grid_propagate(False)

















middle_right_runcard_frame_col1 = tk.Frame(middle_right_runcard_frame, bg='white', width=44, height=int(screen_height-200))
middle_right_runcard_frame_col1.grid(row=0, column=0, sticky="nw")
middle_right_runcard_frame_col1.grid_propagate(False)

middle_right_runcard_frame_col2 = tk.Frame(middle_right_runcard_frame, bg='white', width=432, height=int(screen_height-200))
middle_right_runcard_frame_col2.grid(row=0, column=1, sticky="nw")
middle_right_runcard_frame_col2.grid_propagate(False)



middle_right_runcard_frame_col2_row0 = tk.Frame(middle_right_runcard_frame_col2, bg='white', width=432, height=280)
middle_right_runcard_frame_col2_row0.grid(row=0, column=0, sticky="nw")
middle_right_runcard_frame_col2_row0.grid_propagate(False)

middle_right_runcard_frame_col2_row0_1 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_1.place(x=0, y=0, width=430, height=40)

middle_right_runcard_frame_col2_row0_1_label = tk.Label(middle_right_runcard_frame_col2_row0_1, text="工站生產流程卡 Thẻ quy trình sản xuất", bg="white", bd=0, font=(font_name, 12, "bold"))
middle_right_runcard_frame_col2_row0_1_label.pack(expand=True)

middle_right_runcard_frame_col2_row0_2 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_2.place(x=0, y=36, width=430, height=36)

middle_right_runcard_frame_col2_row0_21 = tk.Frame(middle_right_runcard_frame_col2_row0_2, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_21.place(x=-2, y=-2, width=85, height=36)
middle_right_runcard_frame_col2_row0_21_label = tk.Label(middle_right_runcard_frame_col2_row0_21, text="Ngày:", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_21_label.pack(expand=True)

middle_right_runcard_frame_col2_row0_22 = tk.Frame(middle_right_runcard_frame_col2_row0_2, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_22.place(x=82, y=-2, width=113, height=36)
middle_right_runcard_frame_col2_row0_22_label = tk.Label(middle_right_runcard_frame_col2_row0_22, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_22_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_23 = tk.Frame(middle_right_runcard_frame_col2_row0_2, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_23.place(x=188, y=-2, width=113, height=36)
middle_right_runcard_frame_col2_row0_23_label = tk.Label(middle_right_runcard_frame_col2_row0_23, text="Mã vật tư:", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_23_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_24 = tk.Frame(middle_right_runcard_frame_col2_row0_2, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_24.place(x=300, y=-2, width=128, height=36)
middle_right_runcard_frame_col2_row0_24_label = tk.Label(middle_right_runcard_frame_col2_row0_24, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_24_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_3 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_3.place(x=0, y=70, width=430, height=36)

middle_right_runcard_frame_col2_row0_31 = tk.Frame(middle_right_runcard_frame_col2_row0_3, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_31.place(x=-2, y=-2, width=85, height=36)
middle_right_runcard_frame_col2_row0_31_label = tk.Label(middle_right_runcard_frame_col2_row0_31, text="Xưởng:", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_31_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_32 = tk.Frame(middle_right_runcard_frame_col2_row0_3, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_32.place(x=82, y=-2, width=107, height=36)
middle_right_runcard_frame_col2_row0_32_label = tk.Label(middle_right_runcard_frame_col2_row0_32, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_32_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_33 = tk.Frame(middle_right_runcard_frame_col2_row0_3, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_33.place(x=188, y=-2, width=113, height=36)
middle_right_runcard_frame_col2_row0_33_label = tk.Label(middle_right_runcard_frame_col2_row0_33, text="Mã khách hàng", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_33_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_34 = tk.Frame(middle_right_runcard_frame_col2_row0_3, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_34.place(x=300, y=-2, width=128, height=36)
middle_right_runcard_frame_col2_row0_34_label = tk.Label(middle_right_runcard_frame_col2_row0_34, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_34_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_4 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_4.place(x=0, y=104, width=430, height=71)

middle_right_runcard_frame_col2_row0_41 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_41.place(x=-2, y=-2, width=85, height=36)
middle_right_runcard_frame_col2_row0_41_label = tk.Label(middle_right_runcard_frame_col2_row0_41, text="Máy:", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_41_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_42 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_42.place(x=82, y=-2, width=107, height=36)
middle_right_runcard_frame_col2_row0_42_label = tk.Label(middle_right_runcard_frame_col2_row0_42, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_42_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_43 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_43.place(x=-2, y=33, width=85, height=36)
middle_right_runcard_frame_col2_row0_43_label = tk.Label(middle_right_runcard_frame_col2_row0_43, text="Line", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_43_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_44 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1, highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_44.place(x=82, y=33, width=107, height=36)
middle_right_runcard_frame_col2_row0_44_label = tk.Label(middle_right_runcard_frame_col2_row0_44, text=f"", bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_44_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_45 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_45.place(x=188, y=-2, width=113, height=71)
middle_right_runcard_frame_col2_row0_45_label = tk.Label(middle_right_runcard_frame_col2_row0_45,
                                                         text="Tên viết tắt\nkhách hàng", bg="white", bd=0,
                                                         font=(font_name, 12))
middle_right_runcard_frame_col2_row0_45_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_44 = tk.Frame(middle_right_runcard_frame_col2_row0_4, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_44.place(x=300, y=-2, width=128, height=71)
middle_right_runcard_frame_col2_row0_44_label = tk.Label(middle_right_runcard_frame_col2_row0_44, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_44_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_5 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1,
                                                  highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_5.place(x=0, y=172, width=430, height=71)

middle_right_runcard_frame_col2_row0_51 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_51.place(x=-2, y=-2, width=85, height=36)
middle_right_runcard_frame_col2_row0_51_label = tk.Label(middle_right_runcard_frame_col2_row0_51, text="Công đơn",
                                                         bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_51_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_52 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_52.place(x=82, y=-2, width=107, height=36)
middle_right_runcard_frame_col2_row0_52_label = tk.Label(middle_right_runcard_frame_col2_row0_52, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_52_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_53 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_53.place(x=-2, y=33, width=85, height=36)
middle_right_runcard_frame_col2_row0_53_label = tk.Label(middle_right_runcard_frame_col2_row0_53, text="AQL",
                                                         bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_53_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_54 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_54.place(x=82, y=33, width=107, height=36)
middle_right_runcard_frame_col2_row0_54_label = tk.Label(middle_right_runcard_frame_col2_row0_54, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_54_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_55 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_55.place(x=188, y=-2, width=113, height=71)
middle_right_runcard_frame_col2_row0_55_label = tk.Label(middle_right_runcard_frame_col2_row0_55, text="Loại",
                                                         bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_55_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_54 = tk.Frame(middle_right_runcard_frame_col2_row0_5, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_54.place(x=300, y=-2, width=128, height=71)
middle_right_runcard_frame_col2_row0_54_label = tk.Label(middle_right_runcard_frame_col2_row0_54, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_54_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_6 = tk.Frame(middle_right_runcard_frame_col2_row0, bg='white', bd=1,
                                                  highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_6.place(x=0, y=238, width=430, height=36)

middle_right_runcard_frame_col2_row0_61 = tk.Frame(middle_right_runcard_frame_col2_row0_6, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_61.place(x=-2, y=-2, width=85, height=36)
middle_right_runcard_frame_col2_row0_61_label = tk.Label(middle_right_runcard_frame_col2_row0_61, text="Kiểm tra",
                                                         bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_61_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_62 = tk.Frame(middle_right_runcard_frame_col2_row0_6, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_62.place(x=82, y=-2, width=107, height=36)
middle_right_runcard_frame_col2_row0_62_label = tk.Label(middle_right_runcard_frame_col2_row0_62, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_62_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_63 = tk.Frame(middle_right_runcard_frame_col2_row0_6, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_63.place(x=188, y=-2, width=113, height=36)
middle_right_runcard_frame_col2_row0_63_label = tk.Label(middle_right_runcard_frame_col2_row0_63, text="Kích cỡ",
                                                         bg="white", bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_63_label.pack(expand=True, fill="both")

middle_right_runcard_frame_col2_row0_64 = tk.Frame(middle_right_runcard_frame_col2_row0_6, bg='white', bd=1,
                                                   highlightbackground="black", highlightthickness=1)
middle_right_runcard_frame_col2_row0_64.place(x=300, y=-2, width=128, height=36)
middle_right_runcard_frame_col2_row0_64_label = tk.Label(middle_right_runcard_frame_col2_row0_64, text=f"", bg="white",
                                                         bd=0, font=(font_name, 12))
middle_right_runcard_frame_col2_row0_64_label.pack(expand=True, fill="both")





















middle_right_runcard_frame_col2_row1 = tk.Frame(middle_right_runcard_frame_col2, bg='white', width=432, height=40)
middle_right_runcard_frame_col2_row1.grid(row=1, column=0, sticky="nw")
middle_right_runcard_frame_col2_row1.grid_propagate(False)


middle_right_runcard_frame_col2_row2 = tk.Frame(middle_right_runcard_frame_col2, bg='white', width=432, height=40)
middle_right_runcard_frame_col2_row2.grid(row=2, column=0, sticky="nw")
middle_right_runcard_frame_col2_row2.grid_propagate(False)


middle_right_runcard_frame_col2_row3 = tk.Frame(middle_right_runcard_frame_col2, bg='white', width=432, height=280)
middle_right_runcard_frame_col2_row3.grid(row=3, column=0, sticky="nw")
middle_right_runcard_frame_col2_row3.grid_propagate(False)







middle_right_runcard_frame_col3 = tk.Frame(middle_right_runcard_frame, bg='white', width=44, height=int(screen_height-200))
middle_right_runcard_frame_col3.grid(row=0, column=2, sticky="nw")
middle_right_runcard_frame_col3.grid_propagate(False)


# middle_right_runcard_frame_col3_row0 = tk.Frame(middle_right_runcard_frame_col3, bg='white', width=40, height=28)
# middle_right_runcard_frame_col3_row0.grid(row=0, column=0, sticky="nw")
# middle_right_runcard_frame_col3_row0.grid_propagate(False)
# middle_right_runcard_frame_col3_row0_label = tk.Label(middle_right_runcard_frame_col3_row0, text="Time", font=(font_name, 12, "bold"), bg='white')
# middle_right_runcard_frame_col3_row0_label.grid(row=0, column=0, padx=0, pady=0, sticky="w")





























bottom_frame = tk.Frame(root, bg='#f4f4fe')
bottom_frame.place(relx=0, rely=1.0, height=35, width=screen_width, anchor="sw")

bottom_frame_left_frame = tk.Frame(bottom_frame, bg='#f4f4fe')
bottom_frame_left_frame.place(x=0, y=0, height=35, width=screen_width-160)

error_display_entry = tk.Entry(bottom_frame_left_frame, textvariable=error_msg, font=("Cambria", 12), bg="#f4f4fe", fg=error_fg_color, bd=0, highlightthickness=0, readonlybackground="#f4f4fe", state="readonly")
error_display_entry.place(x=5, y=5, width=screen_width-210, height=25)

bottom_frame_right_frame = tk.Frame(bottom_frame, bg='#f4f4fe')
bottom_frame_right_frame.place(x=screen_width-160, y=0, height=35, width=160)








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
    com_port = get_registry_value("COM2", "")
    if "COM" not in com_port:
        return
    print(f"Thickness COM: {selected_com2.get()}")
    try:
        ser = serial.Serial(selected_com2.get(), baudrate=9600, timeout=0.3)
        while ser.is_open:
            if ser.in_waiting > 0:
                raw_value = ser.readline().decode('utf-8').strip()
                threading.Thread(target=show_error_message, args=(f"{raw_value}", 9, 3000), daemon=True).start()
                if raw_value:
                    match = re.search(r"B\+([\d.]+)", raw_value)
                    if match:
                        thickness_value = float(match.group(1))
                        if current_thickness_entry == entry_thickness_ngon_tay_entry:
                            thickness_value = float(f"{thickness_value / 2:.2f}")
                        elif current_thickness_entry == entry_thickness_dau_ngon_tay_entry:
                            thickness_value = float(f"{thickness_value / 2:.3f}")
                        print(f"Thickness Value: {thickness_value}")
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
                            threading.Thread(target=show_error_message, args=("Chưa nhập giá trị Runcard!", 0, 3000),
                                             daemon=True).start()

    except serial.SerialException as e:
        threading.Thread(target=show_error_message, args=(f"Serial Error: {e}", 0, 3000), daemon=True).start()
    except ValueError:
        threading.Thread(target=show_error_message, args=("Invalid thickness value!", 0, 3000), daemon=True).start()
    finally:
        if ser and ser.is_open:
            ser.close()

def weight_frame_com_port_insert_data():
    global weight_record_log_id
    if "COM" in get_registry_value("COM1", ""):
        print(f"Weight COM: {selected_com1.get()}")
        ser = serial.Serial(selected_com1.get(), baudrate=9600, timeout=0.2)
        try:
            while True:
                if ser.is_open:
                    value = ser.readline().decode('utf-8').strip()
                    if len(value):
                        weight_record_log_id += 1
                        threading.Thread(target=root.after, args=(
                        0, update_com_port_weight_log_display, f"{str(weight_record_log_id).zfill(4)}   {value}"),
                                         daemon=True).start()
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

setting_frame_label = tk.Label(middle_right_setting_frame_top_0_frame, text="Cài đặt", font=(font_name, 18, "bold"), bg='#f0f2f6')
setting_frame_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")



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







selected_com1 = tk.StringVar(value=get_registry_value("COM1", ""))
selected_com2 = tk.StringVar(value=get_registry_value("COM2", ""))


com1_label = tk.Label(middle_right_setting_frame_top_1_frame, text="Trọng lượng:   ", font=(font_name, font_size_ratio), bg='white')
com1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
com1_menu = CustomOptionMenu(middle_right_setting_frame_top_1_frame, selected_com1, "")
com1_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

com2_label = tk.Label(middle_right_setting_frame_top_1_frame, text="Độ dày:", font=(font_name, font_size_ratio), bg='white')
com2_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
com2_menu = CustomOptionMenu(middle_right_setting_frame_top_1_frame, selected_com2, "")
com2_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")

update_com_ports(com1_menu, com2_menu)


def save_setting_frame():
    try:
        threading.Thread(target=show_error_message, args=(f"Save setting", 1, 3000), daemon=True).start()
        set_registry_value("COM1", selected_com1.get())
        set_registry_value("COM2", selected_com2.get())
        set_registry_value("ServerIP", server_ip.get())
        set_registry_value("is_check_runcard", setting_check_runcard_switch.get())
        set_registry_value("is_plant_name", plant_name.get())
        # set_registry_value("is_show_runcard", setting_show_runcard_switch.get())
        switch_middle_left_frame()
        database_test_connection()
        # manage_machine_buttons(middle_right_runcard_frame_col1, get_registry_value('is_plant_name', ""))
        # manage_line_buttons()
        messagebox.showinfo("Success", "Save setting success!\nApplication will close automatically\nRe-open the app")
        root.destroy()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
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
        if "COM" in str(get_registry_value("COM1", "")):
            if weight_com_thread is None or not weight_com_thread.is_alive():
                weight_com_thread = threading.Thread(target=weight_frame_com_port_insert_data, daemon=True)
                weight_com_thread.start()
        if "COM" in str(get_registry_value("COM2", "")):
            if thickness_com_thread is None or not thickness_com_thread.is_alive():
                thickness_com_thread = threading.Thread(target=thickness_frame_com_port_insert_data, daemon=True)
                thickness_com_thread.start()
        set_selected_frame(selected_middle_left_frame.get())
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()

selected_middle_left_frame = tk.StringVar(value=get_selected_frame())
switch_middle_left_frame()
frame_select_label = tk.Label(middle_right_setting_frame_top_1_frame, text="Mặc định:", font=(font_name, font_size_ratio), bg='white')
frame_select_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

frame_select_menu = CustomOptionMenu(middle_right_setting_frame_top_1_frame, selected_middle_left_frame, "Trọng lượng", "Độ dày",command=switch_middle_left_frame)
frame_select_menu.grid(row=2, column=1, padx=5, pady=5, sticky="w")




server_ip_label = tk.Label(middle_right_setting_frame_top_3_frame, text="Server IP:       ", font=(font_name, font_size_ratio), bg='white')
server_ip_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
server_ip = tk.StringVar(value=get_registry_value("ServerIP", ""))
server_ip_entry = tk.Entry(middle_right_setting_frame_top_3_frame, textvariable=server_ip, font=(font_name, font_size_ratio), width=12, bg='white')
server_ip_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")




db_name_label = tk.Label(middle_right_setting_frame_top_3_frame, text="Database:", font=(font_name, font_size_ratio), bg='white')
db_name_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
db_name = tk.StringVar(value=get_registry_value("Database", "PMG_DEVICE"))
db_name_entry = tk.Entry(middle_right_setting_frame_top_3_frame, textvariable=db_name, font=(font_name, font_size_ratio), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
db_name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")




user_id_label = tk.Label(middle_right_setting_frame_top_3_frame, text="User ID:", font=(font_name, font_size_ratio), bg='white')
user_id_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
user_id = tk.StringVar(value=get_registry_value("UserID", "scadauser"))
user_id_entry = tk.Entry(middle_right_setting_frame_top_3_frame, textvariable=user_id, font=(font_name, font_size_ratio), width=12, bg='white', fg='red', readonlybackground='white', state='readonly')
user_id_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")



password_label = tk.Label(middle_right_setting_frame_top_3_frame, text="Password:", font=(font_name, font_size_ratio), bg='white')
password_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
password = tk.StringVar(value=get_registry_value("Password", "pmgscada+123"))
password_entry = tk.Entry(middle_right_setting_frame_top_3_frame, textvariable=password, font=(font_name, font_size_ratio), width=12, show='*', bg='white', fg='red', readonlybackground='white', state='readonly')
password_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")


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
                    # if len(plant_name.get()):
                    #     threading.Thread(target=runcard_machine_line_list, args=(plant_name.get(),), daemon=True).start()
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




def on_enter_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_hover_icon)
def on_leave_database_test_connection_button(event):
    database_test_connection_button.config(image=database_test_connection_button_icon)
database_test_connection_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "test_db.png")).resize((146, 38)))
database_test_connection_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "test_db_hover.png")).resize((146, 38)))
database_test_connection_button = tk.Button(middle_right_setting_frame_top_3_frame, image=database_test_connection_button_icon, command=database_test_connection, bg='#f0f2f6', width=146, height=38, relief="flat", borderwidth=0)
database_test_connection_button.grid(row=4, column=1, padx=5, pady=5, sticky="w")
database_test_connection_button.bind("<Enter>", on_enter_database_test_connection_button)
database_test_connection_button.bind("<Leave>", on_leave_database_test_connection_button)






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
        threading.Thread(target=show_error_message, args=("Chưa nhập giá trị Runcard", 0, 3000), daemon=True).start()





plant_name_label = tk.Label(middle_right_setting_frame_top_5_frame, text="Plant name:", font=(font_name, font_size_ratio), bg='white')
plant_name_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
plant_name = tk.StringVar(value=get_registry_value("is_plant_name", ""))
plant_name_entry = tk.Entry(middle_right_setting_frame_top_5_frame, textvariable=plant_name, font=(font_name, font_size_ratio), width=12, bg='white')
plant_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")



setting_check_runcard_switch_label = tk.Label(middle_right_setting_frame_top_5_frame, text="Checkruncard:", font=(font_name, font_size_ratio), bg='white')
setting_check_runcard_switch_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
setting_check_runcard_switch = tk.StringVar(value=get_registry_value("is_check_runcard", "0"))
setting_check_runcard_switch_button = ttk.Checkbutton(middle_right_setting_frame_top_5_frame, variable=setting_check_runcard_switch, onvalue='1', offvalue='0', style="Switch")
setting_check_runcard_switch_button.grid(row=1, column=1, padx=5, pady=5, sticky="w")


#
# setting_show_runcard_switch_label = tk.Label(middle_right_setting_frame_top_5_frame, text="Showruncard:", font=(font_name, font_size_ratio), bg='white')
# setting_show_runcard_switch_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
# setting_show_runcard_switch = tk.StringVar(value=get_registry_value("is_show_runcard", "0"))
# setting_show_runcard_switch_button = ttk.Checkbutton(middle_right_setting_frame_top_5_frame, variable=setting_show_runcard_switch, onvalue='1', offvalue='0', style="Switch")
# setting_show_runcard_switch_button.grid(row=2, column=1, padx=5, pady=5, sticky="w")
#
#
#
#





def on_enter_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_hover_icon)
def on_leave_top_open_weight_frame_button(event):
    top_open_weight_frame_button.config(image=top_open_weight_frame_button_icon)
top_open_weight_frame_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "weight.png")).resize((156, 36)))
top_open_weight_frame_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "weight_hover.png")).resize((156, 36)))
top_open_weight_frame_button = tk.Button(top_right_frame, image=top_open_weight_frame_button_icon, command=open_weight_frame, bg='#f0f2f6', width=156, height=36, relief="flat", borderwidth=0)
top_open_weight_frame_button.grid(row=0, column=3, padx=10, pady=5, sticky="e")
top_open_weight_frame_button.bind("<Enter>", on_enter_top_open_weight_frame_button)
top_open_weight_frame_button.bind("<Leave>", on_leave_top_open_weight_frame_button)


def on_enter_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_hover_icon)
def on_leave_top_open_thickness_frame_button(event):
    top_open_thickness_frame_button.config(image=top_open_thickness_frame_button_icon)
top_open_thickness_frame_button_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "thickness.png")).resize((156, 36)))
top_open_thickness_frame_button_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "thickness_hover.png")).resize((156, 36)))
top_open_thickness_frame_button = tk.Button(top_right_frame, image=top_open_thickness_frame_button_icon, command=open_thickness_frame, bg='#f0f2f6', width=156, height=36, relief="flat", borderwidth=0)
top_open_thickness_frame_button.grid(row=0, column=4, padx=5, pady=5, sticky="e")
top_open_thickness_frame_button.bind("<Enter>", on_enter_top_open_thickness_frame_button)
top_open_thickness_frame_button.bind("<Leave>", on_leave_top_open_thickness_frame_button)



def thickness_frame_left_entry_delete_button():
    try:
        threading.Thread(target=show_error_message, args=(f"Delete all entries", 9, 3000), daemon=True).start()
        entry_thickness_runcard_id_entry.delete(0, tk.END)
        entry_thickness_cuon_bien_entry.delete(0, tk.END)
        entry_thickness_co_tay_entry.delete(0, tk.END)
        entry_thickness_ban_tay_entry.delete(0, tk.END)
        entry_thickness_ngon_tay_entry.delete(0, tk.END)
        entry_thickness_dau_ngon_tay_entry.delete(0, tk.END)
        entry_thickness_runcard_id_entry.focus_set()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass

def on_enter_thickness_entry_delete_button(event):
    thickness_entry_delete_button.config(image=thickness_entry_delete_hover_icon)
def on_leave_thickness_entry_delete_button(event):
    thickness_entry_delete_button.config(image=thickness_entry_delete_icon)
thickness_entry_delete_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete.png")).resize((117, 32)))
thickness_entry_delete_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_hover.png")).resize((117, 32)))
thickness_entry_delete_button = tk.Button(middle_left_thickness_frame_right_frame, image=thickness_entry_delete_icon, command=thickness_frame_left_entry_delete_button, bg='#f0f2f6', width=117, height=32, relief="flat", borderwidth=0)
thickness_entry_delete_button.place(x=0, y=24)
thickness_entry_delete_button.bind("<Enter>", on_enter_thickness_entry_delete_button)
thickness_entry_delete_button.bind("<Leave>", on_leave_thickness_entry_delete_button)








def weight_frame_left_entry_delete_button():
    try:
        threading.Thread(target=show_error_message, args=(f"Delete all entries", 9, 3000), daemon=True).start()
        entry_weight_weight_value_entry.delete(0, tk.END)
        entry_weight_runcard_id_entry.delete(0, tk.END)
        entry_weight_operator_id_entry.delete(0, tk.END)
        entry_weight_device_name_entry.delete(0, tk.END)
        entry_weight_device_name_entry.focus_set()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


def on_enter_weight_entry_delete_button(event):
    weight_entry_delete_button.config(image=weight_entry_delete_hover_icon)
def on_leave_weight_entry_delete_button(event):
    weight_entry_delete_button.config(image=weight_entry_delete_icon)
weight_entry_delete_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete.png")).resize((117, 32)))
weight_entry_delete_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_hover.png")).resize((117, 32)))
weight_entry_delete_button = tk.Button(middle_left_weight_frame_right_left_frame, image=weight_entry_delete_icon, command=weight_frame_left_entry_delete_button, bg='#f0f2f6', width=117, height=32, relief="flat", borderwidth=0)
weight_entry_delete_button.place(x=0, y=24)
weight_entry_delete_button.bind("<Enter>", on_enter_weight_entry_delete_button)
weight_entry_delete_button.bind("<Leave>", on_leave_weight_entry_delete_button)










def weight_frame_left_entry_delete_all_button():
    try:
        threading.Thread(target=show_error_message, args=(f"Delete everything", 9, 3000), daemon=True).start()
        if hasattr(weight_frame_write_insert_value, "tree"):
            for item in weight_frame_write_insert_value.tree.get_children():
                weight_frame_write_insert_value.tree.delete(item)
        entry_weight_weight_value_entry.delete(0, tk.END)
        entry_weight_runcard_id_entry.delete(0, tk.END)
        entry_weight_operator_id_entry.delete(0, tk.END)
        entry_weight_device_name_entry.delete(0, tk.END)
        entry_weight_device_name_entry.focus_set()
        global com_data_labels
        for label in com_data_labels:
            label.destroy()
        com_data_labels = []
        middle_left_weight_frame_right_frame_log_scrollable_frame.update_idletasks()
        middle_left_weight_frame_right_frame_log_canvas.configure(scrollregion=middle_left_weight_frame_right_frame_log_canvas.bbox("all"))
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


def on_enter_weight_entry_delete_all_button(event):
    weight_entry_delete_all_button.config(image=weight_entry_delete_all_hover_icon)
def on_leave_weight_entry_delete_all_button(event):
    weight_entry_delete_all_button.config(image=weight_entry_delete_all_icon)
weight_entry_delete_all_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all.png")).resize((36, 36)))
weight_entry_delete_all_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all_hover.png")).resize((36, 36)))
weight_entry_delete_all_button = tk.Button(middle_left_weight_frame_right_left_frame, image=weight_entry_delete_all_icon, command=weight_frame_left_entry_delete_all_button, bg='#f0f2f6', width=36, height=36, relief="flat", borderwidth=0)
weight_entry_delete_all_button.place(x=0, y=84)
weight_entry_delete_all_button.bind("<Enter>", on_enter_weight_entry_delete_all_button)
weight_entry_delete_all_button.bind("<Leave>", on_leave_weight_entry_delete_all_button)








def thickness_frame_left_entry_delete_all_button():
    try:
        threading.Thread(target=show_error_message, args=(f"Delete everything", 9, 3000), daemon=True).start()
        if hasattr(thickness_frame_write_insert_value, "tree"):
            for item in thickness_frame_write_insert_value.tree.get_children():
                thickness_frame_write_insert_value.tree.delete(item)
        entry_thickness_runcard_id_entry.delete(0, tk.END)
        entry_thickness_cuon_bien_entry.delete(0, tk.END)
        entry_thickness_co_tay_entry.delete(0, tk.END)
        entry_thickness_ban_tay_entry.delete(0, tk.END)
        entry_thickness_ngon_tay_entry.delete(0, tk.END)
        entry_thickness_dau_ngon_tay_entry.delete(0, tk.END)
        entry_thickness_runcard_id_entry.focus_set()
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


def on_enter_thickness_entry_delete_all_button(event):
    thickness_entry_delete_all_button.config(image=thickness_entry_delete_all_hover_icon)
def on_leave_thickness_entry_delete_all_button(event):
    thickness_entry_delete_all_button.config(image=thickness_entry_delete_all_icon)
thickness_entry_delete_all_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all.png")).resize((36, 36)))
thickness_entry_delete_all_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all_hover.png")).resize((36, 36)))
thickness_entry_delete_all_button = tk.Button(middle_left_thickness_frame_right_frame, image=thickness_entry_delete_all_icon, command=thickness_frame_left_entry_delete_all_button, bg='#f0f2f6', width=36, height=36, relief="flat", borderwidth=0)
thickness_entry_delete_all_button.place(x=0, y=84)
thickness_entry_delete_all_button.bind("<Enter>", on_enter_thickness_entry_delete_all_button)
thickness_entry_delete_all_button.bind("<Leave>", on_leave_thickness_entry_delete_all_button)







def weight_frame_left_entry_delete_all_button():
    try:
        threading.Thread(target=show_error_message, args=(f"Delete everything", 9, 3000), daemon=True).start()
        if hasattr(weight_frame_write_insert_value, "tree"):
            for item in weight_frame_write_insert_value.tree.get_children():
                weight_frame_write_insert_value.tree.delete(item)
        entry_weight_weight_value_entry.delete(0, tk.END)
        entry_weight_runcard_id_entry.delete(0, tk.END)
        entry_weight_operator_id_entry.delete(0, tk.END)
        entry_weight_device_name_entry.delete(0, tk.END)
        entry_weight_device_name_entry.focus_set()
        global com_data_labels
        for label in com_data_labels:
            label.destroy()
        com_data_labels = []
        middle_left_weight_frame_right_frame_log_scrollable_frame.update_idletasks()
        middle_left_weight_frame_right_frame_log_canvas.configure(scrollregion=middle_left_weight_frame_right_frame_log_canvas.bbox("all"))
    except Exception as e:
        threading.Thread(target=show_error_message, args=(f"{e}", 0, 3000), daemon=True).start()
        pass


def on_enter_weight_entry_delete_all_button(event):
    weight_entry_delete_all_button.config(image=weight_entry_delete_all_hover_icon)
def on_leave_weight_entry_delete_all_button(event):
    weight_entry_delete_all_button.config(image=weight_entry_delete_all_icon)
weight_entry_delete_all_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all.png")).resize((36, 36)))
weight_entry_delete_all_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "delete_all_hover.png")).resize((36, 36)))
weight_entry_delete_all_button = tk.Button(middle_left_weight_frame_right_left_frame, image=weight_entry_delete_all_icon, command=weight_frame_left_entry_delete_all_button, bg='#f0f2f6', width=36, height=36, relief="flat", borderwidth=0)
weight_entry_delete_all_button.place(x=0, y=84)
weight_entry_delete_all_button.bind("<Enter>", on_enter_weight_entry_delete_all_button)
weight_entry_delete_all_button.bind("<Leave>", on_leave_weight_entry_delete_all_button)















def on_enter_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_hover_icon)
def on_leave_middle_open_setting_frame_button(event):
    middle_open_setting_frame_button.config(image=setting_icon)
setting_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "setting.png")).resize((42, 42)))
setting_hover_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "setting_hover.png")).resize((42, 42)))
middle_open_setting_frame_button = tk.Button(top_left_frame, image=setting_icon, bd=0, command=open_setting_frame, font=(font_name, font_size_ratio), bg="#f4f4fe")
middle_open_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")
middle_open_setting_frame_button.bind("<Enter>", on_enter_middle_open_setting_frame_button)
middle_open_setting_frame_button.bind("<Leave>", on_leave_middle_open_setting_frame_button)





# def on_enter_middle_open_runcard_frame_button(event):
#     middle_open_runcard_frame_button.config(image=barcode_hover_icon)
# def on_leave_middle_open_runcard_frame_button(event):
#     middle_open_runcard_frame_button.config(image=barcode_icon)
# barcode_icon = ImageTk.PhotoImage(Image.open("icon/barcode.png").resize((42, 42)))
# barcode_hover_icon = ImageTk.PhotoImage(Image.open("icon/barcode_hover.png").resize((42, 42)))
# middle_open_runcard_frame_button = tk.Button(top_left_frame, image=barcode_icon, bd=0, command=open_runcard_frame, font=(font_name, font_size_ratio), bg="#f4f4fe")
# middle_open_runcard_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="w")
# middle_open_runcard_frame_button.bind("<Enter>", on_enter_middle_open_runcard_frame_button)
# middle_open_runcard_frame_button.bind("<Leave>", on_leave_middle_open_runcard_frame_button)




def on_enter_middle_save_setting_frame_button(event):
    middle_save_setting_frame_button.config(image=save_icon_hover)
def on_leave_middle_save_setting_frame_buttonn(event):
    middle_save_setting_frame_button.config(image=save_icon)
save_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "save.png")).resize((134, 34)))
save_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "save_hover.png")).resize((134, 34)))
middle_save_setting_frame_button = tk.Button(middle_right_setting_frame_bottom_frame, image=save_icon, command=save_setting_frame, bg='#fafafa', width=155, height=34, relief="flat", borderwidth=0)
middle_save_setting_frame_button.grid(row=0, column=0, padx=5, pady=5, sticky="n")
middle_save_setting_frame_button.bind("<Enter>", on_enter_middle_save_setting_frame_button)
middle_save_setting_frame_button.bind("<Leave>", on_leave_middle_save_setting_frame_buttonn)





def on_enter_middle_close_setting_frame_button(event):
    middle_close_setting_frame_button.config(image=close_icon_hover)
def on_leave_middle_close_setting_frame_button(event):
    middle_close_setting_frame_button.config(image=close_icon)
close_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "close.png")).resize((134, 34)))
close_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "close_hover.png")).resize((134, 34)))
middle_close_setting_frame_button = tk.Button(middle_right_setting_frame_bottom_frame, image=close_icon, command=close_setting_frame, bg='#fafafa', width=155, height=34, relief="flat", borderwidth=0)
middle_close_setting_frame_button.grid(row=0, column=1, padx=5, pady=5, sticky="n")
middle_close_setting_frame_button.bind("<Enter>", on_enter_middle_close_setting_frame_button)
middle_close_setting_frame_button.bind("<Leave>", on_leave_middle_close_setting_frame_button)


def on_enter_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon_hover)
def on_leave_bottom_exit_button(event):
    bottom_exit_button.config(image=exit_icon)
exit_icon = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "exit.png")).resize((134, 31)))
exit_icon_hover = ImageTk.PhotoImage(Image.open(os.path.join(base_path, "icon", "exit_hover.png")).resize((134, 31)))
bottom_exit_button = tk.Button(bottom_frame_right_frame, image=exit_icon, command=root.destroy, bg='#f4f4fe', width=134, height=31, relief="flat", borderwidth=0)
bottom_exit_button.grid(row=0, column=0, columnspan=5, padx=5, pady=5)
bottom_exit_button.bind("<Enter>", on_enter_bottom_exit_button)
bottom_exit_button.bind("<Leave>", on_leave_bottom_exit_button)










"""Start the update thread"""
update_thread = threading.Thread(target=update_dimensions, daemon=True)
update_thread.start()
root.mainloop()



"""
====> pyinstaller --windowed --onefile --name IPQC --add-data "icon;icon" --add-data "assets;assets" main.py --icon=logo.ico
"""
