import tkinter as tk
import datetime
from tkcalendar import DateEntry

root = tk.Tk()


bg_app_class_color_layer_1 = "#f4f4fe"
bg_app_class_color_layer_2 = "#ffffff"
fg_app_class_color_layer_1 = '#000000'
fg_app_class_color_layer_2 = '#ffffff'
current_date = (datetime.datetime.now() - datetime.timedelta(hours=5) + datetime.timedelta(minutes=22)).date()
calendar = DateEntry(root,
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
    font=("Arial", 12),
    date_pattern='dd-mm-yyyy',
                     tooltipforeground='white')

calendar.pack(padx=10, pady=10)
def on_date_change(event):
    print("Selected date:", calendar.get_date().strftime("%d-%m-%Y"))
calendar.bind("<<DateEntrySelected>>", on_date_change)


root.mainloop()

