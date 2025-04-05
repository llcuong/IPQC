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
import re
import barcode
import os
import sys
from cefpython3 import cefpython as cef
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


