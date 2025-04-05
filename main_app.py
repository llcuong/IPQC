import tkinter as tk
import subprocess
import time
import sys
import win32gui
import win32con

root = tk.Tk()
root.geometry("1000x600")
root.title("App chính với Entry và CEF")

# LEFT PANEL
left_frame = tk.Frame(root, width=300, bg="white")
left_frame.pack(side="left", fill="y")

tk.Label(left_frame, text="Nhập số:", font=("Arial", 14)).pack(pady=10)
entry = tk.Entry(left_frame, font=("Arial", 14))
entry.pack(pady=10)

result_var = tk.StringVar()
tk.Label(left_frame, textvariable=result_var, font=("Arial", 18), fg="blue").pack(pady=10)

def calculate():
    try:
        val = int(entry.get())
        result_var.set(f"Kết quả: {val * 10}")
    except ValueError:
        result_var.set("Lỗi nhập!")

tk.Button(left_frame, text="Tính", font=("Arial", 12), command=calculate).pack(pady=10)

# RIGHT PANEL
right_frame = tk.Frame(root, bg="black", width=700, height=600)
right_frame.pack(side="right", fill="both", expand=True)

# Gửi HWND
root.update_idletasks()
hwnd = right_frame.winfo_id()
proc = subprocess.Popen([sys.executable, "cef_child.py", str(hwnd)])

# 🔄 Chuyển focus về Entry nếu cần
def return_focus():
    entry.focus_set()

tk.Button(left_frame, text="Trả focus về Entry", command=return_focus).pack(pady=10)

root.mainloop()
