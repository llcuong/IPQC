import sys
from cefpython3 import cefpython as cef
import ctypes
import win32gui
import win32con
import time


def main(hwnd_str):
    parent_hwnd = int(hwnd_str)

    cef.Initialize()
    window_info = cef.WindowInfo()
    window_info.SetAsChild(parent_hwnd, [0, 0, 1000, 600])
    browser = cef.CreateBrowserSync(window_info, url="http://10.13.104.181:10000/")

    # Lấy HWND của trình duyệt
    browser_hwnd = browser.GetWindowHandle()

    # Tắt tính năng chiếm focus của CEF
    win32gui.SetWindowLong(browser_hwnd, win32con.GWL_EXSTYLE,
                           win32gui.GetWindowLong(browser_hwnd, win32con.GWL_EXSTYLE) & ~win32con.WS_EX_APPWINDOW)

    # Cho phép thao tác Tkinter
    win32gui.EnableWindow(browser_hwnd, True)  # Nếu cần tắt focus: False
    cef.MessageLoop()
    cef.Shutdown()


if __name__ == "__main__":
    main(sys.argv[1])
