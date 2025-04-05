import tkinter as tk
import sys
import requests
from PIL import Image
from io import BytesIO
from pyzbar.pyzbar import decode
from cefpython3 import cefpython as cef

# ✳️ Chặn Chromium chiếm focus
class FocusHandler:
    def OnSetFocus(self, browser, source):
        print("Blocked Chromium from stealing focus")
        return True  # ✅ Không cho Chromium chiếm focus

# ✳️ Frame chứa Chromium
class BrowserFrame(tk.Frame):
    def __init__(self, master, url):
        super().__init__(master)
        self.url = url
        self.browser = None
        self.bind("<Configure>", self.on_configure)

    def embed_browser(self):
        window_info = cef.WindowInfo()
        rect = [0, 0, self.winfo_width(), self.winfo_height()]
        window_info.SetAsChild(self.winfo_id(), rect)
        self.browser = cef.CreateBrowserSync(window_info, url=self.url)
        self.browser.SetFocusHandler(FocusHandler())

        # 🔗 Bind JS function to Python callback
        bindings = cef.JavascriptBindings()
        bindings.SetFunction("sendImageUrlToPython", self._on_image_url_received)
        self.browser.SetJavascriptBindings(bindings)

    def on_configure(self, _):
        if self.browser is None and self.winfo_width() > 0 and self.winfo_height() > 0:
            self.embed_browser()
        elif self.browser:
            cef.WindowUtils.OnSize(self.winfo_id(), 0, 0, 0)

    def _on_image_url_received(self, img_url):
        print(f"[JS → Python] Image URL received: {img_url}")
        result = scan_barcode_from_url(img_url)
        result_var.set(f"📦 Mã vạch: {result}" if result else "❌ Không đọc được mã vạch")

    def get_image_url_and_decode(self):
        js_code = """
            (function() {
                let img = document.querySelector("img");
                if (img && img.src) {
                    sendImageUrlToPython(img.src);
                } else {
                    sendImageUrlToPython("NO_IMAGE");
                }
            })();
        """
        if self.browser:
            self.browser.GetMainFrame().ExecuteJavascript(js_code)


# ✳️ Quét barcode từ image URL
def scan_barcode_from_url(img_url):
    try:
        response = requests.get(img_url)
        img = Image.open(BytesIO(response.content))
        decoded = decode(img)
        return decoded[0].data.decode("utf-8") if decoded else None
    except Exception as e:
        print(f"Lỗi khi đọc mã vạch: {e}")
        return None

# ✳️ Khởi tạo CEF
sys.excepthook = cef.ExceptHook
cef.Initialize()

# ✳️ Tạo root window
root = tk.Tk()
root.geometry("1000x600")
root.title("CEF Barcode Reader")

# ✳️ Frame bên trái
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

def on_entry_click(event):
    root.after(10, lambda: entry.focus_force())
    print("Entry đã lấy lại focus")

entry.bind("<Button-1>", on_entry_click)

# ✳️ Button: Quét mã vạch từ hình
def scan_barcode():
    browser_frame.get_image_url_and_decode()

tk.Button(left_frame, text="🔍 Quét mã vạch", font=("Arial", 12), command=scan_barcode).pack(pady=10)


# ✳️ Frame bên phải chứa trình duyệt
right_frame = tk.Frame(root, bg="gray")
right_frame.pack(side="right", fill="both", expand=True)

# ✳️ Tạo trình duyệt trong right_frame
browser_frame = BrowserFrame(right_frame, url="http://10.13.104.181:10000/")
browser_frame.pack(fill="both", expand=True)

# ✳️ Vòng lặp CEF không block Tkinter
def cef_loop():
    cef.MessageLoopWork()
    root.after(1, cef_loop)

root.after(1, cef_loop)

def on_close():
    cef.Shutdown()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
