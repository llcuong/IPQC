import tkinter as tk
import webview

def get_embed_position():
    # Wait until the frame is rendered
    root.update_idletasks()

    # Get position and size of the frame
    x = embed_frame.winfo_rootx()
    y = embed_frame.winfo_rooty()
    width = embed_frame.winfo_width()
    height = embed_frame.winfo_height()
    return x, y, width, height

def start_webview():
    x, y, w, h = get_embed_position()
    webview.create_window(
        "Embedded Browser",
        "http://10.13.104.181:10000/",
        x=x,
        y=y,
        width=w,
        height=h,
        frameless=True,
        resizable=False
    )
    webview.start()

root = tk.Tk()
root.geometry("1024x768")
root.title("Main Tkinter App")

embed_frame = tk.Frame(root, width=800, height=600, bg='lightgray')
embed_frame.place(x=100, y=100)

# Delay start to allow window to render, then run webview on main thread
root.after(500, start_webview)

root.mainloop()
