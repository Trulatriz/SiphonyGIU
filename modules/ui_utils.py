import tkinter as tk
from tkinter import ttk

def setup_toplevel(parent, title, geometry=None, resizable=False):
    window = tk.Toplevel(parent)
    window.title(title)
    window.transient(parent)
    window.grab_set()
    if geometry:
        window.geometry(geometry)
    window.resizable(resizable, resizable)
    window.update_idletasks()
    w = window.winfo_width()
    h = window.winfo_height()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    x = px + (pw // 2) - (w // 2)
    y = py + (ph // 2) - (h // 2)
    window.geometry(f"+{x}+{y}")
    return window


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        self.widget.bind("<Enter>", self._show)
        self.widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tip, text=self.text, background="lightyellow", relief="solid", borderwidth=1)
        label.pack(ipadx=4, ipady=2)

    def _hide(self, _event=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None
