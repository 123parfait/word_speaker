# -*- coding: utf-8 -*-
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk


def init_runtime_paths():
    base_dir = Path(__file__).resolve().parent
    for vendor_path in (base_dir / "vendor" / "site-packages",):
        path_str = str(vendor_path)
        if vendor_path.exists() and path_str not in sys.path:
            sys.path.insert(0, path_str)


init_runtime_paths()

from ui.main_view import MainView


def init_style(root):
    root.configure(bg="#f6f7fb")
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TFrame", background="#f6f7fb")
    style.configure("Card.TFrame", background="#ffffff")
    style.configure("TLabel", background="#f6f7fb", foreground="#222")
    style.configure("Card.TLabel", background="#ffffff", foreground="#222")
    style.configure("TButton", padding=(10, 6), background="#f5f6f8", foreground="#222222")
    style.map(
        "TButton",
        background=[("active", "#c6d6ff"), ("pressed", "#a9c1ff")],
        foreground=[("active", "#1e3a8a"), ("pressed", "#1e3a8a")],
    )
    style.configure("Primary.TButton", background="#f5f6f8", foreground="#222222")
    style.map(
        "Primary.TButton",
        background=[("active", "#c6d6ff"), ("pressed", "#a9c1ff")],
        foreground=[("active", "#1e3a8a"), ("pressed", "#1e3a8a")],
    )
    style.configure(
        "CardButton.TButton",
        foreground="#1f2937",
        background="#f5f6f8",
        padding=(18, 12),
    )
    style.map(
        "CardButton.TButton",
        background=[("active", "#c6d6ff"), ("pressed", "#a9c1ff")],
        foreground=[("active", "#1e3a8a"), ("pressed", "#1e3a8a")],
    )
    style.configure("Icon.TButton", padding=(6, 4), background="#f5f6f8", foreground="#222222")
    style.configure("Speed.TButton", padding=(8, 4), background="#f5f6f8", foreground="#374151")
    style.configure(
        "SelectedSpeed.TButton",
        padding=(8, 4),
        background="#cfe3ff",
        foreground="#1e3a8a",
    )
    style.configure("TCheckbutton", background="#f6f7fb")
    style.configure("TRadiobutton", background="#f6f7fb")
    style.configure(
        "WordList.Treeview",
        background="#ffffff",
        fieldbackground="#ffffff",
        foreground="#2b2f36",
        rowheight=70,
        font=("Segoe UI", 13),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "WordList.Treeview",
        background=[("selected", "#d9e8ff")],
        foreground=[("selected", "#1b2736")],
    )
    style.configure(
        "WordList.Treeview.Heading",
        background="#f8fafc",
        foreground="#6b7280",
        font=("Segoe UI", 9, "bold"),
        padding=(6, 6),
        borderwidth=0,
        relief="flat",
    )


def main():
    root = tk.Tk()
    root.title("Word Speaker")
    root.geometry("1480x860")
    root.minsize(1320, 760)
    init_style(root)

    MainView(root).pack(fill="both", expand=True, padx=20, pady=20)
    root.mainloop()


if __name__ == "__main__":
    main()
