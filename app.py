# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk

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
        background=[("active", "#eef2ff"), ("pressed", "#dbe7ff")],
        foreground=[("active", "#1e3a8a"), ("pressed", "#1e3a8a")],
    )
    # Primary uses the same neutral style; color only on interaction
    style.configure("Primary.TButton", background="#f5f6f8", foreground="#222222")
    style.map(
        "Primary.TButton",
        background=[("active", "#eef2ff"), ("pressed", "#dbe7ff")],
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
        background=[("active", "#eef2ff"), ("pressed", "#dbe7ff")],
        foreground=[("active", "#1e3a8a"), ("pressed", "#1e3a8a")],
    )
    style.configure("Speed.TButton", padding=(8, 4), background="#f5f6f8", foreground="#374151")
    style.configure(
        "SelectedSpeed.TButton",
        padding=(8, 4),
        background="#cfe3ff",
        foreground="#1e3a8a",
    )
    style.configure("TCheckbutton", background="#f6f7fb")
    style.configure("TRadiobutton", background="#f6f7fb")


def main():
    root = tk.Tk()
    root.title("Word Speaker")
    init_style(root)

    MainView(root).pack(padx=20, pady=20)
    root.mainloop()


if __name__ == "__main__":
    main()
