import tkinter as tk
from tkinter import ttk
from colors import HITACHI_COLORS
import time


class SplashScreen:
    def __init__(self, parent):
        self.parent = parent

        self.splash = tk.Toplevel(parent)
        self.splash.title("QUALITY TEAM")
        self.splash.geometry("500x400")
        self.splash.configure(bg=HITACHI_COLORS["white"])

        # Remove window decorations
        self.splash.overrideredirect(True)

        # Center the splash screen
        self.center_window()

        # Add content
        self.create_content()

    def center_window(self):
        self.splash.update_idletasks()
        width = 500
        height = 400
        x = (self.splash.winfo_screenwidth() // 2) - (width // 2)
        y = (self.splash.winfo_screenheight() // 2) - (height // 2)
        self.splash.geometry(f'{width}x{height}+{x}+{y}')

    def create_content(self):
        # Main frame
        main_frame = tk.Frame(self.splash, bg=HITACHI_COLORS["white"])
        main_frame.pack(fill="both", expand=True)

        title_frame = tk.Frame(main_frame, bg=HITACHI_COLORS["white"])
        title_frame.pack(pady=50)

        tk.Label(title_frame,
                 text="QUALITY",
                 font=("Arial", 32, "bold"),
                 fg=HITACHI_COLORS["primary_red"],
                 bg=HITACHI_COLORS["white"]).pack()

        tk.Label(title_frame,
                 text="TEAM",
                 font=("Arial", 24, "bold"),
                 fg=HITACHI_COLORS["black"],
                 bg=HITACHI_COLORS["white"]).pack()

        # Application title
        tk.Label(main_frame,
                 text="Digital LPA & ISO Compliance System",
                 font=("Arial", 14),
                 fg=HITACHI_COLORS["dark_gray"],
                 bg=HITACHI_COLORS["white"]).pack(pady=20)

        # Loading bar
        loading_frame = tk.Frame(main_frame, bg=HITACHI_COLORS["white"])
        loading_frame.pack(pady=30)

        self.progress = ttk.Progressbar(loading_frame,
                                        length=300,
                                        mode='indeterminate')
        self.progress.pack()
        self.progress.start(10)

        # Status text
        self.status_label = tk.Label(main_frame,
                                     text="Loading application...",
                                     font=("Arial", 10),
                                     fg=HITACHI_COLORS["text_secondary"],
                                     bg=HITACHI_COLORS["white"])
        self.status_label.pack(pady=10)

        # Version info
        tk.Label(main_frame,
                 text="v1.0.0",
                 font=("Arial", 8),
                 fg=HITACHI_COLORS["medium_gray"],
                 bg=HITACHI_COLORS["white"]).pack(side="bottom", pady=10)

    def update_status(self, text):
        self.status_label.config(text=text)
        self.splash.update()

    def destroy(self):
        self.progress.stop()
        time.sleep(0.5)  # Brief pause before closing
        self.splash.destroy()