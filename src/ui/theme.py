import customtkinter as ctk
from tkinter import ttk

COLORS = {
    "bg": "#F7F8FA",
    "surface": "#FFFFFF",
    "border": "#EEF0F2",
    "primary": "#1B7A3D",
    "primary_dark": "#15632F",
    "primary_muted": "#E8F5EC",
    "field_bg": "#F3F6F4",
    "field_border": "#D8E2DD",
    "chip_bg": "#E6F3EA",
    "chip_text": "#1B7A3D",
    "error": "#D92D20",
    "text": "#1A1D21",
    "text_muted": "#5F6B7A",
    "text_light": "#C8E6C9",
    "info_bg": "#EBF5FB",
    "info_border": "#BEE0F5",
    "info_text": "#1565C0",
}

def apply_base_theme():
    """Applies base customtkinter theme settings."""
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("green")

def setup_ttk_styles(root, font_scale):
    """Configures ttk styles for the application."""
    style = ttk.Style(root)
    style.theme_use("clam")

    # Helper function for scaled fonts
    def sf(n): return max(8, round(n * font_scale))

    base_font_size = sf(13)
    style.configure(".", font=("Segoe UI", base_font_size))

    style.configure("TFrame", background=COLORS["surface"])
    style.configure("Header.TFrame", background=COLORS["primary"])
    style.configure(
        "Header.TLabel",
        background=COLORS["primary"],
        foreground=COLORS["surface"],
        font=("Segoe UI", sf(24), "bold"),
    )
    style.configure("Title.TLabel", font=("Segoe UI", sf(20), "bold"), foreground=COLORS["primary"])
    style.configure("Subtitle.TLabel", font=("Segoe UI", sf(14)), foreground=COLORS["text_muted"])
    style.configure("TLabel", background=COLORS["surface"], foreground=COLORS["text"])

    style.configure("Action.TButton", font=("Segoe UI", sf(13), "bold"), padding=9)
    style.map(
        "Action.TButton",
        background=[("pressed", COLORS["primary_dark"]), ("active", COLORS["primary"])],
        foreground=[("active", COLORS["surface"])],
    )

    style.configure("Primary.TButton", font=("Segoe UI", sf(13), "bold"), padding=9)
    style.map(
        "Primary.TButton",
        background=[("pressed", COLORS["primary_dark"]), ("active", COLORS["primary"])],
        foreground=[("active", COLORS["surface"])],
    )

    style.configure(
        "Menu.TButton",
        font=("Segoe UI", sf(13)),
        padding=11,
        background=COLORS["surface"],
        foreground=COLORS["primary"],
    )
    style.map("Menu.TButton", background=[("active", COLORS["primary_muted"])])

    style.configure(
        "MenuActive.TButton",
        font=("Segoe UI", sf(13), "bold"),
        padding=11,
        background=COLORS["primary"],
        foreground=COLORS["surface"],
    )
    style.map("MenuActive.TButton", background=[("active", COLORS["primary"])])

    style.configure("TSeparator", background=COLORS["border"])

    row_height = max(28, round(34 * font_scale))
    style.configure(
        "Treeview",
        background=COLORS["surface"],
        fieldbackground=COLORS["surface"],
        foreground=COLORS["text"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        rowheight=row_height,
        font=("Segoe UI", sf(13)),
    )
    style.configure(
        "Treeview.Heading",
        background=COLORS["primary_muted"],
        foreground=COLORS["primary"],
        font=("Segoe UI", sf(13), "bold"),
    )
    style.map(
        "Treeview",
        background=[("selected", COLORS["primary"])],
        foreground=[("selected", COLORS["surface"])],
    )
    style.configure("TEntry", fieldbackground=COLORS["field_bg"], padding=8, font=("Segoe UI", sf(13)))
    style.map("TEntry",
        foreground=[("readonly", COLORS["text"]), ("disabled", COLORS["text"])],
        fieldbackground=[("readonly", COLORS["field_bg"]), ("disabled", COLORS["field_bg"])],
    )
    style.configure("TCombobox", padding=8, font=("Segoe UI", sf(13)))
