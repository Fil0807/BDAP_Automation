"""Tkinter Themed Widgets styling configuration for desktop GUI widgets."""

from __future__ import annotations

import sys
from tkinter import font as tkfont
from tkinter import ttk


def resolve_ui_font_family(master=None) -> str:
    """Return a readable UI font available on the current operating system."""
    preferred_by_platform = {
        "darwin": ("Avenir Next", "Helvetica Neue", "Arial"),
        "win32": ("Segoe UI", "Arial", "Tahoma"),
        "linux": ("DejaVu Sans", "Noto Sans", "Liberation Sans", "Arial"),
    }
    platform_key = "linux" if sys.platform.startswith("linux") else sys.platform
    candidates = preferred_by_platform.get(platform_key, ()) + ("TkDefaultFont",)

    try:
        available = {family.lower(): family for family in tkfont.families(master)}
    except Exception:
        available = {}

    for candidate in candidates:
        match = available.get(candidate.lower())
        if match:
            return match

    try:
        return tkfont.nametofont("TkDefaultFont", root=master).actual("family")
    except Exception:
        return "TkDefaultFont"


def ui_font(master=None, size: int = 12, weight: str | None = None) -> tuple[str, int] | tuple[str, int, str]:
    """Build a Tk font tuple using a platform-appropriate family."""
    family = resolve_ui_font_family(master)
    if weight:
        return (family, size, weight)
    return (family, size)


def configure_styles(style: ttk.Style) -> None:
    """Configure all TTK styles for the desktop GUI."""
    try:
        style.theme_use("clam")
    except Exception:
        pass

    master = getattr(style, "master", None)

    # Color palette
    colors = {
        "bg_main": "#eef5f0",
        "bg_panel": "#fbfefc",
        "bg_card": "#f4f9f5",
        "bg_card_label": "#eef6ef",
        "bg_status": "#eaf3ed",
        "bg_drop_inactive": "#f7fcf9",
        "bg_drop_active": "#e9f7ef",
        "border": "#c6d8cc",
        "border_drop": "#9cc6af",
        "border_drop_active": "#1f7a4d",
        "text_title": "#12251c",
        "text_subtitle": "#4f6b5c",
        "text_label": "#1d3f32",
        "text_card_label": "#15382c",
        "text_status": "#2a4f40",
        "text_drop": "#4f6b5c",
        "text_drop_active": "#165d3b",
        "btn_primary": "#1f7a4d",
        "btn_primary_hover": "#16603c",
        "btn_primary_disabled": "#8fb5a1",
        "btn_secondary": "#dce9e0",
        "btn_secondary_hover": "#cfdfd4",
        "progress": "#1f7a4d",
        "progress_track": "#d8e7dd",
        "text_white": "#ffffff",
        "text_disabled": "#eef5f1",
    }

    # Widget configurations
    configs = {
        "App.TFrame": {"background": colors["bg_main"]},
        "Panel.TFrame": {"background": colors["bg_panel"], "bordercolor": colors["border"], "borderwidth": 1, "relief": "solid"},
        "Header.TFrame": {"background": colors["bg_panel"]},
        "Card.TLabelframe": {"background": colors["bg_card"], "bordercolor": colors["border"], "borderwidth": 1, "relief": "solid", "padding": 12},
        "Card.TLabelframe.Label": {"background": colors["bg_card_label"], "foreground": colors["text_card_label"], "font": ui_font(master, 18, "bold")},
        "Field.TLabel": {"background": colors["bg_card"], "foreground": colors["text_label"], "font": ui_font(master, 14, "bold")},
        "Hint.TLabel": {"background": colors["bg_card"], "foreground": colors["text_subtitle"], "font": ui_font(master, 11, "bold")},
        "StatusBox.TFrame": {"background": colors["bg_status"], "bordercolor": "#c5d9cd", "borderwidth": 1, "relief": "solid"},
        "Status.TLabel": {"background": colors["bg_status"], "foreground": colors["text_status"], "font": ui_font(master, 12, "bold")},
        "Progress.TLabel": {"background": colors["bg_status"], "foreground": colors["text_subtitle"], "font": ui_font(master, 10)},
        "Bdap.Horizontal.TProgressbar": {"background": colors["progress"], "troughcolor": colors["progress_track"], "bordercolor": colors["progress_track"], "lightcolor": colors["progress"], "darkcolor": colors["progress"]},
        "DropZone.TLabel": {"background": colors["bg_drop_inactive"], "foreground": colors["text_drop"], "bordercolor": colors["border_drop"], "borderwidth": 2, "relief": "solid", "font": ui_font(master, 12, "bold"), "padding": (12, 10)},
        "DropZoneActive.TLabel": {"background": colors["bg_drop_active"], "foreground": colors["text_drop_active"], "bordercolor": colors["border_drop_active"], "borderwidth": 2, "relief": "solid", "font": ui_font(master, 12, "bold"), "padding": (12, 10)},
        "Primary.TButton": {"background": colors["btn_primary"], "foreground": colors["text_white"], "padding": (16, 10), "font": ui_font(master, 12, "bold"), "borderwidth": 0},
        "Secondary.TButton": {"background": colors["btn_secondary"], "foreground": colors["text_label"], "padding": (14, 8), "font": ui_font(master, 12, "bold"), "borderwidth": 0},
        "Field.TEntry": {"fieldbackground": colors["text_white"], "padding": 6},
        "Field.TCombobox": {"fieldbackground": colors["text_white"], "padding": 6},
        "Title.TLabel": {"background": colors["bg_panel"], "foreground": colors["text_title"], "font": ui_font(master, 30, "bold")},
        "Subtitle.TLabel": {"background": colors["bg_panel"], "foreground": colors["text_subtitle"], "font": ui_font(master, 14)},
    }

    # Apply configurations
    for style_name, config in configs.items():
        style.configure(style_name, **config)

    # Button state maps
    style.map("Primary.TButton", background=[("active", colors["btn_primary_hover"]), ("disabled", colors["btn_primary_disabled"])], foreground=[("disabled", colors["text_disabled"])])
    style.map("Secondary.TButton", background=[("active", colors["btn_secondary_hover"])])
