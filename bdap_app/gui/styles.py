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


def ui_font(master=None, size: int = 11, weight: str | None = None) -> tuple[str, int] | tuple[str, int, str]:
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

    # Modern blue & white color palette
    colors = {
        "bg_main": "#f0f5fa",
        "bg_panel": "#f0f5fa",
        "bg_card": "#ffffff",
        "bg_card_label": "#ffffff",
        "bg_status": "#ffffff",
        "bg_drop_inactive": "#f7faff",
        "bg_drop_active": "#e3effc",
        "border": "#dbe7f5",
        "border_drop": "#93c5fd",
        "border_drop_active": "#1976d2",
        "text_title": "#0f2b5c",
        "text_subtitle": "#5a6e85",
        "text_label": "#1e293b",
        "text_card_label": "#0f2b5c",
        "text_status": "#0f2b5c",
        "text_drop": "#1976d2",
        "text_drop_sub": "#64748b",
        "text_drop_active": "#1565c0",
        "btn_primary": "#1976d2",
        "btn_primary_hover": "#1565c0",
        "btn_primary_disabled": "#90caf9",
        "btn_secondary": "#ffffff",
        "btn_secondary_hover": "#f0f7ff",
        "btn_pill": "#e3effc",
        "btn_pill_hover": "#d5e6fa",
        "btn_pill_text": "#1976d2",
        "progress": "#1976d2",
        "progress_track": "#e2ecf8",
        "text_white": "#ffffff",
        "text_disabled": "#94a3b8",
        "text_placeholder": "#94a3b8",
        "field_border": "#cbd7e6",
        "field_bg": "#ffffff",
    }

    # Widget configurations
    configs = {
        "App.TFrame": {"background": colors["bg_main"]},
        "Panel.TFrame": {"background": colors["bg_main"]},
        "Header.TFrame": {"background": colors["bg_main"]},
        "Card.TFrame": {
            "background": colors["bg_card"],
            "bordercolor": colors["border"],
            "borderwidth": 1,
            "relief": "solid",
        },
        "Field.TLabel": {
            "background": colors["bg_card"],
            "foreground": colors["text_label"],
            "font": ui_font(master, 11, "bold"),
        },
        "Hint.TLabel": {
            "background": colors["bg_card"],
            "foreground": colors["text_drop_sub"],
            "font": ui_font(master, 10),
        },
        "StatusBox.TFrame": {
            "background": colors["bg_status"],
            "bordercolor": colors["border"],
            "borderwidth": 1,
            "relief": "solid",
        },
        "Status.TLabel": {
            "background": colors["bg_status"],
            "foreground": colors["text_status"],
            "font": ui_font(master, 13, "bold"),
        },
        "Progress.TLabel": {
            "background": colors["bg_status"],
            "foreground": colors["text_subtitle"],
            "font": ui_font(master, 10),
        },
        "Bdap.Horizontal.TProgressbar": {
            "background": colors["progress"],
            "troughcolor": colors["progress_track"],
            "bordercolor": colors["progress_track"],
            "lightcolor": colors["progress"],
            "darkcolor": colors["progress"],
        },
        "Primary.TButton": {
            "background": colors["btn_primary"],
            "foreground": colors["text_white"],
            "padding": (16, 9),
            "font": ui_font(master, 11, "bold"),
            "borderwidth": 0,
            "relief": "flat",
        },
        "Secondary.TButton": {
            "background": colors["btn_secondary"],
            "foreground": colors["btn_primary"],
            "bordercolor": colors["btn_primary"],
            "padding": (14, 7),
            "font": ui_font(master, 10, "bold"),
            "borderwidth": 1,
            "relief": "solid",
        },
        "Pill.TMenubutton": {
            "background": colors["btn_pill"],
            "foreground": colors["btn_pill_text"],
            "padding": (12, 6),
            "font": ui_font(master, 10, "bold"),
            "borderwidth": 0,
            "relief": "flat",
        },
        "Field.TEntry": {
            "fieldbackground": colors["field_bg"],
            "foreground": colors["text_label"],
            "bordercolor": colors["field_border"],
            "lightcolor": colors["field_border"],
            "darkcolor": colors["field_border"],
            "padding": (8, 6),
        },
        "Field.TCombobox": {
            "fieldbackground": colors["field_bg"],
            "background": colors["field_bg"],
            "foreground": colors["text_label"],
            "bordercolor": colors["field_border"],
            "lightcolor": colors["field_border"],
            "darkcolor": colors["field_border"],
            "arrowcolor": colors["btn_primary"],
            "padding": (8, 6),
        },
        "Title.TLabel": {
            "background": colors["bg_main"],
            "foreground": colors["text_title"],
            "font": ui_font(master, 24, "bold"),
        },
        "Subtitle.TLabel": {
            "background": colors["bg_main"],
            "foreground": colors["text_subtitle"],
            "font": ui_font(master, 11),
        },
        "CardTitle.TLabel": {
            "background": colors["bg_card"],
            "foreground": colors["text_card_label"],
            "font": ui_font(master, 14, "bold"),
        },
        "TRadiobutton": {
            "background": colors["bg_card"],
            "foreground": colors["text_label"],
            "font": ui_font(master, 10, "bold"),
            "focuscolor": colors["bg_card"],
            "padding": (4, 4),
        },
    }

    # Apply configurations
    for style_name, config in configs.items():
        style.configure(style_name, **config)

    # State maps
    style.map(
        "TRadiobutton",
        background=[("active", colors["bg_card"])],
        foreground=[("active", colors["btn_primary"])],
        indicatorcolor=[
            ("selected", colors["btn_primary"]),
            ("pressed", colors["btn_primary_hover"]),
        ],
    )
    style.map(
        "Primary.TButton",
        background=[("active", colors["btn_primary_hover"]), ("disabled", colors["btn_primary_disabled"])],
        foreground=[("disabled", colors["text_white"])],
    )
    style.map(
        "Secondary.TButton",
        background=[("active", colors["btn_secondary_hover"])],
        foreground=[("active", colors["btn_primary_hover"])],
    )
    style.map(
        "Pill.TMenubutton",
        background=[("active", colors["btn_pill_hover"]), ("pressed", colors["btn_pill_hover"])],
        foreground=[("active", colors["btn_pill_text"])],
    )
    style.map(
        "Field.TCombobox",
        fieldbackground=[("readonly", colors["field_bg"]), ("disabled", "#f8fafc")],
        bordercolor=[("focus", colors["btn_primary"])],
    )
    style.map(
        "Field.TEntry",
        bordercolor=[("focus", colors["btn_primary"])],
    )
