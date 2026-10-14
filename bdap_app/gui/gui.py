"""Interfaccia grafica desktop per il flusso di automazione BDAP."""

from __future__ import annotations

import math
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Import opzionale per il supporto drag-and-drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

# Usa TkinterDnD.Tk se disponibile, altrimenti usa Tk standard
BaseTk = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk

# Allow running this file directly by exposing project root.
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from bdap_app.gui import icons
from bdap_app.gui.service import DesktopAutomationService, GuiFormData
from bdap_app.gui.styles import configure_styles, ui_font
from bdap_app.gui import texts as ui_text
from bdap_app.orchestration.cli_discovery import normalize_token


class PlaceholderStringVar(tk.StringVar):
    """StringVar che restituisce stringa vuota se il valore visualizzato e' il placeholder."""

    def __init__(self, placeholder: str = "", default_value: str = "", *args, **kwargs) -> None:
        self.placeholder = placeholder
        initial = default_value if default_value else placeholder
        super().__init__(value=initial, *args, **kwargs)

    def get(self) -> str:
        val = super().get()
        if val == self.placeholder:
            return ""
        return val

    def set(self, value: str) -> None:
        if not value:
            super().set(self.placeholder)
        else:
            super().set(value)


class RoundedCard(tk.Canvas):
    """Contenitore con angoli arrotondati, sfondo bianco e bordo sottile."""

    def __init__(
        self,
        parent,
        bg_color: str = "#f0f5fa",
        card_color: str = "#ffffff",
        border_color: str = "#dbe7f5",
        radius: int = 12,
        border_width: int = 1,
        padding: tuple[int, int] = (22, 16),
        **kwargs,
    ) -> None:
        super().__init__(parent, bg=bg_color, highlightthickness=0, bd=0, **kwargs)
        self.bg_color = bg_color
        self.card_color = card_color
        self.border_color = border_color
        self.radius = radius
        self.border_width = border_width
        self.pad_x, self.pad_y = padding

        self.inner = tk.Frame(self, bg=card_color)
        self._window_id = self.create_window((self.pad_x, self.pad_y), window=self.inner, anchor="nw")

        self.bind("<Configure>", self._on_canvas_configure)
        self.inner.bind("<Configure>", self._on_inner_configure)

    def _draw_rounded_rect(self, x0: int, y0: int, x1: int, y1: int, r: int) -> None:
        self.delete("card_bg")
        points = [
            x0 + r, y0,
            x1 - r, y0,
            x1, y0,
            x1, y0 + r,
            x1, y1 - r,
            x1, y1,
            x1 - r, y1,
            x0 + r, y1,
            x0, y1,
            x0, y1 - r,
            x0, y0 + r,
            x0, y0,
        ]
        self.create_polygon(
            points,
            smooth=True,
            fill=self.card_color,
            outline=self.border_color,
            width=self.border_width,
            tags="card_bg",
        )
        self.tag_lower("card_bg")

    def _on_canvas_configure(self, event) -> None:
        w, h = event.width, event.height
        if w > 2 * self.pad_x and h > 2 * self.pad_y:
            self._draw_rounded_rect(2, 2, w - 2, h - 2, self.radius)
            self.itemconfigure(self._window_id, width=w - 2 * self.pad_x)

    def _on_inner_configure(self, _event=None) -> None:
        req_h = self.inner.winfo_reqheight() + 2 * self.pad_y + 4
        self.configure(height=req_h)


class DropZoneCanvas(tk.Canvas):
    """Area drag & drop con bordo tratteggiato, icona cloud e testo multilinea."""

    def __init__(self, parent, command=None, **kwargs) -> None:
        super().__init__(parent, bg="#ffffff", highlightthickness=0, bd=0, height=86, **kwargs)
        self.command = command
        self.is_active = False
        self.custom_text: str | None = None
        self.cloud_icon = icons.create_cloud_upload(38)
        self.bind("<Configure>", self._draw)
        self.bind("<Button-1>", lambda _: self.command() if self.command else None)

    def set_active(self, active: bool) -> None:
        self.is_active = active
        self._draw()

    def configure(self, cnf=None, **kw):  # type: ignore
        if "text" in kw:
            self.custom_text = kw.pop("text")
            self._draw()
        if kw or cnf:
            super().configure(cnf, **kw)

    def _draw(self, _event=None) -> None:
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 30 or h < 30:
            return

        bg = "#e3effc" if self.is_active else "#f7faff"
        border = "#1976d2" if self.is_active else "#93c5fd"
        r = 10
        points = [
            2 + r, 2,
            w - 3 - r, 2,
            w - 3, 2,
            w - 3, 2 + r,
            w - 3, h - 3 - r,
            w - 3, h - 3,
            w - 3 - r, h - 3,
            2 + r, h - 3,
            2, h - 3,
            2, h - 3 - r,
            2, 2 + r,
            2, 2,
        ]
        self.create_polygon(points, smooth=True, fill=bg, outline=border, width=1.5, dash=(6, 4))

        cx = w // 2
        cy = h // 2
        text1 = self.custom_text or ui_text.DROP_ZONE_TEXT
        text2 = "oppure usa i pulsanti sopra per selezionare un percorso"

        offset_x = 150
        self.create_image(max(30, cx - offset_x), cy, image=self.cloud_icon)
        self.create_text(
            max(60, cx - offset_x + 35),
            cy - 11,
            text=text1,
            font=("Segoe UI", 11, "bold"),
            fill="#1976d2",
            anchor="w",
        )
        self.create_text(
            max(60, cx - offset_x + 35),
            cy + 13,
            text=text2,
            font=("Segoe UI", 9),
            fill="#64748b",
            anchor="w",
        )


class FolderEntry(tk.Frame):
    """Campo di testo stile moderno con icona cartella a sinistra e placeholder."""

    def __init__(
        self,
        master,
        textvariable: tk.StringVar,
        placeholder: str = "Seleziona o inserisci il percorso della cartella...",
        **kwargs,
    ) -> None:
        super().__init__(
            master,
            bg="#ffffff",
            highlightbackground="#cbd7e6",
            highlightthickness=1,
            bd=0,
            **kwargs,
        )
        self.textvariable = textvariable
        self.placeholder = placeholder
        self.folder_icon = icons.create_folder_icon(16, "#94a3b8")

        self.icon_lbl = tk.Label(self, image=self.folder_icon, bg="#ffffff")
        self.icon_lbl.pack(side="left", padx=(10, 4))

        self.entry = tk.Entry(
            self,
            bg="#ffffff",
            fg="#1e293b",
            bd=0,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(4, 10), pady=7)

        self.entry.bind("<FocusIn>", self._on_focus_in)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<KeyRelease>", self._on_key)

        self.textvariable.trace_add("write", self._on_var_changed)
        self._sync_from_var()

    def _on_focus_in(self, _event=None) -> None:
        if not self.textvariable.get():
            self.entry.delete(0, "end")
            self.entry.configure(fg="#1e293b")

    def _on_focus_out(self, _event=None) -> None:
        if not self.textvariable.get():
            self.entry.delete(0, "end")
            self.entry.insert(0, self.placeholder)
            self.entry.configure(fg="#94a3b8")

    def _on_key(self, _event=None) -> None:
        self.textvariable.set(self.entry.get())

    def _on_var_changed(self, *args) -> None:
        val = self.textvariable.get()
        if self.entry.get() != val:
            self.entry.delete(0, "end")
            if val:
                self.entry.insert(0, val)
                self.entry.configure(fg="#1e293b")
            elif self.focus_get() != self.entry:
                self.entry.insert(0, self.placeholder)
                self.entry.configure(fg="#94a3b8")

    def _sync_from_var(self) -> None:
        val = self.textvariable.get()
        self.entry.delete(0, "end")
        if val:
            self.entry.insert(0, val)
            self.entry.configure(fg="#1e293b")
        else:
            self.entry.insert(0, self.placeholder)
            self.entry.configure(fg="#94a3b8")


class BdapGui(BaseTk):
    """Wrapper GUI desktop attorno alla pipeline di automazione core."""

    def __init__(self) -> None:
        super().__init__()
        self.title(ui_text.APP_TITLE)
        self.geometry("1000x830")
        self.minsize(920, 600)
        self.configure(bg="#f0f5fa")

        # UI state
        self.workspace_var = tk.StringVar(value="")
        self.comune_var = PlaceholderStringVar(placeholder="Scegli un comune...")
        self.year_mode_var = tk.StringVar(value="all")
        self.years_info_var = tk.StringVar(value=ui_text.YEARS_ALL_NO_COMUNE)
        self.years_specific_summary_var = tk.StringVar(value=ui_text.YEARS_SPECIFIC_NONE)
        self.years_label_var = tk.StringVar(value=ui_text.YEARS_SPECIFIC_PLACEHOLDER)
        self.status_var = tk.StringVar(value=ui_text.STATUS_SELECT_DATA)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_label_var = tk.StringVar(value="")
        self.available_years: list[int] = []
        self.selected_years: set[int] = set()
        self.year_toggle_vars: dict[int, tk.BooleanVar] = {}
        self.year_from_var = PlaceholderStringVar(placeholder="AAAA")
        self.year_to_var = PlaceholderStringVar(placeholder="AAAA")

        # Internal state
        self.service = DesktopAutomationService(project_root=PROJECT_ROOT)
        self.drag_drop_enabled = TkinterDnD is not None and DND_FILES is not None
        self._scroll_canvas: tk.Canvas | None = None
        self._scroll_window_id: int | None = None
        self._touchpad_scroll_remainder = 0.0

        # Widget refs
        self.header_subtitle_label: tk.Label | ttk.Label | None = None
        self.status_label: tk.Label | ttk.Label | None = None
        self.progress_frame: tk.Frame | ttk.Frame | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.progress_label: tk.Label | ttk.Label | None = None
        self.workspace_entry: FolderEntry | ttk.Entry | None = None
        self.drop_zone: DropZoneCanvas | tk.Label | None = None
        self.comune_combo: ttk.Combobox | None = None
        self.rb_all_years: ttk.Radiobutton | None = None
        self.rb_range_years: ttk.Radiobutton | None = None
        self.rb_specific_years: ttk.Radiobutton | None = None
        self.years_mode_container: tk.Frame | None = None
        self.panel_mode_all: tk.Frame | None = None
        self.panel_mode_range: tk.Frame | None = None
        self.panel_mode_specific: tk.Frame | None = None
        self.years_menu_button: ttk.Menubutton | None = None
        self.years_menu: tk.Menu | None = None
        self.year_from_combo: ttk.Combobox | None = None
        self.year_to_combo: ttk.Combobox | None = None
        self.run_button: ttk.Button | None = None

        self._setup_styles()
        self._build_ui()
        self.bind("<Configure>", self._on_window_configure)
        self._set_years_options([])

    def _setup_styles(self) -> None:
        """Configura il tema ttk e gli stili visivi condivisi."""
        configure_styles(ttk.Style(self))

    def _build_ui(self) -> None:
        """Costruisce il layout completo dell'interfaccia desktop."""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        container = ttk.Frame(self, style="App.TFrame")
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        # Setup scroll container
        canvas = tk.Canvas(container, bg="#f0f5fa", highlightthickness=0, bd=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)
        self._scroll_canvas = canvas

        scroll_content = ttk.Frame(canvas, style="App.TFrame")
        self._scroll_window_id = canvas.create_window((0, 0), window=scroll_content, anchor="nw")
        scroll_content.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", self._on_canvas_configure)
        canvas.bind("<Enter>", lambda _: self._bind_mousewheel(True))
        canvas.bind("<Leave>", lambda _: self._bind_mousewheel(False))

        # Root content frame with comfortable padding
        root = ttk.Frame(scroll_content, padding=(24, 16), style="App.TFrame")
        root.grid(row=0, column=0, sticky="nsew")
        scroll_content.columnconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        # Build UI sections
        self._build_header(root)
        self._build_data_frame(root)
        self._build_params_frame(root)
        self._build_status_box(root)

    def _build_header(self, parent: ttk.Frame) -> None:
        """Costruisce la sezione header con titolo e illustrazione."""
        header = tk.Frame(parent, bg="#f0f5fa")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(1, weight=1)

        # Left Icon (Document + Cog)
        self._img_header_icon = icons.create_header_main_icon(68)
        lbl_icon = tk.Label(header, image=self._img_header_icon, bg="#f0f5fa")
        lbl_icon.grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="w")

        # Titles
        title_box = tk.Frame(header, bg="#f0f5fa")
        title_box.grid(row=0, column=1, sticky="w", pady=(4, 0))

        lbl_title = tk.Label(
            title_box,
            text=ui_text.HEADER_TITLE,
            font=ui_font(self, 24, "bold"),
            fg="#0f2b5c",
            bg="#f0f5fa",
        )
        lbl_title.pack(anchor="w")

        self.header_subtitle_label = tk.Label(
            title_box,
            text=ui_text.HEADER_SUBTITLE,
            font=ui_font(self, 11),
            fg="#5a6e85",
            bg="#f0f5fa",
        )
        self.header_subtitle_label.pack(anchor="w", pady=(3, 0))

        # Right Illustration
        self._img_header_ill = icons.create_header_illustration(190, 105)
        lbl_ill = tk.Label(header, image=self._img_header_ill, bg="#f0f5fa")
        lbl_ill.grid(row=0, column=2, rowspan=2, sticky="e")

    def _build_data_frame(self, parent: ttk.Frame) -> None:
        """Costruisce la card per il caricamento dei dati."""
        card = RoundedCard(parent, padding=(22, 16))
        card.grid(row=1, column=0, sticky="ew", pady=(0, 14))

        # Card Header: Icon + Title
        h_frame = tk.Frame(card.inner, bg="#ffffff")
        h_frame.pack(fill="x", anchor="w", pady=(0, 14))
        self._img_card_data = icons.create_badge_database(30)
        tk.Label(h_frame, image=self._img_card_data, bg="#ffffff").pack(side="left", padx=(0, 10))
        tk.Label(
            h_frame,
            text=ui_text.CARD_LOADING,
            font=ui_font(self, 14, "bold"),
            fg="#0f2b5c",
            bg="#ffffff",
        ).pack(side="left")

        # Field Label: Workspace root
        lbl_ws = tk.Label(
            card.inner,
            text=ui_text.FIELD_WORKSPACE_ROOT,
            font=ui_font(self, 11, "bold"),
            fg="#1e293b",
            bg="#ffffff",
        )
        lbl_ws.pack(anchor="w", pady=(0, 6))

        # Row: Entry + Sfoglia cartella + Aggiorna comuni
        row = tk.Frame(card.inner, bg="#ffffff")
        row.pack(fill="x", pady=(0, 12))

        self.workspace_entry = FolderEntry(
            row,
            textvariable=self.workspace_var,
            placeholder="Seleziona o inserisci il percorso della cartella...",
        )
        self.workspace_entry.pack(side="left", fill="x", expand=True, padx=(0, 12))

        self._img_btn_folder = icons.create_folder_icon(16, "#1976d2")
        btn_browse = ttk.Button(
            row,
            text=ui_text.BUTTON_BROWSE_FOLDER,
            image=self._img_btn_folder,
            compound="left",
            command=self.select_workspace,
            style="Secondary.TButton",
        )
        btn_browse.pack(side="left", padx=(0, 8))

        self._img_btn_refresh = icons.create_refresh_icon(16, "#ffffff")
        btn_refresh = ttk.Button(
            row,
            text=ui_text.BUTTON_REFRESH_COMUNI,
            image=self._img_btn_refresh,
            compound="left",
            command=self.refresh_comuni,
            style="Primary.TButton",
        )
        btn_refresh.pack(side="left")

        # Drop Zone
        self.drop_zone = DropZoneCanvas(card.inner, command=self.select_workspace)
        self.drop_zone.pack(fill="x", pady=(2, 0))

        if self.drag_drop_enabled:
            self._setup_drop_zone()
        else:
            self.drop_zone.configure(text=ui_text.DROP_ZONE_TEXT_NO_DND)

    def _build_params_frame(self, parent: ttk.Frame) -> None:
        """Costruisce la card dei parametri."""
        card = RoundedCard(parent, padding=(22, 16))
        card.grid(row=2, column=0, sticky="ew", pady=(0, 14))

        # Card Header: Icon + Title
        h_frame = tk.Frame(card.inner, bg="#ffffff")
        h_frame.pack(fill="x", anchor="w", pady=(0, 14))
        self._img_card_params = icons.create_badge_sliders(30)
        tk.Label(h_frame, image=self._img_card_params, bg="#ffffff").pack(side="left", padx=(0, 10))
        tk.Label(
            h_frame,
            text=ui_text.CARD_PARAMETERS,
            font=ui_font(self, 14, "bold"),
            fg="#0f2b5c",
            bg="#ffffff",
        ).pack(side="left")

        # Label: Seleziona il comune *
        lbl_comune_box = tk.Frame(card.inner, bg="#ffffff")
        lbl_comune_box.pack(anchor="w", pady=(0, 6))
        tk.Label(
            lbl_comune_box,
            text="Seleziona il comune ",
            font=ui_font(self, 11, "bold"),
            fg="#1e293b",
            bg="#ffffff",
        ).pack(side="left")
        tk.Label(
            lbl_comune_box,
            text="*",
            font=ui_font(self, 11, "bold"),
            fg="#ef4444",
            bg="#ffffff",
        ).pack(side="left")

        # Comune Combobox
        self.comune_combo = ttk.Combobox(
            card.inner,
            textvariable=self.comune_var,
            state="readonly",
            style="Field.TCombobox",
        )
        self.comune_combo.pack(fill="x", pady=(0, 14), ipady=2)
        self.comune_combo.bind("<<ComboboxSelected>>", self._on_comune_selected)
        self.comune_combo.bind("<Button-1>", self._on_comune_open_attempt, add="+")
        self.comune_combo.bind("<Down>", self._on_comune_open_attempt, add="+")
        self.comune_combo.bind("<space>", self._on_comune_open_attempt, add="+")

        # Seleziona anni row
        years_label_box = tk.Frame(card.inner, bg="#ffffff")
        years_label_box.pack(anchor="w", pady=(0, 6))
        self._img_cal_icon = icons.create_calendar_icon(16, "#1976d2")
        tk.Label(years_label_box, image=self._img_cal_icon, bg="#ffffff").pack(side="left", padx=(0, 8))
        tk.Label(
            years_label_box,
            text=ui_text.FIELD_YEARS,
            font=ui_font(self, 11, "bold"),
            fg="#1e293b",
            bg="#ffffff",
        ).pack(side="left")

        # Radio row per la selezione della modalita
        radio_box = tk.Frame(card.inner, bg="#ffffff")
        radio_box.pack(anchor="w", pady=(0, 8))

        self.rb_all_years = ttk.Radiobutton(
            radio_box,
            text=ui_text.MODE_YEARS_ALL,
            value="all",
            variable=self.year_mode_var,
            command=self._on_year_mode_changed,
        )
        self.rb_all_years.pack(side="left", padx=(0, 24))

        self.rb_range_years = ttk.Radiobutton(
            radio_box,
            text=ui_text.MODE_YEARS_RANGE,
            value="range",
            variable=self.year_mode_var,
            command=self._on_year_mode_changed,
        )
        self.rb_range_years.pack(side="left", padx=(0, 24))

        self.rb_specific_years = ttk.Radiobutton(
            radio_box,
            text=ui_text.MODE_YEARS_SPECIFIC,
            value="specific",
            variable=self.year_mode_var,
            command=self._on_year_mode_changed,
        )
        self.rb_specific_years.pack(side="left")

        # Container dinamico per la modalita attiva
        self.years_mode_container = tk.Frame(
            card.inner,
            bg="#f8fafc",
            highlightbackground="#dbe7f5",
            highlightthickness=1,
            bd=0,
        )
        self.years_mode_container.pack(fill="x", pady=(0, 16))

        # --- Panel ALL: Tutti i presenti ---
        self.panel_mode_all = tk.Frame(self.years_mode_container, bg="#f8fafc", padx=16, pady=12)
        row_all = tk.Frame(self.panel_mode_all, bg="#f8fafc")
        row_all.pack(fill="x", anchor="w")
        self._img_check_circle = icons.create_check_circle_icon(22, "#16a34a")
        tk.Label(row_all, image=self._img_check_circle, bg="#f8fafc").pack(side="left", padx=(0, 10))

        all_text_frame = tk.Frame(row_all, bg="#f8fafc")
        all_text_frame.pack(side="left", fill="x", expand=True)

        tk.Label(
            all_text_frame,
            text=ui_text.YEARS_ALL_INFO_DEFAULT,
            font=ui_font(self, 10, "bold"),
            fg="#0f2b5c",
            bg="#f8fafc",
        ).pack(anchor="w")

        self.lbl_years_detected = tk.Label(
            all_text_frame,
            textvariable=self.years_info_var,
            font=ui_font(self, 10),
            fg="#5a6e85",
            bg="#f8fafc",
        )
        self.lbl_years_detected.pack(anchor="w", pady=(2, 0))

        # --- Panel RANGE: Intervallo di anni ---
        self.panel_mode_range = tk.Frame(self.years_mode_container, bg="#f8fafc", padx=16, pady=12)

        tk.Label(
            self.panel_mode_range,
            text=ui_text.YEARS_RANGE_SUBTITLE,
            font=ui_font(self, 10),
            fg="#64748b",
            bg="#f8fafc",
        ).pack(anchor="w", pady=(0, 8))

        range_inputs = tk.Frame(self.panel_mode_range, bg="#f8fafc")
        range_inputs.pack(fill="x")

        # Anno da
        col1 = tk.Frame(range_inputs, bg="#f8fafc")
        col1.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(
            col1,
            text=ui_text.FIELD_YEAR_FROM,
            font=ui_font(self, 10, "bold"),
            fg="#1e293b",
            bg="#f8fafc",
        ).pack(anchor="w", pady=(0, 4))
        self.year_from_combo = ttk.Combobox(
            col1,
            textvariable=self.year_from_var,
            state="disabled",
            style="Field.TCombobox",
        )
        self.year_from_combo.pack(fill="x", ipady=2)

        # Anno a
        col2 = tk.Frame(range_inputs, bg="#f8fafc")
        col2.pack(side="left", fill="x", expand=True)
        tk.Label(
            col2,
            text=ui_text.FIELD_YEAR_TO,
            font=ui_font(self, 10, "bold"),
            fg="#1e293b",
            bg="#f8fafc",
        ).pack(anchor="w", pady=(0, 4))
        self.year_to_combo = ttk.Combobox(
            col2,
            textvariable=self.year_to_var,
            state="disabled",
            style="Field.TCombobox",
        )
        self.year_to_combo.pack(fill="x", ipady=2)

        # --- Panel SPECIFIC: Anni specifici ---
        self.panel_mode_specific = tk.Frame(self.years_mode_container, bg="#f8fafc", padx=16, pady=12)

        tk.Label(
            self.panel_mode_specific,
            text=ui_text.YEARS_SPECIFIC_SUBTITLE,
            font=ui_font(self, 10),
            fg="#64748b",
            bg="#f8fafc",
        ).pack(anchor="w", pady=(0, 8))

        specific_row = tk.Frame(self.panel_mode_specific, bg="#f8fafc")
        specific_row.pack(fill="x", anchor="w")

        self._img_pill_cal = icons.create_calendar_icon(14, "#1976d2")
        self.years_menu_button = ttk.Menubutton(
            specific_row,
            textvariable=self.years_label_var,
            image=self._img_pill_cal,
            compound="left",
            style="Pill.TMenubutton",
        )
        self.years_menu_button.pack(side="left", padx=(0, 12))
        self.years_menu = tk.Menu(self.years_menu_button, tearoff=False)
        self.years_menu_button["menu"] = self.years_menu

        self.lbl_specific_summary = tk.Label(
            specific_row,
            textvariable=self.years_specific_summary_var,
            font=ui_font(self, 10),
            fg="#5a6e85",
            bg="#f8fafc",
        )
        self.lbl_specific_summary.pack(side="left")

        # Mostra la modalita iniziale (predefinita "all")
        self._on_year_mode_changed()

        # Genera Report Excel Button
        self._img_excel = icons.create_excel_icon(22)
        self.run_button = ttk.Button(
            card.inner,
            text=ui_text.BUTTON_RUN,
            image=self._img_excel,
            compound="left",
            command=self.run_automation,
            style="Primary.TButton",
        )
        self.run_button.pack(anchor="w", pady=(4, 0))

    def _build_status_box(self, parent: ttk.Frame) -> None:
        """Costruisce la card di stato / accordion in basso."""
        card = RoundedCard(parent, padding=(20, 14))
        card.grid(row=3, column=0, sticky="ew", pady=(0, 14))

        # Status header row
        h_frame = tk.Frame(card.inner, bg="#ffffff")
        h_frame.pack(fill="x")

        self._img_card_status = icons.create_badge_database(28)
        tk.Label(h_frame, image=self._img_card_status, bg="#ffffff").pack(side="left", padx=(0, 10))

        self.status_label = tk.Label(
            h_frame,
            textvariable=self.status_var,
            font=ui_font(self, 13, "bold"),
            fg="#0f2b5c",
            bg="#ffffff",
        )
        self.status_label.pack(side="left")

        self._img_chevron = icons.create_chevron_down(16, "#1976d2")
        self.chevron_lbl = tk.Label(h_frame, image=self._img_chevron, bg="#ffffff", cursor="hand2")
        self.chevron_lbl.pack(side="right", padx=(0, 4))

        # Progress elements (inside Card 3, revealed when running)
        self.progress_frame = tk.Frame(card.inner, bg="#ffffff")
        self.progress_frame.pack(fill="x", pady=(10, 0))

        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            style="Bdap.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", pady=(0, 4))

        self.progress_label = tk.Label(
            self.progress_frame,
            textvariable=self.progress_label_var,
            font=ui_font(self, 10),
            fg="#5a6e85",
            bg="#ffffff",
        )
        self.progress_label.pack(anchor="w")

        self._hide_progress()

    def _on_text_wraplength(self, event) -> None:
        """Adatta la lunghezza di wrapping del testo alla larghezza del widget al ridimensionamento."""
        widget = event.widget
        if isinstance(widget, (ttk.Label, tk.Label)):
            widget.configure(wraplength=max(180, event.width - 16))

    def _on_canvas_configure(self, event) -> None:
        """Mantiene la larghezza del contenuto scrollabile sincronizzata con la larghezza del canvas."""
        if self._scroll_window_id is None or self._scroll_canvas is None:
            return
        self._scroll_canvas.itemconfigure(self._scroll_window_id, width=event.width)

    def _on_window_configure(self, _event) -> None:
        """Mantiene il testo lungo leggibile al ridimensionamento della finestra."""
        if self.header_subtitle_label:
            self.header_subtitle_label.configure(wraplength=max(500, self.winfo_width() - 320))
        if self.status_label:
            self.status_label.configure(wraplength=max(500, self.winfo_width() - 140))

    def _bind_mousewheel(self, bind: bool) -> None:
        """Associa/dissocia la rotella del mouse per lo scrolling."""
        if bind:
            self.bind_all("<MouseWheel>", self._on_mousewheel)
            self.bind_all("<Button-4>", self._on_mousewheel)
            self.bind_all("<Button-5>", self._on_mousewheel)
        else:
            self.unbind_all("<MouseWheel>")
            self.unbind_all("<Button-4>")
            self.unbind_all("<Button-5>")

    def _on_mousewheel(self, event) -> None:
        """Gestisce lo scrolling con la rotella su più piattaforme."""
        if not self._scroll_canvas:
            return

        if (step := getattr(event, "num", None)) in (4, 5):
            self._scroll_canvas.yview_scroll(-1 if step == 4 else 1, "units")
            return

        if not event.delta:
            return

        if sys.platform == "darwin":
            total = self._touchpad_scroll_remainder - float(event.delta) / 6.0
            step = int(math.floor(total) if total >= 0 else math.ceil(total))
            self._touchpad_scroll_remainder = total - step
            if step == 0:
                return
        else:
            step = int(-event.delta / 120) or (-1 if event.delta > 0 else 1)

        self._scroll_canvas.yview_scroll(step, "units")

    def _setup_drop_zone(self) -> None:
        """Abilita drag-and-drop di cartelle tramite tkdnd."""
        if not self.drag_drop_enabled or self.drop_zone is None:
            return
        assert DND_FILES is not None
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<DropEnter>>", self._on_drop_enter)
        self.drop_zone.dnd_bind("<<DropLeave>>", self._on_drop_leave)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop_enter(self, _) -> str:
        """Gestisce l'evento di ingresso drag."""
        if isinstance(self.drop_zone, DropZoneCanvas):
            self.drop_zone.set_active(True)
        return "copy"

    def _on_drop_leave(self, _) -> str:
        """Gestisce l'evento di uscita drag."""
        if isinstance(self.drop_zone, DropZoneCanvas):
            self.drop_zone.set_active(False)
        return "copy"

    def _extract_drop_path(self, raw_data: str) -> Path | None:
        """Analizza il payload tkdnd e restituisce la prima directory valida."""
        if not raw_data:
            return None
        for token in self.tk.splitlist(raw_data):
            candidate = Path(token.strip("{}")).expanduser()
            if candidate.exists() and candidate.is_dir():
                return candidate
        return None

    def _on_drop(self, event) -> str:
        """Gestisce la cartella droppata."""
        if isinstance(self.drop_zone, DropZoneCanvas):
            self.drop_zone.set_active(False)
        dropped = self._extract_drop_path(getattr(event, "data", ""))
        if dropped is None:
            self.status_var.set(ui_text.STATUS_DROP_INVALID)
            return "copy"
        self.workspace_var.set(str(dropped))
        self.refresh_comuni()
        return "copy"

    def set_status_threadsafe(self, text: str) -> None:
        """Pianifica l'aggiornamento di stato sul thread principale."""
        self.after(0, lambda: self.status_var.set(text))

    def _show_progress(self) -> None:
        """Mostra la barra di avanzamento."""
        if self.progress_frame is not None:
            self.progress_frame.pack(fill="x", pady=(10, 0))

    def _hide_progress(self) -> None:
        """Nasconde e azzera la barra di avanzamento."""
        self.progress_var.set(0.0)
        self.progress_label_var.set("")
        if self.progress_frame is not None:
            self.progress_frame.pack_forget()

    def set_progress(self, completed: float, total: int, message: str) -> None:
        """Aggiorna la progress bar con valori normalizzati a percentuale."""
        safe_total = max(1, int(total))
        percent = min(max(float(completed) / safe_total * 100.0, 0.0), 100.0)
        self._show_progress()
        self.progress_var.set(percent)
        self.progress_label_var.set(
            ui_text.PROGRESS_LABEL_TEMPLATE.format(percent=percent, message=message)
        )
        self.status_var.set(ui_text.STATUS_RUNNING)

    def set_progress_threadsafe(self, completed: float, total: int, message: str) -> None:
        """Pianifica l'aggiornamento della progress bar sul thread principale."""
        self.after(0, lambda: self.set_progress(completed, total, message))

    def set_running_state(self, running: bool) -> None:
        """Abilita/disabilita i controlli durante l'esecuzione."""
        state = "disabled" if running else "normal"
        if self.run_button:
            self.run_button.configure(state=state)
        for rb in [self.rb_all_years, self.rb_range_years, self.rb_specific_years]:
            if rb:
                rb.configure(state=state)

    def set_running_state_threadsafe(self, running: bool) -> None:
        """Pianifica l'aggiornamento dello stato di esecuzione sul thread principale."""
        self.after(0, lambda: self.set_running_state(running))

    def select_workspace(self) -> None:
        """Permette all'utente di scegliere la directory workspace."""
        selected = filedialog.askdirectory(initialdir=self.workspace_var.get() or str(Path.cwd()))
        if selected:
            self.workspace_var.set(selected)
            self.refresh_comuni()

    def refresh_comuni(self) -> None:
        """Scansiona lo workspace e aggiorna le opzioni dei comuni."""
        raw_path = self.workspace_var.get().strip()
        if not raw_path:
            self._reset_comuni_state()
            self.status_var.set(ui_text.STATUS_SELECT_DATA)
            return

        root = Path(raw_path).expanduser()
        if not root.exists():
            self._reset_comuni_state()
            self.status_var.set(ui_text.STATUS_WORKSPACE_NOT_FOUND)
            return

        comuni = self.service.list_comuni(root)
        if self.comune_combo:
            self.comune_combo["values"] = comuni
        self._set_comune_state(comuni, root)
        self._refresh_years_for_selected_comune()
        self.status_var.set(ui_text.STATUS_READY if comuni else ui_text.NO_COMUNI_FOUND_MESSAGE)

    def _reset_comuni_state(self) -> None:
        """Reimposta il selettore comune allo stato vuoto."""
        if self.comune_combo:
            self.comune_combo["values"] = []
            self.comune_combo.configure(state="disabled")
        self.comune_var.set("")
        self._set_years_options([])

    def _set_comune_state(self, comuni: list[str], workspace_root: Path | None = None) -> None:
        """Imposta lo stato del selettore comune in base ai comuni scoperti."""
        if not self.comune_combo:
            return

        if not comuni:
            self.comune_combo.configure(state="disabled")
            self.comune_var.set("")
            return

        if len(comuni) == 1:
            self.comune_var.set(comuni[0])
            self.comune_combo.configure(state="disabled")
            return

        if workspace_root is not None:
            workspace_name = workspace_root.name
            normalized_workspace = normalize_token(workspace_name)
            for comune in comuni:
                if normalize_token(comune) == normalized_workspace:
                    self.comune_var.set(comune)
                    self.comune_combo.configure(state="readonly")
                    return

        if self.comune_var.get() not in comuni:
            self.comune_var.set("")
        self.comune_combo.configure(state="readonly")

    def _on_year_mode_changed(self) -> None:
        """Cambia la visualizzazione in base alla modalita di selezione anni."""
        if not self.panel_mode_all or not self.panel_mode_range or not self.panel_mode_specific:
            return

        mode = self.year_mode_var.get()
        self.panel_mode_all.pack_forget()
        self.panel_mode_range.pack_forget()
        self.panel_mode_specific.pack_forget()

        if mode == "all":
            self.panel_mode_all.pack(fill="x")
        elif mode == "range":
            self.panel_mode_range.pack(fill="x")
            if self.available_years:
                if not self.year_from_var.get():
                    self.year_from_var.set(str(self.available_years[0]))
                if not self.year_to_var.get():
                    self.year_to_var.set(str(self.available_years[-1]))
        elif mode == "specific":
            self.panel_mode_specific.pack(fill="x")

    def _set_years_options(self, years: list[int]) -> None:
        """Ricostruisce le opzioni degli anni a partire dagli anni scoperti."""
        sorted_years = sorted(set(int(y) for y in years)) if years else []
        self.available_years = sorted_years

        # Aggiorna il testo informativo per la modalita "tutti"
        if sorted_years:
            years_str = ", ".join(str(y) for y in sorted_years)
            self.years_info_var.set(ui_text.YEARS_ALL_DETECTED_TEMPLATE.format(years=years_str))
        else:
            self.years_info_var.set(ui_text.YEARS_ALL_NO_COMUNE)

        # Ricostruisce il menu degli anni specifici
        if self.years_menu:
            self.years_menu.delete(0, "end")
            self.year_toggle_vars = {}

            if sorted_years:
                for year in sorted_years:
                    var = tk.BooleanVar(value=False)
                    self.year_toggle_vars[year] = var
                    self.years_menu.add_checkbutton(
                        label=str(year),
                        variable=var,
                        command=lambda y=year: self._toggle_year(y),
                    )

        self._update_year_range_combos(sorted_years)
        self._select_all_years()
        # Di base e' sempre selezionato "tutti i presenti"
        self.year_mode_var.set("all")
        self._on_year_mode_changed()

    def _update_year_range_combos(self, years: list[int]) -> None:
        """Aggiorna i selettori per l'intervallo di anni."""
        range_values = [str(y) for y in years]
        combo_state = "readonly" if years else "disabled"
        for combo in [self.year_from_combo, self.year_to_combo]:
            if combo is not None:
                combo["values"] = range_values
                combo.configure(state=combo_state)
        if years:
            self.year_from_var.set(str(years[0]))
            self.year_to_var.set(str(years[-1]))
        else:
            self.year_from_var.set("")
            self.year_to_var.set("")

    def _select_all_years(self) -> None:
        """Reimposta la selezione specifica su nessun anno selezionato."""
        self.selected_years.clear()
        for var in self.year_toggle_vars.values():
            var.set(False)
        self.years_label_var.set(ui_text.YEARS_SPECIFIC_PLACEHOLDER)
        self.years_specific_summary_var.set(ui_text.YEARS_SPECIFIC_NONE)

    def _toggle_year(self, year: int) -> None:
        """Attiva/disattiva la selezione di un anno specifico e aggiorna l'etichetta."""
        var = self.year_toggle_vars.get(year)
        if not var:
            return

        if var.get():
            self.selected_years.add(year)
        else:
            self.selected_years.discard(year)

        if not self.selected_years:
            self.years_label_var.set(ui_text.YEARS_SPECIFIC_PLACEHOLDER)
            self.years_specific_summary_var.set(ui_text.YEARS_SPECIFIC_NONE)
            return

        years_str = ", ".join(str(y) for y in sorted(self.selected_years))
        self.years_label_var.set(years_str)
        self.years_specific_summary_var.set(ui_text.YEARS_SPECIFIC_SUMMARY_TEMPLATE.format(years=years_str))

    def _refresh_years_for_selected_comune(self) -> None:
        """Aggiorna gli anni in base al comune selezionato."""
        comune = self.comune_var.get().strip()
        raw_path = self.workspace_var.get().strip()
        workspace_root = Path(raw_path).expanduser() if raw_path else Path("")
        years = self._discover_safe(workspace_root, comune or None)
        self._set_years_options(years)

    def _discover_safe(self, workspace_root: Path, comune: str | None = None) -> list[int]:
        """Scopre gli anni in modo sicuro con fallback al livello workspace."""
        try:
            return self.service.discover_years(workspace_root, comune)
        except Exception:
            try:
                return self.service.discover_years(workspace_root)
            except Exception:
                return []

    def _has_discovered_comuni(self) -> bool:
        """Controlla se la combobox dei comuni ha opzioni."""
        if not self.comune_combo:
            return False
        values = self.comune_combo.cget("values")
        if isinstance(values, str):
            values = self.tk.splitlist(values) if values else ()
        return bool(values)

    def _on_comune_open_attempt(self, _event=None) -> str | None:
        """Impedisce l'apertura di un dropdown vuoto."""
        if self._has_discovered_comuni():
            return None
        self.status_var.set(ui_text.STATUS_LOAD_COMUNI_FIRST)
        self.bell()
        return "break"

    def _on_comune_selected(self, _event=None) -> None:
        """Aggiorna gli anni quando cambia il comune selezionato."""
        self._refresh_years_for_selected_comune()

    def build_args(self):
        """Costruisce gli argomenti della pipeline dallo stato dell'interfaccia."""
        raw_ws = self.workspace_var.get().strip()
        workspace_root = Path(raw_ws).expanduser() if raw_ws else Path("")
        comune = self.comune_var.get().strip()

        mode = self.year_mode_var.get()
        if mode == "all":
            form = GuiFormData(
                workspace_root=workspace_root,
                comune=comune,
                explicit_years=(),
                year_from="",
                year_to="",
                available_years=tuple(self.available_years),
            )
        elif mode == "range":
            yf = self.year_from_var.get().strip()
            yt = self.year_to_var.get().strip()
            if not yf or not yt:
                raise ValueError("Specifica sia 'Anno da' che 'Anno a' per l'intervallo.")
            form = GuiFormData(
                workspace_root=workspace_root,
                comune=comune,
                explicit_years=(),
                year_from=yf,
                year_to=yt,
                available_years=tuple(self.available_years),
            )
        else:  # specific
            if not self.selected_years:
                raise ValueError("Seleziona almeno un anno specifico dal menu oppure scegli 'Tutti i presenti'.")
            form = GuiFormData(
                workspace_root=workspace_root,
                comune=comune,
                explicit_years=tuple(sorted(self.selected_years)),
                year_from="",
                year_to="",
                available_years=tuple(self.available_years),
            )
        return self.service.build_args_from_form(form)

    def run_automation(self) -> None:
        """Esegue la pipeline in un thread di background."""
        try:
            args = self.build_args()
        except Exception as exc:
            messagebox.showerror(ui_text.ERROR_TITLE_INPUT, str(exc))
            return

        self.set_running_state(True)
        self.status_var.set(ui_text.STATUS_RUNNING)
        self.set_progress(0, 1, ui_text.PROGRESS_PREPARING)

        def worker() -> None:
            try:
                self.service.run_pipeline(
                    args,
                    log=lambda _: None,
                    progress=self.set_progress_threadsafe,
                )
                report = getattr(args, "controlli_post_report", None)
                self.after(0, lambda report=report: self._on_run_success(args.comune, report))
            except Exception as exc:
                self.after(0, lambda msg=str(exc): self._on_run_error(msg))
            finally:
                self.set_running_state_threadsafe(False)

        threading.Thread(target=worker, daemon=True).start()

    def _on_run_success(self, comune: str, controlli_report=None) -> None:
        """Mostra il completamento e ripristina lo stato della GUI."""
        self.set_progress(1, 1, ui_text.PROGRESS_DONE)
        self.status_var.set(ui_text.STATUS_DONE)
        message = ui_text.DONE_MESSAGE_TEMPLATE.format(comune=comune)
        if getattr(controlli_report, "has_files", False):
            from bdap_app.support.controlli_post import format_controlli_post_summary

            message = ui_text.DONE_MESSAGE_WITH_CONTROLLI_TEMPLATE.format(
                comune=comune,
                controlli_summary=format_controlli_post_summary(controlli_report),
            )
        messagebox.showinfo(
            ui_text.DONE_TITLE,
            message,
        )
        self._hide_progress()
        self.status_var.set(ui_text.STATUS_READY)

    def _on_run_error(self, message: str) -> None:
        """Mostra l'errore e ripristina lo stato della GUI."""
        self.status_var.set(ui_text.STATUS_ERROR)
        messagebox.showerror(ui_text.ERROR_TITLE_GENERIC, message)
        self._hide_progress()
        self.status_var.set(ui_text.STATUS_READY)


def main() -> None:
    """Start the desktop GUI."""
    app = BdapGui()
    app.mainloop()


if __name__ == "__main__":
    main()
