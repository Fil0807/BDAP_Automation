"""Interfaccia grafica desktop per il flusso di automazione BDAP."""

from __future__ import annotations
import math 
import threading
import sys
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

from bdap_app.gui.service import DesktopAutomationService, GuiFormData
from bdap_app.gui.styles import configure_styles, ui_font
from bdap_app.gui import texts as ui_text
from bdap_app.orchestration.cli_discovery import normalize_token


class BdapGui(BaseTk):
    """Wrapper GUI desktop attorno alla pipeline di automazione core."""

    def __init__(self) -> None:
        super().__init__()
        self.title(ui_text.APP_TITLE)
        self.geometry(ui_text.WINDOW_GEOMETRY)
        self.minsize(900, 560)
        self.configure(bg="#eef5f0")

        # UI state
        self.workspace_var = tk.StringVar(value="")
        self.comune_var = tk.StringVar()
        self.years_label_var = tk.StringVar(value=ui_text.ALL_YEARS_LABEL)
        self.status_var = tk.StringVar(value=ui_text.STATUS_SELECT_DATA)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_label_var = tk.StringVar(value="")
        self.available_years: list[int] = []
        self.selected_years: set[int] = set()
        self.year_toggle_vars: dict[int, tk.BooleanVar] = {}
        self.year_from_var = tk.StringVar(value="")
        self.year_to_var = tk.StringVar(value="")

        # Internal state
        self.service = DesktopAutomationService(project_root=PROJECT_ROOT)
        self.drag_drop_enabled = TkinterDnD is not None and DND_FILES is not None
        self._scroll_canvas: tk.Canvas | None = None
        self._scroll_window_id: int | None = None
        self._touchpad_scroll_remainder = 0.0

        # Widget refs
        self.header_subtitle_label: ttk.Label | None = None
        self.status_label: ttk.Label | None = None
        self.progress_frame: ttk.Frame | None = None
        self.progress_bar: ttk.Progressbar | None = None
        self.progress_label: ttk.Label | None = None
        self.workspace_entry: ttk.Entry | None = None
        self.drop_zone: ttk.Label | None = None
        self.comune_combo: ttk.Combobox | None = None
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
        canvas = tk.Canvas(container, bg="#eef5f0", highlightthickness=0, bd=0)
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

        # Root content frame
        root = ttk.Frame(scroll_content, padding=18, style="App.TFrame")
        root.grid(row=0, column=0, sticky="nsew")
        scroll_content.columnconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        panel = ttk.Frame(root, style="Panel.TFrame", padding=20)
        panel.grid(row=0, column=0, sticky="ew")
        panel.columnconfigure(0, weight=1)

        # Build UI sections
        self._build_header(panel)
        self._build_data_frame(panel)
        self._build_params_frame(panel)
        self._build_status_box(panel)

    def _build_header(self, parent: ttk.Frame) -> None:
        """Costruisce la sezione header."""
        header = ttk.Frame(parent, style="Header.TFrame", padding=(4, 4, 4, 12))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=ui_text.HEADER_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.header_subtitle_label = ttk.Label(header, text=ui_text.HEADER_SUBTITLE, style="Subtitle.TLabel")
        self.header_subtitle_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.header_subtitle_label.bind("<Configure>", self._on_text_wraplength)

    def _build_data_frame(self, parent: ttk.Frame) -> None:
        """Costruisce il pannello per il caricamento dei dati."""
        frame = ttk.LabelFrame(parent, text=ui_text.CARD_LOADING, style="Card.TLabelframe", padding=14)
        frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text=ui_text.FIELD_WORKSPACE_ROOT, style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.workspace_entry = ttk.Entry(frame, textvariable=self.workspace_var, style="Field.TEntry")
        self.workspace_entry.grid(row=0, column=1, sticky="ew", pady=4, ipady=3)

        actions = ttk.Frame(frame)
        actions.grid(row=0, column=2, sticky="e", padx=(8, 0), pady=4)
        ttk.Button(actions, text=ui_text.BUTTON_BROWSE_FOLDER, command=self.select_workspace, style="Secondary.TButton").grid(row=0, column=0, padx=(0, 6))
        ttk.Button(actions, text=ui_text.BUTTON_REFRESH_COMUNI, command=self.refresh_comuni, style="Secondary.TButton").grid(row=0, column=1)

        self.drop_zone = tk.Label(frame, text=ui_text.DROP_ZONE_TEXT, bg="#f7fcf9", fg="#4f6b5c", font=ui_font(self, 12, "bold"), anchor="center", justify="center", relief="solid", borderwidth=2, bd=2, padx=12, pady=10)
        self.drop_zone.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0), ipady=10)
        self.drop_zone.bind("<Configure>", self._on_text_wraplength)
        if self.drag_drop_enabled:
            self._setup_drop_zone()
        else:
            self.drop_zone.configure(text=ui_text.DROP_ZONE_TEXT_NO_DND)

    def _build_params_frame(self, parent: ttk.Frame) -> None:
        """Costruisce il pannello dei parametri."""
        frame = ttk.LabelFrame(parent, text=ui_text.CARD_PARAMETERS, style="Card.TLabelframe", padding=14)
        frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        frame.columnconfigure(1, weight=1)

        # Comune selector
        ttk.Label(frame, text=ui_text.FIELD_COMUNE, style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.comune_combo = ttk.Combobox(frame, textvariable=self.comune_var, state="readonly", style="Field.TCombobox")
        self.comune_combo.grid(row=0, column=1, sticky="ew", pady=4, ipady=3)
        self.comune_combo.bind("<<ComboboxSelected>>", self._on_comune_selected)
        self.comune_combo.bind("<Button-1>", self._on_comune_open_attempt, add="+")
        self.comune_combo.bind("<Down>", self._on_comune_open_attempt, add="+")
        self.comune_combo.bind("<space>", self._on_comune_open_attempt, add="+")

        # Years selector
        ttk.Label(frame, text=ui_text.FIELD_YEARS, style="Field.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        years_frame = ttk.Frame(frame, style="App.TFrame")
        years_frame.grid(row=1, column=1, sticky="w", pady=4)
        self.years_menu_button = ttk.Menubutton(years_frame, textvariable=self.years_label_var, style="Secondary.TButton")
        self.years_menu_button.grid(row=0, column=0, sticky="w")
        self.years_menu = tk.Menu(self.years_menu_button, tearoff=False)
        self.years_menu_button["menu"] = self.years_menu

        # Year range
        ttk.Label(frame, text=ui_text.HINT_YEAR_RANGE, style="Hint.TLabel").grid(row=2, column=1, sticky="w", pady=(2, 0))
        year_range = ttk.Frame(frame, style="App.TFrame")
        year_range.grid(row=3, column=1, sticky="w", pady=(4, 0))

        ttk.Label(year_range, text=ui_text.FIELD_YEAR_FROM, style="Field.TLabel").grid(row=0, column=0, sticky="w")
        self.year_from_combo = ttk.Combobox(year_range, textvariable=self.year_from_var, state="disabled", width=10, style="Field.TCombobox")
        self.year_from_combo.grid(row=1, column=0, sticky="w", padx=(0, 8), ipady=2)

        ttk.Label(year_range, text=ui_text.FIELD_YEAR_TO, style="Field.TLabel").grid(row=0, column=1, sticky="w")
        self.year_to_combo = ttk.Combobox(year_range, textvariable=self.year_to_var, state="disabled", width=10, style="Field.TCombobox")
        self.year_to_combo.grid(row=1, column=1, sticky="w", ipady=2)

        # Run button
        self.run_button = ttk.Button(frame, text=ui_text.BUTTON_RUN, command=self.run_automation, style="Primary.TButton")
        self.run_button.grid(row=4, column=1, sticky="w", pady=(10, 0))

    def _build_status_box(self, parent: ttk.Frame) -> None:
        """Costruisce la casella di stato."""
        box = ttk.Frame(parent, style="StatusBox.TFrame", padding=(12, 10))
        box.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        box.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(box, textvariable=self.status_var, anchor="w", justify="left", style="Status.TLabel")
        self.status_label.grid(row=0, column=0, sticky="ew")
        self.status_label.bind("<Configure>", self._on_text_wraplength)

        self.progress_frame = ttk.Frame(box, style="StatusBox.TFrame")
        self.progress_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.progress_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            style="Bdap.Horizontal.TProgressbar",
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_label = ttk.Label(
            self.progress_frame,
            textvariable=self.progress_label_var,
            anchor="w",
            justify="left",
            style="Progress.TLabel",
        )
        self.progress_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.progress_label.bind("<Configure>", self._on_text_wraplength)
        self._hide_progress()

    def _on_text_wraplength(self, event) -> None:
        """Adatta la lunghezza di wrapping del testo alla larghezza del widget al ridimensionamento."""
        widget = event.widget
        if isinstance(widget, ttk.Label):
            widget.configure(wraplength=max(180, event.width - 16))

    def _on_canvas_configure(self, event) -> None:
        """Mantiene la larghezza del contenuto scrollabile sincronizzata con la larghezza del canvas."""
        if self._scroll_window_id is None:
            return
        self._scroll_canvas.itemconfigure(self._scroll_window_id, width=event.width)

    def _on_window_configure(self, _event) -> None:
        """Mantiene il testo lungo leggibile al ridimensionamento della finestra."""
        if self.header_subtitle_label:
            self.header_subtitle_label.configure(wraplength=max(600, self.winfo_width() - 120))
        if self.status_label:
            self.status_label.configure(wraplength=max(620, self.winfo_width() - 120))

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

        # Linux: 
        if (step := getattr(event, "num", None)) in (4, 5):
            self._scroll_canvas.yview_scroll(-1 if step == 4 else 1, "units")
            return

        # macOS/Windows: 
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
        assert DND_FILES is not None
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<DropEnter>>", self._on_drop_enter)
        self.drop_zone.dnd_bind("<<DropLeave>>", self._on_drop_leave)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop_enter(self, _) -> str:
        """Gestisce l'evento di ingresso drag."""
        self.drop_zone.configure(bg="#e9f7ef", fg="#165d3b", relief="solid", borderwidth=2)
        return "copy"

    def _on_drop_leave(self, _) -> str:
        """Gestisce l'evento di uscita drag."""
        self.drop_zone.configure(bg="#f7fcf9", fg="#4f6b5c", relief="solid", borderwidth=2)
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
        self.drop_zone.configure(bg="#f7fcf9", fg="#4f6b5c", relief="solid", borderwidth=2)
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
            self.progress_frame.grid()

    def _hide_progress(self) -> None:
        """Nasconde e azzera la barra di avanzamento."""
        self.progress_var.set(0.0)
        self.progress_label_var.set("")
        if self.progress_frame is not None:
            self.progress_frame.grid_remove()

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
        """Abilita/disabilita il pulsante di esecuzione."""
        self.run_button.configure(state="disabled" if running else "normal")

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
        root = Path(self.workspace_var.get()).expanduser()
        if not root.exists():
            self._reset_comuni_state()
            self.status_var.set(ui_text.STATUS_WORKSPACE_NOT_FOUND)
            return

        comuni = self.service.list_comuni(root)
        self.comune_combo["values"] = comuni
        self._set_comune_state(comuni, root)
        self._refresh_years_for_selected_comune()
        self.status_var.set(ui_text.STATUS_READY if comuni else ui_text.NO_COMUNI_FOUND_MESSAGE)

    def _reset_comuni_state(self) -> None:
        """Reimposta il selettore comune allo stato vuoto."""
        self.comune_combo["values"] = []
        self.comune_var.set("")
        self.comune_combo.configure(state="disabled")
        self._set_years_options([])

    def _set_comune_state(self, comuni: list[str], workspace_root: Path | None = None) -> None:
        """Imposta lo stato del selettore comune in base ai comuni scoperti.

        Quando vengono scoperti più comuni, prova a fare un match automatico
        con il nome della cartella workspace.
        """
        # CASO 1: Nessun comune scoperto, disabilita il menu e svuotalo
        if not comuni:
            self.comune_combo.configure(state="disabled")
            self.comune_var.set("")
            return

        # CASO 2: Un solo comune scoperto, selezionalo automaticamente e disabilita il menu
        if len(comuni) == 1:
            self.comune_var.set(comuni[0])
            self.comune_combo.configure(state="disabled")
            return

        # CASO 3a: Più comuni scoperti: prova a fare un match automatico con il nome della cartella workspace
        if workspace_root is not None:
            workspace_name = workspace_root.name
            normalized_workspace = normalize_token(workspace_name)
            # Look for a comune that matches the workspace folder name
            for comune in comuni:
                if normalize_token(comune) == normalized_workspace:
                    self.comune_var.set(comune)
                    self.comune_combo.configure(state="readonly")
                    return

        # CASO 3b: Più comuni scoperti ma nessun match con il nome della cartella: lascia il menu vuoto e selezionabile
        if self.comune_var.get() not in comuni:
            self.comune_var.set("")
        self.comune_combo.configure(state="readonly")

    def _set_years_options(self, years: list[int]) -> None:
        """Ricostruisce il menu degli anni a partire dagli anni scoperti."""
        sorted_years = sorted(set(int(y) for y in years)) if years else []
        self.available_years = sorted_years

        # Rebuild menu
        self.years_menu.delete(0, "end")
        self.years_menu.add_command(label=ui_text.ALL_YEARS_LABEL, command=self._select_all_years)
        self.year_toggle_vars = {}

        if sorted_years:
            self.years_menu.add_separator()
            for year in sorted_years:
                var = tk.BooleanVar(value=False)
                self.year_toggle_vars[year] = var
                self.years_menu.add_checkbutton(label=str(year), variable=var, command=lambda y=year: self._toggle_year(y))

        # Update range combos
        self._update_year_range_combos(sorted_years)
        self._select_all_years()

    def _update_year_range_combos(self, years: list[int]) -> None:
        """Aggiorna i selettori per l'intervallo di anni."""
        range_values = [str(y) for y in years]
        combo_state = "readonly" if years else "disabled"
        for combo in [self.year_from_combo, self.year_to_combo]:
            combo["values"] = range_values
            combo.configure(state=combo_state)
        self.year_from_var.set("")
        self.year_to_var.set("")

    def _select_all_years(self) -> None:
        """Reimposta la selezione su tutti gli anni."""
        self.selected_years.clear()
        for var in self.year_toggle_vars.values():
            var.set(False)
        self.years_label_var.set(ui_text.ALL_YEARS_LABEL)

    def _toggle_year(self, year: int) -> None:
        """Attiva/disattiva la selezione di un anno e aggiorna l'etichetta."""
        var = self.year_toggle_vars.get(year)
        if not var:
            return

        if var.get():
            self.selected_years.add(year)
        else:
            self.selected_years.discard(year)

        if not self.selected_years:
            self._select_all_years()
            return

        self.years_label_var.set(",".join(str(y) for y in sorted(self.selected_years)))

    def _refresh_years_for_selected_comune(self) -> None:
        """Aggiorna gli anni in base al comune selezionato."""
        comune = self.comune_var.get().strip()
        workspace_root = Path(self.workspace_var.get()).expanduser()
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
        workspace_root = Path(self.workspace_var.get()).expanduser()
        form = GuiFormData(
            workspace_root=workspace_root,
            comune=self.comune_var.get().strip(),
            explicit_years=tuple(sorted(self.selected_years)),
            year_from=self.year_from_var.get(),
            year_to=self.year_to_var.get(),
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
