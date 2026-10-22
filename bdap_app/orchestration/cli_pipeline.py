"""Helper per l'esecuzione della pipeline nei workflow CLI di BDAP."""

from pathlib import Path
import argparse
import shutil
from typing import Callable, Optional

from openpyxl import load_workbook

import re
from decimal import Decimal
from openpyxl.utils import get_column_letter

from ..core.automation_fcde import populate_fcde_sheet
from ..core.automation_piano_indicatori import populate_piano_indicatori_sheet
from ..core.automation import fill_accounting_analysis, format_inserted_data_cells_center
from ..core.automation_search import find_sheet, find_year_column
from ..core.automation_anagrafica import populate_anagrafica_sheet
from ..core.automation_questionario import _discover_questionario_workbook
from ..support.value_formatter import coerce_numeric, format_value_italian
from ..support.controlli_post import (
    collect_controlli_post_report,
    format_controlli_post_summary,
    write_controlli_post_sheet,
)
from ..support.conclusioni import (
    detect_comune_name,
    detect_population,
    write_conclusioni_sheet,
)
from ..support.comment_builder import fix_all_comments_in_workbook, create_source_comment
from ..support.text_utils import _coerce_year_value
from .cli_discovery import (
    discover_bdap_files_by_year,
    infer_bdap_dir,
    infer_new_analysis_path,
    normalize_token,
    resolve_base_dir_for_comune,
)

TEMPLATE_YEAR_START = 2020
TEMPLATE_YEAR_END = 2026

__all__ = [
    "TEMPLATE_YEAR_START",
    "TEMPLATE_YEAR_END",
    "parse_years_arg",
    "validate_inputs",
    "apply_comune_defaults",
    "ensure_analysis_input",
    "resolve_years_to_process",
    "run_pipeline",
]


def parse_years_arg(years_arg: str) -> list[int]:
    """Analizza una lista separata da virgole di anni in una lista intera ordinata e univoca.

    Esempio: ``"2021, 2023, 2022"`` produce ``[2021, 2022, 2023]``.

    Args:
        years_arg: Stringa contenente gli anni separati da virgole.

    Returns:
        Lista di interi rappresentanti gli anni selezionati, ordinati e privi di duplicati.

    Raises:
        ValueError: Se la stringa passata è vuota o non contiene anni validi.
    """
    years: list[int] = []
    for token in years_arg.split(","):
        token = token.strip()
        if not token:
            continue
        years.append(int(token))
    if not years:
        raise ValueError("--years cannot be empty")
    return sorted(set(years))


def _discover_years_from_workbook(workbook_path: Path) -> list[int]:
    """Scansiona le intestazioni del workbook alla ricerca degli anni supportati.

    Ispeziona le prime 12 righe e 20 colonne di ciascun foglio cercando valori numerici
    riconducibili ad anni (es. 2020..2030).

    Args:
        workbook_path: Percorso del file Excel :class:`pathlib.Path`.

    Returns:
        Lista ordinata di anni unici rilevati nel modello.
    """
    years: set[int] = set()
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            max_rows = min(ws.max_row or 0, 12)
            max_cols = min(ws.max_column or 0, 20)
            for row_idx in range(1, max_rows + 1):
                for col_idx in range(1, max_cols + 1):
                    year_value = _coerce_year_value(ws.cell(row=row_idx, column=col_idx).value)
                    if year_value is not None:
                        years.add(year_value)
    finally:
        wb.close()
    return sorted(years)


def resolve_years_to_process(
    args: argparse.Namespace,
    *,
    template_years: list[int],
    bdap_years: list[int],
) -> list[int]:
    """Risolve l'elenco esatto degli anni richiesti per l'elaborazione corrente.

    Applica le seguenti priorità:
    1. Se è specificato ``--all-years``, restituisce tutti gli anni BDAP disponibili.
    2. Se è specificato ``--years``, elabora l'elenco esplicito.
    3. Se è specificato ``--year``, elabora il singolo anno.
    4. Altrimenti ricorre agli anni individuati nel template, o come fallback agli anni BDAP.

    Args:
        args: Argomenti analizzati dalla CLI o configurati dalla GUI.
        template_years: Elenco degli anni presenti nelle intestazioni del template.
        bdap_years: Elenco degli anni trovati nei file BDAP della cartella comune.

    Returns:
        Lista ordinata di anni da processare.
    """
    if args.all_years:
        return sorted(set(bdap_years))
    if args.years:
        return parse_years_arg(args.years)
    if args.year is not None:
        return [int(args.year)]
    return template_years or sorted(set(bdap_years))


def _should_use_anchor_bdap(args: argparse.Namespace) -> bool:
    """Verifica se utilizzare il file BDAP principale come ancora di fallback ad anno singolo."""
    return args.year is not None and not args.all_years and args.years is None


def validate_inputs(args: argparse.Namespace) -> None:
    """Valida l'esistenza dei file e delle directory di input richiesti per l'elaborazione.

    In modalità multi-anno verifica l'esistenza di ``args.bdap_dir``; in modalità
    anno singolo verifica ``args.bdap``.

    Args:
        args: Argomenti con i percorsi di lavoro.

    Raises:
        FileNotFoundError: Se uno o più percorsi obbligatori non esistono su disco.
    """
    source_analysis = args.template if getattr(args, "template", None) else args.analysis
    required_paths = [source_analysis]
    auto_multi_year_mode = args.all_years or args.years is not None or args.year is None
    if auto_multi_year_mode:
        required_paths.append(args.bdap_dir) # In modalita' multi-anno, richiede l'intera directory BDAP per individuare i file per anno.
    else:
        required_paths.append(args.bdap)

    missing = [str(path) for path in required_paths if not path.exists()]
    if missing:
        missing_list = "\n".join(f"- {path}" for path in missing)
        raise FileNotFoundError(f"Missing input files:\n{missing_list}")


def apply_comune_defaults(args: argparse.Namespace, parser: argparse.ArgumentParser) -> None:
    """Applica i percorsi predefiniti specifici per il comune agli argomenti non specificati esplicitamente.

    Risolve automaticamente la cartella base del comune, la cartella ``1 doc/dati BDAP``,
    il file di analisi e il percorso di output predefinito (``<Nome>_compilato.xlsx``).

    Args:
        args: Argomenti correnti da aggiornare in-place.
        parser: Parser degli argomenti CLI per controllare i valori di default.
    """
    if not args.comune:
        return

    base_dir = resolve_base_dir_for_comune(args.workspace_root, args.comune)
    bdap_dir = infer_bdap_dir(base_dir)
    discovered = discover_bdap_files_by_year(bdap_dir)
    default_bdap = next(iter(discovered.values()), bdap_dir / "Rend.xlsx")

    # Mantiene inalterati i percorsi forniti dall'utente e sovrascrive solo i default del parser.
    if args.analysis == parser.get_default("analysis"):
        args.analysis = infer_new_analysis_path(base_dir)
    if args.bdap_dir == parser.get_default("bdap_dir"):
        args.bdap_dir = bdap_dir
    if args.bdap == parser.get_default("bdap"):
        args.bdap = default_bdap
    if args.output == parser.get_default("output"):
        analysis_name = Path(args.analysis)
        analysis_suffix = analysis_name.suffix or ".xlsx"
        args.output = base_dir / f"{analysis_name.stem}_compilato{analysis_suffix}"


def ensure_analysis_input(args: argparse.Namespace) -> None:
    """Verifica che il file di analisi esista su disco, a meno che non sia specificato un template.

    Args:
        args: Argomenti contenenti i percorsi di analisi o template.

    Raises:
        FileNotFoundError: Se il file di analisi non esiste.
    """
    if getattr(args, "template", None):
        return

    if args.analysis.exists():
        return

    raise FileNotFoundError(f"Analysis file not found: {args.analysis}")


def _resolve_template_source(template_path: Path) -> Path:
    """Risolve il percorso effettivo del template Excel, gestendo eventuali collegamenti .gsheet.

    Args:
        template_path: Percorso del template configurato.

    Returns:
        Percorso valido con estensione ``.xlsx``.

    Raises:
        ValueError: Se il file è un collegamento ``.gsheet`` non esportato in ``.xlsx``.
    """
    if template_path.suffix.lower() != ".gsheet":
        return template_path

    exported_xlsx = template_path.with_suffix(".xlsx")
    if exported_xlsx.exists():
        return exported_xlsx

    raise ValueError(
        "Template .gsheet non utilizzabile direttamente. "
        "Esporta il file Google Sheets in .xlsx e riprova."
    )


def _prepare_output_from_source(args: argparse.Namespace, log: Callable[[str], None]) -> Path:
    """Copia il file sorgente selezionato nel percorso di output prima dell'elaborazione.

    Preserva integralmente il template originale eseguendo una copia fisica su cui
    verranno apportate le modifiche.

    Args:
        args: Argomenti contenenti ``args.analysis``, ``args.template`` e ``args.output``.
        log: Funzione di callback per i log.

    Returns:
        Il percorso :class:`pathlib.Path` del file sorgente effettivamente copiato.

    Raises:
        ValueError: Se il formato del template non è supportato.
        FileNotFoundError: Se il file sorgente non esiste.
    """
    source_path = args.analysis
    if getattr(args, "template", None):
        source_path = _resolve_template_source(args.template)

    if source_path.suffix.lower() == ".gsheet":
        raise ValueError(
            "Template .gsheet non supportato. Usa un file .xlsx come template."
        )
    if not source_path.exists():
        raise FileNotFoundError(f"Missing input file: {source_path}")

    args.output.parent.mkdir(parents=True, exist_ok=True)
   
    # Copia la sorgente nel percorso di output prima dell'elaborazione.
    # Questo preserva il file originale e permette aggiornamenti in-place sul file di output.
    shutil.copy2(source_path, args.output) 
    
    # Se è stato usato un template specifico, registra il percorso sorgente per informare l'utente.
    if getattr(args, "template", None):
        log(f"Template used to generate output: {source_path}")
    return source_path


def run_pipeline(
    args: argparse.Namespace,
    log: Callable[[str], None] = print,
    progress: Optional[Callable[[float, int, str], None]] = None,
) -> None:
    """Esegue la pipeline end-to-end di compilazione per gli anni e i fogli selezionati.

    Questa funzione coordina tutte le fasi di esecuzione:
    1. Convalida input e preparazione output;
    2. Determinazione degli anni e dei fogli da processare;
    3. Popolamento annuale di FCDE, Piano Indicatori e fogli di bilancio;
    4. Sincronizzazione formule e percentuali tra FCDE ed ANALISI CONTABILE;
    5. Estrazione anagrafica da questionario;
    6. Ispezione documentale PDF e scrittura della scheda CONTROLLI-POST;
    7. Formattazione e centraggio celle, ottimizzazione commenti.

    Args:
        args: Namespace contenente tutti i parametri di configurazione (comune, workspace, anni, output).
        log: Funzione di callback per la registrazione dei messaggi di log (default: :func:`print`).
        progress: Funzione di callback opzionale con firma ``(completed, total, message)``
            per aggiornare la barra di avanzamento della GUI.

    Raises:
        ValueError: Se non è possibile determinare gli anni da processare.
        FileNotFoundError: Se file o directory essenziali non sono reperibili.
    """
    def emit_progress(completed: float, total: int, message: str) -> None:
        if progress is None:
            return
        safe_total = max(1, int(total))
        safe_completed = min(max(0.0, float(completed)), float(safe_total))
        progress(safe_completed, safe_total, message)

    emit_progress(0, 1, "Preparazione file...")
    ensure_analysis_input(args)
    validate_inputs(args)
    _prepare_output_from_source(args, log)

    DEDICATED_SHEET_NAMES = {
        "anagrafica",
        "fcde",
        "pianoindicatori",
        "controllipost",
        "mappaturebdap",
        "conclusioni",
    }
    output_workbook = load_workbook(args.output, read_only=True)
    all_sheet_names = list(output_workbook.sheetnames)
    sheets_to_process = [
        s for s in all_sheet_names
        if normalize_token(s) not in DEDICATED_SHEET_NAMES
    ] if getattr(args, "all_sheets", False) else [args.analysis_sheet]
    output_workbook.close()

    source_analysis = args.template if getattr(args, "template", None) else args.analysis
    template_years = _discover_years_from_workbook(source_analysis)

    bdap_by_year = discover_bdap_files_by_year(args.bdap_dir)
    years_to_process = resolve_years_to_process(
        args,
        template_years=template_years,
        bdap_years=list(bdap_by_year.keys()),
    )

    if not years_to_process:
        raise ValueError(
            "No years to process. Provide --year/--years or place BDAP .xlsx files with recognizable year in --bdap-dir (YYYY or YY->20YY for BDAP-like names)."
        )

    total_steps = max(1, len(years_to_process) * max(1, len(sheets_to_process)))
    completed_steps = 0
    emit_progress(
        completed_steps,
        total_steps,
        f"Avvio compilazione: {len(years_to_process)} anni, {len(sheets_to_process)} fogli",
    )

    # tuple (anno, celle_aggiornate, riferimenti_non_risolti) per il riepilogo finale.
    summary: list[tuple[int, int, int]] = []
    # Elabora ciascun anno in modo indipendente così un anno problematico non blocca l'intera esecuzione.
    for year in years_to_process:
        year_updated = 0
        year_unresolved = 0
        emit_progress(
            completed_steps + 0.03,
            total_steps,
            f"Anno {year}: preparazione dati BDAP",
        )
        bdap_file = bdap_by_year.get(year)
        if bdap_file is None:
            if _should_use_anchor_bdap(args) and getattr(args, "bdap", None) is not None:
                bdap_file = args.bdap
                log(
                    f"Year {year}: no year-specific BDAP file found in {args.bdap_dir}; "
                    f"using fallback anchor {bdap_file.name}."
                )
            else:
                log(
                    f"Year {year}: no BDAP file found in {args.bdap_dir}; "
                    "target cells will be left empty."
                )

        if bdap_file is not None:
            populate_fcde_sheet(args.output, bdap_file, year, log=log)
            populate_piano_indicatori_sheet(args.output, bdap_file, year, log=log)

        # In modalità "tutti i fogli", ogni foglio viene trattato come un passaggio di compilazione indipendente.
        for sheet_name in sheets_to_process:
            step_start = completed_steps

            def report_row_progress(row_idx: int, max_row: int, *, year=year, sheet_name=sheet_name, step_start=step_start) -> None:
                if max_row <= 0:
                    fraction = 0.03
                else:
                    fraction = max(0.03, min(float(row_idx) / float(max_row), 1.0))
                emit_progress(
                    step_start + fraction,
                    total_steps,
                    f"Anno {year}, foglio '{sheet_name}': riga {min(row_idx, max_row)}/{max_row}",
                )

            try:
                updated, unresolved = fill_accounting_analysis(
                    analysis_path=args.output,
                    bdap_path=bdap_file,
                    output_path=args.output,
                    analysis_sheet=sheet_name,
                    year=year,
                    clear_when_missing=bdap_file is None,
                    progress=report_row_progress,
                )
                year_updated += updated
                year_unresolved += unresolved
                if getattr(args, "all_sheets", False):
                    log(
                        f"Year {year}, sheet '{sheet_name}': updated cells = {updated}; "
                        f"unresolved references = {unresolved}."
                    )
            except ValueError as exc:
                if "Year column" in str(exc) and getattr(args, "all_sheets", False):
                    log(f"Year {year}, sheet '{sheet_name}': no year column, skipped.")
                else:
                    log(f"Skipping year {year}: {exc}")
            finally:
                completed_steps += 1
                emit_progress(
                    completed_steps,
                    total_steps,
                    f"Anno {year}, foglio '{sheet_name}' completato",
                )

        summary.append((year, year_updated, year_unresolved))
        log(
            f"Year {year} populate completed: updated cells = {year_updated}; "
            f"unresolved references = {year_unresolved}."
        )

        if getattr(args, "template", None) and year_updated == 0:
            # Se il passaggio specifico per l'anno non ha prodotto scritture, riprova su una
            # copia pulita del template e unisci solo la colonna di quell'anno nel workbook
            # finale. Questo evita interferenze di stato tra anni da passaggi precedenti.
            rescue_source = source_analysis
            rescue_work = args.output.parent / f".{args.output.stem}.recheck-{year}.xlsx"
            shutil.copy2(rescue_source, rescue_work)
            rescue_bdap = bdap_by_year.get(year) or args.bdap
            fill_accounting_analysis(
                analysis_path=rescue_work,
                bdap_path=rescue_bdap,
                output_path=rescue_work,
                analysis_sheet=sheets_to_process[0],
                year=year,
                clear_when_missing=False,
            )

            master_wb = load_workbook(args.output)
            rescue_wb = load_workbook(rescue_work)
            try:
                for sheet_name in sheets_to_process:
                    master_ws = find_sheet(master_wb, sheet_name)
                    rescue_ws = find_sheet(rescue_wb, sheet_name)
                    if master_ws is None or rescue_ws is None:
                        continue
                    year_col = find_year_column(master_ws, year)
                    if year_col is None:
                        continue
                    for row_idx in range(1, master_ws.max_row + 1):
                        source_cell = rescue_ws.cell(row=row_idx, column=year_col)
                        target_cell = master_ws.cell(row=row_idx, column=year_col)
                        target_cell.value = source_cell.value
                        if source_cell.has_style:
                            target_cell._style = source_cell._style
                        if source_cell.comment is not None:
                            target_cell.comment = source_cell.comment
            finally:
                rescue_wb.close()
                fix_all_comments_in_workbook(master_wb)
                master_wb.save(args.output)
                master_wb.close()
                try:
                    rescue_work.unlink()
                except Exception:
                    pass

    # Sincronizza formule e valori tra ANALISI CONTABILE ed FCDE (righe 6, 7, 8 in FCDE, righe 25, 26 in AC)
    try:
        sync_fcde_and_analisi_contabile(args.output, years_to_process, log=log)
    except Exception as exc:
        log(f"Avviso: sincronizzazione FCDE / ANALISI CONTABILE non riuscita: {exc}")

    log(f"Final populated output: {args.output}")
    emit_progress(total_steps, total_steps, "Compilazione completata")

    if summary:
        log("\nPopulation summary by year:")
        for year, updated, unresolved in summary:
            log(f"- {year}: updated={updated}, unresolved={unresolved}")

    # Rileva directory di ricerca documenti limitatamente alla cartella del comune
    comune_root = args.output.parent
    if getattr(args, "comune", None) and getattr(args, "workspace_root", None):
        try:
            comune_root = resolve_base_dir_for_comune(args.workspace_root, args.comune)
        except Exception:
            comune_root = args.output.parent

    search_dirs: list[Path] = [comune_root]
    if getattr(args, "bdap_dir", None) and args.bdap_dir.exists():
        search_dirs.append(args.bdap_dir)

    # Popola foglio ANAGRAFICA se presente nel template
    questionario_by_year: dict[int, Path] = {}
    for y, b_path in bdap_by_year.items():
        q_path = _discover_questionario_workbook(b_path, y)
        if q_path is not None:
            questionario_by_year[y] = q_path

    for sdir in search_dirs:
        if sdir.exists():
            for q_file in sdir.rglob("*.xlsx"):
                if q_file.name.startswith("~$") or q_file.name.startswith("._") or "__MACOSX" in q_file.parts:
                    continue
                if "questionario" in q_file.name.lower():
                    y_matches = re.findall(r"(202\d)", q_file.name)
                    if y_matches:
                        y_int = int(y_matches[-1])
                        if y_int not in questionario_by_year:
                            questionario_by_year[y_int] = q_file

    if questionario_by_year:
        try:
            populate_anagrafica_sheet(args.output, questionario_by_year, log=log)
        except Exception as exc:
            log(f"Avviso durante la compilazione del foglio ANAGRAFICA: {exc}")

    controlli_report = collect_controlli_post_report(args.output.parent, years=years_to_process)
    args.controlli_post_report = controlli_report
    if controlli_report.has_files:
        wrote_controlli_sheet = write_controlli_post_sheet(args.output, controlli_report)
        if not wrote_controlli_sheet:
            log("Foglio CONTROLLI-POST non trovato nel template: riepilogo non scritto nel workbook.")
        log("")
        log(format_controlli_post_summary(controlli_report))

    # Aggiorna il foglio CONCLUSIONI (solo se presente in eventuali template legacy)
    has_conclusioni = False
    try:
        check_wb = load_workbook(args.output, read_only=True)
        has_conclusioni = any(s.strip().upper() == "CONCLUSIONI" for s in check_wb.sheetnames)
        check_wb.close()
    except Exception:
        has_conclusioni = False

    if has_conclusioni:
        latest_year = max(years_to_process) if years_to_process else None
        comune_name = detect_comune_name(
            search_dirs=search_dirs,
            bdap_files=bdap_by_year,
            fallback_comune=getattr(args, "comune", None),
        )
        population = detect_population(
            search_dirs=search_dirs,
            target_year=latest_year,
        )
        write_conclusioni_sheet(
            workbook_path=args.output,
            comune_name=comune_name,
            population=population,
            log=log,
        )

    # Allinea al centro sia verticale che orizzontale tutte le celle dei dati inseriti
    try:
        format_inserted_data_cells_center(args.output, log=log)
    except Exception as exc:
        log(f"Avviso durante la formattazione centrata delle celle: {exc}")


def sync_fcde_and_analisi_contabile(
    workbook_path: Path,
    years: list[int],
    log: Optional[Callable[[str], None]] = None,
) -> None:
    """Sincronizza bidirezionalmente i valori e le formule tra i fogli FCDE e ANALISI CONTABILE.

    Esegue le seguenti operazioni per ciascun anno specificato:
    1. Legge la quota di avanzo "Parte disponibile" dal foglio ``ANALISI CONTABILE`` (riga 19);
    2. Scrive nel foglio ``FCDE`` (riga 6) la formula Excel dinamica di collegamento:
       ``='ANALISI CONTABILE'!<Colonna>19``;
    3. Imposta in ``FCDE`` la formula per la percentuale comprensiva della parte disponibile
       (riga 8) e il metodo di calcolo "ordinario" (riga 9);
    4. Sincronizza su ``ANALISI CONTABILE`` la percentuale ordinaria di copertura FCDE (riga 25)
       e la percentuale comprensiva della parte disponibile (riga 26), corredandole con
       commenti di tracciabilità.

    Args:
        workbook_path: Percorso del file Excel di output :class:`pathlib.Path`.
        years: Elenco degli anni per cui applicare la sincronizzazione.
        log: Funzione di log opzionale per messaggi informativi.
    """
    wb = load_workbook(workbook_path)
    try:
        ws_ac = find_sheet(wb, "ANALISI CONTABILE")
        ws_fcde = find_sheet(wb, "FCDE")
        if ws_ac is None or ws_fcde is None:
            return

        for year in years:
            ac_col = find_year_column(ws_ac, year, max_header_rows=3)
            fcde_col = find_year_column(ws_fcde, year, max_header_rows=3)
            if ac_col is None or fcde_col is None:
                continue

            ac_col_letter = get_column_letter(ac_col)
            fcde_col_letter = get_column_letter(fcde_col)

            # 1. Legge Parte disponibile da AC riga 19
            parte_disp_val = ws_ac.cell(row=19, column=ac_col).value
            parte_disp_num = coerce_numeric(parte_disp_val) or Decimal("0")

            # 2. Aggiorna FCDE riga 6 con formula collegata ad AC
            cell_fcde_6 = ws_fcde.cell(row=6, column=fcde_col)
            cell_fcde_6.value = f"='ANALISI CONTABILE'!{ac_col_letter}19"

            # 3. Legge FCDE riga 2 (Importo FCDE) e riga 5 (Totale residui EP)
            fcde_imp = coerce_numeric(ws_fcde.cell(row=2, column=fcde_col).value) or Decimal("0")
            fcde_tot = coerce_numeric(ws_fcde.cell(row=5, column=fcde_col).value)

            # 4. Aggiorna FCDE riga 8 (% FCDE+parte disponibile) e riga 9 (METODO)
            cell_fcde_8 = ws_fcde.cell(row=8, column=fcde_col)
            cell_fcde_8.value = f"=({fcde_col_letter}2+{fcde_col_letter}6)/{fcde_col_letter}5*100"
            ws_fcde.cell(row=9, column=fcde_col).value = "ordinario"

            pct_plus_disp = None
            if fcde_tot not in (None, Decimal("0")):
                pct_plus_disp = ((fcde_imp + parte_disp_num) * Decimal("100")) / fcde_tot

            # 5. Sincronizza su ANALISI CONTABILE riga 25 e riga 26
            cell_fcde_7 = ws_fcde.cell(row=7, column=fcde_col)
            val_7 = coerce_numeric(cell_fcde_7.value)
            if val_7 is not None:
                ws_ac.cell(row=25, column=ac_col).value = format_value_italian(val_7)
                if not ws_ac.cell(row=25, column=ac_col).comment:
                    ws_ac.cell(row=25, column=ac_col).comment = create_source_comment(
                        source_file_name="Template_Analisi.xlsx",
                        sheet_name="FCDE",
                        cell_ref=f"FCDE!{fcde_col_letter}7",
                    )

            if pct_plus_disp is not None:
                ws_ac.cell(row=26, column=ac_col).value = format_value_italian(pct_plus_disp)
                if not ws_ac.cell(row=26, column=ac_col).comment:
                    ws_ac.cell(row=26, column=ac_col).comment = create_source_comment(
                        source_file_name="Template_Analisi.xlsx",
                        sheet_name="FCDE",
                        cell_ref=f"FCDE!{fcde_col_letter}8",
                    )

        fix_all_comments_in_workbook(wb)
        wb.save(workbook_path)
        if log:
            log("FCDE e ANALISI CONTABILE sincronizzati con successo.")
    finally:
        wb.close()
