"""Compilazione automatica del foglio 'Piano Indicatori' dai file BDAP degli indicatori."""

from __future__ import annotations

import logging
from decimal import Decimal
from pathlib import Path
from typing import Callable, Optional

from openpyxl import load_workbook
from openpyxl.comments import Comment

from .automation_indicatori import _discover_aux_workbook
from .automation_search import find_year_column
from ..support.comment_builder import create_source_comment, fix_all_comments_in_workbook
from ..support.text_utils import normalize_text
from ..support.value_formatter import coerce_numeric, format_value_italian

logger = logging.getLogger(__name__)


def _find_indicatori_row_by_code_or_text(ws, code: str, keywords: list[str]) -> Optional[int]:
    """Trova la riga nel foglio 'Piano Indicatori' corrispondente al codice o parole chiave."""
    code_norm = normalize_text(str(code))
    for r in range(1, min(ws.max_row + 1, 40)):
        b_val = normalize_text(str(ws.cell(row=r, column=2).value or ""))
        c_val = normalize_text(str(ws.cell(row=r, column=3).value or ""))
        if code_norm and (b_val == code_norm or code_norm in b_val):
            return r
        for kw in keywords:
            kw_norm = normalize_text(kw)
            if kw_norm and (kw_norm in b_val or kw_norm in c_val):
                return r
    return None


def populate_piano_indicatori_sheet_in_workbook(
    wb_analysis,
    bdap_path: Optional[Path],
    year: int,
    keep_source_reference: bool = True,
) -> bool:
    """Compila le celle dell'anno specificato nel foglio 'Piano Indicatori'."""
    # Trova il foglio Piano Indicatori
    ws_target = None
    for name in wb_analysis.sheetnames:
        if "pianoindicatori" in normalize_text(name):
            ws_target = wb_analysis[name]
            break

    if ws_target is None or bdap_path is None:
        return False

    # Trova la colonna dell'anno
    year_col = find_year_column(ws_target, year)
    if year_col is None:
        return False

    # Localizza il workbook degli indicatori
    ind_file = _discover_aux_workbook(bdap_path, "indicatori", int(year))
    if ind_file is None or not ind_file.exists():
        return False

    try:
        wb_ind = load_workbook(ind_file, data_only=True)
    except Exception as exc:
        logger.warning(f"Impossibile aprire il file indicatori {ind_file}: {exc}")
        return False

    try:
        # Fogli indicatori BDAP
        ws_sint = None
        ws_entr = None
        for sname in wb_ind.sheetnames:
            s_norm = normalize_text(sname)
            if "indicatorisintetici" in s_norm or "rnda" in s_norm:
                ws_sint = wb_ind[sname]
            elif "entrateindicatori" in s_norm or "rndb" in s_norm:
                ws_entr = wb_ind[sname]

        if ws_sint is None:
            return False

        # Valori estratti dal foglio 1 RND a Indicatori sintetici
        # 1.1: riga 13 col D
        val_1_1 = coerce_numeric(ws_sint.cell(row=13, column=4).value) or Decimal("0")
        # 2.8: riga 22 col D
        val_2_8 = coerce_numeric(ws_sint.cell(row=22, column=4).value) or Decimal("0")
        # 3.2: riga 25 col D
        val_3_2 = coerce_numeric(ws_sint.cell(row=25, column=4).value) or Decimal("0")
        # 4.1: riga 27 col D
        val_4_1 = coerce_numeric(ws_sint.cell(row=27, column=4).value) or Decimal("0")
        # 10.3: riga 61 col D
        val_10_3 = coerce_numeric(ws_sint.cell(row=61, column=4).value) or Decimal("0")
        # 12.4: riga 72 col D
        val_12_4 = coerce_numeric(ws_sint.cell(row=72, column=4).value) or Decimal("0")
        # 13.1: riga 74 col D
        val_13_1 = coerce_numeric(ws_sint.cell(row=74, column=4).value) or Decimal("0")
        # 13.2 + 13.3: riga 75 + riga 76 col D
        v_13_2 = coerce_numeric(ws_sint.cell(row=75, column=4).value) or Decimal("0")
        v_13_3 = coerce_numeric(ws_sint.cell(row=76, column=4).value) or Decimal("0")
        val_13_tot = v_13_2 + v_13_3
        # 9.2: riga 57 col D (indicatore tempestività pagamenti)
        val_9_2 = coerce_numeric(ws_sint.cell(row=57, column=4).value) or Decimal("0")

        # Riscossione complessiva dal foglio 2 RND b Entrate (riga 58 colonna H)
        val_risc = None
        if ws_entr is not None:
            # cerca riga Totale entrate
            r_tot = 58
            for r in range(50, min(ws_entr.max_row + 1, 70)):
                if "totale entrate" in str(ws_entr.cell(row=r, column=2).value or "").lower():
                    r_tot = r
                    break
            val_risc = coerce_numeric(ws_entr.cell(row=r_tot, column=8).value) or Decimal("0")

        # Mappatura delle righe del foglio Piano Indicatori
        indicators_map = [
            ("1.1", ["spese rigide"], val_1_1, True, "1 RND a", "D13"),
            ("2.8", ["entrate proprie"], val_2_8, True, "1 RND a", "D22"),
            ("3.2", ["anticipazioni chiuse"], val_3_2, True, "1 RND a", "D25"),
            ("4.1", ["spesa personale"], val_4_1, True, "1 RND a", "D27"),
            ("10.3", ["debiti finanziari"], val_10_3, True, "1 RND a", "D61"),
            ("12.4", ["disavanzo effettivamente"], val_12_4, True, "1 RND a", "D72"),
            ("13.1", ["riconosciuti e finanziati"], val_13_1, True, "1 RND a", "D74"),
            ("13.2 + 13.3", ["in corso di riconoscimento"], val_13_tot, True, "1 RND a", "D75+D76"),
            ("9.2", ["giorni effettivi", "tempestivita"], val_9_2, False, "1 RND a", "D57"),
        ]

        for code, kw_list, val, is_perc, sheet_src, cell_src in indicators_map:
            row_idx = _find_indicatori_row_by_code_or_text(ws_target, code, kw_list)
            if row_idx is not None:
                cell = ws_target.cell(row=row_idx, column=year_col)
                cell.value = format_value_italian(val, percent=is_perc)
                if keep_source_reference:
                    cell.comment = create_source_comment(
                        source_file_name=ind_file.name,
                        sheet_name=sheet_src,
                        cell_ref=cell_src,
                        label=f"Indicatore {code}",
                        header="PIANO DEGLI INDICATORI",
                    )

        # Compila la percentuale di riscossione complessiva se presente come riga distinta
        if val_risc is not None:
            r_risc = _find_indicatori_row_by_code_or_text(
                ws_target,
                "",
                ["riscossione complessiva", "diriscossionecomplessiva"],
            )
            if r_risc is not None:
                cell = ws_target.cell(row=r_risc, column=year_col)
                cell.value = format_value_italian(val_risc, percent=True)
                if keep_source_reference:
                    cell.comment = create_source_comment(
                        source_file_name=ind_file.name,
                        sheet_name="2 RND b",
                        cell_ref="H58",
                        label="Riscossione complessiva",
                        header="PIANO DEGLI INDICATORI",
                    )

            # Verifica parametro riscossione < 47%
            r_param_risc = _find_indicatori_row_by_code_or_text(
                ws_target,
                "",
                ["minore del 47%", "effettiva capacita di riscossione"],
            )
            if r_param_risc is not None:
                cell = ws_target.cell(row=r_param_risc, column=year_col)
                is_deficitario = val_risc < Decimal("47")
                cell.value = "SI" if is_deficitario else "NO"
                if keep_source_reference:
                    cell.comment = create_source_comment(
                        source_file_name=ind_file.name,
                        detail=f"Verifica parametro: {format_value_italian(val_risc, percent=True)} {'<' if is_deficitario else '>='} 47%",
                        header="PIANO DEGLI INDICATORI",
                    )

        # Calcolo verifica finale deficitarietà strutturale (art. 242 TUEL)
        # Un ente è deficitario se almeno la metà dei parametri risulta deficitaria (SI)
        deficitari_count = 0
        total_checked = 0

        # Parametri e rispettive soglie:
        checks = [
            (val_1_1 > Decimal("48")),      # 1.1 > 48%
            (val_2_8 < Decimal("22")),      # 2.8 < 22%
            (val_3_2 > Decimal("0")),       # 3.2 > 0
            (val_10_3 > Decimal("16")),     # 10.3 > 16%
            (val_12_4 > Decimal("1.2")),    # 12.4 > 1.20%
            (val_13_1 > Decimal("1")),      # 13.1 > 1%
            (val_13_tot > Decimal("0.6")),  # 13.2+13.3 > 0.60%
            (val_9_2 > Decimal("0")),       # 9.2 > 0 (ritardo medio)
        ]
        if val_risc is not None:
            checks.append(val_risc < Decimal("47"))  # Riscossione < 47%

        for c in checks:
            total_checked += 1
            if c:
                deficitari_count += 1

        # Riga 16: conteggio parametri deficitari riscontrati (es. "1 su 9")
        r_count = _find_indicatori_row_by_code_or_text(
            ws_target,
            "",
            ["almeno la meta dei parametri", "almeno la metà dei parametri", "sono strutturalmente deficitari"],
        )
        if r_count is not None:
            cell = ws_target.cell(row=r_count, column=year_col)
            cell.value = f"{deficitari_count} su {total_checked}"
            if keep_source_reference:
                cell.comment = create_source_comment(
                    detail=f"Numero parametri deficitari con condizione 'SI': {deficitari_count} su {total_checked}",
                    header="PIANO DEGLI INDICATORI",
                )

        is_ente_deficitario = deficitari_count >= (total_checked / 2)
        r_cond = _find_indicatori_row_by_code_or_text(
            ws_target,
            "",
            ["strutturalmente deficitarie", "condizioni strutturalmente"],
        )
        if r_cond is not None:
            cell = ws_target.cell(row=r_cond, column=year_col)
            cell.value = "SI" if is_ente_deficitario else "NO"
            if keep_source_reference:
                cell.comment = create_source_comment(
                    detail=f"Parametri deficitari riscontrati: {deficitari_count} su {total_checked}",
                    header="PIANO DEGLI INDICATORI",
                )

        # Assicura le descrizioni delle fonti nella colonna K
        _ensure_piano_indicatori_fonte_dati(ws_target)

        return True

    finally:
        wb_ind.close()


def _ensure_piano_indicatori_fonte_dati(ws_target) -> None:
    """Assicura la presenza dell'intestazione e dei valori della colonna K (FONTE DATI)."""
    fonte_col = 11  # Colonna K
    header_cell = ws_target.cell(row=4, column=fonte_col)
    if not header_cell.value:
        header_cell.value = "FONTE DATI"
    if ws_target.column_dimensions["K"].width is None or ws_target.column_dimensions["K"].width < 30:
        ws_target.column_dimensions["K"].width = 42.0

    # Indicatori sintetici (righe con codice indicatore)
    synth_sources = {
        "1.1": "BDAP - Piano degli indicatori sintetici",
        "2.8": "BDAP - Piano degli indicatori sintetici",
        "3.2": "BDAP - Piano degli indicatori sintetici",
        "4.1": "BDAP - Piano degli indicatori sintetici",
        "10.3": "BDAP - Piano degli indicatori sintetici",
        "12.4": "BDAP - Piano degli indicatori sintetici",
        "13.1": "BDAP - Piano degli indicatori sintetici",
        "13.2 + 13.3": "BDAP - Piano degli indicatori sintetici",
        "9.2": "BDAP - Piano degli indicatori sintetici",
    }
    for code, fonte_text in synth_sources.items():
        r = _find_indicatori_row_by_code_or_text(ws_target, code, [])
        if r is not None:
            c = ws_target.cell(row=r, column=fonte_col)
            if not c.value:
                c.value = fonte_text

    # Parametro riscossione < 47%
    r_param_risc = _find_indicatori_row_by_code_or_text(
        ws_target, "", ["minore del 47%", "effettiva capacita di riscossione"]
    )
    if r_param_risc is not None:
        c = ws_target.cell(row=r_param_risc, column=fonte_col)
        if not c.value:
            c.value = "BDAP - Piano degli indicatori analitici"

    # Conteggio parametri deficitari
    r_count = _find_indicatori_row_by_code_or_text(
        ws_target, "", ["almeno la meta dei parametri", "almeno la metà dei parametri", "sono strutturalmente deficitari"]
    )
    if r_count is not None:
        c = ws_target.cell(row=r_count, column=fonte_col)
        if not c.value:
            c.value = "Calcolo parametri deficitari (D.M. 28/12/2018)"

    # Condizione deficitaria complessiva
    r_cond = _find_indicatori_row_by_code_or_text(
        ws_target, "", ["strutturalmente deficitarie", "condizioni strutturalmente"]
    )
    if r_cond is not None:
        c = ws_target.cell(row=r_cond, column=fonte_col)
        if not c.value:
            c.value = "BDAP - Rendiconto (All. l Parametri comuni)"

    # % riscossione complessiva
    r_risc = _find_indicatori_row_by_code_or_text(
        ws_target, "", ["riscossione complessiva", "diriscossionecomplessiva"]
    )
    if r_risc is not None:
        c = ws_target.cell(row=r_risc, column=fonte_col)
        if not c.value:
            c.value = "BDAP - Piano degli indicatori analitici"


def populate_piano_indicatori_sheet(
    analysis_path: Path,
    bdap_file: Path,
    year: int,
    log: Optional[Callable[[str], None]] = None,
) -> bool:
    """Apre il workbook di analisi, compila il foglio 'Piano Indicatori' per l'anno e salva."""
    wb = load_workbook(analysis_path)
    try:
        updated = populate_piano_indicatori_sheet_in_workbook(
            wb,
            bdap_file,
            year,
            keep_source_reference=True,
        )
        if updated:
            fix_all_comments_in_workbook(wb)
            wb.save(analysis_path)
            if log is not None:
                log(f"Year {year}: foglio 'Piano Indicatori' compilato con successo.")
        return updated
    finally:
        wb.close()
