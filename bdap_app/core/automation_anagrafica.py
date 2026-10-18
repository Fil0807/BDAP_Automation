"""Compilazione del foglio ANAGRAFICA a partire dai questionari del rendiconto."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

logger = logging.getLogger(__name__)

ANAGRAFICA_SHEET_NAME = "ANAGRAFICA"


def _parse_italian_date(date_str: Optional[str]) -> Optional[datetime]:
    """Tenta di convertire una stringa di data italiana (GG/MM/AAAA) in datetime."""
    if not date_str or not isinstance(date_str, str):
        return None
    cleaned = date_str.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            pass
    return None


def extract_anagrafica_from_questionario(q_path: Path) -> dict:
    """Estrae i dati anagrafici, delibera e verbale da un file questionario."""
    if not q_path.exists():
        return {}

    wb = openpyxl.load_workbook(q_path, data_only=False)
    sheetname = None
    for s in wb.sheetnames:
        if "ANAGRAFICA" in s.upper() or "SCHEDA_ANAGRAFICA" in s.upper():
            sheetname = s
            break

    ws_dati = wb["DATI"] if "DATI" in wb.sheetnames else None
    res: dict = {}

    if sheetname and ws_dati:
        ws = wb[sheetname]
        for r in range(1, min(40, ws.max_row + 1)):
            for c in range(1, min(12, ws.max_column + 1)):
                v = ws.cell(r, c).value
                if isinstance(v, str) and ("OFFSET(" in v.upper() or "SCARTO(" in v.upper()):
                    m = re.search(r"(?:OFFSET|SCARTO)\([^,]+,(\d+),(\d+)\)", v, re.I)
                    if m:
                        row_off = int(m.group(1))
                        col_off = int(m.group(2))
                        target_row = 2 + row_off
                        target_col = 1 + col_off
                        dati_val = ws_dati.cell(target_row, target_col).value
                        label = ""
                        for prev_c in range(c - 1, 0, -1):
                            lv = ws.cell(r, prev_c).value
                            if lv and not str(lv).startswith("="):
                                label = str(lv).strip().lower()
                                break
                        if label:
                            res[label] = dati_val

    # Leggi o integra direttamente dai metadata/header di DATI
    if ws_dati:
        header_map: dict[str, object] = {}
        for c in range(1, ws_dati.max_column + 1):
            h = ws_dati.cell(1, c).value
            if h:
                header_map[str(h).strip()] = ws_dati.cell(2, c).value

        # Ente
        if not res.get("denominazione"):
            res["denominazione"] = header_map.get("firstname") or header_map.get("S1INFENTE[EnteX]")
        if not res.get("codice fiscale dell'ente:"):
            res["codice fiscale dell'ente:"] = header_map.get("attribute_8") or header_map.get("S1INFENTE[CodFisX]")
        if not res.get("regione:"):
            res["regione:"] = header_map.get("attribute_5")
        if not res.get("provincia:"):
            res["provincia:"] = header_map.get("attribute_6")
        if not res.get("tipo ente"):
            res["tipo ente"] = header_map.get("attribute_1")
        if not res.get("popolazione"):
            res["popolazione"] = header_map.get("attribute_2") or header_map.get("S1INFPOP")

        # Delibera rendiconto
        if not res.get("deliberazione n."):
            res["deliberazione n."] = header_map.get("S1INF40A[Verb]")
        if not res.get("deliberazione data"):
            res["deliberazione data"] = header_map.get("S1INF40A[Data]")

        # Verbale revisore
        if not res.get("verbale n."):
            res["verbale n."] = header_map.get("S1INF30[Verb]")
        if not res.get("verbale data"):
            res["verbale data"] = header_map.get("S1INF30[Data]")

        # Dati Revisore
        if not res.get("nome:"):
            res["nome:"] = header_map.get("INF20ANAREF[Nome]") or header_map.get("S1INF20[Nom]")
        if not res.get("cognome:"):
            res["cognome:"] = header_map.get("INF20ANAREF[Cognome]") or header_map.get("S1INF20[Cogn]")
        if not res.get("codice fiscale"):
            res["codice fiscale"] = header_map.get("INF20ANAREF[CodFis]") or header_map.get("S1INF20[CodFis]")
        if not res.get("indirizzo mail:"):
            res["indirizzo mail:"] = header_map.get("INF20ANAREF[Posta]") or header_map.get("S1INF20[Post]")

    wb.close()
    return res


def populate_anagrafica_sheet(
    workbook_path: Path,
    questionario_files_by_year: dict[int, Path],
    log: Optional[Callable[[str], None]] = None,
) -> bool:
    """Popola il foglio ANAGRAFICA nel workbook specificato."""
    emit = log or (lambda _: None)

    if not workbook_path.exists():
        emit(f"Workbook {workbook_path} non trovato.")
        return False

    wb = openpyxl.load_workbook(workbook_path)
    if ANAGRAFICA_SHEET_NAME not in wb.sheetnames:
        emit(f"Foglio {ANAGRAFICA_SHEET_NAME} non presente nel template.")
        wb.close()
        return False

    ws = wb[ANAGRAFICA_SHEET_NAME]

    # Raccogli anagrafica per tutti gli anni disponibili
    anagrafiche_by_year: dict[int, dict] = {}
    for yr, q_path in sorted(questionario_files_by_year.items()):
        try:
            anag = extract_anagrafica_from_questionario(q_path)
            if anag:
                anagrafiche_by_year[yr] = anag
                emit(f"Estratta anagrafica questionario anno {yr} da {q_path.name}")
        except Exception as e:
            logger.warning(f"Errore estrazione anagrafica anno {yr}: {e}")

    if not anagrafiche_by_year:
        emit("Nessun dato anagrafico estratto dai questionari.")
        wb.close()
        return False

    # 1. Compila Sezione 1: Dati Generali Ente (usando l'anno più recente disponibile)
    latest_year = max(anagrafiche_by_year.keys())
    latest_anag = anagrafiche_by_year[latest_year]

    denominazione = latest_anag.get("denominazione")
    codice_fiscale = latest_anag.get("codice fiscale dell'ente:")
    provincia = latest_anag.get("provincia:")
    regione = latest_anag.get("regione:")
    tipo_ente = latest_anag.get("tipo ente") or "Comune"
    popolazione = latest_anag.get("popolazione")

    # Fallback su altri anni se un campo manca nel più recente
    for a in anagrafiche_by_year.values():
        if not denominazione and a.get("denominazione"):
            denominazione = a.get("denominazione")
        if not codice_fiscale and a.get("codice fiscale dell'ente:"):
            codice_fiscale = a.get("codice fiscale dell'ente:")
        if not provincia and a.get("provincia:"):
            provincia = a.get("provincia:")
        if not regione and a.get("regione:"):
            regione = a.get("regione:")
        if not popolazione and a.get("popolazione"):
            popolazione = a.get("popolazione")

    ws["C5"].value = str(denominazione).strip() if denominazione else None
    ws["C6"].value = str(codice_fiscale).strip() if codice_fiscale else None
    ws["C7"].value = str(provincia).strip() if provincia else None
    ws["C8"].value = str(regione).strip() if regione else None
    ws["C9"].value = str(tipo_ente).strip() if tipo_ente else None
    if popolazione is not None and str(popolazione).strip() not in ("", "0"):
        ws["C10"].value = str(popolazione).strip()

    # 2. Compila Sezione 2: Estremi Deliberazione Rendiconto (righe 14..19 per anni 2020..2025)
    for r in range(14, 20):
        yr_cell = ws.cell(r, 2)
        if not yr_cell.value:
            continue
        try:
            yr = int(str(yr_cell.value).strip())
        except ValueError:
            continue

        anag_yr = anagrafiche_by_year.get(yr, {})
        delib_n = anag_yr.get("deliberazione n.")
        delib_data = anag_yr.get("deliberazione data") or anag_yr.get("del")

        if delib_n is not None and str(delib_n).strip() not in ("", "0"):
            ws.cell(r, 3).value = f"n. {delib_n}" if not str(delib_n).lower().startswith("n.") else str(delib_n)

        if delib_data is not None and str(delib_data).strip() not in ("", "0"):
            ws.cell(r, 4).value = str(delib_data).strip()
            # Calcola tempestività
            d_dt = _parse_italian_date(str(delib_data).strip())
            deadline_dt = datetime(yr + 1, 4, 30)
            if d_dt:
                ws.cell(r, 6).value = "Tempestivo" if d_dt <= deadline_dt else "In ritardo"

    # 3. Compila Sezione 3: Relazione Organo di Revisione (righe 23..28)
    for r in range(23, 29):
        yr_cell = ws.cell(r, 2)
        if not yr_cell.value:
            continue
        try:
            yr = int(str(yr_cell.value).strip())
        except ValueError:
            continue

        anag_yr = anagrafiche_by_year.get(yr, {})
        verb_n = anag_yr.get("verbale n.")
        verb_data = anag_yr.get("verbale data") or anag_yr.get("del ")
        revisore_nome = anag_yr.get("nome:") or ""
        revisore_cognome = anag_yr.get("cognome:") or anag_yr.get("cognome") or ""
        revisore_cf = anag_yr.get("codice fiscale") or ""
        revisore_email = anag_yr.get("indirizzo mail:") or ""
        ruolo = anag_yr.get("dati del referente/responsabile per la compilazione della relazione:") or "Revisore unico"

        if verb_n is not None and str(verb_n).strip() not in ("", "0"):
            ws.cell(r, 3).value = f"n. {verb_n}" if not str(verb_n).lower().startswith("n.") else str(verb_n)

        if verb_data is not None and str(verb_data).strip() not in ("", "0"):
            ws.cell(r, 4).value = str(verb_data).strip()

        ws.cell(r, 5).value = str(ruolo).strip()

        nome_completo = f"{str(revisore_nome).strip()} {str(revisore_cognome).strip()}".strip()
        if nome_completo:
            ws.cell(r, 6).value = nome_completo

        if revisore_cf:
            ws.cell(r, 7).value = str(revisore_cf).strip()

        if revisore_email:
            ws.cell(r, 8).value = str(revisore_email).strip()

    wb.save(workbook_path)
    wb.close()
    emit(f"Foglio ANAGRAFICA popolato con successo in {workbook_path.name}")
    return True
