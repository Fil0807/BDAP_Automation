"""Helper per la scoperta di percorsi e dataset per i workflow CLI di BDAP."""

from pathlib import Path
import re
from typing import Optional
from openpyxl import load_workbook
import sys

# Individua file con nome tipo "Rend. YYYY.xlsx" (case-insensitive, spazi consentiti).
BDAP_YEAR_FILE_PATTERN = re.compile(r"Rend\.\s*(\d{4})\.xlsx$", re.IGNORECASE)
GENERIC_YEAR_PATTERN = re.compile(r"(19\d{2}|20\d{2})")
GENERIC_TWO_DIGIT_YEAR_PATTERN = re.compile(r"(?<!\d)(\d{2})(?!\d)")

__all__ = [
    "BDAP_YEAR_FILE_PATTERN",
    "GENERIC_YEAR_PATTERN",
    "normalize_token",
    "dedupe_paths",
    "infer_new_analysis_path",
    "infer_bdap_dir",
    "discover_comune_base_dirs",
    "resolve_base_dir_for_comune",
    "discover_bdap_files_by_year",
    "resolve_default_template_path",
    "_is_flat_bdap_folder",
    "infer_comune_name_from_bdap_dir",
]

COMUNE_INLINE_PATTERN = re.compile(r"\bcomune\s+di\b\s*[:\-]?\s*(.+)", re.IGNORECASE)
COMUNE_LABEL_PATTERN = re.compile(r"^\s*comune\s+di\s*[:\-]?\s*$", re.IGNORECASE)
GENERIC_COMUNE_VALUE_TOKENS = {
    "nome",
    "ente",
    "denominazione",
    "denominazioneente",
    "comune",
}


def _parse_two_digit_year_from_stem(stem: str) -> Optional[int]:
    """Accetta token anno corto (YY) solo per nomi file BDAP-like e lo mappa a 20YY.

    Ad esempio, converte stringhe come ``"Rend. 26.xlsx"`` nell'anno intero ``2026``.
    Se il nome non contiene parole chiave come ``rend``, ``rendiconto`` o ``bdap``,
    oppure se contiene più di un token a due cifre, restituisce ``None`` per evitare ambiguità.

    Args:
        stem: Nome del file senza estensione (es. ``"Rend. 26"``).

    Returns:
        L'anno a quattro cifre (es. 2026) se identificato univocamente, altrimenti ``None``.
    """
    lower_stem = stem.lower()
    if not any(keyword in lower_stem for keyword in ("rend", "rendiconto", "bdap")):
        return None

    matches = GENERIC_TWO_DIGIT_YEAR_PATTERN.findall(stem)
    if len(matches) != 1:
        return None

    yy = int(matches[0])
    return 2000 + yy


def normalize_token(value: str) -> str:
    """Normalizza etichette e nomi rimuovendo spazi e punteggiatura per un matching robusto.

    Converte l'intera stringa in minuscolo e conserva unicamente i caratteri alfanumerici
    ASCII (a-z, 0-9).

    Args:
        value: Stringa di input (es. ``"Comune di Esempio (BG)"``).

    Returns:
        Stringa pulita e normalizzata (es. ``"comunediesempiobg"``).
    """
    return re.sub(r"[^a-z0-9]", "", value.lower())


EXCLUDED_BDAP_YEAR_FILE_TOKENS = {
    "analisi",
    "compilato",
    "controllipost",
    "debiti",
    "questionario",
    "template",
}


def _bdap_year_file_priority(path: Path) -> Optional[int]:
    """Classifica i soli file annuali BDAP, escludendo workbook ausiliari.

    Assegna un punteggio numerico di priorità (minore è il punteggio, maggiore è la priorità):
    - 0: Nome standard esatto ``Rend. YYYY.xlsx``
    - 1: Presenza di ``rendiconto`` nel nome
    - 2: Presenza di ``rend`` nel nome
    - 3: Presenza di ``bdap`` nel nome

    Args:
        path: Percorso del file da valutare.

    Returns:
        Un intero da 0 a 3 indicante la priorità, oppure ``None`` se il file deve essere escluso.
    """
    normalized_stem = normalize_token(path.stem)
    if any(token in normalized_stem for token in EXCLUDED_BDAP_YEAR_FILE_TOKENS):
        return None
    if BDAP_YEAR_FILE_PATTERN.search(path.name):
        return 0
    if "rendiconto" in normalized_stem:
        return 1
    if "rend" in normalized_stem:
        return 2
    if "bdap" in normalized_stem:
        return 3
    return None


def dedupe_paths(paths: list[Path]) -> list[Path]:
    """Mantiene l'ordine della prima occorrenza rimuovendo duplicati di percorsi fisici.

    Normalizza e risolve i percorsi per evitare che rappresentazioni relative o simboliche
    della medesima cartella vengano elaborate più volte.

    Args:
        paths: Lista di percorsi :class:`pathlib.Path`.

    Returns:
        Nuova lista ordinata contenente solo percorsi unici.
    """
    unique: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        try:
            key = str(path.expanduser().resolve())
        except OSError:
            key = str(path.expanduser())
        if key not in seen:
            unique.append(path)
            seen.add(key)
    return unique


def _has_bdap_year_files(folder: Path) -> bool:
    """Verifica se la cartella specificata contiene direttamente file di rendiconto BDAP.

    Accetta sia nomi rigorosi (es. ``Rend. YYYY.xlsx``) che varianti riconosciute
    da :func:`discover_bdap_files_by_year`.

    Args:
        folder: Cartella da ispezionare.

    Returns:
        ``True`` se sono presenti file annuali BDAP validi, altrimenti ``False``.
    """
    if not folder.exists() or not folder.is_dir():
        return False
    return bool(discover_bdap_files_by_year(folder))


def _is_flat_bdap_folder(folder: Path) -> bool:
    """Verifica se la cartella contiene file BDAP direttamente senza la sottocartella '1 doc'.

    Identifica la modalità "flat" in cui l'utente ha salvato i file Excel del rendiconto
    direttamente nella radice del comune o in una cartella dati senza la gerarchia canonica
    ``<Comune>/1 doc/dati BDAP/``.

    Args:
        folder: Cartella da verificare.

    Returns:
        ``True`` se è una cartella piatta contenente file BDAP, altrimenti ``False``.
    """
    if not _has_bdap_year_files(folder):
        return False
    # Verifica se possiede la struttura standard del comune (sottocartella "1 doc").
    if (folder / "1 doc").exists() or any(
        p.is_dir() and normalize_token(p.name) in {"1doc", "doc1", "1documento", "documento1"}
        for p in folder.iterdir()
    ):
        return False
    return True


def _clean_comune_name(raw: str) -> str:
    """Pulisce le stringhe estratte dai fogli rimuovendo punteggiatura e spazi spuri.

    Args:
        raw: Testo grezzo contenente il nome del comune.

    Returns:
        Stringa ripulita.
    """
    cleaned = re.sub(r"\s+", " ", raw).strip(" .:-\t\n\r")
    return cleaned


def _looks_like_real_comune_name(value: str) -> bool:
    """Verifica se una stringa estratta somiglia al nome reale di un Ente locale.

    Scarta etichette generiche di tabella come "Ente", "Denominazione", "Comune"
    o stringhe troppo corte (meno di 3 caratteri).

    Args:
        value: Stringa candidata.

    Returns:
        ``True`` se il valore è un candidato valido come nome di comune, altrimenti ``False``.
    """
    token = normalize_token(value)
    if not token:
        return False
    if token in GENERIC_COMUNE_VALUE_TOKENS:
        return False
    if len(token) < 3:
        return False
    return True


def _extract_comune_from_workbook(path: Path) -> Optional[str]:
    """Estrae il nome dell'Ente scansionando le prime celle del rendiconto BDAP.

    Cerca l'intestazione standard nella cella B7 (accanto a "Denominazione ente" in A7)
    o mediante pattern regex per frasi come "COMUNE DI ..." nei primi tre fogli.

    Args:
        path: Percorso del file Excel di rendiconto.

    Returns:
        Il nome del comune estratto, oppure ``None`` se non individuato.
    """
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return None

    try:
        for sheet in wb.worksheets[:3]:
            # Nei workbook Rend YYYY il nome del comune spesso è in B7, accanto ad A7 = "Denominazione ente".
            try:
                a7_value = sheet["A7"].value
                b7_value = sheet["B7"].value
            except Exception:
                a7_value = None
                b7_value = None

            max_row = min(sheet.max_row, 40) if sheet.max_row is not None else 40
            max_col = min(sheet.max_column, 12) if sheet.max_column is not None else 12

            if isinstance(a7_value, str) and normalize_token(a7_value) == normalize_token("Denominazione ente"):
                # Scansiona la riga a destra di A7 e conserva il primo valore che sembra un nome di comune reale.
                for col_idx in range(2, max_col + 1):
                    row_value = sheet.cell(row=7, column=col_idx).value
                    if not isinstance(row_value, str):
                        continue

                    candidate = _clean_comune_name(row_value)
                    if not candidate:
                        continue

                    inline_match = COMUNE_INLINE_PATTERN.search(candidate)
                    if inline_match:
                        candidate = _clean_comune_name(inline_match.group(1))

                    if _looks_like_real_comune_name(candidate):
                        return candidate
            for row_idx in range(1, max_row + 1):
                for col_idx in range(1, max_col + 1):
                    value = sheet.cell(row=row_idx, column=col_idx).value
                    if not isinstance(value, str):
                        continue

                    text = value.strip()
                    if not text:
                        continue

                    inline_match = COMUNE_INLINE_PATTERN.search(text)
                    if inline_match:
                        candidate = _clean_comune_name(inline_match.group(1))
                        if candidate:
                            return candidate

                    if COMUNE_LABEL_PATTERN.match(text):
                        right_value = None
                        if col_idx + 1 <= max_col:
                            right_value = sheet.cell(row=row_idx, column=col_idx + 1).value
                        if isinstance(right_value, str):
                            candidate = _clean_comune_name(right_value)
                            if candidate:
                                return candidate
    finally:
        wb.close()

    return None


def infer_comune_name_from_bdap_dir(folder: Path) -> Optional[str]:
    """Inferisce il nome del comune analizzando i rendiconti BDAP presenti in una cartella piatta.

    Viene impiegata quando l'utente seleziona direttamente una cartella contenente i file
    del rendiconto senza una directory nominale genitore. Ispeziona i primi 3 file annuali
    tramite :func:`_extract_comune_from_workbook` e, in caso di esito negativo, effettua un fallback
    sul nome della cartella padre.

    Args:
        folder: Percorso della cartella contenente i file BDAP.

    Returns:
        Il nome del comune se rilevato, altrimenti ``None``.
    """
    if not _is_flat_bdap_folder(folder):
        return None

    # Usa discover_bdap_files_by_year per ottenere solo file BDAP effettivi (non file di analisi)
    bdap_files = discover_bdap_files_by_year(folder)
    for year in sorted(bdap_files.keys())[:3]:  # Prova i primi 3 anni
        workbook_path = bdap_files[year]
        comune_name = _extract_comune_from_workbook(workbook_path)
        if comune_name:
            return comune_name

    # Fallback: prova a usare il nome della cartella padre se sembra un comune
    parent_name = folder.parent.name
    parent_normalized = normalize_token(parent_name)
    if parent_name and parent_normalized not in {"downloads", "desktop", "documents", "dati", "datibdap"}:
        return parent_name

    return None


def infer_new_analysis_path(base_dir: Path) -> Path:
    """Risolve il percorso del file di analisi 'Nuovo Modello' nella directory del comune.

    Cerca in ordine di preferenza:
    1. ``<base_dir>/Analisi Nuovo Modello.xlsx``
    2. Eventuali file ``Analisi*.xlsx`` nella sottocartella ``2 analisi funzionaria``
    3. File corrispondenti al pattern ``Analisi*Nuovo*Modello*.xlsx``
    4. Qualsiasi file ``Analisi*.xlsx`` (esclusi quelli con prefisso ``REV_``)
    5. Ricerca analoga nella cartella superiore se ``base_dir`` è una cartella piatta

    Args:
        base_dir: Directory base del comune.

    Returns:
        Il percorso :class:`pathlib.Path` del file di analisi individuato o del default atteso.
    """
    preferred = base_dir / "Analisi Nuovo Modello.xlsx"
    if preferred.exists():
        return preferred

    preferred_subdir = base_dir / "2 analisi funzionaria"
    if preferred_subdir.exists():
        subdir_candidates = sorted(
            p for p in preferred_subdir.glob("Analisi*.xlsx")
            if not p.name.upper().startswith("REV_")
        )
        if subdir_candidates:
            return subdir_candidates[0]

    candidates = sorted(base_dir.glob("Analisi*Nuovo*Modello*.xlsx"))
    if candidates:
        return candidates[0]

    fallback_candidates = sorted(
        p  for p in base_dir.glob("Analisi*.xlsx")
        if not p.name.upper().startswith("REV_")
    )
    if fallback_candidates:
        return fallback_candidates[0]

    # Fallback: cerca nella cartella padre se base_dir e' una cartella BDAP piatta
    parent_dir = base_dir.parent
    if parent_dir.exists() and parent_dir != base_dir:
        preferred_parent = parent_dir / "Analisi Nuovo Modello.xlsx"
        if preferred_parent.exists():
            return preferred_parent
        
        parent_candidates = sorted(
            p for p in parent_dir.glob("Analisi*.xlsx")
            if not p.name.upper().startswith("REV_")
        )
        if parent_candidates:
            return parent_candidates[0]

    return preferred


def infer_bdap_dir(base_dir: Path) -> Path:
    """Individua la directory dei dati BDAP secondo la gerarchia canonica o varianti note.

    Verifica le seguenti posizioni:

    * ``<base_dir>/1 doc/dati BDAP``
    * La cartella ``base_dir`` stessa (se contiene direttamente file annuali BDAP)
    * Varianti normalizzate della cartella ``1 doc`` (es. ``1doc``, ``doc1``, ``1documento``) contenenti ``datibdap``.

    Args:
        base_dir: Directory base da cui partire.

    Returns:
        Il percorso :class:`pathlib.Path` della cartella contenente i dati BDAP.
    """
    preferred = base_dir / "1 doc" / "dati BDAP"
    if preferred.exists():
        return preferred

    if not base_dir.exists() or not base_dir.is_dir():
        return preferred

    # Supporta dataset piatti in cui i file Rend. YYYY.xlsx sono memorizzati direttamente in base_dir.
    if _has_bdap_year_files(base_dir):
        return base_dir

    doc_dir = base_dir / "1 doc"
    doc_dir_candidates: list[Path] = []
    if doc_dir.exists() and doc_dir.is_dir():
        doc_dir_candidates.append(doc_dir)
    else:
        for candidate in sorted(p for p in base_dir.iterdir() if p.is_dir()):
            token = normalize_token(candidate.name)
            if token in {"1doc", "doc1", "1documento", "documento1"}:
                doc_dir_candidates.append(candidate)

    for doc_candidate in doc_dir_candidates:
        for candidate in sorted(p for p in doc_candidate.iterdir() if p.is_dir()):
            if normalize_token(candidate.name) == "datibdap":
                return candidate

    return preferred


def discover_comune_base_dirs(workspace_root: Path) -> list[Path]:
    """Scopre tutte le radici di cartella comune valide all'interno di un workspace.

    Gestisce diverse modalità di selezione:
    1. Workspace contenente molteplici sottocartelle di comuni;
    2. Cartella del singolo comune selezionata direttamente;
    3. Sottocartella ``1 doc`` o ``dati BDAP`` selezionata direttamente;
    4. Cartella piatta con file BDAP.

    Args:
        workspace_root: Directory selezionata dall'utente.

    Returns:
        Lista ordinata di percorsi :class:`pathlib.Path` corrispondenti alle basi dei comuni trovati.
    """
    candidates: list[Path] = []
    if not workspace_root.exists() or not workspace_root.is_dir():
        return candidates

    # Accetta radici che puntano gia' a una cartella di dataset (livello comune).
    workspace_bdap_dir = infer_bdap_dir(workspace_root)
    if workspace_bdap_dir.exists() and workspace_bdap_dir.is_dir():
        candidates.append(workspace_root)

    normalized_root_name = normalize_token(workspace_root.name)

    # Accetta la selezione diretta di ".../<comune>/1 doc".
    if normalized_root_name in {"1doc", "doc1", "1documento", "documento1"}:
        comune_dir = workspace_root.parent
        bdap_dir = infer_bdap_dir(comune_dir)
        if bdap_dir.exists() and bdap_dir.is_dir():
            candidates.append(comune_dir)

    # Accetta la selezione diretta di ".../<comune>/1 doc/dati BDAP".
    if normalized_root_name == "datibdap":
        # Primo tentativo: fa parte di una struttura valida <comune>/1 doc/dati BDAP?
        doc_dir = workspace_root.parent
        comune_dir = doc_dir.parent
        if doc_dir.exists() and doc_dir.is_dir() and comune_dir.exists() and comune_dir.is_dir():
            if normalize_token(doc_dir.name) in {"1doc", "doc1", "1documento", "documento1"}:
                candidates.append(comune_dir)
        # Secondo tentativo: e' una cartella piatta "dati BDAP" con file BDAP direttamente all'interno?
        elif _has_bdap_year_files(workspace_root):
            # Tratta la cartella "dati BDAP" stessa come un comune
            candidates.append(workspace_root)

    for comune_dir in sorted(p for p in workspace_root.iterdir() if p.is_dir()):
        bdap_dir = infer_bdap_dir(comune_dir)
        if bdap_dir.exists() and bdap_dir.is_dir():
            candidates.append(comune_dir)

    return dedupe_paths(candidates)


def resolve_base_dir_for_comune(workspace_root: Path, comune: str) -> Path:
    """Risolve la cartella base di un comune dato il suo nome, gestendo eventuali ambiguità.

    Confronta il nome richiesto con i nomi delle cartelle individuate e con i nomi
    inferiti dai rendiconti BDAP.

    Args:
        workspace_root: Radice del workspace in cui cercare.
        comune: Nome (o frazione di nome) del comune da individuare.

    Returns:
        Il percorso :class:`pathlib.Path` della cartella base del comune.

    Raises:
        ValueError: Se il parametro ``comune`` è vuoto o se corrisponde a più comuni (ambiguità).
        FileNotFoundError: Se nessun dataset corrisponde al nome indicato o se il workspace non contiene dati.
    """
    normalized_target = normalize_token(comune)
    if not normalized_target:
        raise ValueError("--comune cannot be empty")

    candidates = discover_comune_base_dirs(workspace_root)
    if not candidates:
        raise FileNotFoundError(
            f"No dataset folder found under {workspace_root}. "
            "Expected pattern: <folder>/1 doc/dati BDAP. "
            "You can pass either: workspace root with many comuni, a single comune folder, "
            "the '1 doc' folder, the 'dati BDAP' folder, or a folder containing BDAP .xlsx files with year in filename (YYYY or YY->20YY for BDAP-like names)."
        )

    matched: list[Path] = []
    exact_name_matches: list[Path] = []  # Preferisci corrispondenze esatte con il nome della cartella
    inferred_name_matches: list[Path] = []  # Fallback su corrispondenze con il nome inferito
    
    for path in candidates:
        folder_name_normalized = normalize_token(path.name)
        inferred_name = infer_comune_name_from_bdap_dir(path)
        inferred_name_normalized = normalize_token(inferred_name) if inferred_name else None
        
        # Controlla se il target corrisponde al nome della cartella (match esatto)
        if normalized_target in folder_name_normalized or folder_name_normalized == normalized_target:
            exact_name_matches.append(path)
            matched.append(path)
        # Controlla se il target corrisponde al nome inferito (fallback)
        elif inferred_name_normalized and (normalized_target in inferred_name_normalized or inferred_name_normalized == normalized_target):
            inferred_name_matches.append(path)
            matched.append(path)
    
    # Dai priorita' alle corrispondenze esatte con il nome della cartella
    if exact_name_matches:
        if len(exact_name_matches) == 1:
            return exact_name_matches[0]
        # Corrispondenze esatte multiple: caso realmente ambiguo
        options = "\n".join(f"- {p.name}" for p in exact_name_matches)
        raise ValueError(
            f"Comune '{comune}' is ambiguous. Matching folders:\n{options}"
        )
    
    # Fallback su corrispondenze con il nome inferito
    if inferred_name_matches:
        if len(inferred_name_matches) == 1:
            return inferred_name_matches[0]
        # Corrispondenze inferite multiple: caso anch'esso ambiguo
        options = "\n".join(f"- {p.name}" for p in inferred_name_matches)
        raise ValueError(
            f"Comune '{comune}' is ambiguous. Matching folders:\n{options}"
        )
    
    if len(matched) == 1:
        return matched[0]
    if len(matched) > 1:
        options = "\n".join(f"- {p.name}" for p in matched)
        raise ValueError(
            f"Comune '{comune}' is ambiguous. Matching folders:\n{options}"
        )

    # Nessuna corrispondenza: mostra tutte le opzioni disponibili
    all_options = "\n".join(f"- {p.name}" for p in candidates)
    raise FileNotFoundError(
        f"No dataset folder matched comune '{comune}'. Available folders:\n{all_options}"
    )


def discover_bdap_files_by_year(bdap_dir: Path) -> dict[int, Path]:
    """Scansiona una cartella BDAP e mappa ciascun anno al file di rendiconto più appropriato.

    Se per il medesimo anno esistono più file (es. ``Rend. 2024.xlsx`` e un generico ``BDAP_2024.xlsx``),
    seleziona il file con priorità migliore determinata da :func:`_bdap_year_file_priority`.

    Args:
        bdap_dir: Cartella contenente i file Excel del rendiconto.

    Returns:
        Dizionario ordinato con chiavi gli anni interi e valori i percorsi dei file :class:`pathlib.Path`.
    """
    best_by_year: dict[int, tuple[int, str, Path]] = {}
    if not bdap_dir.exists():
        return {}

    for path in sorted(bdap_dir.glob("*.xlsx")):
        priority = _bdap_year_file_priority(path)
        if priority is None:
            continue

        match = BDAP_YEAR_FILE_PATTERN.search(path.name)
        year: Optional[int] = None
        if match:
            # Estrai l'anno YYYY dal pattern del nome file come "Rend. YYYY.xlsx"
            year = int(match.group(1))
        else:
            generic = GENERIC_YEAR_PATTERN.search(path.stem)
            if generic:
                year = int(generic.group(1))
            else:
                year = _parse_two_digit_year_from_stem(path.stem)
        if year is not None:
            candidate = (priority, path.name.lower(), path)
            if year not in best_by_year or candidate < best_by_year[year]:
                best_by_year[year] = candidate
    return {year: item[2] for year, item in sorted(best_by_year.items())}


def resolve_default_template_path(
    workspace_root: Path,
    project_root: Optional[Path] = None,
) -> Optional[Path]:
    """Risolve il percorso del modello di analisi Excel (Template_Analisi.xlsx).

    Ricerca il template nelle radici delle risorse pacchettizzate (PyInstaller, py2app),
    nel workspace root e nel project root.

    Args:
        workspace_root: Cartella radice del workspace.
        project_root: Percorso facoltativo del progetto.

    Returns:
        Il percorso :class:`pathlib.Path` del template se trovato, altrimenti ``None``.
    """
    roots: list[Path] = [workspace_root]
    
    resource_root = get_resource_root()
    roots.insert(0, resource_root)

    if project_root is None:
        roots.insert(0, project_root)
    
    # Rimuovi percorsi duplicati preservando l'ordine.
    search_roots = dedupe_paths(roots)

    preferred_names = (
        "Template_Analisi.xlsx",
        "Template Analisi Contabile.xlsx",
        "Template_Analisi_Contabile.xlsx",
    )

    for root in search_roots:
        for name in preferred_names:
            candidate = root / name
            if candidate.exists() and candidate.is_file():
                return candidate

    for root in search_roots:
        candidates = sorted(root.glob("Template*Analisi*Contabile*.xlsx"))
        if candidates:
            return candidates[0]

    # Se esiste solo un collegamento a Google Sheets, suggerisci il percorso .xlsx esportato.
    for root in search_roots:
        gsheet_candidates = sorted(root.glob("Template*Analisi*Contabile*.gsheet"))
        if gsheet_candidates:
            return gsheet_candidates[0].with_suffix(".xlsx")

    return None


def get_resource_root() -> Path:
    """Restituisce la cartella contenente le risorse dell'applicazione.

    Gestisce in modo trasparente l'esecuzione normale da sorgente Python,
    l'esecuzione tramite pacchetto PyInstaller (``sys._MEIPASS``) e l'esecuzione
    come bundle macOS py2app.

    Returns:
        Percorso :class:`pathlib.Path` della cartella risorse.
    """
    # PyInstaller
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    
    # py2app
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent.parent / "Resources"
    
    return Path(__file__).resolve().parents[2]