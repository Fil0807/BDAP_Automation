"""Helper per gestire le voci che arrivano da relazione."""

import re
from pathlib import Path
from typing import Optional

from openpyxl.comments import Comment
from ..support.comment_builder import create_source_comment

try:
    import pypdf
except ImportError:
    pypdf = None


def _find_relazione_pdf(bdap_path: Optional[Path], year: int) -> Optional[Path]:
    """Cerca il PDF della relazione del revisore per l'anno richiesto o il più recente disponibile."""
    if bdap_path is None:
        return None

    search_roots = []
    base_dir = bdap_path.parent
    if base_dir.exists():
        search_roots.append(base_dir)
        if base_dir.parent.exists():
            search_roots.append(base_dir.parent)

    candidates = []
    for root in search_roots:
        rel_dir = root / "Relazione revisore"
        if rel_dir.exists() and rel_dir.is_dir():
            candidates.extend(rel_dir.glob("*.pdf"))
        # cerca anche direttamente o ricorsivamente cartelle con 'relazione'
        for sub in root.glob("*[Rr]elazione*"):
            if sub.is_dir():
                candidates.extend(sub.glob("*.pdf"))

    if not candidates:
        return None

    # Prova a matchare l'anno esatto
    year_str = str(year)
    for c in candidates:
        if year_str in c.name:
            return c

    return None


def _check_limiti_personale_in_pdf(pdf_path: Path) -> bool:
    """Verifica se il PDF della relazione attesta il rispetto dei limiti di spesa del personale."""
    if pypdf is None or not pdf_path.exists():
        return False

    try:
        reader = pypdf.PdfReader(str(pdf_path))
        for page in reader.pages:
            text = (page.extract_text() or "").lower()
            if "557" in text and ("rientra" in text or "rispettat" in text or "limiti" in text):
                return True
    except Exception:
        pass
    return False


def process_relazione_row(
    target_cell,
    bdap_path: Optional[Path],
    year: int,
    special_key: Optional[str],
    data_source_mapping: dict,
    keep_source_reference: bool = True,
):
    """Risolve le voci da relazione (es. rispetto limiti spesa personale) se presente il documento."""
    key = str(special_key or "").lower()
    if key in ("spesadelpersonalesostenuta", "spesadelpersonalesostenutarientraneilimitidicui"):
        pdf_path = _find_relazione_pdf(bdap_path, year)
        if pdf_path is not None and _check_limiti_personale_in_pdf(pdf_path):
            target_cell.value = "rispettati i limiti"
            if keep_source_reference:
                target_cell.comment = create_source_comment(
                    source_file_name=pdf_path.name,
                    detail="verificato rispetto limiti spesa personale ex art. 1 c. 557 L. 296/2006",
                    header="RELAZIONE REVISORI",
                )
            return True, 1, None, None, pdf_path, "rispettati i limiti"

    target_cell.value = None
    target_cell.comment = None
    return False, 0, None, None, None, None

