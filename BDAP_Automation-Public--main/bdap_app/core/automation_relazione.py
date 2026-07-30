"""Helper per gestire le voci che arrivano da relazione."""

from pathlib import Path
from typing import Optional


def process_relazione_row(
    target_cell,
    bdap_path: Optional[Path],
    year: int,
    special_key: Optional[str],
    data_source_mapping: dict,
    keep_source_reference: bool = True,
):
    """Lascia la cella vuota; il chiamante registra il riferimento irrisolto."""
    target_cell.value = None
    target_cell.comment = None
    return False, 0, None, None, None, None
