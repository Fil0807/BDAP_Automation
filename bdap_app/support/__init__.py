"""Utility di supporto per elaborazione testi, formattazione e risoluzione regole."""

from .text_utils import normalize_text, label_match_score
from .value_formatter import format_value_italian, coerce_numeric
from .value_resolver import find_bdap_value_by_label, resolve_row_value_from_label
from .conclusioni import (
    clean_comune_name,
    detect_comune_name,
    detect_population,
    write_conclusioni_sheet,
)
from .comment_builder import (
    create_parte_accantonata_comment,
    create_source_comment,
    fix_all_comments_in_workbook,
)

__all__ = [
    "normalize_text",
    "label_match_score",
    "format_value_italian",
    "coerce_numeric",
    "find_bdap_value_by_label",
    "resolve_row_value_from_label",
    "clean_comune_name",
    "detect_comune_name",
    "detect_population",
    "write_conclusioni_sheet",
    "create_source_comment",
    "create_parte_accantonata_comment",
    "fix_all_comments_in_workbook",
]

