"""Utility per normalizzazione e confronto del testo."""

import re
from typing import Optional
from difflib import SequenceMatcher


def normalize_text(value: object) -> str:
    """Normalizza testo libero per un matching più robusto delle etichette."""
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def tokenize_text(value: object) -> list[str]:
    """Estrae i token (parole) dal testo."""
    if value is None:
        return []
    return [token for token in re.findall(r"[a-z0-9]+", str(value).lower()) if token]


def text_options(value: object) -> list[str]:
    """Normalizza una stringa o lista di stringhe in una lista di opzioni testuali."""
    if isinstance(value, str):
        text = value.strip()
        return [text] if text else []
    if isinstance(value, (list, tuple)):
        return [str(item).strip() for item in value if str(item).strip()]
    return []


def first_text_option(value: object) -> Optional[str]:
    """Estrae il primo testo utile da una stringa o lista di stringhe."""
    options = text_options(value)
    return options[0] if options else None


def label_match_score(target_label: str, candidate_label: str) -> float:
    """Calcola la similarità tra due etichette (0.0 - 1.0)."""
    target_norm = normalize_text(target_label)
    candidate_norm = normalize_text(candidate_label)
    if not target_norm or not candidate_norm:
        return 0.0

    if len(candidate_norm) < 6:
        return 0.0

    if target_norm == candidate_norm:
        return 1.0

    contains_bonus = 0.0
    if target_norm in candidate_norm or candidate_norm in target_norm:
        contains_bonus = 0.2

    seq_ratio = SequenceMatcher(None, target_norm, candidate_norm).ratio()

    stopwords = {
        "di", "del", "della", "delle", "dei", "degli", "da", "a", "al", "alla",
        "e", "ed", "o", "con", "per", "su", "in", "il", "lo", "la", "le", "i",
        "parte", "prospetto", "totale", "importo",
    }
    target_tokens = {tok for tok in tokenize_text(target_label) if len(tok) >= 3 and tok not in stopwords}
    candidate_tokens = {tok for tok in tokenize_text(candidate_label) if len(tok) >= 3 and tok not in stopwords}
    if target_tokens:
        token_overlap = len(target_tokens & candidate_tokens) / len(target_tokens)
    else:
        token_overlap = 0.0

    return min(1.0, max(seq_ratio, token_overlap) + contains_bonus)


def parse_bool_flag(raw: object, default: bool = False) -> bool:
    """Interpreta un valore come booleano da rappresentazioni testuali/numeriche comuni.

    Replica il parsing permissivo usato nel codice di automazione: accetta
    booleani, stringhe numeriche e diversi token locali per sì/no.
    """
    if raw is None:
        return default
    if isinstance(raw, bool):
        return raw
    text = normalize_text(raw)
    if text in {"1", "true", "vero", "yes", "si", "s"}:
        return True
    if text in {"0", "false", "falso", "no", "n"}:
        return False
    return default


def _coerce_int_or_default(raw: object, default: int) -> int:
    """Converte un valore grezzo in `int`, ricorrendo al `default` per valori None o non validi."""
    if raw is None:
        return default
    if isinstance(raw, bool):
        return int(raw)
    if isinstance(raw, int):
        return raw
    if isinstance(raw, float):
        return int(raw)
    try:
        text = str(raw).strip()
        if not text:
            return default
        return int(float(text))
    except (TypeError, ValueError):
        return default


def _coerce_year_value(value: object) -> Optional[int]:
    """Coerce un valore in un anno intero se ha l'aspetto di un anno (1900-2100)."""
    if value is None or isinstance(value, bool):
        return None
    text = str(value).strip()
    try:
        val = float(text)
        if 1900 <= val <= 2100 and val.is_integer():
            return int(val)
    except ValueError:
        pass
    return None


def _coerce_year_value_importocassa(value: object) -> Optional[int]:
    """Analizza anni in formati più permissivi usati nelle intestazioni dei questionari."""
    strict = _coerce_year_value(value)
    if strict is not None:
        return strict
    if value is None or isinstance(value, bool):
        return None
    year_attr = getattr(value, "year", None)
    if isinstance(year_attr, int) and 1900 <= year_attr <= 2100:
        return year_attr
    text = str(value).strip()
    matches = re.findall(r"\b(19\d{2}|20\d{2}|2100)\b", text)
    if matches:
        return int(matches[-1])
    return None
