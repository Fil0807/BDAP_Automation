"""Utility per la formattazione dei valori."""

from decimal import Decimal
from typing import Optional


def format_value_italian(value, percent: bool = False, boolean_yes_no: bool = False):
    """Formatta valori numerici secondo la convenzione italiana (virgola per decimali, punto per migliaia).

    Se `percent` è True, tratta i valori piccoli (<=1) come percentuali (es. 0.12 diventa 12%).
    Se `boolean_yes_no` è True, restituisce 'Sì' o 'No' (es. 1 -> 'Sì', 0 -> 'No').
    """
    if value is None:
        return None

    if boolean_yes_no:
        if isinstance(value, bool):
            return "Sì" if value else "No"
        num = coerce_numeric(value)
        if num is not None:
            return "Sì" if num > 0 else "No"
        s_val = str(value).strip().lower()
        if s_val in ("1", "1.0", "si", "sì", "yes", "true", "s", "t"):
            return "Sì"
        if s_val in ("0", "0.0", "no", "false", "n", "f"):
            return "No"
        return str(value).strip()

    string_had_percent = False
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        if not percent:
            return value
        string_had_percent = value.strip().endswith("%")
        numeric_value = coerce_numeric(value)
        if numeric_value is None:
            return value
        value = numeric_value
    
    try:
        if isinstance(value, (int, float)):
            num = Decimal(str(value))
        else:
            num = Decimal(value)
        
        if percent and abs(num) <= 1 and not string_had_percent:
            # Valori come 0.12 arrivano da Excel come rapporto; 12.0 e simili sono gia' percentuali.
            num = num * Decimal("100")

        formatted = format(num, '.2f')
        parts = formatted.split('.')
        if len(parts) == 2:
            integer_part = parts[0]
            decimal_part = parts[1]
            if len(integer_part) > 3:
                reverse = integer_part[::-1]
                grouped = '.'.join([reverse[i:i+3] for i in range(0, len(reverse), 3)])
                integer_part = grouped[::-1]
            result = f"{integer_part},{decimal_part}"
            if percent:
                result = f"{result}%"
            return result
        return formatted
    except (ValueError, TypeError):
        return value


def coerce_numeric(value: object) -> Optional[Decimal]:
    """Converte un valore in `Decimal`, gestendo anche il formato numerico italiano."""
    if value is None:
        return None
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    if not isinstance(value, str):
        return None
    raw = value.strip().replace(" ", "")
    if not raw:
        return None
    if raw.endswith("%"):
        raw = raw[:-1]
        if not raw:
            return None
    raw = raw.replace(".", "").replace(",", ".")
    try:
        return Decimal(raw)
    except Exception:
        return None
