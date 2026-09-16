from decimal import Decimal

from bdap_app.support.value_formatter import coerce_numeric, format_value_italian


def test_percent_ratios_are_scaled() -> None:
    assert format_value_italian(Decimal("0.125"), percent=True) == "12,50%"


def test_percent_values_are_not_scaled_twice() -> None:
    assert format_value_italian(Decimal("12.5"), percent=True) == "12,50%"


def test_percent_text_values_are_normalized() -> None:
    assert format_value_italian("100%", percent=True) == "100,00%"
    assert format_value_italian("85,54%", percent=True) == "85,54%"


def test_coerce_numeric_accepts_percentage_strings() -> None:
    assert coerce_numeric("85,54%") == Decimal("85.54")
    assert coerce_numeric("100 %") == Decimal("100")
