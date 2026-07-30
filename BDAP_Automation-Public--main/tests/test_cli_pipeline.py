from argparse import Namespace

from bdap_app.orchestration.cli_pipeline import resolve_years_to_process


def _args(**overrides):
    values = {"all_years": False, "years": None, "year": None}
    values.update(overrides)
    return Namespace(**values)


def test_explicit_years_do_not_get_template_years_added() -> None:
    assert resolve_years_to_process(
        _args(years="2022,2021"),
        template_years=[2020, 2021],
        bdap_years=[2021, 2022, 2023],
    ) == [2021, 2022]


def test_all_years_means_bdap_years_present() -> None:
    assert resolve_years_to_process(
        _args(all_years=True),
        template_years=[2020, 2021],
        bdap_years=[2021, 2023],
    ) == [2021, 2023]


def test_template_years_are_fallback_only_without_selection() -> None:
    assert resolve_years_to_process(
        _args(),
        template_years=[2020, 2021],
        bdap_years=[2022],
    ) == [2020, 2021]
