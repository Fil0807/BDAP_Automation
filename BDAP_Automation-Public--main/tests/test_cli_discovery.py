from openpyxl import Workbook

from bdap_app.orchestration.cli_discovery import discover_bdap_files_by_year


def _save_workbook(path) -> None:
    wb = Workbook()
    wb.save(path)


def test_discover_bdap_files_ignores_questionari_and_controlli_post(tmp_path) -> None:
    rend_2021 = tmp_path / "2021_rend_ind_658142930546394001_676077.xlsx"
    rend_2024 = tmp_path / "2024_rend_ind_sintetici.xlsx"
    for path in (
        rend_2021,
        tmp_path / "Questionario Bilancio Enti locali Consuntivo 2021.xlsx",
        tmp_path / "Questionario Debiti Fuori Bilancio Periodico 2021.xlsx",
        tmp_path / "Controlli Post - Questionario Bilancio Enti locali Consuntivo 2021.xlsx",
        tmp_path / "2024_Questionario Bilancio Enti locali Consuntivo 2024.xlsx",
        rend_2024,
    ):
        _save_workbook(path)

    discovered = discover_bdap_files_by_year(tmp_path)

    assert discovered == {
        2021: rend_2021,
        2024: rend_2024,
    }


def test_discover_bdap_files_prefers_strict_rend_filename(tmp_path) -> None:
    generic_rend = tmp_path / "2024_rend_ind_sintetici.xlsx"
    strict_rend = tmp_path / "Rend. 2024.xlsx"
    _save_workbook(generic_rend)
    _save_workbook(strict_rend)

    assert discover_bdap_files_by_year(tmp_path)[2024] == strict_rend
