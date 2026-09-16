"""Desktop service layer for the GUI automation flow."""

from __future__ import annotations
import argparse
from dataclasses import dataclass
from pathlib import Path
from bdap_app.orchestration.application_service import AutomationRequest, BdapApplicationService


@dataclass(frozen=True)
class GuiFormData:
    workspace_root: Path
    comune: str
    explicit_years: tuple[int, ...]
    year_from: str
    year_to: str
    available_years: tuple[int, ...]


class DesktopAutomationService:
    """Desktop GUI to automation pipeline adapter."""

    def __init__(self, project_root: Path) -> None:
        self.app_service = BdapApplicationService(project_root=project_root)

    def list_comuni(self, workspace_root: Path) -> list[str]:
        """List comuni from workspace with inferred names."""
        from bdap_app.orchestration.cli_discovery import (
            _is_flat_bdap_folder,
            infer_comune_name_from_bdap_dir,
            normalize_token,
        )

        comuni_paths = self.app_service.list_comuni(workspace_root)
        result = []
        seen_names: set[str] = set()
        for path in comuni_paths:
            comune_name = self._display_comune_name(
                path=path,
                workspace_root=workspace_root,
                is_flat_bdap_folder=_is_flat_bdap_folder,
                infer_comune_name_from_bdap_dir=infer_comune_name_from_bdap_dir,
            )

            normalized_name = normalize_token(comune_name)
            if normalized_name and normalized_name not in seen_names:
                result.append(comune_name)
                seen_names.add(normalized_name)
        return result

    @staticmethod
    def _display_comune_name(
        *,
        path: Path,
        workspace_root: Path,
        is_flat_bdap_folder,
        infer_comune_name_from_bdap_dir,
    ) -> str:
        from bdap_app.orchestration.cli_discovery import normalize_token

        structural_folder = {"1doc", "doc1", "1documento", "documento1", "datibdap"}
        normalized_path_name = normalize_token(path.name)

        if normalized_path_name in structural_folder:
            if normalized_path_name == "datibdap":
                ancestor = path.parent.parent if path.parent.parent != path else path.parent
            else:
                ancestor = path.parent
            return ancestor.name if ancestor.name else path.name

        if is_flat_bdap_folder(path):
            inferred = infer_comune_name_from_bdap_dir(path)
            if inferred:
                return inferred

        if workspace_root.name and normalized_path_name == normalize_token(workspace_root.name):
            return workspace_root.name

        return path.name

    def discover_years(self, workspace_root: Path, comune: str | None = None) -> list[int]:
        """Discover available years, optionally filtered by comune."""
        if comune:
            return self.app_service.discover_years(workspace_root, comune)
        
        years = set()
        for base_dir in self.app_service.list_comuni(workspace_root):
            years.update(self.app_service.discover_years(workspace_root, base_dir.name))
        return sorted(years)

    def build_args_from_form(self, form: GuiFormData) -> argparse.Namespace:
        """Build automation args from GUI form data."""
        comune = form.comune.strip()
        if comune.startswith("[non selezionabile]"):
            raise ValueError(
                "Comune non valido.\n\n"
                "Carica una cartella con il nome del comune (es: Grandola, Sondrio) "
                "oppure una root con cartelle comuni BDAP."
            )

        selected_years = self._resolve_selected_years(
            available_years=form.available_years,
            explicit_years=form.explicit_years,
            year_from=form.year_from,
            year_to=form.year_to,
        )

        request = AutomationRequest(
            workspace_root=form.workspace_root,
            comune=comune,
            selected_years=selected_years,
            all_years=not bool(selected_years),
            analysis_sheet="analisi contabile",
            all_sheets=False,
        )
        return self.app_service.build_args(request)

    @staticmethod
    def _resolve_selected_years(
        *,
        available_years: tuple[int, ...],
        explicit_years: tuple[int, ...],
        year_from: str,
        year_to: str,
    ) -> tuple[int, ...]:
        """Resolve selected years from range or explicit selection."""
        year_from, year_to = year_from.strip(), year_to.strip()

        # Try range selection first
        if year_from and year_to:
            start, end = int(year_from), int(year_to)
            if start > end:
                raise ValueError("Anno da deve essere ≤ Anno a.")
            
            selected = tuple(y for y in available_years if start <= y <= end)
            if not selected:
                raise ValueError("Nessun file BDAP nell'intervallo selezionato.")
            return selected

        if bool(year_from) != bool(year_to):
            raise ValueError("Specifica sia Anno da che Anno a, oppure nessuno.")

        # Fallback to explicit selection
        return tuple(sorted(set(int(y) for y in explicit_years)))

    def run_pipeline(self, args: argparse.Namespace, log=None, progress=None) -> None:
        """Execute automation pipeline from parsed args."""
        emit = log or (lambda _: None)
        self.app_service.run(
            AutomationRequest(
                workspace_root=args.workspace_root,
                comune=args.comune,
                selected_years=tuple(int(y) for y in args.years.split(",")) if args.years else (),
                all_years=bool(args.all_years),
                template=args.template,
                analysis_sheet=args.analysis_sheet,
                all_sheets=bool(args.all_sheets),
                analysis=args.analysis,
                bdap_dir=args.bdap_dir,
                bdap=args.bdap,
                output=args.output,
            ),
            log=emit,
            progress=progress,
        )
