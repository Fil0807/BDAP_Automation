"""Servizio applicativo condiviso usato dai punti d'ingresso CLI/GUI/Web."""

from __future__ import annotations
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from .cli_discovery import (
    discover_bdap_files_by_year,
    discover_comune_base_dirs,
    infer_bdap_dir,
    resolve_base_dir_for_comune,
    resolve_default_template_path,
)
from .cli_pipeline import run_pipeline

# Oggetto richiesta usato per disaccoppiare l'interfaccia del servizio dal parsing degli argomenti CLI e dai default.
@dataclass(frozen=True)
class AutomationRequest:
    workspace_root: Path
    comune: str
    selected_years: tuple[int, ...] = ()
    all_years: bool = False
    template: Optional[Path] = None
    analysis_sheet: str = "analisi contabile"
    all_sheets: bool = True
    analysis: Optional[Path] = None
    bdap_dir: Optional[Path] = None
    bdap: Optional[Path] = None
    output: Optional[Path] = None


class BdapApplicationService:
    # Costruisce gli argomenti a runtime ed esegue la pipeline senza dipendere dai default del parser CLI.
    def __init__(self, project_root: Path) -> None:
        self.project_root = project_root

    def list_comuni(self, workspace_root: Path) -> list[Path]:
        return discover_comune_base_dirs(workspace_root)

    def discover_years(self, workspace_root: Path, comune: str) -> list[int]:
        base_dir = resolve_base_dir_for_comune(workspace_root, comune)
        bdap_dir = infer_bdap_dir(base_dir)
        return sorted(discover_bdap_files_by_year(bdap_dir).keys())

    def build_args(self, request: AutomationRequest) -> argparse.Namespace:
        comune = request.comune.strip()
        if not comune:
            raise ValueError("Inserire il nome del comune")

        base_dir = resolve_base_dir_for_comune(request.workspace_root, comune)
        bdap_dir = request.bdap_dir or infer_bdap_dir(base_dir)
        discovered = discover_bdap_files_by_year(bdap_dir)
        default_bdap = next(iter(discovered.values()), bdap_dir / "Rend.xlsx")

        # Always compile starting from the project template to avoid
        # reusing previously compiled analysis workbooks.
        template = resolve_default_template_path(
            workspace_root=request.workspace_root,
            project_root=self.project_root,
        )
        if template is None:
            raise FileNotFoundError(
                "Template di default non trovato. Verifica che sia presente Template_Analisi.xlsx."
            )
        analysis = Path(template)

        output = request.output
        if output is None:
            template_name = Path(template)
            template_suffix = template_name.suffix or ".xlsx"
            output = base_dir / f"{template_name.stem}_compilato{template_suffix}"

        selected_years = tuple(sorted(set(int(year) for year in request.selected_years)))
        years_arg = ",".join(str(year) for year in selected_years) if selected_years else None

        # Build argparse.Namespace to pass to pipeline, ensuring all necessary fields are populated.
        args = argparse.Namespace()
        args.workspace_root = request.workspace_root
        args.comune = comune
        args.analysis = Path(analysis)
        args.template = Path(template)
        args.bdap_dir = Path(bdap_dir)
        args.bdap = Path(request.bdap) if request.bdap is not None else Path(default_bdap)
        args.output = Path(output)
        args.analysis_sheet = request.analysis_sheet
        args.all_sheets = bool(request.all_sheets)
        args.years = years_arg
        args.year = None
        args.all_years = bool(request.all_years) and not bool(years_arg)
        return args

    def run(
        self,
        request: AutomationRequest,
        log: Optional[Callable[[str], None]] = None,
        progress: Optional[Callable[[float, int, str], None]] = None,
    ) -> argparse.Namespace:
        args = self.build_args(request)
        emit = log if log is not None else (lambda _msg: None)
        run_pipeline(args, log=emit, progress=progress)
        return args
