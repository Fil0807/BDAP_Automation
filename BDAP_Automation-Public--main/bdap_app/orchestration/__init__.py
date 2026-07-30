"""Orchestration and discovery modules for pipeline execution."""

from .application_service import BdapApplicationService, AutomationRequest
from .cli_discovery import discover_comune_base_dirs, resolve_base_dir_for_comune
from .cli_pipeline import run_pipeline

__all__ = [
    "BdapApplicationService",
    "AutomationRequest",
    "discover_comune_base_dirs",
    "resolve_base_dir_for_comune",
    "run_pipeline",
]
