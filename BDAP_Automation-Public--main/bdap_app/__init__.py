"""Locatelli: BDAP → Analisi Contabile automation (GUI Desktop)."""

# Expose primary modules and classes.
from .core.automation import fill_accounting_analysis
from .orchestration.application_service import BdapApplicationService, AutomationRequest
from .orchestration.cli_pipeline import run_pipeline

__all__ = [
    "fill_accounting_analysis",
    "BdapApplicationService",
    "AutomationRequest",
    "run_pipeline",
]
