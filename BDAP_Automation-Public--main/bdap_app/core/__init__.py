"""Core automation engine for BDAP Excel compilation."""

from .automation import fill_accounting_analysis
from .automation_fcde import populate_fcde_sheet, populate_fcde_sheet_in_workbook, read_fcde_percentage_from_workbook

__all__ = [
	"fill_accounting_analysis",
	"populate_fcde_sheet",
	"populate_fcde_sheet_in_workbook",
	"read_fcde_percentage_from_workbook",
]
