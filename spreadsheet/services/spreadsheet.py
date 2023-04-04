from typing import Any, Callable

from api.internal.spreadsheet.renderer import GoogleSheetsRenderer

# from api.internal.spreadsheet.services.api import GoogleSheetsRenderer
from api.internal.spreadsheet.services.api import Spreadsheet
from api.internal.spreadsheet.table_abc import Table

# from api.internal.spreadsheet.services.api import Report


class SpreadsheetService:
    def __init__(self, _spreadsheets_api):
        self._spreadsheets_api = _spreadsheets_api

    def create_spreadsheet(self, domain_permission, filename, report):
        sheet = self._spreadsheets_api.create_spreadsheet(filename, domain_permission)
        spreadsheet_format = report.get_google_spreadsheet_config(sheet, self._spreadsheets_api)
        url = self._get_report_link(sheet, spreadsheet_format, report)
        return url

    def _get_report_link(
        self, spreadsheet: Spreadsheet, format_config: list[tuple[Callable, dict[str, Any]]], report: Table
    ):
        sheets_renderer = GoogleSheetsRenderer(
            report=report,
            spreadsheet_service=self._spreadsheets_api,
            sheet=spreadsheet,
            format_config=format_config,
        )
        content = sheets_renderer.render(spreadsheet.sheet_title)
        return content
