from typing import Callable

from api.internal.spreadsheet.services.api import CellsRange, Spreadsheet, SpreadsheetApiService
from api.internal.spreadsheet.table_abc import ReportRenderer


class GoogleSheetsRenderer(ReportRenderer):
    def __init__(
        self,
        report,
        spreadsheet_service: SpreadsheetApiService,
        sheet: Spreadsheet,
        format_config: list[tuple[Callable, dict[str, any]]],
    ):
        self.report = report
        self.spreadsheet = sheet
        self.spreadsheet_service = spreadsheet_service
        default_config = (
            spreadsheet_service.prepare_set_cells_format,
            {
                "spreadsheet_id": sheet.id,
                "sheet_id": sheet.sheet_id,
                "cells_range": CellsRange(start="A", end="Z"),
                "format_json": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}},
            },
        )
        self.format_config = [default_config] + format_config

    def format_sheet(self, sheet_name):
        report_data = make_list_of_dict(self.report)
        self.spreadsheet_service.prepare_update_rows(sheet_name, self.report.headers, report_data)
        for func, kwargs in self.format_config:
            func(**kwargs)
        self.spreadsheet_service.run_prepared(self.spreadsheet.id)

    def render(self, sheet_name=None) -> str:
        self.format_sheet(sheet_name)
        return self.spreadsheet_service.get_sheet_url(self.spreadsheet.id)


def make_list_of_dict(report) -> list[dict[str, str]]:
    rows = []
    for part in report.parts:
        data = part.get_data()
        for record in data:
            missing_columns = set(report.headers) - set(record.keys())
            for missing_column in missing_columns:
                record[missing_column] = None
        rows.extend(data)

    return rows
