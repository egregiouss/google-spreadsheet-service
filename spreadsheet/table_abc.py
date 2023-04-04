from abc import ABC, abstractmethod

from api.internal.spreadsheet.table_part import TablePart


class Table(ABC):
    def __init__(self):
        self.headers: list[str] = self._get_headers()
        self.parts: list[TablePart] = []

    def parts(self) -> list[TablePart]:
        """Блоки таблицы отчета"""

    @abstractmethod
    def _get_headers(self):
        """Заголовки таблицы отчета"""


class ReportRenderer(ABC):
    @abstractmethod
    def render(self) -> str:
        """Сформировать представление отчета"""
