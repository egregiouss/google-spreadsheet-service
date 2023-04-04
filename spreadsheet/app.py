from api.internal.spreadsheet.services.api import SpreadsheetApiService
from api.internal.spreadsheet.services.spreadsheet import SpreadsheetService

spreadsheet_api_service = SpreadsheetApiService()
spreadsheet_service = SpreadsheetService(spreadsheet_api_service)
