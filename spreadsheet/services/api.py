import apiclient.discovery
import httplib2
from oauth2client.service_account import ServiceAccountCredentials
from pydantic import BaseModel, validator

from config.settings import GOOGLE_CREDS


class CellsRange(BaseModel):
    end: str
    start: str

    @validator("start")
    def start_lte_end(cls, v, values):
        start_letter = v[0]
        end_letter = values["end"][0]
        if ord(start_letter) > ord(end_letter):
            raise ValueError("Start cell should be less than or equal end cell")
        return v

    def __str__(self):
        return f"{self.start}:{self.end}"

    def to_dict(self, spreadsheet_id: str, sheet_id: int) -> dict[str, int]:
        if spreadsheet_id is None:
            raise SheetNotSetError()
        start_cell, end_cell = self.start, self.end
        cells_range = {}
        range_az = range(ord("A"), ord("Z") + 1)
        if ord(start_cell[0]) in range_az and ord(end_cell[0]) in range_az:
            cells_range["startColumnIndex"] = ord(start_cell[0]) - ord("A")
            start_cell = start_cell[1:]
            cells_range["endColumnIndex"] = ord(end_cell[0]) - ord("A") + 1
            end_cell = end_cell[1:]
        if len(start_cell) and len(end_cell) > 0:
            cells_range["startRowIndex"] = int(start_cell) - 1
            cells_range["endRowIndex"] = int(end_cell)
        cells_range["sheetId"] = sheet_id
        return cells_range


class Spreadsheet:
    def __init__(self, spreadsheet_id: str, sheet_id: int, sheet_title: str):
        self.id = spreadsheet_id
        self.sheet_id = sheet_id
        self.sheet_title = sheet_title


def html_color_to_json(html_color):
    if html_color.startswith("#"):
        html_color = html_color[1:]
    return {
        "red": int(html_color[0:2], 16) / 255.0,
        "green": int(html_color[2:4], 16) / 255.0,
        "blue": int(html_color[4:6], 16) / 255.0,
    }


class GoogleDrivePermission:
    def __init__(self, permission_type, role, domain=None, email_address=None):
        self.type = permission_type
        self.role = role
        self.domain = domain
        self.email_address = email_address

    def to_dict(self):
        result = {
            "type": self.type,
            "role": self.role,
        }
        if self.type == "domain" and self.domain is not None:
            result["domain"] = self.domain

        if self.type == "user" and self.email_address is not None:
            result["emailAddress"] = self.email_address

        return result


class SpreadsheetError(Exception):
    pass


class SpreadsheetNotSetError(SpreadsheetError):
    pass


class SheetNotSetError(SpreadsheetError):
    pass


class SpreadsheetApiService:
    def __init__(self):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(
            GOOGLE_CREDS, ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        self.http_auth = self.credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build("sheets", "v4", http=self.http_auth)
        self.drive_service = None
        self.requests = []
        self.value_ranges = []

    def create(
        self,
        title: str,
        sheet_title: str,
        rows: int = 1000,
        cols: int = 26,
        locale: str = "en_US",
        time_zone: str = "Etc/GMT",
    ) -> Spreadsheet:
        spreadsheet = (
            self.service.spreadsheets()
            .create(
                body={
                    "properties": {"title": title, "locale": locale, "timeZone": time_zone},
                    "sheets": [
                        {
                            "properties": {
                                "sheetType": "GRID",
                                "sheetId": 0,
                                "title": sheet_title,
                                "gridProperties": {"rowCount": rows, "columnCount": cols},
                            }
                        }
                    ],
                }
            )
            .execute()
        )
        return Spreadsheet(
            spreadsheet["spreadsheetId"],
            spreadsheet["sheets"][0]["properties"]["sheetId"],
            spreadsheet["sheets"][0]["properties"]["title"],
        )

    def share(self, spreadsheet_id: str, permission: GoogleDrivePermission):
        if spreadsheet_id is None:
            raise SpreadsheetNotSetError()
        if self.drive_service is None:
            self.drive_service = apiclient.discovery.build("drive", "v3", http=self.http_auth)
        self.drive_service.permissions().create(fileId=spreadsheet_id, body=permission.to_dict(), fields="id").execute()

    @staticmethod
    def get_sheet_url(spreadsheet_id: str):
        if spreadsheet_id is None:
            raise SpreadsheetNotSetError()
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/"

    def run_prepared(self, spreadsheet_id: str, value_input_option: str = "USER_ENTERED") -> tuple[list, list]:
        if spreadsheet_id is None:
            raise SpreadsheetNotSetError()
        upd1_res = {"replies": []}
        upd2_res = {"responses": []}
        try:
            if len(self.requests) > 0:
                upd1_res = (
                    self.service.spreadsheets()
                    .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": self.requests})
                    .execute()
                )

            if len(self.value_ranges) > 0:
                upd2_res = (
                    self.service.spreadsheets()
                    .values()
                    .batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body={"valueInputOption": value_input_option, "data": self.value_ranges},
                    )
                    .execute()
                )
        finally:
            self.requests = []
            self.value_ranges = []
        return upd1_res["replies"], upd2_res["responses"]

    def prepare_add_sheet(self, sheet_title: str, rows: int = 1000, cols: int = 26):
        self.requests.append(
            {
                "addSheet": {
                    "properties": {"title": sheet_title, "gridProperties": {"rowCount": rows, "columnCount": cols}}
                }
            }
        )

    def add_sheet(self, spreadsheet: Spreadsheet, sheet_title: str, rows: int = 1000, cols: int = 26) -> int:
        if spreadsheet.id is None:
            raise SpreadsheetNotSetError()
        self.prepare_add_sheet(sheet_title, rows, cols)
        added_sheet = self.run_prepared(spreadsheet.id)[0][0]["addSheet"]["properties"]
        spreadsheet.sheet_id = added_sheet["sheetId"]
        spreadsheet.sheet_title = added_sheet["title"]
        return spreadsheet.sheet_id

    def prepare_set_dimension_pixel_size(self, sheet_id: int, dimension: str, start: int, end: int, pixel_size: int):
        if sheet_id is None:
            raise SheetNotSetError()
        self.requests.append(
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": dimension, "startIndex": start, "endIndex": end},
                    "properties": {"pixelSize": pixel_size},
                    "fields": "pixelSize",
                }
            }
        )

    def prepare_set_columns_width(self, *, sheet_id: int, start: int, end: int, width: int):
        self.prepare_set_dimension_pixel_size(sheet_id, "COLUMNS", start, end + 1, width)

    def prepare_set_rows_height(self, *, sheet_id: int, start: int, end: int, height: int):
        self.prepare_set_dimension_pixel_size(sheet_id, "ROWS", start, end + 1, height)

    def prepare_set_values(
        self, sheet_title: str, cells_range: CellsRange, values: list[list[any]], major_dimension: str = "ROWS"
    ):
        if sheet_title is None:
            raise SheetNotSetError()
        cells_range = str(cells_range)
        self.value_ranges.append(
            {"range": f"{sheet_title}!{cells_range}", "majorDimension": major_dimension, "values": values}
        )

    def prepare_merge_cells(
        self, spreadsheet_id: str, sheet_id: int, cells_range: CellsRange, merge_type: str = "MERGE_ALL"
    ):
        self.requests.append(
            {
                "mergeCells": {
                    "range": cells_range.to_dict(spreadsheet_id, sheet_id),
                    "mergeType": merge_type,
                }
            }
        )

    def prepare_set_cells_format(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        cells_range: CellsRange,
        format_json: dict,
        fields: str = "userEnteredFormat",
    ):

        self.requests.append(
            {
                "repeatCell": {
                    "range": cells_range.to_dict(spreadsheet_id, sheet_id),
                    "cell": {"userEnteredFormat": format_json},
                    "fields": fields,
                }
            }
        )

    def prepare_append_rows(
        self, spreadsheet_id: str, sheet_name, header: list[str], data: list[dict], start_cells: CellsRange = ""
    ):
        values_to_add = [header] + [[o[e] for e in header] for o in data]
        range = f"'{sheet_name}'"

        (
            self.service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption="USER_ENTERED",
                body={"values": values_to_add},
            )
            .execute()
        )

    def create_spreadsheet(
        self, filename: str, domain_permission: GoogleDrivePermission, main_sheet_name: str = "Основная"
    ):
        spreadsheet = self.create(filename, main_sheet_name)
        self.share(spreadsheet.id, domain_permission)
        self.run_prepared(spreadsheet.id)

        return spreadsheet

    def get_spreadsheet_by_id(self, spreadsheet_id):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return Spreadsheet(
            spreadsheet["spreadsheetId"],
            spreadsheet["sheets"][0]["properties"]["sheetId"],
            spreadsheet["sheets"][0]["properties"]["title"],
        )

    def prepare_change_spreadsheet_title(self, title):
        self.requests.append({"updateSpreadsheetProperties": {"properties": {"title": title}, "fields": "title"}})

    def prepare_update_rows(self, sheet_name, header: list[str], data: list[dict], major_dimension: str = "ROWS"):
        values_to_add = [header] + [[o[e] for e in header] for o in data]
        range = f"'{sheet_name}'"

        self.value_ranges.append({"range": f"{range}", "majorDimension": major_dimension, "values": values_to_add})
