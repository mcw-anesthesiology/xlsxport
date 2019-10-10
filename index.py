from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from http import HTTPStatus
from http.server import BaseHTTPRequestHandler

import json, os, re, tempfile

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class handler(BaseHTTPRequestHandler):
    def add_cors_headers(self, origin):
        self.send_header("Access-Control-Allowed-Methods", "POST")
        self.send_header("Access-Control-Allow-Origin", origin)

    def test_origin(self, origin):
        allowed_origins = os.getenv("ALLOWED_ORIGINS")
        if allowed_origins is None or allowed_origins == "*":
            return True
        else:
            allowed_origins = allowed_origins.split(" ")
            for allowed_origin in allowed_origins:
                if re.search(r"{}$".format(allowed_origin), origin):
                    return True

        return False

    def do_OPTIONS(self):
        origin = self.headers.get("origin")
        if self.test_origin(origin):
            self.add_cors_headers(origin)
        else:
            self.add_cors_headers(os.getenv("ALLOWED_ORIGINS"))

        self.send_response(HTTPStatus.OK)
        self.end_headers()

    def do_POST(self):
        origin = self.headers.get("origin")

        if not self.test_origin(origin):
            self.send_error(HTTPStatus.BAD_REQUEST)
            return

        content_length = self.headers.get("content-length")
        if not content_length:
            self.send_error(HTTPStatus.LENGTH_REQUIRED)
            return

        user_input = self.rfile.read(int(content_length))

        rows = json.loads(user_input)

        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Export")
        for in_row in rows:
            out_row = []
            for cell_value in in_row:
                if isinstance(cell_value, dict):
                    cell = WriteOnlyCell(ws, value=cell_value["value"])
                    style = cell_value.get("style")
                    if style:
                        cell.style = style
                    out_row.append(cell)
                else:
                    out_row.append(cell_value)
            ws.append(out_row)

        self.send_response(HTTPStatus.OK)
        self.add_cors_headers(origin)
        self.send_header("Content-Type", XLSX_MIME)
        self.send_header("Content-Disposition", 'attachment; filename="export.xlsx"')
        self.end_headers()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            wb.save(tmp)
            tmp.seek(0)
            self.wfile.write(tmp.read())

        return
