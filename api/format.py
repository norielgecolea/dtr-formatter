from __future__ import annotations

import cgi
import json
import sys
import tempfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from formatter import convert_dtr  # noqa: E402


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self._json({"error": "Use POST with a DTR Excel file."}, HTTPStatus.METHOD_NOT_ALLOWED)

    def do_POST(self):
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
            },
        )

        field = form["file"] if "file" in form else None
        if field is None or not getattr(field, "filename", ""):
            self._json({"error": "Please upload a DTR Excel file."}, HTTPStatus.BAD_REQUEST)
            return

        original_name = Path(field.filename).name
        suffix = Path(original_name).suffix or ".xls"
        with tempfile.TemporaryDirectory(dir="/tmp") as tmp_dir:
            tmp_root = Path(tmp_dir)
            input_path = tmp_root / f"upload{suffix}"
            input_path.write_bytes(field.file.read())

            try:
                output_path, _ = convert_dtr(input_path, original_name, tmp_root)
                body = output_path.read_bytes()
            except Exception as exc:
                self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
                return

            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{output_path.name}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def _json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
