from __future__ import annotations

import cgi
import json
import os
import tempfile
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

from formatter import convert_dtr


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_POST(self):
        if self.path != "/api/format":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

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
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(field.file.read())
            input_path = Path(tmp.name)

        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                output_path, _ = convert_dtr(input_path, original_name, Path(tmp_dir))
                body = output_path.read_bytes()
        except Exception as exc:
            self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        finally:
            try:
                input_path.unlink()
            except OSError:
                pass

        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{output_path.name}"')
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.startswith("/outputs/"):
            name = Path(unquote(self.path.removeprefix("/outputs/"))).name
            path = OUTPUT_DIR / name
            if not path.exists():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{name}"')
            self.send_header("Content-Length", str(path.stat().st_size))
            self.end_headers()
            with path.open("rb") as fh:
                self.wfile.write(fh.read())
            return

        if self.path == "/":
            self.path = "/index.html"
        super().do_GET()

    def _json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    port = int(os.environ.get("PORT", "5174"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"DTR Formatter running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
