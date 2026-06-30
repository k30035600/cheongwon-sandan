#!/usr/bin/env python3
"""로컬 HTTP 서버 — / · /청원지구_포털 → 청원지구_포털.html"""
from __future__ import annotations

import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
PORT = 8080
PORTAL_FILE = "청원지구_포털.html"
PORTAL_URL = "/" + quote(PORTAL_FILE, safe="")
ENTRY_PATHS = frozenset({
    "/",
    "/index.html",
    "/청원지구_포털",
    f"/{PORTAL_FILE.replace('.html', '')}",
})


def norm_path(raw: str) -> str:
    path = unquote(raw.split("?", 1)[0])
    if not path.startswith("/"):
        path = "/" + path
    if len(path) > 1 and path.endswith("/"):
        path = path.rstrip("/")
    return path


class PortalHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def do_GET(self) -> None:
        path = norm_path(self.path)
        if path in ENTRY_PATHS:
            qs = self.path.split("?", 1)[1] if "?" in self.path else ""
            loc = PORTAL_URL + (f"?{qs}" if qs else "")
            self.send_response(302)
            self.send_header("Location", loc)
            self.end_headers()
            return
        super().do_GET()

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("%s - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), fmt % args))


def main() -> None:
    host = "127.0.0.1"
    with ThreadingHTTPServer((host, PORT), PortalHandler) as httpd:
        url = f"http://localhost:{PORT}/청원지구_포털"
        print("")
        print(f"  {url}")
        print("  (종료: Ctrl+C)")
        print("")
        httpd.serve_forever()


if __name__ == "__main__":
    main()
