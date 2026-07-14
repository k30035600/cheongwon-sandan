#!/usr/bin/env python3
"""로컬 HTTP 서버 — / · /청원지구_포털 → 청원지구_포털.html · MD 저장 API"""
from __future__ import annotations

import json
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


def resolve_md_rel(rel: str) -> Path | None:
    rel = unquote(rel).replace("\\", "/").lstrip("/")
    if not rel or ".." in rel.split("/"):
        return None
    if not rel.lower().endswith(".md"):
        return None
    full = (ROOT / rel).resolve()
    root_res = ROOT.resolve()
    try:
        full.relative_to(root_res)
    except ValueError:
        return None
    return full


class PortalHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def _send_json(self, code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self) -> bytes:
        n = int(self.headers.get("Content-Length", "0") or "0")
        return self.rfile.read(n) if n else b""

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

    def do_PUT(self) -> None:
        if norm_path(self.path) != "/api/md":
            self.send_error(404)
            return
        try:
            data = json.loads(self._read_body().decode("utf-8"))
            rel = str(data.get("path", ""))
            content = data.get("content")
            if content is None:
                self._send_json(400, {"ok": False, "error": "content 없음"})
                return
            target = resolve_md_rel(rel)
            if not target:
                self._send_json(400, {"ok": False, "error": "허용되지 않는 경로"})
                return
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(str(content), encoding="utf-8", newline="\n")
            self._send_json(200, {"ok": True, "path": rel.replace("\\", "/")})
        except json.JSONDecodeError:
            self._send_json(400, {"ok": False, "error": "JSON 오류"})
        except OSError as e:
            self._send_json(500, {"ok": False, "error": str(e)})

    def do_OPTIONS(self) -> None:
        if norm_path(self.path) == "/api/md":
            self.send_response(204)
            self.send_header("Allow", "PUT, OPTIONS")
            self.end_headers()
            return
        self.send_error(404)

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
