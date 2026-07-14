"""청원지구 부동산개발 절차서 PDF 생성 (사업수지분석표 부록 병합).

- 절차서 md → PDF (md_to_pdf, 표지 포함)
- 사업수지분석표 html → PDF (Playwright, .wrap 페이지 그대로)
- 두 PDF를 병합해 절차서 PDF 뒤에 사업수지분석표를 부록으로 첨부

사용:
  python -X utf8 tools/build_procedure_pdf.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).resolve().parent))

import md_to_pdf  # noqa: E402
from pypdf import PdfReader, PdfWriter  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
BASE = ROOT / "00_부동산개발"
MD = BASE / "청원지구_부동산개발_절차서.md"
SUJI_HTML = BASE / "청원지구_사업수지분석표(추정).html"
OUT = BASE / "청원지구_부동산개발_절차서.pdf"


def html_to_pdf(html_path: Path, pdf_path: Path) -> None:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(html_path.as_uri(), wait_until="networkidle")
        page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            prefer_css_page_size=True,
        )
        browser.close()


def main() -> int:
    if not MD.is_file():
        print(f"오류: 절차서 md 없음 {MD}", file=sys.stderr)
        return 1
    if not SUJI_HTML.is_file():
        print(f"오류: 사업수지 html 없음 {SUJI_HTML}", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory(prefix="procpdf_") as tmp:
        tmp = Path(tmp)
        proc_pdf = tmp / "proc.pdf"
        suji_pdf = tmp / "suji.pdf"

        md_to_pdf.convert(MD, proc_pdf)
        html_to_pdf(SUJI_HTML, suji_pdf)

        writer = PdfWriter()
        for pdf in (proc_pdf, suji_pdf):
            reader = PdfReader(str(pdf))
            for pg in reader.pages:
                writer.add_page(pg)
        with open(OUT, "wb") as fp:
            writer.write(fp)

    n = len(PdfReader(str(OUT)).pages)
    print(f"OK  {OUT}  ({n}p, 사업수지 부록 포함)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
