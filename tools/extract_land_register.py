#!/usr/bin/env python3
"""토지·임야대장 PDF OCR → 토지조서 HTML 갱신용 데이터 추출.
스캔 PDF는 EasyOCR 사용(최초 1회 모델 다운로드)."""
from __future__ import annotations
import io
import re
import sys
from pathlib import Path
import fitz
import numpy as np
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "01_토지조서"
OUT_TXT = ROOT / "tools" / "_land_ocr"

# 파일명 → 지번
LOT_FROM_NAME = {
    "임야대장_청원리 산175-1.pdf": ("산175-1", "임야대장"),
    "임야대장_청원리 산175-2.pdf": ("산175-2", "임야대장"),
    "임야대장_청원리 산175.pdf": ("산175", "임야대장"),
    "임야대장_청원리 산177-1.pdf": ("산177-1", "임야대장"),
    "임야대장_청원리 산177-4.pdf": ("산177-4", "임야대장"),
    "임야대장_청원리 산177-5.pdf": ("산177-5", "임야대장"),
    "임야대장_청원리 산178.pdf": ("산178", "임야대장"),
    "임야대장_청원리 산183.pdf": ("산183", "임야대장"),
    "토지대장_청원리 513-8.pdf": ("513-8", "토지대장"),
    "토지대장_청원리 513-9.pdf": ("513-9", "토지대장"),
}


def ocr_page(pdf: Path, reader) -> str:
    doc = fitz.open(pdf)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    text = "\n".join(reader.readtext(np.array(img), detail=0, paragraph=True))
    doc.close()
    return text


def parse_register(text: str, lot: str) -> dict:
    d: dict = {"lot": lot, "raw_snip": text[:500]}
    m = re.search(r"4159131025-\d+-\d+", text.replace(" ", ""))
    if m:
        d["pnu"] = m.group(0)
    # 지목
    if re.search(r"\(05\)\s*임야|지목[^\n]*임야", text):
        d["jimok"] = "임야"
    elif re.search(r"\(01\)\s*전|\(511\)", text):
        d["jimok"] = "전"
    else:
        d["jimok"] = "확인"
    # 면적 — *9 , 535*1 / *10,880*1 / 568* 패턴
    areas = []
    for m in re.finditer(r"\*?\s*([\d,]+)\s*\*?\s*1?", text):
        s = m.group(1).replace(",", "")
        if s.isdigit():
            v = int(s)
            if 100 <= v <= 50000:
                areas.append(v)
    # 최신 면적: 변동이력 중 큰 값 또는 마지막 (51)2026 근처
    if areas:
        d["area_m2"] = max(areas[-3:]) if len(areas) >= 3 else max(areas)
    # 소유자 — 마지막 (03) 소유권이전 계열
    owners = []
    for m in re.finditer(
        r"\(03\)\s*소유권이\s*전\s*([가-힣A-Za-z\s]+?(?:주식회사|회사|외\s*\d+인)?)\s*(\d{6}-[\d\*]+)?",
        text,
    ):
        name = re.sub(r"\s+", "", m.group(1))
        if len(name) >= 2:
            owners.append(name)
    for m in re.finditer(r"국\s*\(\s*산림\s*청\s*\)", text):
        owners.append("국(산림청)")
    if owners:
        d["owner"] = owners[-1]
    elif "심재필" in text:
        d["owner"] = "심재필"
    # 2026 개별공시지가
    prices = [int(x.replace(",", "").replace(" ", "")) for x in re.findall(r"(\d{2,3}\s*,?\s*\d{3})", text) if 30000 <= int(x.replace(",", "").replace(" ", "")) <= 90000]
    if prices:
        d["price_2026"] = prices[-1]
    return d


def main() -> None:
    try:
        import easyocr
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "easyocr", "-q"])
        import easyocr

    reader = easyocr.Reader(["ko", "en"], gpu=False, verbose=False)
    OUT_TXT.mkdir(parents=True, exist_ok=True)
    rows = []
    for pdf in sorted(BASE.glob("*.pdf")):
        lot, kind = LOT_FROM_NAME.get(pdf.name, (pdf.stem, "?"))
        text = ocr_page(pdf, reader)
        (OUT_TXT / f"{lot.replace('/', '_')}.txt").write_text(text, encoding="utf-8")
        row = parse_register(text, lot)
        row["kind"] = kind
        row["source"] = pdf.name
        rows.append(row)
        print(f"{lot:10} | {row.get('jimok','?'):4} | {row.get('area_m2','?'):>6} | {row.get('owner','?')}")

    total = sum(r.get("area_m2", 0) for r in rows if isinstance(r.get("area_m2"), int))
    print(f"\n지적공부 면적 합계(10필지): {total:,} ㎡")


if __name__ == "__main__":
    main()
