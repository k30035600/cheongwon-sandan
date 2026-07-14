#!/usr/bin/env python3
"""포털 본문 중 작성 표기를 삼아건설(주)로 통일."""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
PORTAL = ROOT / "청원지구_포털.html"
AUTHOR = "삼아건설(주)"


def collect_mains(html: str) -> set[Path]:
    paths: set[str] = set()
    for m in re.finditer(r'main:\s*"([^"]+)"', html):
        paths.add(m.group(1).split("?")[0])
    for m in re.finditer(r'main:\s*JIN_BASE\s*\+\s*"([^"]+)"', html):
        paths.add("00_진명개발(주)" + m.group(1))
    for m in re.finditer(r'main:\s*B\s*\+\s*"([^"]+)"', html):
        paths.add("05_내역서" + m.group(1))
    for m in re.finditer(r'main:\s*JMY_BASE\s*\+\s*"([^"]+)"', html):
        paths.add("09_공사지명원" + m.group(1))
    for m in re.finditer(r'main:\s*JMY_BASE\s*\+\s*"/"\s*\+\s*toMd\(html\)', html):
        pass
    # workPath(공종, file)
    for m in re.finditer(r'workPath\("([^"]+)",\s*"([^"]+)"\)', html):
        paths.add(f"05_내역서/내역서작업/{m.group(1)}/{m.group(2)}")
    # drop nav md paths commonly linked as main via toMd
    for name in (
        "건설업_안내.md",
        "계약검토_안내.md",
        "관계도_안내.md",
        "하도급사_시공참여조건.md",
        "외부수집/외부수집_안내.md",
        "지명원_근일건설/근일건설_지명요약.md",
        "지명원_희상건설/희상건설_지명요약.md",
    ):
        paths.add("09_공사지명원/" + name)
    out: set[Path] = set()
    for p in paths:
        fp = ROOT / p
        if fp.is_file():
            out.add(fp)
        # html twin of md
        if p.endswith(".md"):
            h = ROOT / p.replace(".md", ".html")
            if h.is_file():
                out.add(h)
    # 분묘 + 절차서 명시
    for extra in (
        "01_토지조서/청원지구_분묘기지권.md",
        "00_부동산개발/청원지구_부동산개발_절차서.md",
        "00_부동산개발/청원지구_사업수지분석표(추정).html",
        "07_타견적/실행예산서/토목조경_실행예산서_보고서.html",
        "09_공사지명원/공사비PF_준비서류.html",
    ):
        fp = ROOT / extra
        if fp.is_file():
            out.add(fp)
    return out


def patch_text(text: str, path: Path) -> str | None:
    orig = text
    # 표/필드: 작성 | 삼아건설 → 삼아건설(주)
    text = text.replace("| **작성** | **삼아건설** |", f"| **작성** | **{AUTHOR}** |")
    text = text.replace("삼아건설 작성", f"{AUTHOR} 작성")
    # 이미 (주)가 있으면 중복 방지
    text = text.replace(f"{AUTHOR}(주)", AUTHOR)
    text = text.replace("삼아건설(주)(주)", AUTHOR)

    # 머리말 **작성일** … 뒤에 작성이 없으면 추가 (md)
    if path.suffix == ".md" and re.search(r"\*\*작성일\*\*", text) and "작성자" not in text and f"**작성** | **{AUTHOR}**" not in text:
        # 첫 메타 줄에 | **작성** 삼아건설(주) 추가
        def add_author_line(m: re.Match) -> str:
            line = m.group(0)
            if AUTHOR in line or "**작성**" in line:
                return line
            return line.rstrip() + f" | **작성** {AUTHOR}"

        text = re.sub(
            r"^\*\*작성일\*\*[^\n]+$",
            add_author_line,
            text,
            count=1,
            flags=re.M,
        )

    # html meta: <b>작성일</b> … → 작성 추가
    if path.suffix == ".html" and "<b>작성일</b>" in text and AUTHOR not in text.split("<b>작성일</b>")[1][:200]:
        text = re.sub(
            r"(<b>작성일</b>\s*[^<\n]+)",
            rf"\1 &nbsp;|&nbsp; <b>작성</b> {AUTHOR}",
            text,
            count=1,
        )

    # 말미 "YYYY. M. D. 작성. 끝." → 삼아건설(주) 작성
    text = re.sub(
        r"(\d{4}\.\s*\d{1,2}\.\s*\d{1,2}\.)\s*작성(\.?\s*끝\.)",
        rf"\1 {AUTHOR} 작성\2",
        text,
    )
    # "작성 2026. …" 꼬리 (이미 기관 없으면)
    text = re.sub(
        r"(?<![가-힣\)])작성\s+(20\d{2}\.\s*\d{1,2}(?:\.\s*\d{1,2}\.)?)",
        rf"{AUTHOR} 작성 \1",
        text,
    )

    # 절차서 표지: 작성 | 날짜 → 작성자 행 추가
    if "부동산개발_절차서" in path.name and "작성자" not in text:
        old = (
            '<tr><td style="padding:6px 16px;color:#5a6675;text-align:right;border:none;">작성</td>'
            '<td style="padding:6px 16px;border:none;">2026. 7.</td></tr>'
        )
        new = (
            f'<tr><td style="padding:6px 16px;color:#5a6675;text-align:right;border:none;">작성</td>'
            f'<td style="padding:6px 16px;font-weight:700;border:none;">{AUTHOR}</td></tr>'
            '<tr><td style="padding:6px 16px;color:#5a6675;text-align:right;border:none;">작성일</td>'
            '<td style="padding:6px 16px;border:none;">2026. 7.</td></tr>'
        )
        if old in text:
            text = text.replace(old, new)

    if text == orig:
        return None
    return text


def main() -> int:
    html = PORTAL.read_text(encoding="utf-8")
    files = collect_mains(html)
    n = 0
    for fp in sorted(files):
        try:
            t = fp.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        # 작성 관련 표기가 있는 파일만
        if not re.search(r"작성", t):
            continue
        new = patch_text(t, fp)
        if new is None:
            continue
        fp.write_text(new, encoding="utf-8")
        print(f"  {fp.relative_to(ROOT)}")
        n += 1

    # 포털 푸터
    foot_old = "화성 청원지구 포털 · 좌측 본문 · 우측 근거자료 · 2026. 6. 29."
    foot_new = f"화성 청원지구 포털 · 좌측 본문 · 우측 근거자료 · 작성 {AUTHOR} · 2026. 6. 29."
    if foot_old in html and AUTHOR not in html[html.find("footer.bar") : html.find("footer.bar") + 500]:
        html2 = html.replace(foot_old, foot_new)
        # safer: replace footer content
        html2 = re.sub(
            r'(<footer class="bar">)[^<]+(</footer>)',
            rf"\1화성 청원지구 포털 · 좌측 본문 · 우측 근거자료 · 작성 {AUTHOR} · 2026. 6. 29.\2",
            html,
            count=1,
        )
        PORTAL.write_text(html2, encoding="utf-8")
        print("  청원지구_포털.html (footer)")
        n += 1

    print(f"갱신 {n}건")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
