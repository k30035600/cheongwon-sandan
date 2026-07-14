"""정적 HTML 문서를 순수 Markdown(.md)으로 일괄 변환.

- 표준 라이브러리(html.parser)만 사용 — 외부 의존 없음.
- <head>/<script>/<style>/<nav>/<footer> 등 비본문 제거.
- 제목·문단·목록·표(GFM, colspan/rowspan 확장)·인용·코드·링크·강조 변환.
- 플로우차트 등 CSS 전용 장식(화살표)은 텍스트로 남지 않음 → 필요한 문서는 수기 보정.

사용:  python -X utf8 tools/html_to_md.py            (기본 목록 전체 변환)
       python -X utf8 tools/html_to_md.py <파일...>   (지정 파일만 변환)
"""
import sys
import re
from pathlib import Path
from html.parser import HTMLParser
from html import unescape

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent

# 변환 대상(정적 문서). 자동 생성 문서는 제외.
DEFAULT_TARGETS = [
    "00_진명개발(주)/진명개발_팩트체크.html",
    "00_부동산개발/청원지구_부동산개발_절차서.html",
    "01_토지조서/청원지구_토지조서.html",
    "02_환경영향평가/청원지구_전략환경영향평가_요약.html",
    "03_지구단위계획/청원지구_지구단위계획_관련자료.html",
    "04_인허가도면/인허가도면_안내.html",
    "05_내역서/공내역서/공내역서_안내.html",
    "09_공사지명원/건설업_안내.html",
    "09_공사지명원/관계도_안내.html",
    "09_공사지명원/희상건설_지명요약.html",
    "09_공사지명원/근일건설_지명요약.html",
    "09_공사지명원/하도급사_시공참여조건.html",
    "09_공사지명원/계약검토_안내.html",
    "09_공사지명원/외부수집/외부수집_안내.html",
]

VOID = {"br", "img", "hr", "meta", "link", "input", "col", "area", "base"}
SKIP_TREE = {"script", "style", "head", "noscript", "svg"}
BLOCK = {"p", "div", "section", "article", "header", "footer", "table", "ul",
         "ol", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote",
         "pre", "thead", "tbody", "tfoot", "figure", "figcaption"}


class Node:
    __slots__ = ("tag", "attrs", "children", "text")

    def __init__(self, tag, attrs=None):
        self.tag = tag
        self.attrs = dict(attrs or {})
        self.children = []
        self.text = None  # 텍스트 노드일 때만


class DOM(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("#root")
        self.stack = [self.root]
        self.skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if self.skip_depth:
            if tag not in VOID:
                self.skip_depth += 1
            return
        if tag in SKIP_TREE:
            self.skip_depth = 1
            return
        node = Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        if self.skip_depth or tag in SKIP_TREE:
            return
        self.stack[-1].children.append(Node(tag, attrs))

    def handle_endtag(self, tag):
        if self.skip_depth:
            self.skip_depth -= 1
            return
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break

    def handle_data(self, data):
        if self.skip_depth:
            return
        t = Node("#text")
        t.text = data
        self.stack[-1].children.append(t)


def _cls(node):
    return node.attrs.get("class", "")


# ── 인라인 변환 (한 줄 텍스트) ──
def inline(node):
    if node.tag == "#text":
        return re.sub(r"\s+", " ", node.text or "")
    tag = node.tag
    if tag == "br":
        return "\n"
    if tag == "img":
        alt = node.attrs.get("alt", "")
        src = node.attrs.get("src", "")
        return f"![{alt}]({src})" if src else ""
    inner = "".join(inline(c) for c in node.children)
    if tag in ("strong", "b"):
        s = inner.strip()
        return f"**{s}**" if s else ""
    if tag in ("em", "i"):
        s = inner.strip()
        return f"*{s}*" if s else ""
    if tag == "code":
        return f"`{inner.strip()}`"
    if tag == "a":
        href = node.attrs.get("href", "")
        txt = inner.strip()
        if not txt:
            return ""
        if not href or href.startswith("#") or href.startswith("javascript"):
            return txt
        return f"[{txt}]({href})"
    return inner


def cell_text(node):
    """표 셀 → 한 줄. 개행/파이프 정리."""
    txt = "".join(inline(c) for c in node.children)
    txt = txt.replace("|", "\\|")
    txt = re.sub(r"\s*\n\s*", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def _rows_of(node):
    rows = []
    for c in node.children:
        if c.tag == "tr":
            rows.append(c)
        elif c.tag in ("thead", "tbody", "tfoot"):
            rows.extend(_rows_of(c))
    return rows


def convert_table(node):
    trs = _rows_of(node)
    if not trs:
        return ""
    # 그리드 확장 (colspan/rowspan)
    grid = []
    carry = {}  # col -> [remaining, value]
    ncols = 0
    header_rows = 0
    # thead 존재 시 헤더 행 수
    for c in node.children:
        if c.tag == "thead":
            header_rows = len(_rows_of(c))
            break
    for tr in trs:
        row = []
        col = 0
        cells = [c for c in tr.children if c.tag in ("td", "th")]
        ci = 0
        while ci < len(cells) or col in carry:
            if col in carry:
                rem, val = carry[col]
                row.append(val)
                rem -= 1
                if rem <= 0:
                    del carry[col]
                else:
                    carry[col] = [rem, val]
                col += 1
                continue
            cell = cells[ci]
            ci += 1
            val = cell_text(cell)
            try:
                cs = int(cell.attrs.get("colspan", 1))
            except ValueError:
                cs = 1
            try:
                rs = int(cell.attrs.get("rowspan", 1))
            except ValueError:
                rs = 1
            for k in range(cs):
                row.append(val if k == 0 else "")
                if rs > 1:
                    carry[col] = [rs - 1, val if k == 0 else ""]
                col += 1
        grid.append(row)
        ncols = max(ncols, len(row))
    if ncols == 0:
        return ""
    for row in grid:
        row += [""] * (ncols - len(row))
    if header_rows == 0:
        header = grid[0]
        body = grid[1:]
    else:
        # 여러 헤더 행은 하나로 합침
        header = [" ".join(filter(None, (grid[r][c] for r in range(header_rows)))).strip()
                  for c in range(ncols)]
        body = grid[header_rows:]
    header = [h if h.strip() else " " for h in header]
    out = ["| " + " | ".join(header) + " |",
           "| " + " | ".join(["---"] * ncols) + " |"]
    for row in body:
        out.append("| " + " | ".join(row) + " |")
    return "\n".join(out) + "\n"


def convert_list(node, ordered, depth=0):
    lines = []
    idx = 1
    for c in node.children:
        if c.tag != "li":
            continue
        # li 안의 블록/인라인 분리
        prefix = ("  " * depth) + (f"{idx}. " if ordered else "- ")
        sub_lists = [x for x in c.children if x.tag in ("ul", "ol")]
        inline_parts = "".join(
            inline(x) for x in c.children if x.tag not in ("ul", "ol")
        )
        inline_parts = re.sub(r"\s*\n\s*", " ", inline_parts)
        inline_parts = re.sub(r"\s+", " ", inline_parts).strip()
        lines.append(prefix + inline_parts)
        for sl in sub_lists:
            lines.append(convert_list(sl, sl.tag == "ol", depth + 1))
        idx += 1
    return "\n".join(l for l in lines if l.strip())


def block(node):
    """블록 요소 → md 문단(들). 리스트[str] 반환."""
    tag = node.tag
    if tag in SKIP_TREE:
        return []
    if node.tag == "#text":
        t = re.sub(r"\s+", " ", node.text or "")
        return [t] if t.strip() else []
    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        lvl = int(tag[1])
        s = "".join(inline(c) for c in node.children).strip()
        return [("#" * lvl) + " " + s] if s else []
    if tag == "table":
        # 인쇄 여백용 레이아웃 래퍼(table.page-box)는 표가 아니라 컨테이너 → 내용 언랩
        if "page-box" in _cls(node):
            out = []
            for tr in _rows_of(node):
                for cell in tr.children:
                    if cell.tag in ("td", "th"):
                        for c in cell.children:
                            out.extend(block(c))
            return out
        t = convert_table(node)
        return [t] if t.strip() else []
    if tag in ("ul", "ol"):
        t = convert_list(node, tag == "ol")
        return [t] if t.strip() else []
    if tag == "blockquote":
        inner = []
        for c in node.children:
            inner.extend(block(c))
        text = "\n\n".join(inner).strip()
        if not text:
            return []
        return ["\n".join("> " + ln for ln in text.split("\n"))]
    if tag == "pre":
        code = "".join(_raw_text(c) for c in node.children).rstrip("\n")
        return ["```\n" + code + "\n```"]
    if tag == "hr":
        return ["---"]
    if tag in ("p", "figcaption", "li"):
        s = "".join(inline(c) for c in node.children)
        s = re.sub(r"[ \t]*\n[ \t]*", "\n", s).strip()
        return [s] if s else []
    # div/section 등 컨테이너: 자식이 블록이면 재귀, 아니면 인라인 묶음
    child_blocks = any(
        (c.tag in BLOCK or c.tag == "hr") for c in node.children
    )
    if child_blocks:
        out = []
        pending_inline = []
        for c in node.children:
            if c.tag in BLOCK or c.tag == "hr":
                if pending_inline:
                    s = "".join(pending_inline)
                    s = re.sub(r"[ \t]*\n[ \t]*", "\n", s).strip()
                    if s:
                        out.append(s)
                    pending_inline = []
                out.extend(block(c))
            else:
                pending_inline.append(inline(c))
        if pending_inline:
            s = "".join(pending_inline)
            s = re.sub(r"[ \t]*\n[ \t]*", "\n", s).strip()
            if s:
                out.append(s)
        return out
    s = "".join(inline(c) for c in node.children)
    s = re.sub(r"[ \t]*\n[ \t]*", "\n", s).strip()
    return [s] if s else []


def _raw_text(node):
    if node.tag == "#text":
        return node.text or ""
    return "".join(_raw_text(c) for c in node.children)


def find(node, tag):
    if node.tag == tag:
        return node
    for c in node.children:
        r = find(c, tag)
        if r:
            return r
    return None


def convert_html(html_text):
    dom = DOM()
    dom.feed(html_text)
    body = find(dom.root, "body") or dom.root
    blocks = []
    for c in body.children:
        blocks.extend(block(c))
    # 정리: 빈 블록 제거, 문단 간 한 줄 공백
    md = "\n\n".join(b.strip() for b in blocks if b.strip())
    md = unescape(md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


def main(argv):
    targets = argv[1:] if len(argv) > 1 else DEFAULT_TARGETS
    done, missing = [], []
    for rel in targets:
        src = (ROOT / rel).resolve()
        if not src.exists():
            missing.append(rel)
            continue
        html_text = src.read_text(encoding="utf-8")
        md = convert_html(html_text)
        out = src.with_suffix(".md")
        out.write_text(md, encoding="utf-8")
        done.append(out.relative_to(ROOT).as_posix())
    print(f"[변환] {len(done)}개")
    for p in done:
        print("  +", p)
    if missing:
        print(f"[없음] {len(missing)}개")
        for p in missing:
            print("  !", p)


if __name__ == "__main__":
    main(sys.argv)
