#!/usr/bin/env python3
"""01 조경 — forestinfo(재료) + 조경일위2024(식재 노무·경비) + 푸른조경(시설 미매칭).

입력: 05_내역서/공내역서/01_화성 청원지구 조경.XLS
       05_내역서/공내역서/조경시설물/01_260620화성시마도면청원리청원지구 조경.XLS (푸른조경 단가)
출력: 05_내역서/내역서작업/조경/01_화성 청원지구 조경_표준단가산출.xlsx (갱신)
"""
from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

TOOLS = Path(__file__).resolve().parent
ROOT = TOOLS.parent
sys.path.insert(0, str(TOOLS))

import apply_standard_prices as asp  # noqa: E402
import match_forest_tree_unmatched as mft  # noqa: E402

BASE = ROOT / "05_내역서"
SRC = BASE / "공내역서" / "01_화성 청원지구 조경.XLS"
PUREUN_SRC = BASE / "공내역서" / "조경시설물" / "01_260620화성시마도면청원리청원지구 조경.XLS"
OUT = BASE / "내역서작업" / "조경" / "01_화성 청원지구 조경_표준단가산출.xlsx"
OUT_MD = OUT.with_name(OUT.stem + "_요약.md")
LANDSCAPE_CSV = BASE / "일위대가DB" / "조경" / "조경표준일위대가_2024.csv"
R_BAND_RE = re.compile(r"R(\d+)\s*~\s*R?(\d+)|R(\d+)", re.I)
H_BAND_RE = re.compile(r"H([\d.]+)\s*~\s*([\d.]+)|H([\d.]+)\s*미만|H([\d.]+)\s*이하", re.I)


def parse_dims(spec: str) -> dict[str, float]:
    return mft.parse_dims(spec)


def load_landscape_planting() -> list[dict]:
    rows: list[dict] = []
    if not LANDSCAPE_CSV.exists():
        return rows
    with LANDSCAPE_CSV.open(encoding="utf-8-sig", newline="") as f:
        for rec in csv.DictReader(f):
            name = (rec.get("품명") or "").strip()
            if "식재" not in name:
                continue
            try:
                mat = float(rec.get("재료비") or 0)
                lab = float(rec.get("노무비") or 0)
                exp = float(rec.get("경비") or 0)
                tot = float(rec.get("합계") or 0) or mat + lab + exp
            except ValueError:
                continue
            if tot <= 0:
                continue
            rows.append({
                "name": name,
                "spec": (rec.get("규격") or "").strip(),
                "unit": (rec.get("단위") or "주").strip(),
                "mat": mat,
                "lab": lab,
                "exp": exp,
                "total": tot,
            })
    return rows


def _r_in_band(r: float, spec: str) -> bool:
    m = R_BAND_RE.search(spec.replace(" ", ""))
    if not m:
        return False
    if m.group(3):
        return abs(r - float(m.group(3))) < 0.01
    lo, hi = float(m.group(1)), float(m.group(2))
    return lo <= r <= hi


def _h_in_band(h: float, spec: str) -> bool:
    s = spec.replace(" ", "")
    if "미만" in spec and (m := re.search(r"H([\d.]+)\s*미만", spec)):
        return h < float(m.group(1))
    if "이하" in spec and (m := re.search(r"H([\d.]+)\s*이하", spec)):
        return h <= float(m.group(1))
    m = H_BAND_RE.search(s)
    if not m:
        return False
    if m.group(3):
        return h <= float(m.group(3))
    if m.group(4):
        return h < float(m.group(4))
    lo, hi = float(m.group(1)), float(m.group(2))
    return lo <= h <= hi


def pick_planting(spec: str, tree_name: str) -> dict | None:
    rows = load_landscape_planting()
    if not rows:
        return None
    dims = parse_dims(spec)
    h = dims.get("H")
    r = dims.get("R")

    if h is not None and h <= 0.35:
        cands = [x for x in rows if "관목식재" in x["name"]]
        for row in cands:
            if _h_in_band(h, row["spec"]):
                return row
        return cands[0] if cands else None

    if r is not None:
        cands = [x for x in rows if x["name"] == "수목식재" and "R" in x["spec"]]
        for row in cands:
            if _r_in_band(r, row["spec"]):
                return row
        # R 근사 — 가장 가까운 구간
        best, best_d = None, 999.0
        for row in cands:
            m = R_BAND_RE.search(row["spec"].replace(" ", ""))
            if not m:
                continue
            mid = float(m.group(3)) if m.group(3) else (float(m.group(1)) + float(m.group(2))) / 2
            d = abs(r - mid)
            if d < best_d:
                best_d, best = d, row
        if best:
            return best

    if h is not None:
        for key, row_name in [(3.0, "수목식재(인력) H3.0이하"), (5.0, "수목식재(기계) H5.0이하")]:
            if h <= key:
                for row in rows:
                    if row["name"].startswith(row_name.split()[0]) and _h_in_band(h, row["spec"]):
                        return row
        for row in rows:
            if row["name"] == "수목식재" and "2.0이하" in row["spec"]:
                return row

    for row in rows:
        if row["name"] == "수목식재" and "B5" in row["spec"]:
            return row
    return rows[0] if rows else None


def norm_item_key(name: str, spec: str, unit: str) -> tuple[str, str, str]:
    n = re.sub(r"\s+", "", name.strip())
    s = re.sub(r"\s+", "", spec.strip())
    u = asp.norm_unit(unit.strip())
    return n, s, u


def load_pureun_prices(src: Path = PUREUN_SRC) -> dict[tuple[str, str, str], dict]:
    """푸른조경 조경 XLS — 품명·규격·단위별 단가(재·노·경·합)."""
    if not src.exists():
        return {}
    import xlrd

    wb = xlrd.open_workbook(str(src))
    sh = wb.sheet_by_name("내역서")
    index: dict[tuple[str, str, str], dict] = {}
    for r in range(3, sh.nrows):
        name = str(sh.cell_value(r, 0)).strip()
        spec = str(sh.cell_value(r, 1)).strip()
        qty = sh.cell_value(r, 2) if sh.cell_type(r, 2) == xlrd.XL_CELL_NUMBER else None
        unit = str(sh.cell_value(r, 3)).strip()
        if not (qty and unit and unit != "식" and float(qty) > 0):
            continue
        try:
            total_u = float(sh.cell_value(r, 4) or 0)
            mat_u = float(sh.cell_value(r, 6) or 0)
            lab_u = float(sh.cell_value(r, 8) or 0)
            exp_u = float(sh.cell_value(r, 10) or 0)
        except (TypeError, ValueError):
            continue
        if total_u <= 0 and mat_u + lab_u + exp_u <= 0:
            continue
        if total_u <= 0:
            total_u = mat_u + lab_u + exp_u
        key = norm_item_key(name, spec, unit)
        index[key] = {
            "code": "푸른조경",
            "name": re.sub(r"^\s+", "", name),
            "spec": spec,
            "unit": unit,
            "mat": mat_u,
            "lab": lab_u,
            "exp": exp_u,
            "total": total_u,
            "date": "푸른조경 2026.6.20",
        }
    return index


def match_pureun(item: dict, pureun: dict[tuple[str, str, str], dict]) -> dict | None:
    key = norm_item_key(item["name"], item["spec"], item["unit"])
    return pureun.get(key)


def combine_forest_planting(forest: dict, plant: dict | None, item: dict) -> dict:
    mat = float(forest["mat"])
    lab = float(plant["lab"]) if plant else 0.0
    exp = float(plant["exp"]) if plant else 0.0
    # forestinfo=묘목(재료), 조경일위=식재 노무·경비(재료 중복 제외)
    total = mat + lab + exp
    spec_note = f"forestinfo {forest.get('date','')} + 조경일위 {plant['name'] if plant else '—'}"
    return {
        "code": f"FOREST|{forest['code']}",
        "name": mft.clean_tree_name(item["name"]),
        "spec": item["spec"],
        "unit": "주",
        "mat": mat,
        "lab": lab,
        "exp": exp,
        "total": total,
        "date": spec_note,
    }


def build_jo_results(
    items: list[dict],
    prices: list[dict],
    forest_prices: list[dict],
    src_name: str,
    pureun: dict[tuple[str, str, str], dict] | None = None,
):
    matched_rows = []
    unmatched_rows = []
    by_row: dict[int, dict] = {}
    totals = {"mat": 0.0, "lab": 0.0, "exp": 0.0, "sum": 0.0}
    section_totals: dict[str, dict] = {}
    forest_applied = 0
    pureun_applied = 0
    pureun = pureun or {}

    for item in items:
        sec = item["section"]
        section_totals.setdefault(
            sec, {"mat": 0.0, "lab": 0.0, "exp": 0.0, "sum": 0.0, "matched": 0, "items": 0}
        )
        section_totals[sec]["items"] += 1

        manual = asp.imok_outsource_price(item, src_name)
        if manual:
            price, score, term, terms = manual, 1.0, "수동위탁", ["위탁"]
        else:
            price, score, term, terms = asp.find_best_match(item, prices)

        forest_used = False
        pureun_used = False
        if (not price or score < asp.THRESHOLD) and item["unit"] == "주":
            fp, fs, fn = mft.match_tree(
                {"name": mft.clean_tree_name(item["name"]), "spec": item["spec"], "unit": "주"},
                forest_prices,
            )
            if fp and fs >= asp.THRESHOLD:
                plant = pick_planting(item["spec"], item["name"])
                price = combine_forest_planting(fp, plant, item)
                score = fs
                term = fn
                terms = ["forestinfo", "조경일위2024"]
                forest_used = True
                forest_applied += 1

        if not price or score < asp.THRESHOLD:
            pp = match_pureun(item, pureun)
            if pp:
                price = pp
                score = 1.0
                term = "푸른조경"
                terms = ["푸른조경"]
                pureun_used = True
                pureun_applied += 1

        base = {
            **item,
            "status": "미매칭",
            "match_score": round(score, 3) if score >= 0 else None,
            "terms": ", ".join(terms[:5]) if isinstance(terms, list) else str(terms),
        }

        if not price or score < asp.THRESHOLD:
            unmatched_rows.append({**base, "score": score})
            by_row[item["row"]] = base
            continue

        amts = asp.calc_amounts(item, price)
        if manual:
            confidence, status = "수동", "매칭"
        elif pureun_used:
            confidence, status = "푸른조경", "매칭"
        elif forest_used:
            status = "매칭" if score >= asp.REVIEW_THRESHOLD else "검토"
            confidence = "forestinfo" if status == "매칭" else "forestinfo·검토"
        else:
            confidence = "높음" if score >= asp.REVIEW_THRESHOLD else "검토"
            status = "매칭" if score >= asp.REVIEW_THRESHOLD else "검토"

        row = {
            **item,
            **amts,
            "status": status,
            "match_score": round(score, 3),
            "match_term": term,
            "price_code": price["code"],
            "price_name": price["name"],
            "price_spec": price.get("spec", ""),
            "price_unit": price["unit"],
            "mat_unit": price["mat"],
            "lab_unit": price["lab"],
            "exp_unit": price["exp"],
            "total_unit": price["total"],
            "confidence": confidence,
        }
        matched_rows.append(row)
        by_row[item["row"]] = row
        for key, val in zip(
            ["mat", "lab", "exp", "sum"],
            [amts["mat_amt"], amts["lab_amt"], amts["exp_amt"], amts["sum_amt"]],
        ):
            totals[key] += val
            section_totals[sec][key] += val
        section_totals[sec]["matched"] += 1

    review_rows = [r for r in matched_rows if r["status"] == "검토"]
    integrated = [by_row.get(it["row"], {**it, "status": "미매칭"}) for it in items]
    return (
        matched_rows, unmatched_rows, review_rows, integrated,
        totals, section_totals, forest_applied, pureun_applied,
    )


def build_pureun_basis_results(src: Path = PUREUN_SRC):
    """푸른조경 XLS 내역 금액(재·노·경)을 그대로 직접공사비로 사용."""
    import xlrd

    if not src.exists():
        raise FileNotFoundError(src)
    items, _ = asp.load_estimate(src)
    sh = xlrd.open_workbook(str(src)).sheet_by_name("내역서")
    matched_rows: list[dict] = []
    totals = {"mat": 0.0, "lab": 0.0, "exp": 0.0, "sum": 0.0}
    section_totals: dict[str, dict] = {}

    for item in items:
        r = item["row"] - 1
        mat_a = float(sh.cell_value(r, 7) or 0)
        lab_a = float(sh.cell_value(r, 9) or 0)
        exp_a = float(sh.cell_value(r, 11) or 0)
        sum_a = float(sh.cell_value(r, 5) or 0) or mat_a + lab_a + exp_a
        qty = float(item["qty"] or 1) or 1.0
        mat_u = float(sh.cell_value(r, 6) or 0)
        lab_u = float(sh.cell_value(r, 8) or 0)
        exp_u = float(sh.cell_value(r, 10) or 0)
        tot_u = float(sh.cell_value(r, 4) or 0) or mat_u + lab_u + exp_u

        sec = item["section"]
        section_totals.setdefault(
            sec, {"mat": 0.0, "lab": 0.0, "exp": 0.0, "sum": 0.0, "matched": 0, "items": 0}
        )
        section_totals[sec]["items"] += 1
        section_totals[sec]["matched"] += 1

        row = {
            **item,
            "mat_amt": mat_a,
            "lab_amt": lab_a,
            "exp_amt": exp_a,
            "sum_amt": sum_a,
            "status": "매칭",
            "match_score": 1.0,
            "match_term": "푸른조경",
            "price_code": "푸른조경",
            "price_name": item["name"],
            "price_spec": item["spec"],
            "price_unit": item["unit"],
            "mat_unit": mat_u if mat_u else mat_a / qty,
            "lab_unit": lab_u if lab_u else lab_a / qty,
            "exp_unit": exp_u if exp_u else exp_a / qty,
            "total_unit": tot_u if tot_u else sum_a / qty,
            "confidence": "푸른조경",
        }
        matched_rows.append(row)
        for key, val in zip(["mat", "lab", "exp", "sum"], [mat_a, lab_a, exp_a, sum_a]):
            totals[key] += val
            section_totals[sec][key] += val

    integrated = matched_rows
    return matched_rows, [], [], integrated, totals, section_totals


def main():
    import argparse

    ap = argparse.ArgumentParser(description="01 조경 표준단가산출")
    ap.add_argument(
        "--basis",
        choices=("pureun", "hybrid"),
        default="pureun",
        help="pureun=푸른조경(260620) 내역 금액 그대로 / hybrid=표준+forestinfo+푸른조경 폴백",
    )
    args = ap.parse_args()

    src_name = PUREUN_SRC.name if args.basis == "pureun" else SRC.name
    estimate_src = PUREUN_SRC if args.basis == "pureun" else SRC

    if args.basis == "pureun":
        if not PUREUN_SRC.exists():
            print(f"[중단] 푸른조경 원본 없음: {PUREUN_SRC}")
            sys.exit(1)
        matched, unmatched, review, integrated, totals, section_totals = build_pureun_basis_results(PUREUN_SRC)
        price_date = f"푸른조경(2026.6.20) 내역 금액 — {PUREUN_SRC.name}"
        n_forest = n_pureun = len(matched)
    else:
        if not SRC.exists():
            print(f"[중단] 원본 없음: {SRC}")
            sys.exit(1)
        if not mft.CSV_PATH.exists():
            print("[중단] 조경수_관측시세.csv 없음 — fetch_forest_tree_prices.py 실행")
            sys.exit(1)

        market = asp.load_market_csv(asp.MARKET_2026, "표준시장단가2026") or asp.load_prices()
        sijang = asp.load_market_csv(asp.SIJANG_2026, "시장시공가격2026")
        ildae = asp.load_ildae_prices()
        landscape = asp.load_landscape_ildae()
        mulga = asp.load_mulga()
        jojadang = asp.load_jojadang()
        forest = mft.load_prices()
        prices = asp.precompute(market + sijang + ildae + mulga + jojadang + landscape)

        if not PUREUN_SRC.exists():
            print(f"[경고] 푸른조경 원본 없음: {PUREUN_SRC.name}")
        pureun = load_pureun_prices()

        items, _ = asp.load_estimate(SRC)
        matched, unmatched, review, integrated, totals, section_totals, n_forest, n_pureun = build_jo_results(
            items, prices, forest, SRC.name, pureun
        )
        price_date = (
            f"표준시장단가2026 + forestinfo 조경수(주 {n_forest}건) + 조경일위2024 식재노무"
            f" + 푸른조경 시설({n_pureun}건)"
        )

    saved = asp.write_xlsx(
        matched, unmatched, review, integrated, totals, section_totals, price_date, OUT, src_name
    )
    asp.write_md(matched, unmatched, review, totals, section_totals, price_date, src_name, OUT_MD)

    matched_ok = len([r for r in matched if r["status"] == "매칭"])
    if args.basis == "pureun":
        print(f"푸른조경 기준 — 항목 {len(matched)}건 / 직접공사비(순공사비) 합계")
    else:
        print(f"forestinfo+조경일위 반영: {n_forest}건(주) / 푸른조경: {n_pureun}건")
        print(f"항목 {len(integrated)} / 매칭 {matched_ok} / 검토 {len(review)} / 미매칭 {len(unmatched)}")
    print(f"재료비 {totals['mat']:,.0f}  노무비 {totals['lab']:,.0f}  경비 {totals['exp']:,.0f}  합계 {totals['sum']:,.0f}")
    print(f"저장: {saved}")


if __name__ == "__main__":
    main()
