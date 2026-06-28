#!/usr/bin/env python3
"""01 조경 — forestinfo(재료) + 조경일위2024(식재 노무·경비) 미매칭(주) 반영.

입력: 05_내역서/공내역서/01_화성 청원지구 조경.XLS
출력: 05_내역서/내역서작업/01_화성 청원지구 조경_표준단가산출.xlsx (갱신)
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
OUT = BASE / "내역서작업" / "01_화성 청원지구 조경_표준단가산출.xlsx"
OUT_MD = OUT.with_name(OUT.stem + "_요약.md")
LANDSCAPE_CSV = BASE / "일위대가DB" / "조경표준일위대가_2024.csv"

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


def build_jo_results(items: list[dict], prices: list[dict], forest_prices: list[dict], src_name: str):
    matched_rows = []
    unmatched_rows = []
    by_row: dict[int, dict] = {}
    totals = {"mat": 0.0, "lab": 0.0, "exp": 0.0, "sum": 0.0}
    section_totals: dict[str, dict] = {}
    forest_applied = 0

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
        confidence = "높음" if score >= asp.REVIEW_THRESHOLD else "검토"
        if manual:
            confidence, status = "수동", "매칭"
        elif forest_used:
            status = "매칭" if score >= asp.REVIEW_THRESHOLD else "검토"
            confidence = "forestinfo" if status == "매칭" else "forestinfo·검토"
        else:
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
    return matched_rows, unmatched_rows, review_rows, integrated, totals, section_totals, forest_applied


def main():
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

    items, _ = asp.load_estimate(SRC)
    src_name = SRC.name
    matched, unmatched, review, integrated, totals, section_totals, n_forest = build_jo_results(
        items, prices, forest, src_name
    )

    price_date = (
        f"표준시장단가2026 + forestinfo 조경수(주 {n_forest}건) + 조경일위2024 식재노무"
    )
    saved = asp.write_xlsx(
        matched, unmatched, review, integrated, totals, section_totals, price_date, OUT, src_name
    )
    asp.write_md(matched, unmatched, review, totals, section_totals, price_date, src_name, OUT_MD)

    matched_ok = len([r for r in matched if r["status"] == "매칭"])
    print(f"forestinfo+조경일위 반영: {n_forest}건(주)")
    print(f"항목 {len(items)} / 매칭 {matched_ok} / 검토 {len(review)} / 미매칭 {len(unmatched)}")
    print(f"재료비 {totals['mat']:,.0f}  노무비 {totals['lab']:,.0f}  경비 {totals['exp']:,.0f}  합계 {totals['sum']:,.0f}")
    print(f"저장: {saved}")


if __name__ == "__main__":
    main()
