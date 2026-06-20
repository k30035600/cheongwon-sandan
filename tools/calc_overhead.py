#!/usr/bin/env python3
"""직접공사비 → 총공사비(도급액) 제비율 자동계산.

- **RATES** — 국가계약 예정가격 작성기준 표준 예시(2025 토목 일반)
- **ELECTRIC_RATES** — 03_화성 청원지구 전기설비.xlsx 「원가」 시트 (총괄표 원가계산서 적용)
"""
from __future__ import annotations

import sys
from typing import Any

sys.stdout.reconfigure(encoding="utf-8")

# --- 표준 예시 요율 (발주처 확정 시 교체) ---
RATES: dict[str, float] = {
    "간접노무비": 0.127,
    "산재보험료": 0.0370,
    "고용보험료": 0.0127,
    "건강보험료": 0.03545,
    "연금보험료": 0.0450,
    "노인장기요양": 0.1295,
    "산업안전보건관리비": 0.0197,
    "기타경비": 0.050,
    "환경보전비": 0.005,
    "일반관리비": 0.050,
    "이윤": 0.100,
    "부가가치세": 0.10,
}

RATE_TABLE: list[tuple[str, str, str]] = [
    ("간접노무비", "간접노무비", "× 직접노무비"),
    ("산재보험료", "산재보험료", "× 노무비계(직접+간접)"),
    ("고용보험료", "고용보험료", "× 노무비계"),
    ("국민건강보험료", "건강보험료", "× 직접노무비"),
    ("국민연금보험료", "연금보험료", "× 직접노무비"),
    ("노인장기요양보험료", "노인장기요양", "× 건강보험료"),
    ("산업안전보건관리비", "산업안전보건관리비", "× (재료비+직접노무비)"),
    ("기타경비", "기타경비", "× (재료비+노무비계)"),
    ("환경보전비", "환경보전비", "× 직접공사비"),
    ("일반관리비", "일반관리비", "× 순공사원가"),
    ("이윤", "이윤", "× (노무비계+경비계+일반관리비), 재료비 제외"),
    ("부가가치세", "부가가치세", "× 공급가액"),
]

# 03_화성 청원지구 전기설비.xlsx 「원가」 시트 (2026-06-19 총괄표 적용)
ELECTRIC_RATES: dict[str, float] = {
    "간접노무비": 0.167,
    "산재보험료": 0.0356,
    "고용보험료": 0.0101,
    "퇴직공제": 0.023,
    "안전관리비": 0.0207,
    "기타경비": 0.053,
    "건강보험료": 0.03595,
    "연금보험료": 0.0475,
    "노인장기요양": 0.1314,
    "석면분담금": 0.00006,
    "임금채권부담금": 0.0009,
    "일반관리비": 0.065,
    "이윤": 0.15,
    "부가가치세": 0.10,
}

ELECTRIC_RATE_TABLE: list[tuple[str, str, str]] = [
    ("간접노무비", "간접노무비", "× 직접노무비"),
    ("산재보험료", "산재보험료", "× 노무비(직접+간접)"),
    ("고용보험료", "고용보험료", "× 노무비(직접+간접)"),
    ("퇴직금(퇴직공제부금)", "퇴직공제", "× 직접노무비"),
    ("안전관리비", "안전관리비", "× (재료비+직접노무비)"),
    ("기타경비", "기타경비", "× (재료비+노무비계)"),
    ("국민건강보험료", "건강보험료", "× 직접노무비"),
    ("국민연금보험료", "연금보험료", "× 직접노무비"),
    ("노인장기요양보험료", "노인장기요양", "× 건강보험료"),
    ("석면분담금", "석면분담금", "× 노무비(직접+간접)"),
    ("임금채권부담금", "임금채권부담금", "× 노무비(직접+간접)"),
    ("일반관리비", "일반관리비", "× (재료+노무+경비)"),
    ("이윤", "이윤", "× (노무+경비+일반관리비), 재료비 제외"),
    ("부가가치세", "부가가치세", "× 공급가액"),
]


def _row(
    step: str,
    name: str,
    amount: float,
    formula: str = "",
    *,
    bold: bool = False,
    total: bool = False,
    indent: int = 0,
) -> dict[str, Any]:
    return {
        "step": step,
        "name": name,
        "amount": round(amount),
        "formula": formula,
        "bold": bold,
        "total": total,
        "indent": indent,
    }


def compute_cost_statement(
    mat: float,
    lab: float,
    exp: float,
    rates: dict[str, float] | None = None,
) -> dict[str, Any]:
    """직접공사비(재·노·경) → 원가계산서 산출."""
    R = rates or RATES
    M, L, E = float(mat), float(lab), float(exp)
    D = M + L + E

    iln = L * R["간접노무비"]
    N = L + iln

    sanjae = N * R["산재보험료"]
    goyong = N * R["고용보험료"]
    gangang = L * R["건강보험료"]
    yeongeum = L * R["연금보험료"]
    nojang = gangang * R["노인장기요양"]
    anjeon = (M + L) * R["산업안전보건관리비"]
    gita = (M + N) * R["기타경비"]
    hwankyung = D * R["환경보전비"]
    sanchul = sanjae + goyong + gangang + yeongeum + nojang + anjeon + gita + hwankyung

    gyeongbi_gye = E + sanchul
    sun = M + N + gyeongbi_gye
    ilban = sun * R["일반관리비"]
    iyun = (N + gyeongbi_gye + ilban) * R["이윤"]
    gonggeup = sun + ilban + iyun
    vat = gonggeup * R["부가가치세"]
    dogeup = gonggeup + vat

    pct = lambda k: f"{R[k] * 100:g}%"

    rows = [
        _row("①", "직접공사비", D, "재료비+직접노무비+직접경비", bold=True, total=True),
        _row("", "· 재료비", M, indent=1),
        _row("", "· 직접노무비", L, indent=1),
        _row("", "· 직접경비", E, indent=1),
        _row("②", "간접노무비", iln, f"직접노무비 × {pct('간접노무비')}"),
        _row("", "노무비계(직접+간접)", N, bold=True),
        _row("③", "산재보험료", sanjae, f"노무비계 × {pct('산재보험료')}"),
        _row("③", "고용보험료", goyong, f"노무비계 × {pct('고용보험료')}"),
        _row("③", "국민건강보험료", gangang, f"직접노무비 × {pct('건강보험료')}"),
        _row("③", "국민연금보험료", yeongeum, f"직접노무비 × {pct('연금보험료')}"),
        _row("③", "노인장기요양보험료", nojang, f"건강보험료 × {pct('노인장기요양')}"),
        _row("③", "산업안전보건관리비", anjeon, f"(재료비+직접노무비) × {pct('산업안전보건관리비')}"),
        _row("③", "기타경비", gita, f"(재료비+노무비계) × {pct('기타경비')}"),
        _row("③", "환경보전비", hwankyung, f"직접공사비 × {pct('환경보전비')}"),
        _row("", "경비계(직접경비+산출경비)", gyeongbi_gye, bold=True),
        _row("⑥", "순공사원가", sun, "재료비+노무비계+경비계", bold=True, total=True),
        _row("⑦", "일반관리비", ilban, f"순공사원가 × {pct('일반관리비')}"),
        _row("⑧", "이윤", iyun, f"(노무비계+경비계+일반관리비) × {pct('이윤')}"),
        _row("", "공급가액", gonggeup, "순공사원가+일반관리비+이윤", bold=True),
        _row("", "부가가치세", vat, f"공급가액 × {pct('부가가치세')}"),
        _row(
            "★",
            "도급액(총공사비)",
            dogeup,
            f"공급가액 + 부가세 (직접비의 {dogeup / D:.3f}배)" if D else "",
            bold=True,
            total=True,
        ),
    ]

    rate_rows = [
        (label, R[key], basis)
        for label, key, basis in RATE_TABLE
    ]

    return {
        "mat": round(M),
        "lab": round(L),
        "exp": round(E),
        "direct": round(D),
        "indirect_labor": round(iln),
        "labor_total": round(N),
        "expense_total": round(gyeongbi_gye),
        "net_cost": round(sun),
        "general_admin": round(ilban),
        "profit": round(iyun),
        "supply": round(gonggeup),
        "vat": round(vat),
        "contract": round(dogeup),
        "multiplier": dogeup / D if D else 0,
        "rows": rows,
        "rate_rows": rate_rows,
        "rates": dict(R),
    }


def compute_cost_statement_electric(
    mat: float,
    lab: float,
    exp: float,
    rates: dict[str, float] | None = None,
) -> dict[str, Any]:
    """직접공사비 + 03 전기 원가 시트 제비율 → 원가계산서 산출."""
    R = rates or ELECTRIC_RATES
    M, L, E = float(mat), float(lab), float(exp)
    D = M + L + E

    iln = L * R["간접노무비"]
    N = L + iln

    sanjae = N * R["산재보험료"]
    goyong = N * R["고용보험료"]
    toejik = L * R["퇴직공제"]
    anjeon = (M + L) * R["안전관리비"]
    gita = (M + N) * R["기타경비"]
    gangang = L * R["건강보험료"]
    yeongeum = L * R["연금보험료"]
    nojang = gangang * R["노인장기요양"]
    seokmyeon = N * R["석면분담금"]
    imgeum = N * R["임금채권부담금"]
    sanchul = sanjae + goyong + toejik + anjeon + gita + gangang + yeongeum + nojang + seokmyeon + imgeum

    gyeongbi_gye = E + sanchul
    sun = M + N + gyeongbi_gye
    ilban = sun * R["일반관리비"]
    iyun = (N + gyeongbi_gye + ilban) * R["이윤"]
    gonggeup = sun + ilban + iyun
    vat = gonggeup * R["부가가치세"]
    dogeup = gonggeup + vat

    pct = lambda k: f"{R[k] * 100:g}%"

    rows = [
        _row("①", "직접공사비", D, "재료비+직접노무비+직접경비", bold=True, total=True),
        _row("", "· 재료비", M, indent=1),
        _row("", "· 직접노무비", L, indent=1),
        _row("", "· 직접경비", E, indent=1),
        _row("②", "간접노무비", iln, f"직접노무비 × {pct('간접노무비')}"),
        _row("", "노무비계(직접+간접)", N, bold=True),
        _row("③", "산재보험료", sanjae, f"노무비 × {pct('산재보험료')}"),
        _row("③", "고용보험료", goyong, f"노무비 × {pct('고용보험료')}"),
        _row("③", "퇴직금(퇴직공제부금)", toejik, f"직접노무비 × {pct('퇴직공제')}"),
        _row("③", "안전관리비", anjeon, f"(재료비+직접노무비) × {pct('안전관리비')}"),
        _row("③", "기타경비", gita, f"(재료비+노무비계) × {pct('기타경비')}"),
        _row("③", "국민건강보험료", gangang, f"직접노무비 × {pct('건강보험료')}"),
        _row("③", "국민연금보험료", yeongeum, f"직접노무비 × {pct('연금보험료')}"),
        _row("③", "노인장기요양보험료", nojang, f"건강보험료 × {pct('노인장기요양')}"),
        _row("③", "석면분담금", seokmyeon, f"노무비 × {pct('석면분담금')}"),
        _row("③", "임금채권부담금", imgeum, f"노무비 × {pct('임금채권부담금')}"),
        _row("", "경비계(직접경비+산출경비)", gyeongbi_gye, bold=True),
        _row("⑥", "순공사원가", sun, "재료비+노무비계+경비계", bold=True, total=True),
        _row("⑦", "일반관리비", ilban, f"(재료+노무+경비) × {pct('일반관리비')}"),
        _row("⑧", "이윤", iyun, f"(노무+경비+일반관리비) × {pct('이윤')}"),
        _row("", "공급가액", gonggeup, "순공사원가+일반관리비+이윤", bold=True),
        _row("", "부가가치세", vat, f"공급가액 × {pct('부가가치세')}"),
        _row(
            "★",
            "도급액(총공사비)",
            dogeup,
            f"공급가액 + 부가세 (직접비의 {dogeup / D:.3f}배)" if D else "",
            bold=True,
            total=True,
        ),
    ]

    rate_rows = [(label, R[key], basis) for label, key, basis in ELECTRIC_RATE_TABLE]

    return {
        "mat": round(M),
        "lab": round(L),
        "exp": round(E),
        "direct": round(D),
        "indirect_labor": round(iln),
        "labor_total": round(N),
        "expense_total": round(gyeongbi_gye),
        "net_cost": round(sun),
        "general_admin": round(ilban),
        "profit": round(iyun),
        "supply": round(gonggeup),
        "vat": round(vat),
        "contract": round(dogeup),
        "multiplier": dogeup / D if D else 0,
        "rows": rows,
        "rate_rows": rate_rows,
        "rates": dict(R),
        "rate_source": "03 전기 원가",
    }


# --- 토목공사 간접공사비 적용기준(2026.4.13) — 한국표준품셈정보원 공식 기준(현행) ---
# 공사기간 13~36개월·직접공사비 50억-300억·종합건설업 기준.
# 2026.4.13(입찰공고분부터) 개정: 간접노무비·기타경비 인상, 건설기계대여대금
#   지급보증 발급수수료(법정부담금) 신설. 보험료(산재·고용·건강·연금·노인장기요양)는
#   법정 고정으로 종전과 동일. 일반관리비·이윤은 추정가격 구간별 자동 선택.
# 산업안전보건관리비는 대상액(재료비+직접노무비) 구간에 따라 자동 선택(기초액 포함).
CIVIL_TOMOK_RATES: dict[str, float] = {
    "간접노무비": 0.197,    # 토목 13~36개월·50-300억 (260413, 종전 16.6→19.7)
    "산재보험료": 0.0356,
    "고용보험료": 0.0101,   # 고시금액~140억 미만(7등급)
    "건강보험료": 0.03595,
    "연금보험료": 0.0475,
    "노인장기요양": 0.1314,
    "기타경비": 0.069,      # 토목 13~36개월·50-300억 (260413, 종전 6.1→6.9)
    "건설기계대여보증": 0.004,  # 건설기계대여대금 지급보증(법정부담금) 토목 0.4% × 직접공사비
    "환경보전비": 0.009,    # 토목 도로(0.9%) — 공사종류별 상이(상하수도 0.5·플랜트 0.4 등)
    "부가가치세": 0.10,
    # 일반관리비·이윤은 추정가격 구간별 _ilban_iyun_civil()로 자동 선택
}

CIVIL_JOGYEONG_RATES: dict[str, float] = {
    "간접노무비": 0.191,    # 조경 13~36개월·50-300억 (260413, 종전 16.2→19.1)
    "산재보험료": 0.0356,
    "고용보험료": 0.0101,
    "건강보험료": 0.03595,
    "연금보험료": 0.0475,
    "노인장기요양": 0.1314,
    "기타경비": 0.063,      # 조경 13~36개월·50-300억 (260413, 종전 5.6→6.3)
    "건설기계대여보증": 0.0018,  # 조경공사 0.18% × 직접공사비
    "환경보전비": 0.005,    # 조경 — 별도 기준 없어 0.5% 보수 적용
    "부가가치세": 0.10,
}

CIVIL_RATE_TABLE: list[tuple[str, str, str]] = [
    ("간접노무비", "간접노무비", "× 직접노무비"),
    ("산재보험료", "산재보험료", "× 노무비계(직접+간접)"),
    ("고용보험료", "고용보험료", "× 노무비계"),
    ("국민건강보험료", "건강보험료", "× 직접노무비"),
    ("국민연금보험료", "연금보험료", "× 직접노무비"),
    ("노인장기요양보험료", "노인장기요양", "× 건강보험료"),
    ("산업안전보건관리비", "_산안비", "× (재료비+직접노무비) + 기초액(구간별)"),
    ("기타경비", "기타경비", "× (재료비+노무비계)"),
    ("건설기계대여대금 지급보증", "건설기계대여보증", "× 직접공사비(법정부담금)"),
    ("환경보전비", "환경보전비", "× 직접공사비"),
    ("일반관리비", "_일반관리비", "× 순공사원가(추정가격 구간별)"),
    ("이윤", "_이윤", "× (노무비계+경비계+일반관리비), 추정가격 구간별"),
    ("부가가치세", "부가가치세", "× 공급가액"),
]


def _sanan_civil(base_amt: float) -> tuple[float, int]:
    """토목 산업안전보건관리비 — 대상액(재+직노) 구간별 (율, 기초액원)."""
    if base_amt < 500_000_000:           # 5억 미만
        return 0.0315, 0
    if base_amt < 5_000_000_000:         # 5억 ~ 50억 미만
        return 0.0253, 3_300_000
    return 0.026, 0                      # 50억 이상(800억 미만)


def _ilban_iyun_civil(est_price: float) -> tuple[float, float]:
    """추정가격 구간별 (일반관리비율, 이윤율) — 260413 기준.

    50억 미만 8%/15%, 50~300억 6.5%/12%, 300~1000억 5%/10%, 1000억 이상 4.5%/9%.
    """
    if est_price < 5_000_000_000:          # 50억 미만
        return 0.08, 0.15
    if est_price < 30_000_000_000:         # 50억 ~ 300억 미만
        return 0.065, 0.12
    if est_price < 100_000_000_000:        # 300억 ~ 1000억 미만
        return 0.05, 0.10
    return 0.045, 0.09                     # 1000억 이상


def compute_cost_statement_civil(
    mat: float,
    lab: float,
    exp: float,
    rates: dict[str, float] | None = None,
    *,
    est_price: float | None = None,
    rate_source: str = "토목공사 간접공사비 기준(2026.4.13)",
) -> dict[str, Any]:
    """직접공사비 + 토목공사 간접공사비 적용기준(2026.4.13) → 원가계산서 산출.

    est_price(추정가격)로 일반관리비·이윤 구간을 자동 선택한다(미지정 시 직접공사비×1.7 추정).
    """
    R = rates or CIVIL_TOMOK_RATES
    M, L, E = float(mat), float(lab), float(exp)
    D = M + L + E

    iln = L * R["간접노무비"]
    N = L + iln

    sanjae = N * R["산재보험료"]
    goyong = N * R["고용보험료"]
    gangang = L * R["건강보험료"]
    yeongeum = L * R["연금보험료"]
    nojang = gangang * R["노인장기요양"]
    sanan_rate, sanan_base = _sanan_civil(M + L)
    anjeon = (M + L) * sanan_rate + sanan_base
    gita = (M + N) * R["기타경비"]
    gigye = D * R.get("건설기계대여보증", 0.0)
    hwankyung = D * R["환경보전비"]
    sanchul = sanjae + goyong + gangang + yeongeum + nojang + anjeon + gita + gigye + hwankyung

    gyeongbi_gye = E + sanchul
    sun = M + N + gyeongbi_gye
    ilban_rate, iyun_rate = _ilban_iyun_civil(est_price if est_price is not None else D * 1.7)
    ilban = sun * ilban_rate
    iyun = (N + gyeongbi_gye + ilban) * iyun_rate
    gonggeup = sun + ilban + iyun
    vat = gonggeup * R["부가가치세"]
    dogeup = gonggeup + vat

    pct = lambda k: f"{R[k] * 100:g}%"
    sanan_formula = (
        f"(재료비+직접노무비) × {sanan_rate * 100:g}%"
        + (f" + 기초액 {sanan_base:,}원" if sanan_base else "")
    )

    rows = [
        _row("①", "직접공사비", D, "재료비+직접노무비+직접경비", bold=True, total=True),
        _row("", "· 재료비", M, indent=1),
        _row("", "· 직접노무비", L, indent=1),
        _row("", "· 직접경비", E, indent=1),
        _row("②", "간접노무비", iln, f"직접노무비 × {pct('간접노무비')}"),
        _row("", "노무비계(직접+간접)", N, bold=True),
        _row("③", "산재보험료", sanjae, f"노무비계 × {pct('산재보험료')}"),
        _row("③", "고용보험료", goyong, f"노무비계 × {pct('고용보험료')}"),
        _row("③", "국민건강보험료", gangang, f"직접노무비 × {pct('건강보험료')}"),
        _row("③", "국민연금보험료", yeongeum, f"직접노무비 × {pct('연금보험료')}"),
        _row("③", "노인장기요양보험료", nojang, f"건강보험료 × {pct('노인장기요양')}"),
        _row("③", "산업안전보건관리비", anjeon, sanan_formula),
        _row("③", "기타경비", gita, f"(재료비+노무비계) × {pct('기타경비')}"),
        _row("③", "건설기계대여대금 지급보증", gigye, f"직접공사비 × {pct('건설기계대여보증')} (법정부담금)"),
        _row("③", "환경보전비", hwankyung, f"직접공사비 × {pct('환경보전비')}"),
        _row("", "경비계(직접경비+산출경비)", gyeongbi_gye, bold=True),
        _row("⑥", "순공사원가", sun, "재료비+노무비계+경비계", bold=True, total=True),
        _row("⑦", "일반관리비", ilban, f"순공사원가 × {ilban_rate * 100:g}% (추정가격 구간)"),
        _row("⑧", "이윤", iyun, f"(노무비계+경비계+일반관리비) × {iyun_rate * 100:g}% (추정가격 구간)"),
        _row("", "공급가액", gonggeup, "순공사원가+일반관리비+이윤", bold=True),
        _row("", "부가가치세", vat, f"공급가액 × {pct('부가가치세')}"),
        _row(
            "★",
            "도급액(총공사비)",
            dogeup,
            f"공급가액 + 부가세 (직접비의 {dogeup / D:.3f}배)" if D else "",
            bold=True,
            total=True,
        ),
    ]

    rate_rows = []
    for label, key, basis in CIVIL_RATE_TABLE:
        if key == "_산안비":
            rate_rows.append((label, sanan_rate, basis))
        elif key == "_일반관리비":
            rate_rows.append((label, ilban_rate, basis))
        elif key == "_이윤":
            rate_rows.append((label, iyun_rate, basis))
        else:
            rate_rows.append((label, R[key], basis))

    return {
        "mat": round(M),
        "lab": round(L),
        "exp": round(E),
        "direct": round(D),
        "indirect_labor": round(iln),
        "labor_total": round(N),
        "expense_total": round(gyeongbi_gye),
        "net_cost": round(sun),
        "general_admin": round(ilban),
        "profit": round(iyun),
        "supply": round(gonggeup),
        "vat": round(vat),
        "contract": round(dogeup),
        "multiplier": dogeup / D if D else 0,
        "rows": rows,
        "rate_rows": rate_rows,
        "rates": dict(R),
        "rate_source": rate_source,
    }


def won(x: float) -> str:
    return f"{round(x):,}"


def main() -> None:
    M, L, E = 1_551_879_570, 4_015_732_330, 679_892_807
    cs = compute_cost_statement_electric(M, L, E)
    print(f"[{cs.get('rate_source', '전기요율')}]")
    for r in cs["rows"]:
        prefix = "  " * r["indent"]
        tag = r["step"]
        print(f"{tag:>2} {prefix}{r['name']:<22} {won(r['amount']):>17}  {r['formula']}")
    print(f"\n직접공사비 대비 도급액 배율: {cs['multiplier']:.3f}")


if __name__ == "__main__":
    main()
