"""05_내역서 — 공종별 하위 폴더 경로 (SSOT)."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "05_내역서"
WORK = BASE / "내역서작업"
ILW = BASE / "일위대가DB"
POOM = BASE / "표준품셈"
COMMON = "_공통"

GONGJONG = ("회전교차로", "진입도로", "토목", "조경", "폐기물", "전기", COMMON)


def work(gj: str, name: str) -> Path:
    return WORK / gj / name


def ilw_path(gj: str, name: str) -> Path:
    return ILW / gj / name


# 표준단가산출 xlsx (공종 → 파일명)
STD_XLSX: dict[str, str] = {
    "회전교차로": "05_화성 청원로(회전교차로)_표준단가산출.xlsx",
    "진입도로": "04_화성 청원지구 진입도로 실시설계_표준단가산출.xlsx",
    "토목": "01_화성 청원지구 토목_표준단가산출.xlsx",
    "조경": "01_화성 청원지구 조경_표준단가산출.xlsx",
    "전기": "02_화성 청원지구 전기설비_표준단가산출.xlsx",
    "폐기물": "07_화성 청원지구 건설폐기물처리_표준단가산출.xlsx",
}

# 06 개발행위 — 토목 폴더
STD_XLSX_DEV = "06_화성 청원지구 산업유통형 개발행위_표준단가산출.xlsx"

# 파일명 → 공종 (내역서작업)
WORK_FILE_GJ: dict[str, str] = {
    **{v: k for k, v in STD_XLSX.items()},
    "06_화성 청원지구 산업유통형 개발행위_표준단가산출.xlsx": "토목",
    "06_화성 청원지구 산업유통형 개발행위_표준단가산출_요약.md": "토목",
}


def resolve_work(name: str) -> Path:
    """내역서작업 하위 공종 폴더에서 파일 경로 탐색."""
    if name in WORK_FILE_GJ:
        p = WORK / WORK_FILE_GJ[name] / name
        if p.exists():
            return p
    for p in WORK.rglob(name):
        if p.is_file():
            return p
    return WORK / name


def resolve_ilw(name: str) -> Path:
    """일위대가DB 하위 경로 탐색."""
    for p in ILW.rglob(name):
        if p.is_file() or (p.is_dir() and name.endswith("/")):
            return p
    return ILW / name


def is_sanan_indirect_excluded(section, name) -> bool:
    """산업안전보건관리비·안전관리비 — 간접비(제경비) 별도 적용 → 미매칭 집계 제외."""
    sec = str(section or "")
    nm = str(name or "")
    if "산업안전보건관리비" in sec:
        return True
    if "안전관리비" in nm or "안전관리계획" in nm:
        return True
    return False


def is_qc_mgmt_excluded(section, name) -> bool:
    """품질관리비 — 별도 산정(14개월 등) → 미확정(_ka) 집계 제외."""
    sec = str(section or "")
    nm = str(name or "")
    if "품질관리비" in sec:
        return True
    if "품질시험" in nm or "품질관리자" in nm:
        return True
    return False


# 토목(조경) 시트 「1. 토공」 — 실질 조경 범위(01 조경에서 확인)
JOGYEONG_CROSSCHECK_NAMES = frozenset({
    "가). 교목 병해방제",
    "나). 교목 충해방제",
    "다). 가로지지대",
})


def is_jogyeong_crosscheck_item(section, name) -> bool:
    """토목 통합본에 포함된 조경성 항목 — 01 조경에서 확인."""
    sec = str(section or "").replace(" ", "")
    nm = str(name or "").strip()
    if "토공" not in sec and "1.토" not in sec:
        return False
    return nm in JOGYEONG_CROSSCHECK_NAMES or any(
        k in nm for k in ("교목 병해방제", "교목 충해방제", "가로지지대")
    )


def is_ka_pending_excluded(section, name) -> bool:
    """미확정(_ka) 집계·목록 제외 — 간접비·품질관리·조경 확인 대상."""
    return (
        is_sanan_indirect_excluded(section, name)
        or is_qc_mgmt_excluded(section, name)
        or is_jogyeong_crosscheck_item(section, name)
    )


GAYEONGHYEON_STATUS = "가영현"
MATCHED_STATUS_BLANK = ""  # _ka ◆ 내역 일괄 — 매칭 확정 행은 상태 빈칸


def normalize_ka_detail_status(status) -> str:
    """_ka 상세 「상태」 — 매칭은 빈칸."""
    s = str(status or "").strip()
    if s == "매칭":
        return MATCHED_STATUS_BLANK
    return s


def is_ka_status_excluded(status) -> bool:
    """집계 제외 — 매칭(빈칸)·원본."""
    s = str(status or "").strip()
    return s in ("", "매칭", "원본")


def is_ka_gayeonghyeon(status, *, has_sum: bool = False) -> bool:
    """가영현 — 상태 명시 또는 검토·미매칭 중 합계(단가·금액) 입력."""
    s = str(status or "").strip()
    if GAYEONGHYEON_STATUS in s:
        return True
    if s in ("검토", "미매칭", "미산출") and has_sum:
        return True
    return False


def classify_ka_mihwakjeong(status, *, has_sum: bool = False) -> str | None:
    """미확정 집계·표시 — 검토·미매칭만. 가영현·매칭·원본 제외."""
    if is_ka_status_excluded(status):
        return None
    if is_ka_gayeonghyeon(status, has_sum=has_sum):
        return None
    s = str(status or "").strip()
    if s in ("미매칭", "미산출"):
        return "미매칭"
    if s == "검토":
        return "검토"
    return None


def _cell_num(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def detail_row_has_sum(
    mat_u=None,
    lab_u=None,
    exp_u=None,
    mat_amt=None,
    lab_amt=None,
    exp_amt=None,
    unit_sum=None,
    total_amt=None,
) -> bool:
    """◆ 내역 일괄 — 합계단가·합계금액·단가·금액 중 하나라도 0이 아니면 True."""
    if abs(_cell_num(total_amt)) > 0 or abs(_cell_num(unit_sum)) > 0:
        return True
    parts = (mat_u, lab_u, exp_u, mat_amt, lab_amt, exp_amt)
    if any(abs(_cell_num(x)) > 0 for x in parts):
        return True
    return False
