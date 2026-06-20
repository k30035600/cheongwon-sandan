"""서울시설공단 2024 조경 표준설계 일위대가 XLS → 조경표준일위대가_2024.csv 변환.

원본: 05_내역서/일위대가DB/2024표준설계대가_데이터베이스(하반기노임)최종.xls
대상 시트: 「일위대가목록(총괄)」 (품명·규격·단위·재료비·노무비·경비·합계 단가)
출력: 05_내역서/일위대가DB/조경표준일위대가_2024.csv
"""
import csv, sys
from pathlib import Path
import xlrd
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "05_내역서" / "일위대가DB" / "2024표준설계대가_데이터베이스(하반기노임)최종.xls"
OUT = ROOT / "05_내역서" / "일위대가DB" / "조경표준일위대가_2024.csv"
SHEET = "일위대가목록(총괄)"


def num(v):
    try:
        f = float(v)
        return f
    except (TypeError, ValueError):
        return 0.0


def main():
    wb = xlrd.open_workbook(str(SRC))
    sh = wb.sheet_by_name(SHEET)
    hdr = [str(sh.cell_value(0, c)).replace(" ", "").strip() for c in range(sh.ncols)]
    print("헤더:", [(i, h) for i, h in enumerate(hdr) if h])

    def find(name):
        for i, h in enumerate(hdr):
            if name in h:
                return i
        return -1
    c_div = find("구분")
    c_name = find("품명")
    c_spec = find("규격")
    c_unit = find("단위")
    c_mat = find("재료비")
    c_lab = find("노무비")
    c_exp = find("경비")
    c_sum = find("합계")
    print(f"열 매핑 구분={c_div} 품명={c_name} 규격={c_spec} 단위={c_unit} "
          f"재료비={c_mat} 노무비={c_lab} 경비={c_exp} 합계={c_sum}")

    rows = []
    cur_div = ""
    for r in range(1, sh.nrows):
        div = str(sh.cell_value(r, c_div)).strip() if c_div >= 0 else ""
        if div:
            cur_div = div
        name = str(sh.cell_value(r, c_name)).strip() if c_name >= 0 else ""
        spec = str(sh.cell_value(r, c_spec)).strip() if c_spec >= 0 else ""
        unit = str(sh.cell_value(r, c_unit)).strip() if c_unit >= 0 else ""
        if not name or not unit:
            continue
        mat = num(sh.cell_value(r, c_mat)) if c_mat >= 0 else 0.0
        lab = num(sh.cell_value(r, c_lab)) if c_lab >= 0 else 0.0
        exp = num(sh.cell_value(r, c_exp)) if c_exp >= 0 else 0.0
        tot = num(sh.cell_value(r, c_sum)) if c_sum >= 0 else 0.0
        if tot <= 0:
            tot = mat + lab + exp
        if tot <= 0:
            continue
        rows.append([cur_div, name, spec, unit, int(mat), int(lab), int(exp), int(tot)])

    with OUT.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["구분", "품명", "규격", "단위", "재료비", "노무비", "경비", "합계"])
        w.writerows(rows)
    print(f"변환 완료: {len(rows)}건 → {OUT}")
    print("샘플:")
    for row in rows[:3] + [x for x in rows if "수목식재" in x[1]][:3] + [x for x in rows if "잔디" in x[1]][:2]:
        print("   ", row)


if __name__ == "__main__":
    main()
