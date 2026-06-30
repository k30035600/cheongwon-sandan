# cheongwon-sandan — 화성 청원지구 공내역서 표준단가 산출

화성시 마도면 청원리 산175-2번지 일원(산업유통형 지구단위계획, 80,617㎡) 공내역서의
**표준단가 산출·매칭·집계·검증** 도구 모음과 결과 포털.

## 결과 요약 (`총괄표.xlsx` 기준 · 2026. 6. 19.)

| 항목 | 값 |
|---|---:|
| 직접공사비(01·02·04·05·06·07) | **7,136,910,628원 (71.37억)** |
| 재료비 / 노무비 / 경비 | 20.20억 / 43.13억 / 8.04억 |
| 매칭 / 검토 / 미매칭(조경 제외) / 전체 | 932 / 157 / 131 / 1,220 (매칭률 76.4%) |
| 도급액(총공사비, 개략) | 118.95억 (직접비 1.667배) |

> SSOT: `05_내역서/내역서작업/_공통/총괄표.md` · `portal_stats.js`(포털 자동 로드).
> 갱신: `python -X utf8 tools/build_consolidated_summary.py` 후 `build_summary_html.py` · `build_report_html.py`.

## 구성

- `청원지구_포털.html` — 통합 열람 포털(좌: 본문 / 우: 근거자료)
- `tools/` — 산출·매칭·집계·검증 파이썬 (`tools/README.md` · 일회성은 `tools/_archive/`)
  - `apply_standard_prices.py` · `apply_confirmed_prices.py` — 표준단가/확정단가 반영
  - `build_consolidated_summary.py` · `build_summary_html.py` · `build_report_html.py` — 총괄표·종합보고서
  - `build_gongjong_cost_html.py` — 회전·진입·토목조경·지구단위 직접·간접비 HTML
  - `build_ilwidae_worksheet.py` · `build_review_*_ilwidae.py` — 일위대가·품셈 산출
  - `fix_outlier_prices.py` · `build_validity_report.py` — 이상치 교정·품목검증
- `05_내역서/` — 산출 결과 문서(총괄표·종합보고서·요약). 대용량 원본/DB는 추적 제외.

## 비고

- 엑셀(`.xls`·`.xlsx`)은 포털 「근거자료」 미리보기용으로 추적한다.
  원본 PDF·CSV(단가 DB)·이미지는 용량 절감을 위해 `.gitignore`로 제외한다.
- 금액은 직접공사비(재료·노무·경비) 추정이며 제경비·부가가치세는 「원가계산서」에서 별도 산정.
