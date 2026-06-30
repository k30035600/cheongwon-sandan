# tools — 표준단가 산출 파이프라인

화성 청원지구 공내역서 **표준단가 산출·매칭·집계·검증** 스크립트.

> 일회성·디버그·적용 완료 패치는 `tools/_archive/` 참조.

## 파이프라인 (순서)

| 단계 | 스크립트 |
|---|---|
| 단가 산출 | `apply_standard_prices.py` · `apply_electric_prices.py` · `apply_waste_prices.py` · `apply_jo_forest_prices.py` |
| 확정 반영 | `apply_confirmed_prices.py` |
| 일위대가·미매칭 | `build_ilwidae_worksheet.py` · `build_unmatched_worksheet.py` · `build_review_*_ilwidae.py` |
| 집계 | `build_consolidated_summary.py`(→ `portal_stats.js`) · `build_summary_html.py` · `build_report_html.py`(→ `portal_ka_stats.js`) · `build_gongjong_cost_html.py` · `build_dan_ga_ban_yeong.py` |
| 검증·교정 | `fix_outlier_prices.py` · `build_validity_report.py` · `compare_um_vs_ka.py` |
| DB 구축 | `build_ildae_db.py` · `build_landscape_ildae.py` · `build_nojng_wages.py` · `build_jojadang_materials.py` |
| 외부 단가 | `fetch_g2b_price.py` · `fetch_forest_tree_prices.py` · `fetch_kseis_data.py` · `fetch_shopmall.py` |
| 점검·리포트 | `match_g2b_unmatched.py` · `match_forest_tree_unmatched.py` · `probe_kpa_unmatched.py` · `export_unmatched_review.py` |
| 토공·공정 | `calc_togong_schedule.py` · `report_togong_haul.py` · `report_gongjong_qty.py` |
| 원가·비교 | `calc_overhead.py` · `compare_cost_rates.py` |

## 유틸

| 스크립트 | 용도 |
|---|---|
| `naeyeok_gongjong.py` | 공종별 경로 SSOT (`resolve_work`) |
| `reorganize_05_by_gongjong.py` | 내역서작업·DB·품셈 공종 하위 분류 |
| `patch_05_gongjong_paths.py` | 공종 경로 문자열 일괄 갱신 |
| `extract_jo_planting.py` | 조경 식재 내역 추출 |
| `cleanup_oneoff.py` | 일회성 스크립트 `_archive` 이동 |
| `serve_portal.py` | 로컬 서버 (`/청원지구_포털` → 포털 HTML) |

끝.
