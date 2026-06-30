# tools/_archive — 일회성·점검용 스크립트 보관

> 2026. 6. 29. 정리. **재실행이 필요 없는** 디버그·감사·폴더 재편·패치 스크립트.

## 디버그·감사 (`_*.py`)

| 파일 | 용도 |
|---|---|
| `_audit_deploy_assets.py` | 배포 asset 점검 |
| `_audit_portal_full.py` | 포털 링크 전수 |
| `_audit_portal_tabs.py` | 포털 탭 점검 |
| `_audit_rowpatch.py` | 행 패치 오류 감사 |
| `_check_row_align.py` | 통합내역 행 정렬 |
| `_diag_rowsrc.py` | 행 출처 추적 |
| `_dump_shopmall.py` | 쇼핑몰 API 덤프 |
| `_factcheck_confirmed.py` | 확정단가 팩트체크 |
| `_list_tree_unmatched.py` | 조경수 미매칭 목록 |
| `_probe_ilwidae.py` | 일위대가 프로브 |
| `_probe_shopmall2.py` | 쇼핑몰 API 프로브 |
| `_scan_lowscore_conf.py` | 저점수 확정 스캔 |
| `_scan_outliers.py` | 이상치 스캔 |
| `_scan_rebar.py` | 철근 단가 스캔 |
| `_trace_item.py` | 품목 추적 |
| `_trace_row666.py` | 특정 행 추적 |
| `_validity_check.py` | 품목검증 보조 |
| `_verify_confirmed.py` | 확정단가 검증 |

## 폴더 재편·경로 패치 (적용 완료)

| 파일 | 용도 |
|---|---|
| `reorganize_05_naeyeokseo.py` | 4분류 1차 이동 |
| `move_final_outputs_to_root.py` | 산출물 루트 이동(역정리 전) |
| `restore_05_four_folders.py` | 4분류 복원 |
| `patch_05_work_paths.py` | 4분류 경로 패치 |
| `patch_05_root_outputs.py` | 루트 산출 경로 패치 |

## 일회성 데이터·공내역서 패치

| 파일 | 용도 |
|---|---|
| `patch_qc_months.py` | 품질관리 34→14개월 XLS 패치 |
| `patch_jo_facility_formulas.py` | 조경시설물 20건 xlsx 수식 |
| `split_01_tok_jo.py` | 01 통합원본 → 토목·조경 분리 |
| `clean_xls_sheets.py` | XLS 시트 정리 |
| `extract_land_register.py` | 지적/OCR 추출 |
| `build_pe_tarp_ilwidae.py` | PE천막 일위대가 산출 |
| `build_crane_ilwidae.py` | 크레인10톤 일위대가 산출 |

끝.
