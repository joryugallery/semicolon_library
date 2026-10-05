"""ADE-45 Excel 업로드 계약 — fixture 생성 및 검증 스크립트.

최소 크기의 합성 .xlsx fixture를 만들고(개인정보·실제 기록 없음),
실제 `library_etl.pipeline.preprocess()`에 그대로 통과시켜
`docs/excel-upload-contract.md`에 적힌 결과와 같은지 확인한다.

사용법
------
    # 저장소 루트에서
    python scripts/make_excel_fixtures.py
    # 또는 fixture 출력 위치를 지정
    python scripts/make_excel_fixtures.py --out contracts/fixtures/excel

성공하면 종료 코드 0, 하나라도 계약과 다르면 1을 돌려준다.
CI에 넣을 때는 종료 코드만 보면 된다.

출력
----
    <out>/*.xlsx                        fixture 파일
    <out>/../excel-validation-results.json   실행 결과(케이스별 실제 동작)
"""
from __future__ import annotations

import argparse
import json
import sys
import traceback
from pathlib import Path

from openpyxl import Workbook

# 저장소 루트를 import 경로에 올린다 (scripts/ 의 부모)
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from library_etl.pipeline import preprocess, DataError, REQUIRED, HOURS  # noqa: E402

HDR = list(REQUIRED)  # 수집일자, 게이트, 통로ID, 전체_IN, 전체_OUT, IN_08..OUT_23
FRONT, BACK = "자료실.정문", "자료실.후문"
PF, PB = "10.0.0.1A", "10.0.0.2A"  # 합성 통로ID (실제 장비값 아님)

D1, D2 = "2026-01-06", "2026-01-07"  # 화, 수
D_MON = "2026-01-05"                 # 월 (휴관일 가정)


# ── 행 구성 도우미 ────────────────────────────────────────────────────────
def row(date, gate, passage, hours, total_in=None, total_out=None):
    """hours: {시: (IN, OUT)}. 전체_IN/OUT 기본값은 시간대 합(= 경고 없음)."""
    cells = {"수집일자": date, "게이트": gate, "통로ID": passage}
    for h in HOURS:
        i, o = hours.get(h, (0, 0))
        cells[f"IN_{h:02}"] = i
        cells[f"OUT_{h:02}"] = o
    cells["전체_IN"] = sum(v[0] for v in hours.values()) if total_in is None else total_in
    cells["전체_OUT"] = sum(v[1] for v in hours.values()) if total_out is None else total_out
    return cells


def writer(out: Path):
    def write(name, rows, header=None, sheets=("Data",)):
        wb = Workbook()
        wb.remove(wb.active)
        hdr = header or HDR
        for s in sheets:
            ws = wb.create_sheet(s)
            ws.append(hdr)
            if s == sheets[0]:
                for r in rows:
                    ws.append([r.get(c) for c in hdr])
        path = out / name
        wb.save(path)
        return path
    return write


def build_cases(write):
    ok_front = row(D1, FRONT, PF, {9: (5, 1), 10: (4, 6), 11: (1, 3)})
    ok_back = row(D1, BACK, PB, {9: (3, 0), 10: (2, 4), 11: (0, 1)})
    d2_front = row(D2, FRONT, PF, {9: (6, 2), 10: (3, 7)})
    d2_back = row(D2, BACK, PB, {9: (2, 1), 10: (1, 2)})
    summary_sum = {**row(D1, FRONT, PF, {}), "수집일자": "전체", "게이트": "그룹", "통로ID": "합계"}
    summary_avg = {**row(D1, FRONT, PF, {}), "수집일자": "전체", "게이트": "그룹", "통로ID": "평균"}

    # (id, 파일, 의도, 기대, 빌더, preprocess kwargs)
    C = []

    def case(cid, intent, expect, path_fn, **kw):
        C.append(dict(id=cid, file=path_fn().name, intent=intent, expect=expect, kwargs=kw))

    # 정상 ------------------------------------------------------------------
    case("OK-01", "정상 최소 파일 (1일 × 2게이트)", "accept",
         lambda: write("valid-minimal.xlsx", [ok_front, ok_back]), partial_dates=[])
    case("OK-02", "정상 2일치 — 최신일 수집 완료 미확인 경고", "accept+warn",
         lambda: write("valid-two-dates.xlsx", [ok_front, ok_back, d2_front, d2_back]))
    case("OK-03", "원본 합계·평균 행 포함 — 건너뜀", "accept",
         lambda: write("valid-summary-rows.xlsx", [summary_sum, summary_avg, ok_front, ok_back]),
         partial_dates=[])
    case("OK-04", "완전히 빈 행 포함 — 건너뜀", "accept",
         lambda: write("valid-blank-row.xlsx", [ok_front, {}, ok_back]), partial_dates=[])
    case("OK-05", "비필수 열 추가(시작IN·종료IN) — 무시", "accept",
         lambda: write("valid-extra-column.xlsx",
                       [{**ok_front, "시작IN": 100, "종료IN": 110},
                        {**ok_back, "시작IN": 200, "종료IN": 205}],
                       header=HDR + ["시작IN", "종료IN"]), partial_dates=[])

    # 경고 ------------------------------------------------------------------
    case("WARN-01", "시간대 합 ≠ 전체_IN/OUT — 경고 후 시간대 값으로 반영", "accept+warn",
         lambda: write("warn-hourly-mismatch.xlsx",
                       [row(D1, FRONT, PF, {9: (5, 1)}, total_in=12, total_out=9), ok_back]),
         partial_dates=[])
    case("WARN-02", "partial_dates 미지정 — 최신 날짜를 부분 데이터로 표시", "accept+warn",
         lambda: write("valid-minimal.xlsx", [ok_front, ok_back]))

    # 거부 ------------------------------------------------------------------
    case("ERR-01", "필수 열 누락 (IN_08)", "reject",
         lambda: write("invalid-missing-column.xlsx", [ok_front, ok_back],
                       header=[c for c in HDR if c != "IN_08"]))
    case("ERR-02", "필수 열 중복 (전체_IN 2회)", "reject",
         lambda: write("invalid-duplicate-column.xlsx", [ok_front, ok_back],
                       header=HDR + ["전체_IN"]))
    case("ERR-03", "같은 날짜·게이트 2행", "reject",
         lambda: write("invalid-duplicate-row.xlsx", [ok_front, ok_back, ok_front]))
    case("ERR-04", "음수 입출입 수", "reject",
         lambda: write("invalid-negative-count.xlsx", [{**ok_front, "IN_09": -1}, ok_back]))
    case("ERR-05", "결측 입출입 수 (빈 칸)", "reject",
         lambda: write("invalid-blank-count.xlsx", [{**ok_front, "IN_10": None}, ok_back]))
    case("ERR-06", "소수 입출입 수 (2.5)", "reject",
         lambda: write("invalid-decimal-count.xlsx", [{**ok_front, "IN_09": 2.5}, ok_back]))
    case("ERR-07", "날짜 형식 위반 (2026/01/06)", "reject",
         lambda: write("invalid-date-format.xlsx", [{**ok_front, "수집일자": "2026/01/06"}, ok_back]))
    case("ERR-08", "존재하지 않는 날짜 (2026-02-30)", "reject",
         lambda: write("invalid-date-impossible.xlsx",
                       [{**ok_front, "수집일자": "2026-02-30"}, ok_back]))
    case("ERR-09", "알 수 없는 게이트 (자료실.옆문)", "reject",
         lambda: write("invalid-gate.xlsx", [{**ok_front, "게이트": "자료실.옆문"}, ok_back]))
    case("ERR-10", "통로ID가 수식 (=1+1)", "reject",
         lambda: write("invalid-passage-formula.xlsx", [{**ok_front, "통로ID": "=1+1"}, ok_back]))
    case("ERR-11", "통로ID 공백", "reject",
         lambda: write("invalid-passage-empty.xlsx", [{**ok_front, "통로ID": "  "}, ok_back]))
    case("ERR-12", "데이터 행 없음 (헤더만)", "reject",
         lambda: write("invalid-header-only.xlsx", []))
    case("ERR-13", "시트 2개 + sheet 미지정", "reject",
         lambda: write("invalid-multi-sheet.xlsx", [ok_front, ok_back], sheets=("Data", "Sheet2")))
    case("ERR-14", "논리값 입출입 수 (TRUE)", "reject",
         lambda: write("invalid-bool-count.xlsx", [{**ok_front, "IN_09": True}, ok_back]))
    case("ERR-15", "지정한 부분 데이터 날짜가 파일에 없음", "reject",
         lambda: write("valid-minimal.xlsx", [ok_front, ok_back]), partial_dates=["2026-01-09"])
    case("ERR-16", "source_name이 경로", "reject",
         lambda: write("valid-minimal.xlsx", [ok_front, ok_back]),
         partial_dates=[], source_name="a/b.xlsx")

    # 계약상 통과하지만 주의가 필요한 경계 ------------------------------------
    case("EDGE-01", "모든 시간대 0 — 원본 그대로 통과", "accept",
         lambda: write("edge-all-zero.xlsx",
                       [row(D1, FRONT, PF, {}), row(D1, BACK, PB, {})]), partial_dates=[])
    case("EDGE-02", "휴관일(월) 데이터 — 전처리에서 걸러내지 않음", "accept",
         lambda: write("gap-monday.xlsx",
                       [row(D_MON, FRONT, PF, {9: (1, 0)}), row(D_MON, BACK, PB, {})]),
         partial_dates=[])
    case("EDGE-03", "23시에만 통행 — 운영시간 검사 없음", "accept",
         lambda: write("gap-hour-23.xlsx",
                       [row(D1, FRONT, PF, {23: (4, 4)}), row(D1, BACK, PB, {})]),
         partial_dates=[])
    case("EDGE-04", "같은 통로ID가 두 게이트로 보고됨(이중 집계) — 현재 미검출", "accept",
         lambda: write("gap-gate-row-duplicated.xlsx",
                       [ok_front, {**ok_back, "통로ID": PF}]), partial_dates=[])
    return C


# ── records 계층(API와 공유하는 검증) ────────────────────────────────────
def base_records(**over):
    """(date, gate) 1쌍 × 16시간 = 16 레코드. over로 특정 필드를 덮어쓴다."""
    out = []
    for h in HOURS:
        r = dict(date=D1, day_of_week="Tue", gate="front", gate_name=FRONT,
                 passage_id=PF, hour=h, in_count=1 if h == 9 else 0,
                 out_count=1 if h == 10 else 0, total_in=1, total_out=1,
                 is_partial=False, source_file="fixture.xlsx")
        out.append(r)
    for key, value in over.items():
        out[0][key] = value
    return out


def records_cases():
    return [
        ("REC-01", "정상 records 16건", "accept", base_records()),
        ("REC-02", "out_count=null 허용 (신뢰 못 하는 칸)", "accept",
         base_records(out_count=None)),
        ("REC-03", "in_count=null 거부", "reject", base_records(in_count=None)),
        ("REC-04", "quality_note 추가 필드 허용", "accept",
         base_records(quality_note="HOURLY_TOTAL_MISMATCH")),
        ("REC-05", "is_closed_day 추가 필드 허용", "accept",
         base_records(is_closed_day=True)),
        ("REC-06", "요일이 날짜와 불일치", "reject", base_records(day_of_week="Mon")),
        ("REC-07", "hour=7 (운영시간 밖)", "reject", base_records(hour=7)),
        ("REC-08", "hour=24 (운영시간 밖)", "reject", base_records(hour=24)),
        ("REC-09", "정의되지 않은 필드", "reject", base_records(unexpected="x")),
        ("REC-10", "같은 date+gate 안에서 total_in 불일치", "reject",
         base_records(total_in=99)),
        ("REC-11", "같은 date+gate+hour 중복", "reject",
         base_records() + base_records()[:1]),
        # validate_records는 빈 배열을 통과시킨다. 거부는 RecordsPublisher.publish()에서
        # 일어난다 ('빈 records 배열은 업로드할 수 없습니다.'). API 테스트는 publish 경로로 확인할 것.
        ("REC-12", "빈 배열 — validate_records 통과, 거부는 publish()", "accept", []),
    ]


def run_records(verbose=True):
    from backend.domain import validate_records, DataError as CommonDataError
    results, failures = [], []
    if verbose:
        print(f"\n{'ID':9} {'기대':9} {'실제':9} 의도 / 메시지")
        print("-" * 120)
    for cid, intent, expect, payload in records_cases():
        try:
            clean = validate_records(payload)
            actual, message = "accept", f"{len(clean)}건 통과"
        except CommonDataError as exc:
            actual, message = "reject", str(exc)[:70]
        except Exception as exc:  # noqa: BLE001
            actual, message = "crash", f"{type(exc).__name__}: {exc}"[:70]
        ok = actual == expect
        if not ok:
            failures.append((cid, expect, actual))
        if verbose:
            print(f"{cid:9} {expect:9} {actual:9} {'OK' if ok else 'MISMATCH'} {intent} — {message}")
        results.append(dict(id=cid, intent=intent, expect=expect, actual=actual, message=message))
    return results, failures


def run(out: Path, results_path: Path, verbose=True):
    out.mkdir(parents=True, exist_ok=True)
    cases = build_cases(writer(out))
    results, failures = [], []
    if verbose:
        print(f"{'ID':9} {'파일':34} {'기대':13} {'실제':13} {'코드':22} 요약")
        print("-" * 120)
    for c in cases:
        path = out / c["file"]
        try:
            records, report = preprocess(path, **c["kwargs"])
            codes = sorted(report["warning_counts"])
            actual = "accept+warn" if codes else "accept"
            code = ",".join(codes)
            summary = f"{report['source_rows']}행→{report['output_rows']}레코드"
            detail = dict(records=len(records),
                          report={k: report[k] for k in
                                  ("source_rows", "output_rows", "dates", "partial_dates",
                                   "warning_counts", "skipped_rows")})
        except DataError as exc:
            actual, code = "reject", exc.code
            summary = str(exc)[:46]
            detail = dict(message=str(exc), issues=exc.issues[:3])
        except Exception as exc:  # noqa: BLE001
            actual, code, summary = "crash", type(exc).__name__, str(exc)[:46]
            detail = dict(traceback=traceback.format_exc()[-300:])
        ok = actual == c["expect"]
        if not ok:
            failures.append((c["id"], c["expect"], actual, code))
        if verbose:
            print(f"{c['id']:9} {c['file']:34} {c['expect']:13} {actual:13} {code:22} "
                  f"{'OK' if ok else 'MISMATCH'} {summary}")
        results.append(dict(**{k: c[k] for k in ("id", "file", "intent", "expect")},
                            kwargs={k: v for k, v in c["kwargs"].items()},
                            actual=actual, code=code, detail=detail))
    rec_results, rec_failures = run_records(verbose=verbose)
    results_path.parent.mkdir(parents=True, exist_ok=True)
    results_path.write_text(json.dumps(
        {"sha": "928a337bfa2e9f84549df876156f931033bdfc75",
         "excel_cases": results, "records_cases": rec_results},
        ensure_ascii=False, indent=2), encoding="utf-8")
    failures += rec_failures
    if verbose:
        print(f"\nfixture {len({c['file'] for c in cases})}개 / Excel 케이스 {len(cases)}개 → {out}")
        print(f"records 케이스 {len(rec_results)}개")
        print(f"결과 → {results_path}")
        if failures:
            print("\n계약과 다른 케이스:")
            for f in failures:
                print("  ", f)
        else:
            print("모든 케이스가 계약 문서와 일치합니다.")
    return failures


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(ROOT / "contracts" / "fixtures" / "excel"))
    ap.add_argument("--results", default=str(ROOT / "contracts" / "excel-validation-results.json"))
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    sys.exit(1 if run(Path(args.out), Path(args.results), verbose=not args.quiet) else 0)
