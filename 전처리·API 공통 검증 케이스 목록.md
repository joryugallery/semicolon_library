# 전처리 · API 공통 검증 케이스 목록

**대상 SHA** `928a337bfa2e9f84549df876156f931033bdfc75`
**규칙 원문** [`docs/excel-upload-contract.md`](../docs/excel-upload-contract.md)
**마지막 실행** 2026-10-05 · 39개 케이스 전부 계약과 일치

이 문서는 규칙을 설명하지 않습니다. 규칙은 계약 문서에 있고, 이 문서는 **그 규칙이 실제로 어떻게
동작하는지 확인한 케이스 목록**입니다. 전처리 테스트(박수용)와 API 테스트(이가영)는 같은 ID를 써서
같은 입력에 같은 결과를 기대하면 됩니다.

---

## 0. 실행 방법

```bash
# 저장소 루트에서
python scripts/make_excel_fixtures.py
```

- fixture 24개를 `contracts/fixtures/excel/` 에 새로 만들고, 실제 `preprocess()` 와
  `validate_records()` 에 그대로 통과시킵니다.
- 아래 표의 기대값과 하나라도 다르면 **종료 코드 1** 과 함께 어긋난 케이스 ID를 출력합니다.
  CI에서는 종료 코드만 보면 됩니다.
- 케이스별 실제 결과는 `contracts/excel-validation-results.json` 에 기계가 읽을 형태로 남습니다.
  (테스트에서 기대값을 하드코딩하지 않고 이 파일을 읽어 쓰는 것도 가능합니다.)

fixture는 전부 합성 데이터입니다. 날짜는 2026-01-05~07, 통로ID는 `10.0.0.1A` / `10.0.0.2A`,
사람 수는 한 자리 숫자입니다. **개인정보도, 실제 수집 기록도 들어 있지 않습니다.**

---

## 1. Excel 업로드 계층 (`library_etl.pipeline.preprocess`)

`기대` 열의 의미
- `accept` — 레코드가 나오고 경고 없음
- `accept+warn` — 레코드가 나오지만 `report['warning_counts']` 에 경고가 있음
- `reject` — `DataError` 발생, **파일 전체가 반영되지 않음**

| ID | fixture | 입력 요지 | 호출 인자 | 기대 | 전처리 코드 | API 코드 / HTTP |
|---|---|---|---|---|---|---|
| OK-01 | `valid-minimal.xlsx` | 1일 × 2게이트, 정상 | `partial_dates=[]` | accept | — | 200 |
| OK-02 | `valid-two-dates.xlsx` | 2일치 | (기본) | accept+warn | `PARTIAL_DATE_UNCONFIRMED` | 200 + warning |
| OK-03 | `valid-summary-rows.xlsx` | 원본 합계·평균 행 포함 | `partial_dates=[]` | accept | — | 200 |
| OK-04 | `valid-blank-row.xlsx` | 완전히 빈 행 포함 | `partial_dates=[]` | accept | — | 200 |
| OK-05 | `valid-extra-column.xlsx` | 비필수 열(`시작IN`·`종료IN`) 추가 | `partial_dates=[]` | accept | — | 200 |
| WARN-01 | `warn-hourly-mismatch.xlsx` | 시간대 합 ≠ `전체_IN`/`전체_OUT` | `partial_dates=[]` | accept+warn | `HOURLY_TOTAL_MISMATCH` | 200 + warning |
| WARN-02 | `valid-minimal.xlsx` | 수집 완료 여부 미확인 | (기본) | accept+warn | `PARTIAL_DATE_UNCONFIRMED` | 200 + warning |
| ERR-01 | `invalid-missing-column.xlsx` | 필수 열 누락 (`IN_08`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-02 | `invalid-duplicate-column.xlsx` | 필수 열 중복 (`전체_IN` 2회) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-03 | `invalid-duplicate-row.xlsx` | 같은 날짜·게이트 2행 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-04 | `invalid-negative-count.xlsx` | 음수 입출입 수 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-05 | `invalid-blank-count.xlsx` | 빈 칸 입출입 수 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-06 | `invalid-decimal-count.xlsx` | 소수 입출입 수 (2.5) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-07 | `invalid-date-format.xlsx` | 날짜 형식 위반 (`2026/01/06`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-08 | `invalid-date-impossible.xlsx` | 없는 날짜 (`2026-02-30`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-09 | `invalid-gate.xlsx` | 알 수 없는 게이트 (`자료실.옆문`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-10 | `invalid-passage-formula.xlsx` | 통로ID가 수식 (`=1+1`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-11 | `invalid-passage-empty.xlsx` | 통로ID 공백 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-12 | `invalid-header-only.xlsx` | 데이터 행 없음 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-13 | `invalid-multi-sheet.xlsx` | 시트 2개 + `sheet` 미지정 | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-14 | `invalid-bool-count.xlsx` | 논리값 입출입 수 (`TRUE`) | (기본) | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-15 | `valid-minimal.xlsx` | 파일에 없는 날짜를 부분 데이터로 지정 | `partial_dates=["2026-01-09"]` | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| ERR-16 | `valid-minimal.xlsx` | `source_name` 이 경로 | `source_name="a/b.xlsx"` | reject | `EXCEL_FORMAT_ERROR` | `INVALID_EXCEL` / 422 |
| EDGE-01 | `edge-all-zero.xlsx` | 모든 시간대 0 | `partial_dates=[]` | accept | — | 200 |
| EDGE-02 | `gap-monday.xlsx` | 휴관일(월) 데이터 | `partial_dates=[]` | accept | — | 200 |
| EDGE-03 | `gap-hour-23.xlsx` | 23시에만 통행 | `partial_dates=[]` | accept | — | 200 |
| EDGE-04 | `gap-gate-row-duplicated.xlsx` | 같은 통로ID가 두 게이트로 보고 | `partial_dates=[]` | accept | — | 200 |

### 이 표에서 반드시 같이 읽어야 할 세 가지

1. **거부 코드는 한 종류입니다.** ERR-01~16은 원인이 모두 다르지만 전처리는 전부
   `EXCEL_FORMAT_ERROR`, API는 전부 `INVALID_EXCEL` / 422 로 내려옵니다. 어떤 줄의 무엇이
   문제였는지는 `DataError.issues` 와 API 응답의 `details[].reason` 에만 있습니다.
   **테스트는 코드만 보고 통과시키지 말고, `issues` 안의 `row`·`field` 를 함께 확인하세요.**
   그렇지 않으면 "다른 이유로 거부된" 파일도 통과합니다.
2. **EDGE-01~04은 통과가 정답이지만 안심할 결과는 아닙니다.** 전처리는 원본을 그대로
   보존하므로 휴관일 데이터, 운영시간 밖 통행, 전부 0인 행을 걸러내지 않습니다. EDGE-04는
   같은 통로를 두 게이트로 이중 집계해도 막히지 않는 **현재 구멍**입니다
   (계약 문서 §11-1). 지금은 `accept` 가 맞으니 그렇게 기대값을 쓰고, 규칙이 바뀌면 이 네
   줄부터 고치면 됩니다.
3. **`partial_dates` 를 생략하면 항상 경고가 붙습니다.** OK-01과 WARN-02는 같은 파일인데
   인자만 다릅니다. "정상 파일인데 경고가 없어야 한다"는 테스트를 쓸 거라면
   `partial_dates=[]` 를 명시해야 합니다.

---

## 2. records 계층 (`backend.domain.validate_records`)

Excel을 거치지 않고 `/api/v1/admin/records` 로 바로 올라오는 경로이고, Excel 경로의 마지막
단계이기도 합니다. 즉 **여기 규칙은 두 경로가 공유합니다.**

기준 입력은 `(2026-01-06, front)` 1쌍 × 16시간 = 16 레코드이고, 각 케이스는 첫 레코드의 한
필드만 바꿉니다.

| ID | 바꾼 것 | 기대 | 메시지 |
|---|---|---|---|
| REC-01 | (없음) | accept | 16건 통과 |
| REC-02 | `out_count = null` | accept | 16건 통과 |
| REC-03 | `in_count = null` | reject | `in_count 결측/음수/비정수` |
| REC-04 | `quality_note` 추가 | accept | 16건 통과 |
| REC-05 | `is_closed_day` 추가 | accept | 16건 통과 |
| REC-06 | `day_of_week = "Mon"` (날짜와 불일치) | reject | `요일 또는 게이트 오류` |
| REC-07 | `hour = 7` | reject | `hour는 8~23 정수입니다.` |
| REC-08 | `hour = 24` | reject | `hour는 8~23 정수입니다.` |
| REC-09 | 정의되지 않은 필드 추가 | reject | `필수 12개 필드와 선택 품질 필드만 허용합니다.` |
| REC-10 | `total_in = 99` (같은 날짜·게이트 안에서 불일치) | reject | `날짜/게이트 내 일일 합계 또는 partial 불일치` |
| REC-11 | 같은 `date+gate+hour` 중복 | reject | `중복 date+gate+hour` |
| REC-12 | 빈 배열 | **accept** | 0건 통과 |

### API 테스트가 걸려 넘어질 두 지점

- **REC-02 / REC-03 비대칭.** records는 `out_count = null` 을 허용하지만 `in_count = null` 은
  거부합니다. 그런데 **Excel 경로로는 `out_count = null` 을 만들 수 없습니다** — 빈 칸은
  ERR-05로 거부되기 때문입니다. 따라서 "믿을 수 없는 OUT 한 칸"은 지금 Excel로 표현할 방법이
  없습니다(계약 문서 §11-2). API 테스트는 `null` 을 받아들이도록 쓰되, 그 입력이 Excel에서는
  올 수 없다는 점을 주석으로 남겨 주세요.
- **REC-12 는 거부가 아닙니다.** `validate_records([])` 는 통과합니다. 빈 배열 거부는
  `RecordsPublisher.publish()` 에서 `빈 records 배열은 업로드할 수 없습니다.` 로 일어납니다.
  빈 배열 테스트는 **반드시 publish 경로(또는 엔드포인트)로** 작성해야 하고,
  `validate_records` 단위 테스트로 쓰면 틀립니다.
- REC-11 의 중복은 API에서 `PROCESSING_ERROR` 가 아니라 `GATE_ROW_DUPLICATED` / 422 로
  내려옵니다. `backend/app.py` 가 메시지에 `중복 date+gate+hour` 가 들어있는지로 구분하기
  때문입니다. 즉 **이 코드는 메시지 문자열에 의존합니다.** 메시지를 고치면 코드가 조용히
  `PROCESSING_ERROR` 로 바뀝니다. 메시지를 건드리는 PR에서는 이 케이스를 꼭 돌리세요.

---

## 3. 전송 계층 (`/api/v1/admin/records`)

이 구간은 HTTP 요청이 필요해서 fixture 스크립트가 돌리지 않습니다. 아래는 코드에서 읽은
기대값이며(`backend/admin_upload.py`, `backend/app.py`), **API 테스트에서 채워야 할 몫**입니다.

| ID | 요청 | 기대 코드 | HTTP |
|---|---|---|---|
| API-01 | `Authorization` 헤더 없음 / `Bearer ` 아님 / 토큰 불일치 | `UNAUTHORIZED` | 401 |
| API-02 | `Content-Type` 이 `application/json` 아님 | `UNSUPPORTED_MEDIA_TYPE` | 415 |
| API-03 | `Content-Length` 가 50MB 초과 | `PAYLOAD_TOO_LARGE` | 413 |
| API-04 | 본문이 50MB 초과(헤더 없이 스트리밍) | `PAYLOAD_TOO_LARGE` | 413 |
| API-05 | `Content-Length` 가 숫자가 아님 | `INVALID_REQUEST` | 400 |
| API-06 | 본문이 JSON이 아님 | `INVALID_JSON` | 400 |
| API-07 | records 스키마 위반 (REC-03 등) | `PROCESSING_ERROR` | 422 |
| API-08 | `date+gate+hour` 중복 (REC-11) | `GATE_ROW_DUPLICATED` | 422 |
| API-09 | 버전 관리 데이터가 이미 있음 | `PUBLISH_IN_PROGRESS` | 409 |
| API-10 | 교체 중 디스크 오류 — **기존 정상본 유지** | `PUBLISH_FAILED` | 500 |
| API-11 | 업로드 기능이 꺼진 실행 환경 | `UPLOAD_DISABLED` | 503 |
| API-12 | 같은 내용 재업로드 | 200, `changed: false` | 200 |

API-12는 실패가 아닙니다. 같은 스냅샷이면 백업도 교체도 하지 않고 `changed: false` 로
돌려줍니다. "두 번 올렸는데 백업이 안 생긴다"는 버그가 아닙니다.

---

## 4. 오류 코드 한눈에

| 코드 | 어디서 | HTTP | 뜻 |
|---|---|---|---|
| `EXCEL_FORMAT_ERROR` | `pipeline.preprocess` | — | Excel 거부(내부 코드) |
| `INVALID_EXCEL` | `versions.upload_excel` | 422 | 위를 감싼 API 코드 |
| `HOURLY_TOTAL_MISMATCH` | 경고 | 200 | 시간대 합 ≠ 전체. 시간대 값으로 반영 |
| `PARTIAL_DATE_UNCONFIRMED` | 경고 | 200 | 수집 완료 미확인. 최신일을 부분 데이터로 표시 |
| `PROCESSING_ERROR` | records 검증 | 422 | 스키마 위반 |
| `GATE_ROW_DUPLICATED` | records 검증 | 422 | `date+gate+hour` 중복 |
| `UNAUTHORIZED` | 전송 | 401 | 토큰 없음/불일치 |
| `UNSUPPORTED_MEDIA_TYPE` | 전송 | 415 | JSON 아님 |
| `PAYLOAD_TOO_LARGE` | 전송 | 413 | 50MB 초과 |
| `INVALID_JSON` / `INVALID_REQUEST` | 전송 | 400 | 본문/헤더 오류 |
| `PUBLISH_IN_PROGRESS` | 발행 | 409 | 버전 관리 경로와 충돌 |
| `PUBLISH_FAILED` | 발행 | 500 | 교체 실패, 기존본 유지 |
| `UPLOAD_DISABLED` | 발행 | 503 | 기능 비활성 |

---

## 5. 테스트를 같은 규칙으로 쓰기 위한 약속

1. **케이스 ID를 테스트 이름에 넣습니다.** `test_err_05_blank_count_rejected` 처럼요. 전처리와
   API 양쪽에서 같은 ID가 보이면 같은 규칙을 보고 있다는 뜻입니다.
2. **fixture를 직접 만들지 않습니다.** 새 케이스가 필요하면
   `scripts/make_excel_fixtures.py` 의 `build_cases()` 에 한 줄 추가하고 스크립트를 돌립니다.
   손으로 만든 `.xlsx` 는 리뷰에서 무엇이 달라졌는지 볼 수 없습니다.
3. **거부 테스트는 코드 + `issues` 를 같이 봅니다.** (§1의 주의 1)
4. **경고 테스트는 `warning_counts` 의 키로 봅니다.** 메시지 문장은 바뀔 수 있습니다.
5. **기대값을 바꿀 때는 계약 문서를 같은 PR에서 고칩니다.** 이 목록과
   `docs/excel-upload-contract.md` 가 어긋나면 그 PR은 되돌려야 합니다.

---

## 6. 실행 기록

| 날짜 | SHA | Excel | records | 결과 |
|---|---|---|---|---|
| 2026-10-05 | `928a337` | 27 | 12 | 39/39 계약과 일치 |

처음 돌렸을 때 두 건이 어긋났고, 둘 다 **계약 문서가 아니라 제 기대값이 틀린 것**이었습니다.
기록으로 남깁니다.

- EDGE-04: 같은 게이트를 두 번 넣어 만들었더니 기존 `(날짜, 게이트)` 중복 검사에 걸렸습니다.
  실제 구멍은 *게이트가 다른* 이중 집계였고, 그렇게 고치자 통과했습니다.
- REC-12: 빈 배열이 거부될 줄 알았는데 `validate_records` 는 통과시킵니다. 거부 지점이
  `publish()` 라는 것을 이 케이스로 확인했습니다.

---

## 7. 아직 케이스로 막지 못하는 것

계약 문서 §11의 미해결 항목과 같습니다. 팀 결정이 나면 여기 케이스를 추가합니다.

| 항목 | 지금 | 결정이 필요한 것 |
|---|---|---|
| 같은 통로의 이중 집계 (EDGE-04) | 통과 | 경고로 잡을지, 거부할지 |
| 믿을 수 없는 한 칸 (REC-02/03) | Excel로 표현 불가 | 빈 칸 허용 vs 별도 입력 |
| 휴관일·운영시간 밖 데이터 (EDGE-02/03) | 통과 | 전처리에서 걸러낼지, 조회에서 가릴지 |
| 전송 계층 (API-01~12) | 미실행 | API 테스트에서 채우기 |
