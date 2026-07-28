#!/usr/bin/env python3
"""
조달청 나라장터 종합쇼핑몰 물품정보 서비스(ShoppingMallPrdctInfoService)를 조회해서
지정한 업체(기본값: 두원전자통신)가 등록한 보안용 카메라가 몇 종류인지 확인하는 스크립트.

data.go.kr 활용신청: https://www.data.go.kr/data/15129471/openapi.do
Endpoint: https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService

사용법 (CLI):
    export DATA_GO_KR_SERVICE_KEY="발급받은 서비스키(디코딩된 일반 인증키)"
    python3 scripts/check_doowon_cameras.py
    python3 scripts/check_doowon_cameras.py --corp-name "다른업체명" --keyword "CCTV"
    python3 scripts/check_doowon_cameras.py --output 결과.xlsx

GUI로 실행하려면 scripts/check_doowon_cameras_gui.py 를 실행하세요.

서비스키는 절대 코드에 하드코딩하거나 커밋하지 마세요. 환경변수로만 주입합니다.

매칭된 물품 목록은 --output으로 지정한 경로(기본값: 두원전자통신_보안용카메라.xlsx)에 저장됩니다.
openpyxl이 설치되어 있으면 엑셀(.xlsx)로, 없으면 같은 이름의 .csv로 대신 저장합니다.
엑셀로 저장하려면: pip install openpyxl

주의: 실제 오퍼레이션명/요청·응답 필드명은 data.go.kr의 참고문서
("조달청_OpenAPI참고자료_조달청 나라장터쇼핑몰물품목록정보서비스 1.3.docx")로 확인된 값을 사용합니다
(getThptyUcntrctPrdctInfoList, cntrctCorpNm).
"""

import argparse
import csv
import datetime
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

BASE_URL = "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService"

# cntrctCorpNm 필터가 서버에서 무시되어 전국 데이터를 전부 받아오게 되는 사고를 막기 위한 안전장치.
MAX_ITEMS = 5000

# 후보 오퍼레이션 목록 (실제로 확인된 getThptyUcntrctPrdctInfoList를 우선 사용)
CANDIDATE_OPERATIONS = [
    "getThptyUcntrctPrdctInfoList",  # 제3자단가계약 물품 목록
]

# 응답에서 품명(카메라 종류 구분)으로 실제 확인된 필드
NAME_FIELD = "prdctClsfcNoNm"  # 물품분류번호명 (품명)

# 키워드 검색 대상 필드. 품명(대분류)만으로는 "안내전광판"처럼 더 구체적인 단어가
# 걸리지 않을 수 있어서, 물품식별번호명(모델/규격 설명)도 같이 검사한다.
KEYWORD_FIELDS = [NAME_FIELD, "prdctIdntNoNm"]

EXPORT_COLUMNS = [
    ("계약업체명", "cntrctCorpNm"),
    ("품명", "prdctClsfcNoNm"),
    ("규격/모델명", "prdctSpecNm"),
    ("물품식별번호", "prdctIdntNo"),
    ("제조사", "prdctMakrNm"),
    ("원산지", "prdctOrgplceNm"),
    ("계약방법", "cntrctMthdNm"),
    ("계약번호", "shopngCntrctNo"),
    ("계약일자", "cntrctDate"),
    ("계약시작일", "cntrctBgnDate"),
    ("계약종료일", "cntrctEndDate"),
    ("계약금액", "cntrctPrceAmt"),
    ("단위", "prdctUnit"),
]


def fetch(
    operation: str,
    service_key: str,
    params: dict,
    timeout: int = 90,
    retries: int = 5,
    log=print,
) -> dict:
    query = dict(params)
    query["serviceKey"] = service_key
    url = f"{BASE_URL}/{operation}?{urllib.parse.urlencode(query)}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})

    # 502/503/504는 게이트웨이/서버가 일시적으로 과부하일 때 나는 오류라 재시도할 가치가 있다.
    RETRYABLE_HTTP_CODES = (502, 503, 504)

    last_error = None
    for attempt in range(1, retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
            break
        except TimeoutError as e:
            last_error = e
            log(f"  (타임아웃, {attempt}/{retries}번째 시도 실패: {e})")
        except urllib.error.HTTPError as e:
            if e.code not in RETRYABLE_HTTP_CODES:
                raise
            last_error = e
            log(f"  (HTTP {e.code} {e.reason}, {attempt}/{retries}번째 시도 실패)")
        else:
            continue
        if attempt < retries:
            wait = min(5 * attempt, 20)
            log(f"  {wait}초 대기 후 재시도합니다...")
            time.sleep(wait)
    else:
        raise last_error

    try:
        return {"raw": raw, "json": json.loads(raw)}
    except json.JSONDecodeError:
        return {"raw": raw, "json": None}


def extract_items(payload: dict):
    """data.go.kr 표준 응답 포맷(response.body.items)에서 item 리스트를 뽑아본다."""
    if not payload:
        return []
    body = (
        payload.get("response", {}).get("body", {})
        if isinstance(payload.get("response"), dict)
        else None
    )
    if not body:
        return []
    items = body.get("items")
    if isinstance(items, dict):
        item = items.get("item", [])
        return item if isinstance(item, list) else [item]
    if isinstance(items, list):
        return items
    return []


def run_query(
    corp_name: str,
    keyword: str,
    begin_date: str,
    end_date: str,
    service_key: str,
    num_of_rows: int = 999,
    log=print,
) -> dict:
    """API를 조회하고 회사명/키워드로 필터링한 결과를 dict로 반환한다.

    반환값: {"ok": bool, "error": str|None, "all_items": [...],
             "camera_items": [...], "distinct_names": [...], "name_field": str|None}
    """
    base_params = {
        "numOfRows": str(num_of_rows),
        "type": "json",
        "cntrctCorpNm": corp_name,
        "inqryDiv": "1",
        "inqryBgnDate": begin_date,
        "inqryEndDate": end_date,
    }

    working_operation = None
    all_items = []

    for operation in CANDIDATE_OPERATIONS:
        log(f"\n=== 오퍼레이션 시도: {operation} ===")
        page_no = 1
        operation_items = []
        total_count = None
        while True:
            params = dict(base_params, pageNo=str(page_no))
            try:
                result = fetch(operation, service_key, params, log=log)
            except urllib.error.HTTPError as e:
                log(f"HTTP 오류: {e.code} {e.reason}")
                break
            except urllib.error.URLError as e:
                log(f"네트워크 오류: {e.reason}")
                break
            except TimeoutError as e:
                log(f"타임아웃 (재시도 모두 실패): {e}")
                break

            payload = result["json"]
            if payload is None:
                log("JSON이 아닌 응답 (앞부분 500자):")
                log(result["raw"][:500])
                break

            header = payload.get("response", {}).get("header", {})
            result_code = header.get("resultCode")
            result_msg = header.get("resultMsg")
            log(f"[page {page_no}] resultCode={result_code} resultMsg={result_msg}")

            if result_code not in ("00", "0", None):
                break

            body = payload.get("response", {}).get("body", {})
            if total_count is None:
                total_count = body.get("totalCount")
            items = extract_items(payload)
            log(f"[page {page_no}] 조회된 item 수: {len(items)} (totalCount={total_count})")
            if not items:
                break
            if page_no == 1:
                working_operation = operation
            operation_items.extend(items)

            if total_count is not None and len(operation_items) >= int(total_count):
                break
            if len(items) < num_of_rows:
                break
            if len(operation_items) >= MAX_ITEMS:
                log(
                    f"경고: {MAX_ITEMS}건 이상 조회되어 중단합니다. "
                    "cntrctCorpNm 필터가 서버에서 적용되지 않았을 수 있습니다."
                )
                break
            page_no += 1
            time.sleep(1)  # 연속 요청으로 서버에 부담을 주지 않도록 짧게 대기

        if working_operation:
            all_items = operation_items
            break

    # 서버가 cntrctCorpNm 파라미터를 무시하고 전체 목록을 반환하는 경우에 대비해
    # 응답에 실제로 들어있는 회사명 필드로 다시 한번 걸러낸다.
    corp_field = next(
        (f for f in ("cntrctCorpNm", "corpNm") if all_items and f in all_items[0]), None
    )
    if corp_field:
        before = len(all_items)
        all_items = [item for item in all_items if corp_name in str(item.get(corp_field, ""))]
        log(
            f"\n'{corp_field}' 필드 기준으로 '{corp_name}' 클라이언트측 재필터링: "
            f"{before}건 -> {len(all_items)}건"
        )

    if not working_operation:
        log(
            "\n어떤 오퍼레이션도 정상 응답(item 포함)을 주지 않았습니다. "
            "resultCode/resultMsg를 참고해서 파라미터를 조정해주세요."
        )
        return {"ok": False, "error": "no_working_operation", "all_items": [],
                "camera_items": [], "distinct_names": [], "name_field": None}

    log(f"\n성공한 오퍼레이션: {working_operation}")

    if not all_items:
        log(f"'{corp_name}' 이름과 일치하는 등록 물품이 없습니다.")
        log("회사명 표기가 다를 수 있습니다(예: '(주)두원전자통신' 등). 다른 표기로 다시 시도해보세요.")
        return {"ok": False, "error": "no_matching_corp", "all_items": [],
                "camera_items": [], "distinct_names": [], "name_field": None}

    if NAME_FIELD not in all_items[0]:
        log(f"'{NAME_FIELD}' 필드를 찾지 못했습니다. item 전체 필드 목록:")
        log(str(list(all_items[0].keys())))
        return {"ok": False, "error": "no_name_field", "all_items": all_items,
                "camera_items": [], "distinct_names": [], "name_field": None}

    name_field = NAME_FIELD
    all_categories = sorted({str(item.get(name_field, "")) for item in all_items})

    def matches_keyword(item):
        return any(keyword in str(item.get(f, "")) for f in KEYWORD_FIELDS)

    camera_items = [item for item in all_items if matches_keyword(item)]
    distinct_names = sorted({str(item.get(name_field, "")) for item in camera_items})

    log(f"\n'{corp_name}' 전체 등록 물품 수: {len(all_items)}")
    log(f"조회 기간 내 등록된 전체 품명({name_field}) 종류 ({len(all_categories)}개):")
    for category in all_categories:
        log(f"  - {category}")

    log(f"\n'{keyword}'({'/'.join(KEYWORD_FIELDS)} 중 포함) 물품 수: {len(camera_items)}")
    if not camera_items:
        log(
            f"'{keyword}'와(과) 일치하는 품명/규격이 없습니다. "
            "위 전체 품명 목록에서 정확한 표기를 확인해 --keyword 값을 맞춰보세요. "
            "조회 기간(--begin-date/--end-date) 밖의 계약이라 안 보일 수도 있습니다."
        )
    log(f"'{keyword}' 관련 물품 종류(고유 {name_field} 개수): {len(distinct_names)}")
    for name in distinct_names:
        log(f"  - {name}")

    return {
        "ok": True,
        "error": None,
        "all_items": all_items,
        "camera_items": camera_items,
        "distinct_names": distinct_names,
        "name_field": name_field,
    }


# 이 필드는 문자열이 아니라 숫자로 저장하고, 엑셀에서 천단위 구분(회계 서식)으로 표시한다.
NUMERIC_EXPORT_FIELDS = {"cntrctPrceAmt"}
ACCOUNTING_NUMBER_FORMAT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'


def write_output(camera_items, corp_name, keyword, begin_date, end_date, output_path, log=print):
    """카메라 목록을 엑셀(.xlsx)로 저장한다. openpyxl이 없으면 CSV로 대신 저장한다."""
    try:
        import openpyxl
    except ImportError:
        openpyxl = None

    def cell_value(item, field):
        value = item.get(field, "")
        if field in NUMERIC_EXPORT_FIELDS and value not in ("", None):
            try:
                return int(value)
            except (TypeError, ValueError):
                pass
        return value

    rows = [[cell_value(item, field) for _, field in EXPORT_COLUMNS] for item in camera_items]
    headers = [header for header, _ in EXPORT_COLUMNS]

    if openpyxl is not None:
        path = output_path
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "카메라목록"
        ws.append(headers)
        for row in rows:
            ws.append(row)

        numeric_cols = [
            col_idx for col_idx, (_, field) in enumerate(EXPORT_COLUMNS, start=1)
            if field in NUMERIC_EXPORT_FIELDS
        ]
        for col_idx in numeric_cols:
            for row_idx in range(2, len(rows) + 2):
                ws.cell(row=row_idx, column=col_idx).number_format = ACCOUNTING_NUMBER_FORMAT

        for col_idx, header in enumerate(headers, start=1):
            width = max(len(header), *(len(str(r[col_idx - 1])) for r in rows)) if rows else len(header)
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(width + 2, 60)

        summary = wb.create_sheet("요약")
        summary.append(["업체명", corp_name])
        summary.append(["필터 키워드", keyword])
        summary.append(["조회 기간", f"{begin_date} ~ {end_date}"])
        summary.append(["매칭 물품 수", len(camera_items)])

        wb.save(path)
        log(f"\n엑셀 파일로 저장했습니다: {path}")
        return path
    else:
        path = output_path
        if not path.lower().endswith(".csv"):
            path = os.path.splitext(path)[0] + ".csv"
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
        log(f"\nopenpyxl이 설치되어 있지 않아 CSV로 저장했습니다: {path}")
        log("엑셀(.xlsx)로 저장하려면 'pip install openpyxl' 실행 후 다시 실행하세요.")
        return path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corp-name", default="두원전자통신", help="조회할 업체명")
    parser.add_argument("--keyword", default="보안용카메라", help="품명/규격에서 필터링할 키워드")
    parser.add_argument(
        "--num-of-rows",
        type=int,
        default=999,
        help="페이지당 조회 건수. 업체 하나의 물품 수는 보통 999건 이내라 한 번에 끝나는 게 더 빠름. "
             "타임아웃이 잦으면 낮춰서 시도해보세요",
    )
    parser.add_argument(
        "--begin-date",
        default=(datetime.date.today() - datetime.timedelta(days=365)).strftime("%Y%m%d"),
        help="조회 시작일자(YYYYMMDD). 기본값: 오늘로부터 1년 전",
    )
    parser.add_argument(
        "--end-date",
        default=datetime.date.today().strftime("%Y%m%d"),
        help="조회 종료일자(YYYYMMDD)",
    )
    parser.add_argument(
        "--output",
        default="두원전자통신_보안용카메라.xlsx",
        help="결과를 저장할 엑셀 파일 경로 (openpyxl 미설치 시 같은 이름의 .csv로 대신 저장)",
    )
    args = parser.parse_args()

    service_key = os.environ.get("DATA_GO_KR_SERVICE_KEY")
    if not service_key:
        print("환경변수 DATA_GO_KR_SERVICE_KEY 가 설정되어 있지 않습니다.", file=sys.stderr)
        print('예: export DATA_GO_KR_SERVICE_KEY="발급받은 서비스키"', file=sys.stderr)
        sys.exit(1)

    result = run_query(
        args.corp_name,
        args.keyword,
        args.begin_date,
        args.end_date,
        service_key,
        num_of_rows=args.num_of_rows,
    )

    if not result["ok"]:
        sys.exit(2)

    write_output(
        result["camera_items"],
        args.corp_name,
        args.keyword,
        args.begin_date,
        args.end_date,
        args.output,
    )


if __name__ == "__main__":
    main()
