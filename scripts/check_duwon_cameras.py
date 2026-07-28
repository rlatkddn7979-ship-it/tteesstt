#!/usr/bin/env python3
"""
조달청 나라장터 종합쇼핑몰 물품정보 서비스(ShoppingMallPrdctInfoService)를 조회해서
지정한 업체(기본값: 두원전자통신)가 등록한 보안용 카메라가 몇 종류인지 확인하는 스크립트.

data.go.kr 활용신청: https://www.data.go.kr/data/15129471/openapi.do
Endpoint: https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService

사용법:
    export DATA_GO_KR_SERVICE_KEY="발급받은 서비스키(디코딩된 일반 인증키)"
    python3 scripts/check_duwon_cameras.py
    python3 scripts/check_duwon_cameras.py --corp-name "다른업체명" --keyword "CCTV"

서비스키는 절대 코드에 하드코딩하거나 커밋하지 마세요. 환경변수로만 주입합니다.

주의: 실제 오퍼레이션명/요청·응답 필드명은 data.go.kr의 참고문서
("조달청_OpenAPI참고자료_조달청 나라장터쇼핑몰물품목록정보서비스 1.3.docx")로 검증되지 않았습니다.
이 스크립트는 여러 후보 오퍼레이션을 순서대로 호출해보면서, 어떤 오퍼레이션이
실제로 서비스되는지/필수 파라미터가 무엇인지 응답(에러 메시지 포함)을 그대로 출력합니다.
성공하는 오퍼레이션을 찾으면 그 결과로 카메라 종류 수를 집계합니다.
"""

import argparse
import datetime
import json
import os
import sys
import urllib.parse
import urllib.request

BASE_URL = "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService"

# 후보 오퍼레이션 목록 (문서로 확인 전까지는 추정치이므로 순서대로 시도한다)
CANDIDATE_OPERATIONS = [
    "getThptyUnyPrceBassApnetPrdlstInfoList",  # 제3자단가계약 기준단가 적용 물품 목록
]

# 응답에서 "카메라 종류"를 식별할 때 후보가 될 필드명들 (문서 확인 전 추정)
NAME_FIELD_CANDIDATES = [
    "prdctIdntNoNm",  # 물품식별번호명 (모델/규격명)
    "prdctClsfcNoNm",  # 물품분류번호명 (품명)
    "dtilPrdctClsfcNoNm",
    "prdctNm",
]


def fetch(operation: str, service_key: str, params: dict, timeout: int = 20) -> dict:
    query = dict(params)
    query["serviceKey"] = service_key
    url = f"{BASE_URL}/{operation}?{urllib.parse.urlencode(query)}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
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


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corp-name", default="두원전자통신", help="조회할 업체명")
    parser.add_argument("--keyword", default="보안용카메라", help="품명/규격에서 필터링할 키워드")
    parser.add_argument("--num-of-rows", type=int, default=999)
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
    args = parser.parse_args()

    service_key = os.environ.get("DATA_GO_KR_SERVICE_KEY")
    if not service_key:
        print("환경변수 DATA_GO_KR_SERVICE_KEY 가 설정되어 있지 않습니다.", file=sys.stderr)
        print('예: export DATA_GO_KR_SERVICE_KEY="발급받은 서비스키"', file=sys.stderr)
        sys.exit(1)

    base_params = {
        "pageNo": "1",
        "numOfRows": str(args.num_of_rows),
        "type": "json",
        "corpNm": args.corp_name,
        "inqryDiv": "1",
        "inqryBgnDate": args.begin_date,
        "inqryEndDate": args.end_date,
    }

    working_operation = None
    all_items = []

    for operation in CANDIDATE_OPERATIONS:
        print(f"\n=== 오퍼레이션 시도: {operation} ===")
        try:
            result = fetch(operation, service_key, base_params)
        except urllib.error.HTTPError as e:
            print(f"HTTP 오류: {e.code} {e.reason}")
            continue
        except urllib.error.URLError as e:
            print(f"네트워크 오류: {e.reason}")
            continue

        payload = result["json"]
        if payload is None:
            print("JSON이 아닌 응답 (앞부분 500자):")
            print(result["raw"][:500])
            continue

        header = payload.get("response", {}).get("header", {})
        result_code = header.get("resultCode")
        result_msg = header.get("resultMsg")
        print(f"resultCode={result_code} resultMsg={result_msg}")

        if result_code not in ("00", "0", None):
            # 에러 메시지 자체가 필수 파라미터/오퍼레이션 존재 여부에 대한 단서가 된다.
            continue

        items = extract_items(payload)
        print(f"조회된 item 수: {len(items)}")
        if items:
            print("첫 item 샘플:")
            print(json.dumps(items[0], ensure_ascii=False, indent=2))
            working_operation = operation
            all_items = items
            break

    if not working_operation:
        print(
            "\n어떤 오퍼레이션도 정상 응답(item 포함)을 주지 않았습니다. "
            "위에 출력된 resultCode/resultMsg를 참고 문서와 대조해서 "
            "오퍼레이션명이나 필수 파라미터(inqryDiv/inqryBgnDate 등)를 조정해주세요."
        )
        sys.exit(2)

    print(f"\n성공한 오퍼레이션: {working_operation}")

    # 카메라 종류 집계: 이름 계열 필드 중 실제로 존재하는 것을 찾아서 키워드로 필터링
    name_field = next(
        (f for f in NAME_FIELD_CANDIDATES if f in all_items[0]), None
    )
    if not name_field:
        print("품명/규격 관련 필드를 찾지 못했습니다. item 전체 필드 목록:")
        print(list(all_items[0].keys()))
        sys.exit(3)

    print(f"품명 필드로 '{name_field}' 사용")

    camera_items = [
        item for item in all_items if args.keyword in str(item.get(name_field, ""))
    ]
    distinct_names = sorted({str(item.get(name_field, "")) for item in camera_items})

    print(f"\n'{args.corp_name}' 전체 등록 물품 수: {len(all_items)}")
    print(f"'{args.keyword}' 포함 물품 수: {len(camera_items)}")
    print(f"'{args.keyword}' 관련 물품 종류(고유 {name_field} 개수): {len(distinct_names)}")
    for name in distinct_names:
        print(f"  - {name}")


if __name__ == "__main__":
    main()
