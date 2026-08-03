#!/usr/bin/env python3
"""
진단용 스크립트: 물품식별번호(prdctIdntNo) 하나로 직접 조회해서,
이 오퍼레이션(getThptyUcntrctPrdctInfoList)이 그 물품을 알고 있는지 확인한다.

업체명/날짜 필터를 전혀 걸지 않고 prdctIdntNo만으로 조회하기 때문에,
"날짜 범위를 넓혀도 특정 품목이 안 나오는" 문제가
1) 날짜 필터 때문인지
2) 애초에 이 오퍼레이션 자체가 그 물품을 반환하지 않는 것인지
를 구분하는 데 쓴다.

사용법 (CLI):
    export DATA_GO_KR_SERVICE_KEY="발급받은 서비스키"
    python3 scripts/check_product_by_id.py 24101309

GUI로 실행하려면 scripts/check_product_by_id_gui.py 를 실행하세요.
"""

import argparse
import json
import sys

import check_doowon_cameras as core


def lookup_product(product_id: str, service_key: str, log=print) -> dict:
    """물품식별번호만으로 직접 조회한다.

    반환값: {"ok": bool, "items": [...], "total_count": int|None}
    """
    operation = core.CANDIDATE_OPERATIONS[0]
    params = {
        "numOfRows": "10",
        "pageNo": "1",
        "type": "json",
        "prdctIdntNo": product_id,
    }

    log(f"=== prdctIdntNo={product_id} 로 날짜/업체명 필터 없이 직접 조회 ===")
    result = core.fetch(operation, service_key, params, log=log)
    payload = result["json"]

    if payload is None:
        log("\nJSON이 아닌 응답 (앞부분 1000자):")
        log(result["raw"][:1000])
        return {"ok": False, "items": [], "total_count": None}

    header = payload.get("response", {}).get("header", {})
    log(f"\nresultCode={header.get('resultCode')} resultMsg={header.get('resultMsg')}")

    body = payload.get("response", {}).get("body", {})
    total_count = body.get("totalCount")
    items = core.extract_items(payload)
    log(f"totalCount={total_count}, 반환된 item 수={len(items)}")

    if not items:
        log(
            "\n=> 이 물품식별번호로는 아무것도 안 나옵니다. "
            "즉 getThptyUcntrctPrdctInfoList 오퍼레이션 자체가 이 물품을 모른다는 뜻입니다 "
            "(날짜/업체명 필터 문제가 아닙니다)."
        )
        return {"ok": True, "items": [], "total_count": total_count}

    log("\n=> 조회된 item(s):")
    for item in items:
        log(json.dumps(item, ensure_ascii=False, indent=2))

    return {"ok": True, "items": items, "total_count": total_count}


def main():
    parser = argparse.ArgumentParser(description="물품식별번호로 직접 조회 (진단용)")
    parser.add_argument("product_id", help="물품식별번호 (예: 24101309)")
    parser.add_argument(
        "--service-key",
        default=None,
        help="서비스키 (생략하면 DATA_GO_KR_SERVICE_KEY 환경변수를 사용)",
    )
    args = parser.parse_args()

    service_key = args.service_key or core.os.environ.get("DATA_GO_KR_SERVICE_KEY", "")
    if not service_key:
        print("서비스키가 필요합니다. --service-key 또는 DATA_GO_KR_SERVICE_KEY 환경변수를 설정해주세요.")
        sys.exit(1)

    lookup_product(args.product_id, service_key, log=print)


if __name__ == "__main__":
    main()
