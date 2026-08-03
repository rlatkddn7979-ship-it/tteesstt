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


def lookup_product(product_id: str, service_key: str, num_of_rows: int = 100, log=print) -> dict:
    """물품식별번호만으로 직접 조회한다. 결과가 numOfRows보다 많으면 페이지를 넘겨가며
    전부 가져온다 (한 물품식별번호에 대해 계약업체/제조사가 다른 레코드가 여러 건 있을 수 있어서,
    1페이지만 보면 뒤쪽 페이지에 있는 레코드를 놓칠 수 있다).

    반환값: {"ok": bool, "items": [...], "total_count": int|None}
    """
    operation = core.CANDIDATE_OPERATIONS[0]
    log(f"=== prdctIdntNo={product_id} 로 날짜/업체명 필터 없이 직접 조회 ===")

    all_items = []
    page_no = 1
    total_count = None
    while True:
        params = {
            "numOfRows": str(num_of_rows),
            "pageNo": str(page_no),
            "type": "json",
            "prdctIdntNo": product_id,
        }
        result = core.fetch(operation, service_key, params, log=log)
        payload = result["json"]

        if payload is None:
            log("\nJSON이 아닌 응답 (앞부분 1000자):")
            log(result["raw"][:1000])
            return {"ok": False, "items": all_items, "total_count": total_count}

        header = payload.get("response", {}).get("header", {})
        result_code = header.get("resultCode")
        log(f"\n[page {page_no}] resultCode={result_code} resultMsg={header.get('resultMsg')}")
        if result_code not in ("00", "0", None):
            break

        body = payload.get("response", {}).get("body", {})
        if total_count is None:
            total_count = body.get("totalCount")
        items = core.extract_items(payload)
        log(f"[page {page_no}] 조회된 item 수: {len(items)} (totalCount={total_count})")
        all_items.extend(items)

        if not items:
            break
        if total_count is not None and len(all_items) >= int(total_count):
            break
        if len(items) < num_of_rows:
            break
        page_no += 1

    if not all_items:
        log(
            "\n=> 이 물품식별번호로는 아무것도 안 나옵니다. "
            "즉 getThptyUcntrctPrdctInfoList 오퍼레이션 자체가 이 물품을 모른다는 뜻입니다 "
            "(날짜/업체명 필터 문제가 아닙니다)."
        )
        return {"ok": True, "items": [], "total_count": total_count}

    log(f"\n=> 조회된 item(s) (전체 {len(all_items)}건):")
    for item in all_items:
        log(json.dumps(item, ensure_ascii=False, indent=2))

    return {"ok": True, "items": all_items, "total_count": total_count}


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
