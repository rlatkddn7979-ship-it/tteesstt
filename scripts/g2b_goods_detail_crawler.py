"""
나라장터 목록정보시스템(goods.g2b.go.kr) 상품 상세페이지에서
지정한 항목을 추출하여 CSV/텍스트/JSON으로 저장하는 스크립트

대상 예시:
    https://goods.g2b.go.kr:8053/search/productSearchView.do?goodsClsfcNo=4617162201&goodsIdntfcNo=26181419

추출 항목 (문서 표와 동일):
    물품목록번호, 물품분류번호, 물품식별번호, 품명, 세부품명, 단위,
    상품원산지국가명, 모델명, 용도, 구성, 옵션/기타

사용법:
    1. pip install requests beautifulsoup4
    2. python g2b_goods_detail_crawler.py
    3. 실행이 끝나면 생성된 g2b_output.json 파일을 Claude에게 다시 업로드
       -> Claude가 그 JSON을 읽어서 docx 문서에 표를 자동으로 추가해 드립니다.

주의:
    - 정부 사이트이므로 과도한 요청은 피하고, 요청 간 딜레이(REQUEST_DELAY)를 유지하세요.
    - 포트번호(:8053)가 URL에 반드시 포함되어야 합니다.
"""

import csv
import json
import time
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://goods.g2b.go.kr:8053/search/productSearchView.do"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
REQUEST_DELAY = 0.5  # 요청 사이 대기 시간(초)
OUTPUT_FILE = "g2b_output.csv"
OUTPUT_FILE_VERTICAL = "g2b_output_paste.txt"
OUTPUT_FILE_JSON = "g2b_output.json"

# 조회할 (goodsClsfcNo, goodsIdntfcNo) 목록
FIXED_GOODS_CLSFC_NO = "4617162201"

GOODS_IDNTFC_NO_LIST = [
    "25673585", "25674932", "25254919", "25416305", "25554490", "25512907",
    "25673588", "25254921", "25254912", "25254926", "25983397", "25551146",
    "25254918", "25512909", "25417155", "25983396", "25254923", "25673586",
    "25254920", "25254915", "25673587", "25674931", "25254917", "25686245",
    "25254922", "25254916", "25416306", "25551147", "25254924", "25983402",
    "25254928", "25512905", "25254914", "25254932", "25254927", "25983403",
    "25417156", "25254913", "25512906", "25254925", "25724701", "25254933",
    "25983401", "25512908", "25983393", "25549621", "25254929", "25674933",
    "25512910", "25983400", "25983398", "25983394", "25254931", "25254930",
    "25983395", "25674934", "25549620", "25383348",
]

TARGETS = [
    {"goodsClsfcNo": FIXED_GOODS_CLSFC_NO, "goodsIdntfcNo": no}
    for no in GOODS_IDNTFC_NO_LIST
]

# 추출할 항목명 (출력 컬럼 순서) : 페이지에서 매칭할 라벨 후보들
FIELD_LABEL_CANDIDATES = {
    "물품목록번호": ["물품목록번호"],
    "물품분류번호": ["물품분류번호"],
    "물품식별번호": ["물품식별번호"],
    "품명": ["품명"],
    "세부품명": ["세부품명", "세부품명번호"],
    "단위": ["단위"],
    "상품원산지국가명": ["상품원산지국가명"],
    "모델명": ["모델명"],
    "용도": ["용도"],
    "구성": ["구성"],
    "옵션/기타": ["옵션/기타", "옵션기타"],
}


def fetch_html(params: dict) -> str:
    resp = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=10)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    return resp.text


def extract_label_value_pairs(html: str) -> dict:
    """
    페이지 내 모든 표(table)에서 '라벨: 값' 형태의 행을 찾아 dict로 반환한다.
    - 공통속성정보 표: th(라벨) + td(값) 형태. 첫 행은 이미지 셀(rowspan)이
      td로 같이 잡히므로, th 1개인 경우 마지막 td를 값으로 사용한다.
    - 개별속성정보 표(속성명|속성값|측정단위): th 없이 td로만 구성되므로,
      표 안에 '속성명' 헤더가 있으면 각 행의 첫 td를 라벨, 두 번째 td를 값으로 사용한다.
    """
    soup = BeautifulSoup(html, "html.parser")
    pairs = {}

    for table in soup.select("table"):
        table_headers = [th.get_text(strip=True) for th in table.select("th")]
        is_attr_table = "속성명" in table_headers

        for tr in table.select("tr"):
            th_list = [th.get_text(strip=True) for th in tr.select("th")]
            td_list = [td.get_text(strip=True) for td in tr.select("td")]

            if th_list and not is_attr_table:
                if len(th_list) == 1 and td_list:
                    # 이미지 셀 등이 앞에 섞여올 수 있으므로 마지막 td를 값으로 사용
                    pairs[th_list[0]] = td_list[-1]
                elif len(th_list) == len(td_list):
                    # th, td 개수가 같은 가로 반복형 (예: 모델명 | 값 | 상품브랜드명 | 값 ...)
                    for label, value in zip(th_list, td_list):
                        pairs[label] = value
            elif is_attr_table and not th_list and len(td_list) >= 2:
                label, value = td_list[0], td_list[1]
                if label and label != "속성명":
                    pairs[label] = value

    return pairs


def extract_fields(html: str) -> dict:
    pairs = extract_label_value_pairs(html)

    row = {}
    for field, candidates in FIELD_LABEL_CANDIDATES.items():
        value = ""
        for cand in candidates:
            if cand in pairs and pairs[cand]:
                value = pairs[cand]
                break
        row[field] = value

    return row


if __name__ == "__main__":
    fieldnames = list(FIELD_LABEL_CANDIDATES.keys())
    rows = []
    fail_list = []

    total = len(TARGETS)
    for i, target in enumerate(TARGETS, start=1):
        tag = f"[{i}/{total}] goodsIdntfcNo={target['goodsIdntfcNo']}"
        try:
            html = fetch_html(target)
            row = extract_fields(html)
            row["_요청goodsIdntfcNo"] = target["goodsIdntfcNo"]  # 매칭 확인용
            rows.append(row)
            print(f"{tag} -> 완료")
        except requests.RequestException as e:
            print(f"{tag} -> 실패: {e}")
            fail_list.append(target["goodsIdntfcNo"])

        time.sleep(REQUEST_DELAY)

    # 1) CSV (표 형태, 항목=열, 한 상품=한 행) - 여러 건을 한 표로 볼 때 사용
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames + ["_요청goodsIdntfcNo"])
        writer.writeheader()
        writer.writerows(rows)

    # 2) 값만 세로로 나열 (라벨 없이, 항목 순서대로 한 줄씩)
    with open(OUTPUT_FILE_VERTICAL, "w", encoding="utf-8-sig") as f:
        for row in rows:
            for field in fieldnames:
                f.write(f"{row.get(field, '')}\n")
            f.write("\n")  # 상품 간 구분용 빈 줄

    # 3) JSON (각 상품을 dict로, 프로그램에서 다시 읽어 문서 표에 채울 때 사용)
    with open(OUTPUT_FILE_JSON, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)

    print(f"\n총 {total}건 중 {len(rows)}건을 저장했습니다.")
    print(f" - 표 형태(가로): {OUTPUT_FILE}")
    print(f" - 붙여넣기용(값만, 항목순서대로 한 줄씩): {OUTPUT_FILE_VERTICAL}")
    print(f" - JSON(문서 자동 채우기용): {OUTPUT_FILE_JSON}")
    if fail_list:
        print(f"실패한 goodsIdntfcNo: {', '.join(fail_list)}")
