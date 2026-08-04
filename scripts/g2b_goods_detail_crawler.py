"""
나라장터 목록정보시스템(goods.g2b.go.kr) 상품 상세페이지에서
지정한 항목을 추출하여 CSV/텍스트/JSON으로 저장하는 스크립트

대상 예시:
    https://goods.g2b.go.kr:8053/search/productSearchView.do?goodsClsfcNo=4617162201&goodsIdntfcNo=26181419

추출 항목:
    물품식별번호, 모델명, 옵션/기타, 제조업체명, 촬상소자, 촬영소자화소수, 형태
    (FIELD_LABEL_CANDIDATES에서 이 목록을 바꾸면 추출 항목도 바뀐다)

사용법 (CLI):
    1. pip install requests beautifulsoup4
    2. python g2b_goods_detail_crawler.py
    3. 실행이 끝나면 생성된 g2b_output_*.json 파일을 Claude에게 다시 업로드
       -> Claude가 그 JSON을 읽어서 docx 문서에 표를 자동으로 추가해 드립니다.

GUI로 실행하려면 g2b_goods_detail_crawler_gui.py 를 실행하세요. 물품식별번호
목록을 화면에서 붙여넣기로 바로 바꿔서 조회할 수 있습니다.

주의:
    - 정부 사이트이므로 과도한 요청은 피하고, 요청 간 딜레이(REQUEST_DELAY)를 유지하세요.
    - 포트번호(:8053)가 URL에 반드시 포함되어야 합니다.
"""

import csv
import datetime
import json
import os
import re
import time
import unicodedata
import uuid

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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve_output_path(path: str) -> str:
    return path if os.path.isabs(path) else os.path.join(SCRIPT_DIR, path)


def _timestamped_path(path):
    # 파일명에 이미 실행 시각이 들어있을 수 있으므로(자동 생성 파일명), 짧은 임의 문자열을
    # 붙여 항상 새로운 이름이 되게 한다.
    root, ext = os.path.splitext(path)
    suffix = uuid.uuid4().hex[:6]
    return f"{root}_retry{suffix}{ext}"


def _write_with_fallback(write_fn, path, log):
    """write_fn(경로)로 저장을 시도한다. 잠겨 있어 실패하면 타임스탬프를 붙인
    새 파일명으로 한 번 더 시도한다 (원인을 진단하는 대신 확실히 진행되게 한다)."""
    try:
        write_fn(path)
        return path
    except PermissionError:
        fallback = _timestamped_path(path)
        log(
            f"   '{path}' 파일을 저장할 수 없어(다른 프로그램에서 사용 중이거나 잠겨 있을 수 있음) "
            f"'{fallback}'(으)로 대신 저장합니다."
        )
        try:
            write_fn(fallback)
        except PermissionError as e:
            raise PermissionError(
                f"'{fallback}'로도 저장할 수 없습니다. 해당 폴더에 쓰기 권한이 없거나 "
                "(예: OneDrive 동기화 중, 보안 프로그램 차단 등) 폴더 자체의 문제일 수 있습니다. "
                "다른 폴더로 저장 경로를 바꿔서 시도해보세요."
            ) from e
        return fallback


def _display_width(text: str) -> int:
    """엑셀 열 너비 계산용 텍스트 폭. 한글/한자 등 동아시아 넓은 문자는 라틴 문자의
    약 2배 폭으로 표시되는데, 그냥 len()으로 계산하면 절반 정도로 좁게 잡혀서
    한글이 많이 들어간 열이 실제로는 잘려 보이는 문제가 있었다."""
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in str(text))


# 조회할 (goodsClsfcNo, goodsIdntfcNo) 목록 (CLI로 직접 실행할 때 쓰는 기본값)
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

# 추출할 항목명 (출력 컬럼 순서) : 페이지에서 매칭할 라벨 후보들
FIELD_LABEL_CANDIDATES = {
    "물품식별번호": ["물품식별번호"],
    "모델명": ["모델명"],
    "옵션/기타": ["옵션/기타", "옵션기타"],
    "제조업체명": ["제조업체명"],
    "촬상소자": ["촬상소자"],
    "촬영소자화소수": ["촬영소자화소수"],
    "형태": ["형태"],
}

# "옵션/기타" 안에 "배율:광학40배줌/디지털12배줌, 초점거리:6.6~264mm" 처럼 쉼표로
# 구분된 항목들이 같이 적혀 있어서, "배율:" 다음부터 다음 쉼표 전까지만 뽑아서
# 별도 컬럼으로 만든다 (줌 없는 고정렌즈 카메라는 "배율" 항목 자체가 없어서 빈 값).
ZOOM_RATIO_PATTERN = re.compile(r"배율\s*[:：]\s*([^,，]+)")


def extract_zoom_ratio(option_text: str) -> str:
    if not option_text:
        return ""
    m = ZOOM_RATIO_PATTERN.search(option_text)
    return m.group(1).strip() if m else ""


FIELDNAMES = ["물품식별번호", "모델명", "옵션/기타", "배율", "제조업체명", "촬상소자", "촬영소자화소수", "형태"]


def parse_goods_idntfc_no_list(text: str) -> list:
    """줄바꿈/쉼표/공백으로 구분된 텍스트에서 물품식별번호 목록을 뽑아낸다.
    (GUI 텍스트 박스에 그대로 붙여넣은 내용을 파싱하는 용도)"""
    raw = text.replace(",", "\n").splitlines()
    result = []
    seen = set()
    for line in raw:
        no = line.strip()
        if no and no not in seen:
            seen.add(no)
            result.append(no)
    return result


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
    """페이지에서 FIELD_LABEL_CANDIDATES에 지정된 항목만 뽑아온다."""
    pairs = extract_label_value_pairs(html)

    row = {}
    for field, candidates in FIELD_LABEL_CANDIDATES.items():
        value = ""
        for cand in candidates:
            if cand in pairs and pairs[cand]:
                value = pairs[cand]
                break
        row[field] = value

    row["배율"] = extract_zoom_ratio(row.get("옵션/기타", ""))

    return row


def run_crawl(goods_clsfc_no: str, goods_idntfc_no_list: list, request_delay: float = REQUEST_DELAY, log=print) -> dict:
    """goods_idntfc_no_list를 순회하며 상세페이지를 하나씩 조회한다.

    반환값: {"rows": [...], "fail_list": [...]}
    """
    targets = [
        {"goodsClsfcNo": goods_clsfc_no, "goodsIdntfcNo": no}
        for no in goods_idntfc_no_list
    ]

    rows = []
    fail_list = []
    total = len(targets)
    for i, target in enumerate(targets, start=1):
        tag = f"[{i}/{total}] goodsIdntfcNo={target['goodsIdntfcNo']}"
        try:
            html = fetch_html(target)
            row = extract_fields(html)
            row["_요청goodsIdntfcNo"] = target["goodsIdntfcNo"]  # 매칭 확인용
            rows.append(row)
            log(f"{tag} -> 완료")
        except requests.RequestException as e:
            log(f"{tag} -> 실패: {e}")
            fail_list.append(target["goodsIdntfcNo"])

        if i < total:
            time.sleep(request_delay)

    return {"rows": rows, "fail_list": fail_list}


def default_output_basename() -> str:
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"g2b_output_{timestamp}"


def write_outputs(rows: list, fail_list: list, basename: str = None, log=print) -> dict:
    """표 형태(엑셀 또는 CSV), TXT(붙여넣기용), JSON 세 가지 형식으로 결과를 저장한다.

    표 형태는 openpyxl이 설치되어 있으면 .xlsx로 저장하고, 텍스트 길이에 맞춰
    열 너비를 자동으로 넓힌다. openpyxl이 없으면 .csv로 대신 저장한다(열 너비 개념이
    없는 순수 텍스트 형식이라, 엑셀에서 열었을 때 너비가 자동으로 맞춰지지는 않는다).
    제조업체명이 이전 행과 달라지는 지점마다 빈 줄을 한 줄 끼워서 업체별로
    구분하기 쉽게 한다.

    반환값: {"table": 경로, "txt": 경로, "json": 경로}
    """
    basename = basename or default_output_basename()
    txt_path = _resolve_output_path(f"{basename}_paste.txt")
    json_path = _resolve_output_path(f"{basename}.json")

    log(f"\n결과 저장을 시작합니다 (총 {len(rows) + len(fail_list)}건 중 {len(rows)}건 성공)...")

    try:
        import openpyxl
    except ImportError:
        openpyxl = None

    table_headers = FIELDNAMES + ["_요청goodsIdntfcNo"]
    table_rows = []
    prev_maker = None
    for i, row in enumerate(rows):
        maker = row.get("제조업체명", "")
        if i > 0 and maker != prev_maker:
            table_rows.append([""] * len(table_headers))  # 제조업체명이 바뀌는 구간에 빈 줄
        table_rows.append([row.get(field, "") for field in table_headers])
        prev_maker = maker

    def save_csv(p):
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(table_headers)
            writer.writerows(table_rows)

    def save_xlsx(p):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "결과"
        ws.append(table_headers)
        for row in table_rows:
            ws.append(row)
        for col_idx, header in enumerate(table_headers, start=1):
            width = (
                max(_display_width(header), *(_display_width(r[col_idx - 1]) for r in table_rows))
                if table_rows else _display_width(header)
            )
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(width + 2, 60)
        wb.save(p)

    def save_txt(p):
        with open(p, "w", encoding="utf-8-sig") as f:
            for row in rows:
                for field in FIELDNAMES:
                    f.write(f"{row.get(field, '')}\n")
                f.write("\n")  # 상품 간 구분용 빈 줄

    def save_json(p):
        with open(p, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)

    if openpyxl is not None:
        table_path = _resolve_output_path(f"{basename}.xlsx")
        log(f" - 표 형태(엑셀, 열 너비 자동조정) 저장 시도: {table_path}")
        table_saved = _write_with_fallback(save_xlsx, table_path, log)
        if not os.path.exists(table_saved):
            # wb.save()는 예외 없이 끝났는데 실제 파일이 없는 경우(보안 프로그램이 저장
            # 직후 오피스 문서 형식만 조용히 지우는 경우가 있었음). CSV로 대신 저장한다.
            log(
                f"\n엑셀 저장은 오류 없이 끝났는데 실제 파일이 생기지 않았습니다: {table_saved} "
                "(보안 프로그램이 저장 직후 오피스 문서 형식만 조용히 삭제하고 있을 수 있습니다). "
                "CSV로 대신 저장합니다."
            )
            table_path = _resolve_output_path(f"{basename}.csv")
            log(f" - 표 형태(CSV) 저장 시도: {table_path}")
            table_saved = _write_with_fallback(save_csv, table_path, log)
    else:
        table_path = _resolve_output_path(f"{basename}.csv")
        log(f" - 표 형태(CSV, openpyxl 미설치로 열 너비 자동조정 불가) 저장 시도: {table_path}")
        log("   'pip install openpyxl' 설치 후 다시 실행하면 열 너비가 자동으로 맞춰진 엑셀로 저장됩니다.")
        table_saved = _write_with_fallback(save_csv, table_path, log)

    log(f" - 붙여넣기용(값만) 저장 시도: {txt_path}")
    txt_saved = _write_with_fallback(save_txt, txt_path, log)
    log(f" - JSON 저장 시도: {json_path}")
    json_saved = _write_with_fallback(save_json, json_path, log)

    log(f"\n결과를 저장했습니다: {table_saved}, {txt_saved}, {json_saved}")
    if fail_list:
        log(f"실패한 goodsIdntfcNo: {', '.join(fail_list)}")

    return {"table": table_saved, "txt": txt_saved, "json": json_saved}


if __name__ == "__main__":
    result = run_crawl(FIXED_GOODS_CLSFC_NO, GOODS_IDNTFC_NO_LIST, log=print)
    write_outputs(result["rows"], result["fail_list"], log=print)
