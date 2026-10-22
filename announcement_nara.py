
import csv
import json
import math
import time
import urllib.parse
import urllib.request
import urllib.error

# ==========================================
# 1. 사용자 설정
# ==========================================

api_key = "f517321cb1a13fda947ad1dc491ae6851eb1cd051ebede0edae5d548ac228ed5"

api_url = (
    "https://apis.data.go.kr/1230000/ad/"
    "BidPublicInfoService/"
    "getBidPblancListInfoServcPPSSrch"
)

start_date = "20261015"
end_date = "20261016"

industry_code = "1169"
exclude_closed = "Y"

page_size = 100

output_file = (
    f"나라장터_학술연구용역_{start_date}_{end_date}.csv"
)

# ==========================================
# 2. 출력 컬럼과 API 필드 매핑
# ==========================================

field_map = {
    "순번": None,
    "공고구분": "ntceKindNm",
    "공고게시일시": "bidNtceDt",
    "입찰공고번호-차수": None,
    "참조번호": "refNo",
    "입찰공고명": "bidNtceNm",
    "공고기관": "ntceInsttNm",
    "수요기관": "dminsttNm",
    "입찰참가등록 마감일시": "bidQlfctRgstDt",
    "입찰개시일시": "bidBeginDt",
    "입찰마감일시": "bidClseDt",
    "개찰일시": "opengDt",
    "계약방법": "cntrctCnclsMthdNm",
    "입찰방식": "bidMethdNm",
    "추정가격": "presmptPrce",
    "업종코드": None,
    "업종명": None,
    "공고상세URL": "bidNtceDtlUrl",
}

columns = list(field_map.keys())

# ==========================================
# 3. API 호출
# ==========================================

def fetch_page(page_no):

    params = {
        "ServiceKey": api_key,
        "type": "json",
        "numOfRows": page_size,
        "pageNo": page_no,
        "inqryDiv": "1",
        "inqryBgnDt": start_date + "0000",
        "inqryEndDt": end_date + "2359",
        "indstrytyCd": industry_code,
        "bidClseExcpYn": exclude_closed,
    }

    url = api_url + "?" + urllib.parse.urlencode(params)

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/json",
        }
    )

    for attempt in range(3):
        try:
            with urllib.request.urlopen(
                request, timeout=40
            ) as response:
                raw = response.read()

            data = json.loads(raw.decode("utf-8-sig"))
            break

        except (
            urllib.error.URLError,
            TimeoutError
        ):
            if attempt == 2:
                raise
            time.sleep(2 * (attempt + 1))

        except json.JSONDecodeError:
            raise RuntimeError(
                "JSON 응답이 아닙니다. "
                "API 인증키 또는 응답 형식을 확인하세요."
            )

    header = data.get("response", {}).get("header", {})
    body = data.get("response", {}).get("body", {})

    result_code = str(header.get("resultCode", ""))
    result_msg = header.get("resultMsg", "")

    if result_code not in ("00", "0"):
        raise RuntimeError(
            f"API 오류: {result_code} / {result_msg}"
        )

    total_count = int(body.get("totalCount", 0))

    items = body.get("items") or {}
    if isinstance(items, dict):
        records = items.get("item", [])
    elif isinstance(items, list):
        records = items
    else:
        records = []

    if isinstance(records, dict):
        records = [records]

    return total_count, records


# ==========================================
# 4. 컬럼 변환
# ==========================================

def convert_record(record, number):

    converted = {}

    for column, api_field in field_map.items():
        if api_field is None:
            converted[column] = ""
        else:
            converted[column] = record.get(
                api_field, ""
            ) or ""

    converted["순번"] = number

    bid_no = record.get("bidNtceNo", "") or ""
    bid_ord = record.get("bidNtceOrd", "") or ""

    converted["입찰공고번호-차수"] = (
        f"{bid_no}-{bid_ord}"
        if bid_no and bid_ord
        else bid_no
    )

    # 업종코드는 검색조건을 기록
    # 개별 공고의 전체 허용 업종 목록을 의미하지 않음
    converted["업종코드"] = industry_code
    converted["업종명"] = "학술·연구용역"

    return converted


# ==========================================
# 5. 전체 페이지 수집
# ==========================================

def main():

    if not api_key or api_key == "여기에_API_KEY_입력":
        raise ValueError("api_key에 인증키를 입력하세요.")

    total_count, first_records = fetch_page(1)

    all_records = list(first_records)
    total_pages = math.ceil(total_count / page_size)

    print("===== 나라장터 입찰공고 조회 =====")
    print(f"공고게시일: {start_date} ~ {end_date}")
    print(f"업종코드: {industry_code} (학술·연구용역)")
    print(f"입찰마감제외: {exclude_closed}")
    print(f"API 검색 결과: {total_count}건")

    for page in range(2, total_pages + 1):

        _, records = fetch_page(page)
        all_records.extend(records)

        print(
            f"수집 진행: {page}/{total_pages}페이지"
        )

        time.sleep(0.2)

    if len(all_records) != total_count:
        raise RuntimeError(
            f"수집 건수 불일치: "
            f"전체 {total_count}건, "
            f"수집 {len(all_records)}건"
        )

    converted_records = [
        convert_record(record, index)
        for index, record in enumerate(
            all_records, start=1
        )
    ]

    # ======================================
    # 6. CSV 저장
    # ======================================

    with open(
        output_file,
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=columns
        )

        writer.writeheader()
        writer.writerows(converted_records)

    print("\n===== 수집 완료 =====")
    print(f"최종 수집 건수: {len(converted_records)}건")
    print(f"저장 파일: {output_file}")

    print("\n===== 공고 미리보기 =====")

    for record in converted_records[:10]:
        print(
            record["순번"],
            record["공고게시일시"],
            record["입찰공고명"],
            record["입찰마감일시"]
        )


if __name__ == "__main__":
    main()
