
import csv
import math
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

# ==========================================
# 1. 사용자 설정
# ==========================================

api_key = "여기에_API_KEY_입력"

api_url = (
    "https://apis.data.go.kr/1690000/"
    "BidPblancInfoService/"
    "getDmstcCmpetBidPblancList"
)

start_date = "20261015"
end_date = "20261016"

page_size = 100

output_file = (
    f"D2B_입찰공고_{start_date}_{end_date}.csv"
)

# ==========================================
# 2. 출력 컬럼
# ==========================================

columns = [
    "순번",
    "입찰구분",
    "공고구분",
    "공고일자",
    "G2B공고번호-차수",
    "통합참조번호",
    "판단번호",
    "입찰건명",
    "발주기관",
    "사전심사 마감일시",
    "입찰참가등록 마감일시",
    "입찰서/견적서제출 마감일시",
    "계약방법",
    "입찰형태",
    "기초예가",
]


# ==========================================
# 3. XML 처리
# ==========================================

def local_name(tag):
    return tag.split("}")[-1]


def find_text(root, name):
    for element in root.iter():
        if local_name(element.tag) == name:
            return (element.text or "").strip()
    return None


def parse_items(root):
    records = []

    for element in root.iter():
        if local_name(element.tag) != "item":
            continue

        record = {
            local_name(child.tag): (
                child.text.strip()
                if child.text else ""
            )
            for child in element
        }

        records.append(record)

    return records


# ==========================================
# 4. API 호출
# ==========================================

def fetch_page(page_no):
    params = {
        "serviceKey": api_key,
        "pageNo": str(page_no),
        "numOfRows": str(page_size),
        "anmtDateBegin": start_date,
        "anmtDateEnd": end_date,
    }

    url = (
        api_url + "?"
        + urllib.parse.urlencode(params)
    )

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0"
        }
    )

    with urllib.request.urlopen(
        request, timeout=30
    ) as response:
        raw_data = response.read()

    root = ET.fromstring(raw_data)

    result_code = find_text(root, "resultCode")
    result_msg = find_text(root, "resultMsg")

    if result_code is not None and result_code not in ("00", "0"):
        raise RuntimeError(
            f"API 오류: {result_code} / {result_msg}"
        )

    total_text = find_text(root, "totalCount")

    if total_text is None:
        raise RuntimeError(
            "전체 결과 수를 확인할 수 없습니다.\n"
            + raw_data.decode(
                "utf-8", errors="replace"
            )[:1000]
        )

    total_count = int(total_text)
    records = parse_items(root)

    return total_count, records


# ==========================================
# 5. 컬럼 변환
# ==========================================

def convert_record(record, number):

    g2b_no = record.get(
        "g2bPblancNo", ""
    )

    g2b_odr = record.get(
        "g2bPblancOdr", ""
    )

    if g2b_no and g2b_odr:
        g2b_number = f"{g2b_no}-{g2b_odr}"
    else:
        g2b_number = g2b_no

    return {
        "순번": number,

        "입찰구분": record.get(
            "busiDivs", ""
        ),

        "공고구분": record.get(
            "pblancSe", ""
        ),

        "공고일자": record.get(
            "pblancDate", ""
        ),

        "G2B공고번호-차수": g2b_number,

        # API에 해당 필드 없음
        "통합참조번호": "",

        "판단번호": record.get(
            "dcsNo", ""
        ),

        "입찰건명": record.get(
            "bidNm", ""
        ),

        "발주기관": record.get(
            "ornt", ""
        ),

        # API에 해당 필드 없음
        "사전심사 마감일시": "",

        "입찰참가등록 마감일시": record.get(
            "bidPartcptRegistClosDt", ""
        ),

        "입찰서/견적서제출 마감일시": record.get(
            "biddocPresentnClosDt", ""
        ),

        "계약방법": record.get(
            "cntrctMth", ""
        ),

        "입찰형태": record.get(
            "bidStle", ""
        ),

        "기초예가": record.get(
            "bsicExpt", ""
        ),
    }


# ==========================================
# 6. 전체 페이지 수집
# ==========================================

def main():

    total_count, first_records = fetch_page(1)

    all_records = list(first_records)

    total_pages = math.ceil(
        total_count / page_size
    )

    print("===== 입찰공고 조회 =====")
    print(f"조회 시작일: {start_date}")
    print(f"조회 종료일: {end_date}")
    print(f"전체 공고 수: {total_count}건")

    for page in range(2, total_pages + 1):

        _, records = fetch_page(page)

        all_records.extend(records)

        print(
            f"페이지 수집: {page}/{total_pages}"
        )

    if not all_records:
        print("조회된 공고가 없습니다.")
        return

    converted_records = [
        convert_record(record, index)
        for index, record in enumerate(
            all_records, start=1
        )
    ]

    # ======================================
    # 7. CSV 저장
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
    print(f"수집 건수: {len(converted_records)}건")
    print(f"저장 파일: {output_file}")

    if len(all_records) != total_count:
        print(
            "주의: API 전체 건수와 수집 건수가 "
            "일치하지 않습니다."
        )

    print("\n===== 공고 미리보기 =====")

    for record in converted_records[:5]:
        print(
            record["순번"],
            record["공고일자"],
            record["입찰건명"],
            record["발주기관"]
        )


if __name__ == "__main__":
    main()
