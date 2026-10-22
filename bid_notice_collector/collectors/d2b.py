"""방위사업청 D2B 국내 경쟁입찰공고 수집."""

import xml.etree.ElementTree as ET
from http_client import fetch_bytes, collect_pages

D2B_LIST_URL = "https://www.d2b.go.kr/psb/bid/serviceBidAnnounceList.do?key=137"

API_URL = (
    "https://apis.data.go.kr/1690000/BidPblancInfoService/"
    "getDmstcCmpetBidPblancList"
)


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def find_text(root, name):
    for elem in root.iter():
        if local_name(elem.tag) == name:
            return (elem.text or "").strip()
    return None


def parse_items(root):
    result = []
    for element in root.iter():
        if local_name(element.tag) != "item":
            continue
        result.append({local_name(child.tag): (child.text or "").strip() for child in element})
    return result


def convert(item, display_name):
    # D2B의 G2B번호가 없는 경우, 자체 공고번호를 사용합니다.
    g2b_no = item.get("g2bPblancNo") or ""
    g2b_ord = item.get("g2bPblancOdr") or ""
    number = g2b_no or (item.get("pblancNo") or "")
    order = g2b_ord if g2b_no else (item.get("pblancOdr") or "")
    return {
        "공고위치": display_name,
        "공고번호": f"{number}-{order}" if number and order else number,
        "공고구분": item.get("pblancSe") or "",
        "공고일자": item.get("pblancDate") or "",
        "공고명": item.get("bidNm") or "",
        "공고기관(발주기관)": item.get("ornt") or "",
        "수요기관": "",  # 이 목록 API 응답에 없음
        "추정가격(원)": "",  # bsicExpt는 '기초예비가격'이며 추정가격 아님
        "입찰참가등록 마감일시": item.get("bidPartcptRegistClosDt") or "",
        "상세보기 주소": D2B_LIST_URL,  # 개별 공고 링크가 아닌 검색 목록 페이지
    }


def collect(api_key, start_date, end_date, settings):
    page_size = settings["page_size"]
    start_token = start_date.strftime("%Y%m%d")
    end_token = end_date.strftime("%Y%m%d")

    def fetch(page):
        params = {
            "serviceKey": api_key,
            "pageNo": page,
            "numOfRows": page_size,
            "anmtDateBegin": start_token,
            "anmtDateEnd": end_token,
        }
        raw = fetch_bytes(API_URL, params)
        try:
            root = ET.fromstring(raw)
            code = find_text(root, "resultCode")
            if code is not None and code not in ("00", "0"):
                raise RuntimeError(f"D2B API 오류: {code} / {find_text(root, 'resultMsg')}")
            total_text = find_text(root, "totalCount")
            if total_text is None:
                raise RuntimeError("D2B API 응답에 totalCount가 없습니다. 인증/응답 구조를 확인하세요.")
            return int(total_text), parse_items(root)
        except ET.ParseError as exc:
            raise RuntimeError("D2B API가 올바른 XML을 반환하지 않았습니다.") from exc

    raw_records = collect_pages(fetch, page_size, "D2B")
    return [convert(row, settings["d2b"]["display_name"]) for row in raw_records]
