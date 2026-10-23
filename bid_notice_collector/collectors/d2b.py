# 파일 역할: D2B 국내 경쟁입찰 공고를 날짜 범위로 조회하고 XML 응답을 공통 컬럼으로 변환합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""방위사업청 D2B 국내 경쟁입찰공고 수집."""

import xml.etree.ElementTree as ET
from http_client import fetch_bytes, collect_pages

D2B_LIST_URL = "https://www.d2b.go.kr/psb/bid/serviceBidAnnounceList.do?key=137"

API_URL = (
    "https://apis.data.go.kr/1690000/BidPblancInfoService/"
    "getDmstcCmpetBidPblancList"
)


def local_name(tag):
    # XML 네임스페이스가 붙은 {주소}태그에서도 실제 태그명만 비교합니다.
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


def collect(api_key, start_date, end_date, settings, on_page=None):
    page_size = settings["page_size"]
    start_token = start_date.strftime("%Y%m%d")
    end_token = end_date.strftime("%Y%m%d")

    # 이 API에는 나라장터의 업종 1169·입찰마감제외 조건을 전달하지 않습니다.
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

    # UI에는 원본 API 필드 대신 화면·출력에서 사용하는 공통 컬럼으로 전달합니다.
    def notify_page(items, page, pages, total):
        if on_page is not None:
            on_page([convert(item, settings["d2b"]["display_name"]) for item in items], page, pages, total)

    raw_records = collect_pages(fetch, page_size, "D2B", on_page=notify_page)
    return [convert(row, settings["d2b"]["display_name"]) for row in raw_records]
