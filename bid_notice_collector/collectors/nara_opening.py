# 파일 역할: 나라장터 용역 개찰결과를 개찰일 기준으로 조회하고 화면 표시용 컬럼으로 변환합니다.
# 마지막 업데이트: 2026-11-11
# 작업자: SMI 장지수

"""나라장터(G2B) 용역 개찰결과 조회 (조달청 나라장터 낙찰정보서비스)."""

from datetime import timedelta

from collectors.nara import parse_page
from http_client import collect_pages, fetch_bytes
from ui_helpers import format_price

API_URL = (
    "https://apis.data.go.kr/1230000/as/ScsbidInfoService/"
    "getOpengResultListInfoServcPPSSrch"
)
DETAIL_URL = "https://www.g2b.go.kr/link/PNPE027_01/single/?bidPbancNo={no}&bidPbancOrd={order}"

# 이 API는 한 번에 최대 7일까지만 조회할 수 있습니다(8일 이상은 07 입력범위값 초과).
MAX_DAYS_PER_REQUEST = 7

COLUMNS = [
    "공고번호", "공고명", "공고기관", "수요기관", "개찰일시", "진행상태", "비고",
    "참가업체수", "1순위 업체", "투찰금액(원)", "투찰률(%)", "상세보기 주소",
]
# 화면 검색어가 비교하는 컬럼입니다.
SEARCH_COLUMNS = ("공고명", "공고기관", "수요기관", "1순위 업체")


def date_chunks(start_date, end_date, days=MAX_DAYS_PER_REQUEST):
    """Split an inclusive date range into ranges of at most `days` days."""
    current = start_date
    while current <= end_date:
        last = min(current + timedelta(days=days - 1), end_date)
        yield current, last
        current = last + timedelta(days=1)


def first_bidder(info):
    # opengCorpInfo: 업체명^사업자번호^대표자^투찰금액^투찰률. 협상계약은 금액·투찰률이 없거나 0입니다.
    parts = [part.strip() for part in str(info or "").split("^")]
    parts += [""] * (5 - len(parts))
    name, _, _, amount, rate = parts[:5]
    try:
        rate = "" if not rate or float(rate) == 0 else rate
    except ValueError:
        pass
    try:
        amount = "" if not amount or float(amount) == 0 else format_price(amount)
    except ValueError:
        pass
    return name, amount, rate


def convert(item):
    no = str(item.get("bidNtceNo") or "").strip()
    order = str(item.get("bidNtceOrd") or "").strip()
    name, amount, rate = first_bidder(item.get("opengCorpInfo"))
    return {
        "공고번호": f"{no}-{order}" if no and order else no,
        "공고명": item.get("bidNtceNm") or "",
        "공고기관": item.get("ntceInsttNm") or "",
        "수요기관": item.get("dminsttNm") or "",
        "개찰일시": str(item.get("opengDt") or "")[:16],
        "진행상태": item.get("progrsDivCdNm") or "",
        # 유찰 사유(단일응찰·무응찰 등)가 들어옵니다. 개찰완료와 같은 반복 문구는 비워 둡니다.
        "비고": "" if item.get("opengRsltNtcCntnts") == item.get("progrsDivCdNm")
                else item.get("opengRsltNtcCntnts") or "",
        "참가업체수": str(item.get("prtcptCnum") or "0"),
        "1순위 업체": name,
        "투찰금액(원)": amount,
        "투찰률(%)": rate,
        "상세보기 주소": DETAIL_URL.format(no=no, order=order) if no and order else "",
    }


def matches(row, query):
    """Case-insensitive OR search; separate several words with commas."""
    terms = [term.strip().casefold() for term in str(query or "").split(",") if term.strip()]
    if not terms:
        return True
    text = " ".join(str(row.get(column) or "") for column in SEARCH_COLUMNS).casefold()
    return any(term in text for term in terms)


def collect(api_key, start_date, end_date, industry_code=None, page_size=100, on_page=None):
    """Return opening results whose opening date is within the inclusive range.

    on_page(rows, done_chunks, total_chunks) receives converted rows as each page arrives.
    """
    if start_date > end_date:
        raise ValueError("시작일은 종료일보다 늦을 수 없습니다.")
    chunks = list(date_chunks(start_date, end_date))
    results = []
    for index, (first, last) in enumerate(chunks, 1):
        def fetch(page, first=first, last=last):
            params = {
                "ServiceKey": api_key,
                "type": "json",
                "numOfRows": page_size,
                "pageNo": page,
                "inqryDiv": "2",  # 개찰일시 기준
                "inqryBgnDt": first.strftime("%Y%m%d") + "0000",
                "inqryEndDt": last.strftime("%Y%m%d") + "2359",
            }
            if industry_code:
                params["indstrytyCd"] = industry_code
            return parse_page(fetch_bytes(API_URL, params), "나라장터 개찰결과")

        def notify(items, page, pages, total, index=index):
            if on_page is not None:
                on_page([convert(item) for item in items], index, len(chunks))

        results.extend(convert(item) for item in collect_pages(fetch, page_size, "나라장터 개찰결과", notify))
    # 최근 개찰이 위로 오도록 정렬합니다.
    results.sort(key=lambda row: row["개찰일시"], reverse=True)
    return results
