# 파일 역할: 나라장터 용역 공고를 날짜·업종·마감 조건으로 조회하고 JSON 응답을 공통 컬럼으로 변환합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""나라장터(G2B) 학술·연구용역 입찰공고 수집."""

import json
from config import COLUMNS
from http_client import fetch_bytes, collect_pages

API_URL = (
    "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
    "getBidPblancListInfoServcPPSSrch"
)


def convert(item, display_name):
    no = str(item.get("bidNtceNo") or "").strip()
    order = str(item.get("bidNtceOrd") or "").strip()
    return {
        "공고위치": display_name,
        "공고번호": f"{no}-{order}" if no and order else no,
        "공고구분": item.get("ntceKindNm") or "",
        "공고일자": item.get("bidNtceDt") or "",
        "공고명": item.get("bidNtceNm") or "",
        "공고기관(발주기관)": item.get("ntceInsttNm") or "",
        "수요기관": item.get("dminsttNm") or "",
        "추정가격(원)": item.get("presmptPrce") or "",
        "입찰참가등록 마감일시": item.get("bidQlfctRgstDt") or "",
        "상세보기 주소": item.get("bidNtceDtlUrl") or "",
    }


def collect(api_key, start_date, end_date, settings, on_page=None):
    page_size = settings["page_size"]
    source_settings = settings["nara"]
    start_token = start_date.strftime("%Y%m%d")
    end_token = end_date.strftime("%Y%m%d")

    def fetch(page):
        params = {
            "ServiceKey": api_key,
            "type": "json",
            "numOfRows": page_size,
            "pageNo": page,
            "inqryDiv": "1",  # 공고게시일시
            # 시작일 00:00부터 종료일 23:59까지 조회합니다. 업종·마감 조건은 서버에서 적용합니다.
            "inqryBgnDt": start_token + "0000",
            "inqryEndDt": end_token + "2359",
            "indstrytyCd": source_settings.get("industry_code", "1169"),
            "bidClseExcpYn": source_settings.get("exclude_closed", "Y"),
        }
        raw = fetch_bytes(API_URL, params)
        try:
            data = json.loads(raw.decode("utf-8-sig"))
            response = data["response"]
            header = response["header"]
            if str(header.get("resultCode", "")) not in ("00", "0"):
                raise RuntimeError(f"나라장터 API 오류: {header.get('resultCode')} / {header.get('resultMsg')}")
            body = response["body"]
            count = int(body["totalCount"])
            # 응답의 items/item이 목록 또는 단일 객체인 경우를 모두 목록으로 통일합니다.
            container = body.get("items") or {}
            if isinstance(container, dict):
                records = container.get("item") or []
            elif isinstance(container, list):
                records = container
            else:
                records = []
            if isinstance(records, dict):
                records = [records]
            if not isinstance(records, list):
                raise ValueError("item의 자료형이 예상과 다릅니다.")
            return count, records
        except (json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
            raise RuntimeError("나라장터 API 응답 형식 확인 필요 (JSON 구조/인증키 확인)") from exc

    # UI에는 원본 API 필드 대신 화면·출력에서 사용하는 공통 컬럼으로 전달합니다.
    def notify_page(items, page, pages, total):
        if on_page is not None:
            on_page([convert(item, source_settings["display_name"]) for item in items], page, pages, total)

    raw_records = collect_pages(fetch, page_size, "나라장터", on_page=notify_page)
    return [convert(row, source_settings["display_name"]) for row in raw_records]
