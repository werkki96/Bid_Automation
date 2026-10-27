# 파일 역할: HTTPS 요청의 일시적 오류를 재시도하고 API의 전체 페이지를 수집·검증합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Common HTTPS retrieval and transient-error retries."""

import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


def fetch_bytes(base_url, params, *, timeout=40, retries=3):
    # Keys/parameters are never written to logs or exceptions.
    url = base_url + "?" + urlencode(params)
    request = Request(url, headers={"User-Agent": "BidNoticeCollector/1.0"})
    for attempt in range(retries):
        try:
            with urlopen(request, timeout=timeout) as response:
                return response.read()
        # 요청 제한·일시적 서버 장애만 재시도합니다. 인증 오류 등은 즉시 호출부에 알립니다.
        except HTTPError as exc:
            if exc.code not in (429, 500, 502, 503, 504) or attempt == retries - 1:
                raise RuntimeError(f"API 서버 HTTP 오류: {exc.code}") from None
        except (URLError, TimeoutError) as exc:
            if attempt == retries - 1:
                raise RuntimeError(f"API 통신 실패: {type(exc).__name__}") from None
        time.sleep(2 * (attempt + 1))
    raise RuntimeError("API 통신에 실패했습니다.")


def collect_pages(fetch_one, page_size, label, on_page=None):
    # 첫 페이지의 전체 건수로 필요한 페이지 수를 계산합니다. 정상적인 0건과 빈 응답 오류를 구분합니다.
    total, first = fetch_one(1)
    if total < 0:
        raise ValueError("API가 잘못된 전체 건수를 반환했습니다.")
    if total == 0:
        if first:
            raise RuntimeError(f"{label}: 전체 0건이나 실제 데이터가 있습니다.")
        return []
    if not first:
        raise RuntimeError(f"{label}: 검색 결과는 존재하지만 첫 페이지가 비어 있습니다.")
    result = list(first)
    pages = (total + page_size - 1) // page_size
    # 페이지 진행 상황을 먼저 출력한 뒤 공고명을 출력해 수집 흐름을 바로 확인합니다.
    print(f"  {label}: 검색 결과 {total}건 / {pages}페이지")
    if on_page is not None:
        on_page(first, 1, pages, total)
    for page in range(2, pages + 1):
        _, records = fetch_one(page)
        if not records:
            raise RuntimeError(f"{label}: {page}페이지 응답이 비어 있습니다.")
        result.extend(records)
        print(f"  {label}: {page}/{pages}페이지 수집")
        if on_page is not None:
            on_page(records, page, pages, total)
    # 건수 불일치는 실패로 처리합니다. 건수가 같아도 개별 공고의 중복 여부까지 보장하지는 않습니다.
    if len(result) != total:
        raise RuntimeError(f"{label}: 수집 누락 또는 중복 가능성 (API {total}건 / 수집 {len(result)}건)")
    return result
