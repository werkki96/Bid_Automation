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
        except HTTPError as exc:
            if exc.code not in (429, 500, 502, 503, 504) or attempt == retries - 1:
                raise RuntimeError(f"API 서버 HTTP 오류: {exc.code}") from None
        except (URLError, TimeoutError) as exc:
            if attempt == retries - 1:
                raise RuntimeError(f"API 통신 실패: {type(exc).__name__}") from None
        time.sleep(2 * (attempt + 1))
    raise RuntimeError("API 통신에 실패했습니다.")


def collect_pages(fetch_one, page_size, label, on_page=None):
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
    if on_page is not None:
        on_page(first, 1, 1 + (total - 1) // page_size, total)
    pages = (total + page_size - 1) // page_size
    print(f"  {label}: 검색 결과 {total}건 / {pages}페이지")
    for page in range(2, pages + 1):
        _, records = fetch_one(page)
        if not records:
            raise RuntimeError(f"{label}: {page}페이지 응답이 비어 있습니다.")
        result.extend(records)
        if on_page is not None:
            on_page(records, page, pages, total)
        print(f"  {label}: {page}/{pages}페이지 수집")
    if len(result) != total:
        raise RuntimeError(f"{label}: 수집 누락 또는 중복 가능성 (API {total}건 / 수집 {len(result)}건)")
    return result
