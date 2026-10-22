"""Configuration and date handling. Standard library only."""

from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
from urllib.parse import unquote

BASE_DIR = Path(__file__).resolve().parent
KST = timezone(timedelta(hours=9))

COLUMNS = [
    "공고위치", "공고번호", "공고구분", "공고일자", "공고명",
    "공고기관(발주기관)", "수요기관", "추정가격(원)",
    "입찰참가등록 마감일시", "상세보기 주소",
]


def read_env(path=None):
    """Load simple .env KEY=VALUE lines without an external dependency."""
    target = Path(path) if path is not None else BASE_DIR / ".env"
    if not target.is_file():
        return {}
    values = {}
    for lineno, line in enumerate(target.read_text(encoding="utf-8-sig").splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):].strip()
        if "=" not in line:
            raise ValueError(f".env 파일 {lineno}행: KEY=VALUE 형태가 아닙니다.")
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key] = value
    return values


def api_keys(env=None):
    """Environment variables take precedence over .env. Decode URL-encoded keys once."""
    env = read_env() if env is None else env

    def choose(name):
        value = os.environ.get(name) or env.get(name) or os.environ.get("API_KEY") or env.get("API_KEY", "")
        return unquote(value.strip())

    return {"nara": choose("NARA_API_KEY"), "d2b": choose("D2B_API_KEY")}


def load_settings(path=None):
    target = Path(path) if path else BASE_DIR / "settings.json"
    data = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("settings.json의 최상위 구조는 객체여야 합니다.")
    includes = data.get("include_keywords", [])
    if not isinstance(includes, list) or any(not isinstance(x, str) for x in includes):
        raise ValueError("include_keywords에는 문자열 목록을 입력해야 합니다.")
    keywords = data.get("exclude_keywords", [])
    search_fields = data.get("keyword_search_columns", ["공고명"])
    if not isinstance(keywords, list) or any(not isinstance(x, str) for x in keywords):
        raise ValueError("exclude_keywords에는 문자열 목록을 입력해야 합니다.")
    if not isinstance(search_fields, list) or any(x not in COLUMNS for x in search_fields):
        raise ValueError("keyword_search_columns에 올바른 출력 컬럼명을 입력하세요.")
    date_settings = data.get("date_range", {})
    if not isinstance(date_settings, dict):
        raise ValueError("date_range는 start_date/end_date를 가진 객체여야 합니다.")
    for field in ("start_date", "end_date"):
        value = date_settings.get(field)
        if value is not None and not isinstance(value, str):
            raise ValueError(f"date_range.{field}에는 YYYY-MM-DD 문자열 또는 null을 입력하세요.")
    # Both explicit dates must be set; otherwise the default is yesterday ~ today.
    resolve_date_range(data)
    page_size = data.get("page_size", 100)
    if not isinstance(page_size, int) or isinstance(page_size, bool) or not 1 <= page_size <= 999:
        raise ValueError("page_size는 1~999의 정수여야 합니다.")
    for source in ("nara", "d2b"):
        if source not in data or not isinstance(data[source], dict):
            raise ValueError(f"settings.json에 {source} 설정이 필요합니다.")
        name = data[source].get("display_name")
        if not isinstance(name, str) or not name.strip():
            raise ValueError(f"{source}.display_name을 설정하세요.")
    return data


def _parse_date(value, label):
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label}은 YYYY-MM-DD 형식의 실제 날짜여야 합니다: {value!r}") from exc


def resolve_date_range(settings, explicit_start=None, explicit_end=None, now=None):
    """Return inclusive (start, end) dates. Defaults to yesterday ~ today in KST.

    Settings may specify both dates (or neither); a CLI override must give both.
    """
    if explicit_start is not None or explicit_end is not None:
        if not explicit_start or not explicit_end:
            raise ValueError("직접 날짜를 지정할 때는 시작일과 종료일을 모두 입력하세요.")
        first, last = explicit_start, explicit_end
    else:
        configured = settings.get("date_range", {})
        if not isinstance(configured, dict):
            raise ValueError("date_range는 객체여야 합니다.")
        first, last = configured.get("start_date"), configured.get("end_date")
        if first is None and last is None:
            current = now if now is not None else datetime.now(KST)
            today = current.astimezone(KST).date()
            return today - timedelta(days=1), today
        if not first or not last:
            raise ValueError("date_range.start_date와 end_date를 모두 설정하거나 모두 null로 두세요.")
    first_day = _parse_date(first, "시작일")
    last_day = _parse_date(last, "종료일")
    if first_day > last_day:
        raise ValueError("시작일은 종료일보다 늦을 수 없습니다.")
    return first_day, last_day
