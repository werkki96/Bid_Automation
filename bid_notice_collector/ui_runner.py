"""Shared collection engine for Streamlit and Tkinter, streaming each API page."""
from pathlib import Path
from config import BASE_DIR, api_keys, load_settings, resolve_date_range
from collectors import nara, d2b
from exporter import filter_keywords, save_csv


def parse_keywords(value):
    """Comma or newline separated keywords; OR condition."""
    if isinstance(value, str):
        return [v.strip() for v in value.replace('\n', ',').split(',') if v.strip()]
    return [str(v).strip() for v in value if str(v).strip()]


def run_ui(start_date, end_date, keywords, emit, sources=('nara', 'd2b')):
    """emit(event_dict) is called in a worker thread. No UI operations here."""
    settings = load_settings()
    first, last = resolve_date_range(settings, start_date, end_date)
    terms = parse_keywords(keywords)
    keys = api_keys()
    output = (BASE_DIR / settings.get('output_directory', 'output')).resolve()
    if not output.is_relative_to(BASE_DIR.resolve()):
        raise ValueError('결과물 저장 경로는 프로젝트 내부여야 합니다.')
    filename_date = f'{first:%Y%m%d}_{last:%Y%m%d}'
    jobs = [('nara', '나라장터', nara.collect), ('d2b', 'D2B', d2b.collect)]
    failures = []
    for name, label, collector in jobs:
        if name not in sources:
            continue
        emit({'type':'status', 'source':name, 'message':f'{label} 수집 시작'})
        if not keys.get(name):
            failures.append(name)
            emit({'type':'error', 'source':name, 'message':f'{name.upper()}_API_KEY를 .env에 설정하세요.'})
            continue
        shown = []
        def on_page(rows, page, pages, total):
            matched, _ = filter_keywords(
                rows,
                settings.get('exclude_keywords', []),
                settings.get('keyword_search_columns', ['공고명']),
                include_keywords=terms,
            )
            shown.extend(matched)
            emit({'type':'page', 'source':name, 'page':page, 'pages':pages,
                  'total':total, 'rows':matched, 'matched':len(shown)})
        try:
            # Collectors stream records per page and validate complete retrieval.
            collector(keys[name], first, last, settings, on_page=on_page)
            path = output / f'{label}_{filename_date}.csv'
            save_csv(path, shown)
            emit({'type':'saved', 'source':name, 'path':str(path), 'count':len(shown)})
        except Exception as exc:
            failures.append(name)
            emit({'type':'error', 'source':name, 'message':f'{type(exc).__name__}: {exc}'})
    emit({'type':'done', 'failures':failures})
