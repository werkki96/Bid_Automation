# 파일 역할: 웹·데스크톱 UI가 공유하는 수집 실행기입니다. 페이지 이벤트 전달, 필터링과 Excel 저장을 담당합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Shared collection engine for Streamlit and Tkinter, streaming each API page."""
from pathlib import Path
from config import BASE_DIR, api_keys, load_settings, resolve_date_range
from collectors import nara, d2b
from exporter import filter_keywords, mark_excluded, save_excel
from highlighter import highlight_reasons


def parse_keywords(value):
    """Comma or newline separated keywords; OR condition."""
    if isinstance(value, str):
        return [v.strip() for v in value.replace('\n', ',').split(',') if v.strip()]
    return [str(v).strip() for v in value if str(v).strip()]


def run_ui(start_date, end_date, keywords, emit, sources=('nara', 'd2b'),
           exclude_keywords=None):
    """emit(event_dict) is called in a worker thread. No UI operations here."""
    settings = load_settings()
    first, last = resolve_date_range(settings, start_date, end_date)
    terms = parse_keywords(keywords)
    configured_ui_excludes = settings.get('exclude_keywords', [])
    ui_excludes = parse_keywords(
        configured_ui_excludes if exclude_keywords is None else exclude_keywords
    )
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
        # 화면 이벤트는 수집 스레드에서 전송하며, 화면 갱신은 각 UI의 메인 스레드가 담당합니다.
        shown = []
        def on_page(rows, page, pages, total):
            matched, _ = filter_keywords(
                rows,
                [],
                settings.get('keyword_search_columns', ['공고명']),
                include_keywords=terms,
            )
            search_columns = settings.get('keyword_search_columns', ['공고명'])
            matched = mark_excluded(matched, ui_excludes, search_columns)
            # 화면에도 저장 파일과 같은 규칙의 강조 근거를 전달합니다.
            matched = [dict(row, **{"강조 근거": "; ".join(highlight_reasons(
                row.get("공고명"), settings.get("highlight_keywords", []),
                settings.get("highlight_rules", {})))}) for row in matched]
            shown.extend(matched)
            emit({'type':'page', 'source':name, 'page':page, 'pages':pages,
                  'total':total, 'rows':matched, 'matched':len(shown)})
        try:
            # Collectors stream records per page and validate complete retrieval.
            collector(keys[name], first, last, settings, on_page=on_page)
            path = output / f'{label}_{filename_date}.xlsx'
            # CLI와 동일한 저장기로 강조·정렬·근거·헤더·테두리·자동필터를 적용합니다.
            save_excel(path, shown, settings.get("highlight_keywords", []),
                       settings.get("highlight_rules", {}))
            emit({'type':'saved', 'source':name, 'path':str(path), 'count':len(shown)})
        except Exception as exc:
            failures.append(name)
            emit({'type':'error', 'source':name, 'message':f'{type(exc).__name__}: {exc}'})
    emit({'type':'done', 'failures':failures})
