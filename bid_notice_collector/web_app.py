# 파일 역할: 검색 조건, 진행 현황, 강조 공고와 Excel 다운로드를 제공하는 Streamlit 화면입니다.
# 마지막 업데이트: 2026-10-27
# 작업자: SMI 장지수

"""Start with: streamlit run web_app.py"""
import copy
import queue
import threading
from datetime import timedelta, datetime
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import streamlit as st

from config import KST, load_settings, save_ui_keywords
from exporter import EXCEL_COLUMNS
from ui_runner import parse_keywords, run_ui
from ui_helpers import format_price

st.set_page_config(page_title='SMI | 입찰공고', page_icon='📋', layout='wide')
st.markdown('''<style>
.stApp {background: #f5f7fb; color: #172b4d;}
[data-testid="stSidebar"] {background: #fff; border-right: 1px solid #e2e8f0;}
.block-container {padding-top: 4rem; padding-bottom: 2rem; max-width: 1600px;}
h1, h2, h3 {letter-spacing: -.035em;}
.smi-hero {padding: 28px 32px; border-radius: 20px; color: white;
 background: linear-gradient(115deg, #142d4e, #20567b); margin-bottom: 24px;}
.smi-hero small {color: #b5d7ee; letter-spacing: .16em; font-weight: 700;}
.smi-hero h1 {color: #fff; font-size: 2rem; padding: 8px 0; margin: 0;}
.smi-hero p {color: #d5e4ef; margin: 0; line-height: 1.7;}
[data-testid="stMetric"] {background: white; border: 1px solid #e2e8f0;
 border-radius: 14px; padding: 16px 20px;}
[data-testid="stMetricValue"] {color: #173e68;}
[data-testid="stBaseButton-primary"] {background: #2260a4; border-color: #2260a4;}
[data-testid="stBaseButton-primary"]:hover {background: #174b85; border-color: #174b85;}
.stButton button, .stDownloadButton button {border-radius: 9px; min-height: 42px;}
.smi-empty {background: white; border: 1px dashed #cbd5e1; border-radius: 16px;
 padding: 46px 24px; text-align: center; color: #52637a; margin-top: 16px;}
.smi-empty h3 {color: #203f62;}
@media(max-width: 700px) {.smi-hero {padding: 24px;} .smi-hero h1 {font-size: 1.55rem;}}
</style>''', unsafe_allow_html=True)

settings = load_settings()
now = datetime.now(KST).date()
st.session_state.setdefault('web_highlight_keywords', list(settings.get('highlight_keywords', [])))
st.session_state.setdefault('web_exclude_keywords', list(settings.get('exclude_keywords', [])))
st.session_state.setdefault('job', None)
st.session_state.setdefault('start', now - timedelta(days=1))
st.session_state.setdefault('end', now)
running = bool(st.session_state.job and st.session_state.job['running'])


def set_period(days):
    today = datetime.now(KST).date()
    st.session_state.start = today - timedelta(days=days - 1)
    st.session_state.end = today


def _exclude_editor_text(values):
    return '\n'.join(' + '.join(item) if isinstance(item, list) else str(item) for item in values)


@st.dialog('강조 키워드 편집')
def edit_highlight_keywords():
    st.caption('관심 분야의 공고를 찾는 추가 키워드입니다. 적용 후 다음 검색부터 사용합니다.')
    value = st.text_area('쉼표 또는 줄바꿈으로 구분',
                         value='\n'.join(st.session_state.web_highlight_keywords),
                         height=260, key='highlight_editor')
    if st.button('변경 적용', type='primary', key='apply_highlight_editor'):
        st.session_state.web_highlight_keywords = parse_keywords(value)
        st.session_state.pop('highlight_editor', None)
        st.rerun()


@st.dialog('강조 제외 조건 편집')
def edit_exclude_keywords():
    st.caption('일치하는 공고는 결과에 남겨두고 노란색 강조만 해제합니다.')
    value = st.text_area('한 줄에 한 조건 · 모두 일치: + · 하나라도 일치: |',
                         value=_exclude_editor_text(st.session_state.web_exclude_keywords),
                         height=300, key='exclude_editor')
    st.caption('예: 학생|학교|교육 + 연수 → 학생·학교·교육 중 하나와 연수가 함께 포함된 경우')
    if st.button('변경 적용', type='primary', key='apply_exclude_editor'):
        edited = []
        for line in value.splitlines():
            parts = [part.strip() for part in line.split('+') if part.strip()]
            if parts:
                edited.append(parts if len(parts) > 1 else parts[0])
        st.session_state.web_exclude_keywords = edited
        st.session_state.pop('exclude_editor', None)
        st.rerun()


with st.sidebar:
    st.markdown('### SMI · 공고 검색')
    st.caption('조건 설정 → 공고 검토 → Excel 저장')
    st.divider()
    st.markdown('#### 조회 기간')
    period_cols = st.columns(3)
    for col, label, days in zip(period_cols, ['오늘', '2일', '7일'], [1, 2, 7]):
        col.button(label, on_click=set_period, args=(days,), disabled=running, use_container_width=True)
    start = st.date_input('시작일', key='start', disabled=running)
    end = st.date_input('종료일', key='end', disabled=running)
    st.caption('한국시간(KST) 기준 · 시작일과 종료일 포함')
    sources = st.multiselect('검색 대상', ['나라장터', 'D2B'], default=['나라장터', 'D2B'], disabled=running)
    keywords = st.text_input('포함 검색어', value=', '.join(settings.get('include_keywords', [])),
                             placeholder='예: 연구, 인공지능, 정책', key='keywords', disabled=running,
                             help='쉼표로 구분한 검색어 중 하나라도 포함된 공고를 찾습니다. 비워두면 전체를 조회합니다.')
    st.divider()
    search_clicked = st.button('수집 중…' if running else '공고 검색 시작', type='primary',
                                disabled=running, use_container_width=True)
    st.divider()
    st.markdown('#### 검토 기준')
    st.caption(f"강조 키워드 {len(st.session_state.web_highlight_keywords)}개 · "
               f"강조 제외 조건 {len(st.session_state.web_exclude_keywords)}개")
    editor_cols = st.columns(2)
    if editor_cols[0].button('강조 편집', disabled=running, use_container_width=True):
        edit_highlight_keywords()
    if editor_cols[1].button('강조 제외 편집', disabled=running, use_container_width=True):
        edit_exclude_keywords()
    if st.button('현재 키워드를 기본값으로 저장', disabled=running, use_container_width=True):
        try:
            save_ui_keywords(parse_keywords(keywords), st.session_state.web_exclude_keywords,
                             st.session_state.web_highlight_keywords)
            st.success('다음 실행에도 사용할 기본값을 저장했습니다.')
        except (OSError, ValueError, TypeError) as exc:
            st.error(f'저장 실패: {exc}')
    st.caption('편집한 조건은 다음 검색에 적용됩니다. 기본값 저장은 선택 사항입니다.')
    st.divider()

st.markdown('''<div class="smi-hero"><small>SMI · BID NOTICE WORKSPACE</small>
<h1>검토할 공고를 한눈에.</h1><p>나라장터와 국방전자조달의 공고를 모으고,<br>
관심 공고를 강조해 빠르게 검토합니다.</p></div>''', unsafe_allow_html=True)

if search_clicked:
    if start > end:
        st.error('조회 기간을 확인해 주세요. 시작일은 종료일보다 늦을 수 없습니다.')
    elif not sources:
        st.error('검색 대상을 하나 이상 선택해 주세요.')
    else:
        events = queue.Queue()
        selected = tuple('nara' if source == '나라장터' else 'd2b' for source in sources)
        st.session_state.job = {
            'running': True, 'queue': events, 'rows': [], 'by_source': {'nara': [], 'd2b': []},
            'logs': [], 'files': [], 'errors': [], 'failures': [],
            'statuses': {source: '대기 중' for source in selected},
            'period': f'{start:%Y.%m.%d} — {end:%Y.%m.%d}',
        }
        # 작업 스레드에는 검색 순간의 값을 복사해 전달합니다. 화면 상태를 직접 읽지 않습니다.
        args = (start.isoformat(), end.isoformat(), keywords, events.put, selected,
                copy.deepcopy(st.session_state.web_exclude_keywords),
                list(st.session_state.web_highlight_keywords))
        def work(arguments):
            try:
                run_ui(*arguments)
            except Exception as exc:
                events.put({'type': 'error', 'message': str(exc)})
                events.put({'type': 'done', 'failures': ['setup']})
        threading.Thread(target=work, args=(args,), daemon=True).start()
        st.rerun()


@st.fragment(run_every='1s')
def results():
    job = st.session_state.job
    if job is None:
        a, b, c = st.columns(3)
        a.metric('수집 공고', '—')
        b.metric('강조 공고', '—')
        c.metric('저장된 Excel', '—')
        st.markdown('''<div class="smi-empty"><h3>첫 검색을 시작해 보세요</h3>
        <p>왼쪽에서 조회 기간과 검색 대상을 선택하세요.<br>
        검색어를 비워두면 해당 기간의 전체 공고를 확인할 수 있습니다.</p></div>''', unsafe_allow_html=True)
        return
    completed = False
    while True:
        try:
            event = job['queue'].get_nowait()
        except queue.Empty:
            break
        typ = event['type']
        source = event.get('source')
        if typ == 'page':
            job['rows'].extend(event['rows'])
            job['by_source'].setdefault(source, []).extend(event['rows'])
            job['statuses'][source] = f"{event['page']}/{event['pages']}페이지 · {event['matched']}건"
            job['logs'].append(f"{source}: {job['statuses'][source]}")
        elif typ == 'saved':
            job['files'].append(event['path'])
            job['statuses'][source] = '저장 완료'
            job['logs'].append(f"저장 완료: {Path(event['path']).name} ({event['count']}건)")
        elif typ == 'error':
            job['errors'].append(event['message'])
            job['logs'].append('오류: ' + event['message'])
            if source:
                job['statuses'][source] = '확인 필요'
        elif typ == 'status':
            job['statuses'][source] = '수집 중'
            job['logs'].append(event['message'])
        elif typ == 'done':
            job['running'] = False
            job['failures'] = event['failures']
            completed = True
    # 사이드바의 검색 버튼도 다시 활성화되도록 완료 시 한 번만 전체 화면을 갱신합니다.
    if completed:
        st.rerun()
    a, b, c = st.columns(3)
    a.metric('수집 공고', f"{len(job['rows']):,}건")
    b.metric('강조 공고', f"{sum(bool(row.get('강조 근거')) for row in job['rows']):,}건")
    c.metric('저장된 Excel', f"{len(job['files'])}개")
    st.caption(job['period'] + ' · KST')
    if job['running']:
        st.info('공고를 수집하고 있습니다. 도착한 결과부터 표시하며, 수집 완료 후 Excel이 저장됩니다.')
    elif job['failures']:
        st.warning('일부 공고를 수집하지 못했습니다. 아래 오류를 확인한 뒤 다시 검색해 주세요.')
    else:
        st.success('수집이 완료되었습니다. 강조 공고를 검토하고 Excel을 다운로드하세요.')
    for error in job['errors']:
        st.error(error)
    controls = st.columns([2, 1])
    find = controls[0].text_input('결과 내 검색', placeholder='공고명 또는 기관명으로 찾기', key='result_find')
    highlighted_only = controls[1].checkbox('강조 공고만 보기', key='highlighted_only')
    st.caption('노란색은 검토 후보입니다. 화면 필터는 저장된 Excel 전체 결과에 영향을 주지 않습니다.')
    tabs = st.tabs(['나라장터', '국방전자조달 · D2B'])
    for tab, source in zip(tabs, ('nara', 'd2b')):
        with tab:
            if source not in job['statuses']:
                st.caption('이번 검색에서 선택하지 않은 대상입니다.')
                continue
            source_rows = job['by_source'].get(source, [])
            rows = [row for row in source_rows if (not highlighted_only or row.get('강조 근거'))
                    and (not find.strip() or find.strip().casefold() in ' '.join(
                        str(row.get(col) or '') for col in ('공고명', '공고기관(발주기관)', '수요기관')).casefold())]
            rows.sort(key=lambda row: not bool(row.get('강조 근거')))
            st.caption(f"{job['statuses'][source]} · 전체 {len(source_rows):,}건 중 {len(rows):,}건 표시")
            if not rows:
                st.info('표시할 공고가 없습니다. 수집 중이라면 잠시 기다리거나 화면 검색 조건을 확인해 주세요.')
                continue
            frame = pd.DataFrame(rows, columns=EXCEL_COLUMNS).fillna('')
            frame["추정가격(원)"] = frame["추정가격(원)"].map(format_price)
            def style_row(row):
                return ['background-color: #fff1b8; color: #503e00' if col == '공고명' and row['강조 근거'] else ''
                        for col in EXCEL_COLUMNS]
            st.dataframe(frame.style.apply(style_row, axis=1), column_order=EXCEL_COLUMNS,
                         column_config={'공고명': st.column_config.TextColumn(width='large'),
                                        '강조 근거': st.column_config.TextColumn(width='large'),
                                        '추정가격(원)': st.column_config.TextColumn(width='medium'),
                                        '상세보기 주소': st.column_config.LinkColumn('상세보기', display_text='공고 열기', validate=r'^https?://.+')},
                         height=420, use_container_width=True, hide_index=True)
            with st.expander('공고 상세 확인 · 새 탭으로 열기'):
                selected_index = st.selectbox('확인할 공고', range(len(rows)),
                    format_func=lambda index: f"{index + 1}. {rows[index].get('공고명', '')}",
                    key=f'detail_{source}')
                selected_row = rows[selected_index]
                st.write(selected_row.get('공고명', ''))
                st.caption(selected_row.get('강조 근거') or '자동 강조 조건에 해당하지 않는 일반 공고입니다.')
                url = str(selected_row.get('상세보기 주소') or '').strip()
                parsed = urlparse(url)
                if parsed.scheme in ('http', 'https') and parsed.netloc:
                    # Streamlit의 link_button은 target=_blank로 새 탭을 엽니다.
                    st.link_button('공고 원문 새 탭으로 열기', url)
                else:
                    st.caption('이 공고에는 유효한 상세보기 주소가 제공되지 않았습니다.')
    if job['files']:
        st.markdown('#### Excel 내려받기')
        download_cols = st.columns(2)
        for index, path in enumerate(job['files']):
            file = Path(path)
            if file.is_file():
                download_cols[index % 2].download_button(file.name, file.read_bytes(), file_name=file.name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=path, use_container_width=True)
    with st.expander('상세 진행 기록'):
        st.text('\n'.join(job['logs'][-80:]))


results()
