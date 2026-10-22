"""Start with: streamlit run web_app.py"""
import queue
import threading
from datetime import timedelta, datetime
import streamlit as st
from config import KST, COLUMNS
from ui_runner import run_ui

st.set_page_config(page_title='입찰공고 통합 검색', layout='wide')
st.title('입찰공고 통합 검색')
st.caption('나라장터 학술·연구용역(1169) · D2B 경쟁입찰공고')
now = datetime.now(KST).date()
start = st.date_input('검색 시작일', value=now-timedelta(days=1), key='start')
end = st.date_input('검색 종료일', value=now, key='end')
keywords = st.text_input('포함 검색 키워드 (쉼표 구분, OR 검색)', key='keywords')
sources = st.multiselect('검색 대상', ['나라장터', 'D2B'], default=['나라장터','D2B'])
if 'job' not in st.session_state:
    st.session_state.job = None
if st.button('검색 시작', disabled=bool(st.session_state.job and st.session_state.job['running'])):
    if start > end:
        st.error('시작일은 종료일보다 늦을 수 없습니다.')
    elif not sources:
        st.error('검색 대상을 하나 이상 선택하세요.')
    else:
        events = queue.Queue()
        state = {'running': True, 'queue': events, 'rows': [], 'logs': [], 'files': []}
        st.session_state.job = state
        selected = tuple('nara' if s=='나라장터' else 'd2b' for s in sources)
        def work():
            try:
                run_ui(start.isoformat(), end.isoformat(), keywords, events.put, selected)
            except Exception as exc:
                events.put({'type':'error', 'message':str(exc)})
                events.put({'type':'done', 'failures':['setup']})
        threading.Thread(target=work, daemon=True).start()
        st.rerun()

@st.fragment(run_every='1s')
def results():
    job = st.session_state.job
    if job is None:
        st.info('검색을 실행하면 수집된 공고가 화면에 추가됩니다.')
        return
    events = job['queue']
    while True:
        try:
            event = events.get_nowait()
        except queue.Empty:
            break
        typ = event['type']
        if typ == 'page':
            job['rows'].extend(event['rows'])
            job['logs'].append(f"{event['source']}: {event['page']}/{event['pages']}페이지 · 표시 {event['matched']}건")
        elif typ == 'saved':
            job['files'].append(event['path'])
            job['logs'].append(f"저장: {event['path']} ({event['count']}건)")
        elif typ == 'error':
            job['logs'].append('오류: '+event['message'])
        elif typ == 'status':
            job['logs'].append(event['message'])
        elif typ == 'done':
            job['running'] = False
            job['logs'].append('수집 종료' + (' (일부 실패)' if event['failures'] else ''))
    st.subheader(f"검색 결과: {len(job['rows'])}건" + (' · 수집 중' if job['running'] else ''))
    st.dataframe(job['rows'], column_order=COLUMNS, use_container_width=True, hide_index=True)
    for path in job['files']:
        from pathlib import Path
        file = Path(path)
        if file.is_file():
            st.download_button(f'{file.name} 다운로드', file.read_bytes(), file_name=file.name, mime='text/csv', key=path)
    with st.expander('진행 기록', expanded=job['running']):
        st.text('\n'.join(job['logs'][-80:]))

results()
