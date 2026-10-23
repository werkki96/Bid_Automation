# 파일 역할: Streamlit 화면에서 공고를 조회하고 진행 상태, 표와 Excel 다운로드를 제공합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Start with: streamlit run web_app.py"""
import queue
import threading
from datetime import timedelta, datetime
import streamlit as st
import pandas as pd
from config import KST
from exporter import EXCEL_COLUMNS
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

# 전체 화면을 다시 실행하지 않고 큐의 새 이벤트만 주기적으로 화면에 반영합니다.
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
    # 전체 수집 결과에서 강조 공고를 먼저 표시하며 같은 그룹의 수집 순서를 유지합니다.
    rows = sorted(job['rows'], key=lambda row: not bool(row.get('강조 근거')))
    frame = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    def style_row(row):
        return ['background-color: #FFFF00' if col == '공고명' and row['강조 근거'] else ''
                for col in EXCEL_COLUMNS]
    st.dataframe(frame.style.apply(style_row, axis=1), column_order=EXCEL_COLUMNS,
                 use_container_width=True, hide_index=True)
    for path in job['files']:
        from pathlib import Path
        file = Path(path)
        if file.is_file():
            st.download_button(f'{file.name} 다운로드', file.read_bytes(), file_name=file.name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key=path)
    with st.expander('진행 기록', expanded=job['running']):
        st.text('\n'.join(job['logs'][-80:]))

results()
