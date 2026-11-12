# 파일 역할: Streamlit 검색 결과의 금액, 출처 분리, 필터와 새 탭 링크를 검증합니다.
# 마지막 업데이트: 2026-10-27
# 작업자: SMI 장지수

import importlib.util
from pathlib import Path
import queue
import unittest


@unittest.skipUnless(importlib.util.find_spec('streamlit'), 'requirements-web.txt 설치 필요')
class WebAppTests(unittest.TestCase):
    def test_result_display_and_filters(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'web_app.py')).run()
        self.assertFalse(app.exception)
        events = queue.Queue()
        events.put({'type': 'page', 'source': 'nara', 'page': 1, 'pages': 1, 'matched': 2,
                    'rows': [{'공고명': '일반 공고', '공고위치': '사용자표시명', '강조 근거': '', '추정가격(원)': '1234567'},
                             {'공고명': 'AI 연구', '공고위치': '사용자표시명', '강조 근거': 'AI',
                              '추정가격(원)': '153670610', '상세보기 주소': 'https://example.org/bid'}]})
        events.put({'type': 'done', 'failures': []})
        app.session_state['job'] = {
            'running': True, 'queue': events, 'rows': [], 'by_source': {'nara': [], 'd2b': []},
            'statuses': {'nara': '수집 중'}, 'logs': [], 'files': [], 'errors': [],
            'failures': [], 'period': '2026.10.26 — 2026.10.27',
        }
        app.run()
        self.assertFalse(app.exception)
        self.assertEqual(app.dataframe[0].value['추정가격(원)'].tolist(), ['153,670,610', '1,234,567'])
        self.assertEqual(app.dataframe[0].value['공고명'].tolist(), ['AI 연구', '일반 공고'])
        self.assertFalse(next(b for b in app.button if b.label == '공고 검색 시작').disabled)
        self.assertEqual(app.get('link_button')[0].proto.url, 'https://example.org/bid')
        app.checkbox(key='highlighted_only').check().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.dataframe[0].value), 1)
        app.text_input(key='result_find').set_value('없는공고').run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.dataframe), 0)
