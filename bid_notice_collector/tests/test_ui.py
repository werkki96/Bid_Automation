# 파일 역할: 페이지 이벤트와 UI 공통 실행기의 키워드 필터·Excel 저장을 모의 응답으로 검증합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch
import ui_runner
from http_client import collect_pages
from openpyxl import load_workbook
from exporter import EXCEL_COLUMNS

class UITests(unittest.TestCase):
    def test_pages_stream_in_order(self):
        emitted=[]
        def fetch(n):
            return 3, ([{'v':1},{'v':2}] if n == 1 else [{'v':3}])
        rows=collect_pages(fetch, 2, 'test', on_page=lambda items,page,pages,total:emitted.append((page,pages,total,len(items))))
        self.assertEqual(len(rows), 3)
        self.assertEqual(emitted,[(1,2,3,2),(2,2,3,1)])

    def test_keywords_parser(self):
        self.assertEqual(ui_runner.parse_keywords('조사, 연구\n분석'), ['조사','연구','분석'])

    def test_ui_runner_filters_and_emits(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)
            settings={'page_size':100,'output_directory':'output','include_keywords':[], 'exclude_keywords':['제외'], 'keyword_search_columns':['공고명'], 'nara':{'display_name':'나라장터'},'d2b':{'display_name':'조달청'}}
            records=[{'공고명':'연구 사업'}, {'공고명':'단순 구매'}, {'공고명':'연구 제외 건'}]
            def fake_collector(key, start, end, config, on_page=None):
                self.assertEqual((start,end),(date(2026,10,12),date(2026,10,14)))
                on_page(records[:2],1,2,3)
                on_page(records[2:],2,2,3)
                return records
            events=[]
            with patch.object(ui_runner,'BASE_DIR',root),patch.object(ui_runner,'load_settings',return_value=settings),patch.object(ui_runner,'api_keys',return_value={'nara':'fake','d2b':'fake'}),patch.object(ui_runner.nara,'collect',side_effect=fake_collector),patch.object(ui_runner.d2b,'collect',side_effect=fake_collector):
                ui_runner.run_ui('2026-10-12','2026-10-14','연구',events.append)
            self.assertEqual(len([e for e in events if e['type']=='page']),4)
            self.assertEqual(len([e for e in events if e['type']=='saved']),2)
            self.assertEqual(sum(len(e['rows']) for e in events if e['type']=='page'),2)
            self.assertTrue((root/'output'/'나라장터_20261012_20261014.xlsx').exists())
            self.assertEqual(events[-1],{'type':'done','failures':[]})

    def test_excel_format_and_highlights_across_pages(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            settings = {
                'output_directory': 'output', 'exclude_keywords': ['제외'],
                'keyword_search_columns': ['공고명'],
                'highlight_keywords': ['AI'], 'highlight_rules': {'과업': ['연구']},
            }
            records = [{'공고명': '일반 용역'}, {'공고명': 'AI 구축'},
                       {'공고명': '정책 연구'}, {'공고명': 'AI 제외'}, {'공고명': '=1+1'}]
            def collect(key, first, last, settings, on_page):
                on_page(records[:2], 1, 2, 5)
                on_page(records[2:], 2, 2, 5)
                return records
            events = []
            with patch.object(ui_runner, 'BASE_DIR', root), \
                 patch.object(ui_runner, 'load_settings', return_value=settings), \
                 patch.object(ui_runner, 'api_keys', return_value={'nara': 'test', 'd2b': 'test'}), \
                 patch.object(ui_runner.nara, 'collect', side_effect=collect), \
                 patch.object(ui_runner.d2b, 'collect', side_effect=collect):
                ui_runner.run_ui('2026-10-12', '2026-10-14', '', events.append)
            saved = [e for e in events if e['type'] == 'saved']
            self.assertEqual(len(saved), 2)
            self.assertFalse(list((root / 'output').glob('*.csv')))
            for event in saved:
                self.assertEqual(event['count'], 4)
                workbook = load_workbook(event['path'])
                sheet = workbook.active
                self.assertEqual([c.value for c in sheet[1]], EXCEL_COLUMNS)
                self.assertEqual([c.value for c in sheet['E'][1:]],
                                 ['AI 구축', '정책 연구', '일반 용역', '=1+1'])
                self.assertEqual(sheet['E2'].fill.fgColor.rgb, 'FFFFFF00')
                self.assertEqual(sheet['E3'].fill.fgColor.rgb, 'FFFFFF00')
                self.assertIsNone(sheet['E4'].fill.patternType)
                self.assertEqual(sheet['K2'].value, '지정 키워드: AI')
                self.assertEqual(sheet['K3'].value, '과업: 연구')
                self.assertEqual(sheet['A1'].fill.fgColor.rgb, 'FFD9D9D9')
                self.assertEqual(sheet['E2'].border.left.style, 'thin')
                self.assertEqual(sheet.freeze_panes, 'A2')
                self.assertEqual(sheet.auto_filter.ref, 'A1:K5')
                self.assertEqual(sheet['E5'].data_type, 's')
                workbook.close()
            page_rows = [r for e in events if e['type'] == 'page' for r in e['rows']]
            self.assertTrue(all('강조 근거' in r for r in page_rows))
            self.assertTrue(any(r['강조 근거'] == '과업: 연구' for r in page_rows))
            self.assertNotIn('강조 근거', records[0])

    def test_failed_collection_preserves_existing_excel(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / 'output'
            output.mkdir()
            path = output / '나라장터_20261012_20261014.xlsx'
            path.write_bytes(b'existing result')
            events = []
            def fail(key, first, last, settings, on_page):
                on_page([{'공고명': '연구'}], 1, 2, 2)
                raise RuntimeError('second page failed')
            with patch.object(ui_runner, 'BASE_DIR', root), \
                 patch.object(ui_runner, 'load_settings', return_value={}), \
                 patch.object(ui_runner, 'api_keys', return_value={'nara': 'test'}), \
                 patch.object(ui_runner.nara, 'collect', side_effect=fail):
                ui_runner.run_ui('2026-10-12', '2026-10-14', '', events.append, ('nara',))
            self.assertEqual(path.read_bytes(), b'existing result')
            self.assertFalse(any(e['type'] == 'saved' for e in events))
            self.assertEqual(events[-1]['failures'], ['nara'])

if __name__=='__main__':
    unittest.main()
