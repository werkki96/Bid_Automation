# 파일 역할: 페이지 이벤트와 UI 공통 실행기의 키워드 필터, 저장 버튼용 Excel·원본 형식 저장을 모의 응답으로 검증합니다.
# 마지막 업데이트: 2026-11-02
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

from openpyxl.utils import get_column_letter
# 컬럼이 추가돼도 같은 열을 검사하도록 위치를 이름으로 계산합니다.
TITLE_COL = get_column_letter(EXCEL_COLUMNS.index("공고명") + 1)
REASON_COL = get_column_letter(EXCEL_COLUMNS.index("강조 근거") + 1)
LAST_COL = get_column_letter(len(EXCEL_COLUMNS))

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
            self.assertEqual([e['source'] for e in events if e['type']=='collected'],['nara','d2b'])
            self.assertEqual(sum(len(e['rows']) for e in events if e['type']=='page'),4)
            # 저장 버튼을 누르기 전에는 Excel 파일을 만들지 않습니다.
            self.assertFalse((root/'output').exists())
            self.assertEqual(events[-1],{'type':'done','failures':[]})

    def test_excel_format_and_highlights_across_pages(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            settings = {
                'output_directory': 'output', 'exclude_keywords': ['제외'],
                'keyword_search_columns': ['공고명'],
                'highlight_keywords': ['AI'], 'highlight_rules': {'과업': ['연구']},
            }
            records = [{'공고명': '일반 용역'}, {'공고명': 'AI 연구'},
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
                collected = [e for e in events if e['type'] == 'collected']
                self.assertEqual([e['count'] for e in collected], [5, 5])
                saved = [
                    ui_runner.save_ui_results(
                        source, [r for e in events if e['type'] == 'page' and e['source'] == source
                                 for r in e['rows']],
                        '2026-10-12', '2026-10-14')
                    for source in ('nara', 'd2b')
                ]
            self.assertEqual([path.name for path in saved],
                             ['나라장터_20261012_20261014.xlsx', 'D2B_20261012_20261014.xlsx'])
            self.assertFalse(list((root / 'output').glob('*.csv')))
            for path in saved:
                workbook = load_workbook(path)
                sheet = workbook.active
                self.assertEqual([c.value for c in sheet[1]], EXCEL_COLUMNS)
                self.assertEqual([c.value for c in sheet[TITLE_COL][1:]],
                                 ['AI 연구', '일반 용역', '정책 연구', 'AI 제외', '=1+1'])
                self.assertEqual(sheet[f'{TITLE_COL}2'].fill.fgColor.rgb, 'FFFFFF00')
                self.assertIsNone(sheet[f'{TITLE_COL}3'].fill.patternType)
                self.assertIsNone(sheet[f'{TITLE_COL}4'].fill.patternType)
                self.assertEqual(sheet[f'{REASON_COL}2'].value, '지정 키워드: AI; 과업: 연구')
                self.assertIsNone(sheet[f'{REASON_COL}3'].value)
                self.assertEqual(sheet['A1'].fill.fgColor.rgb, 'FFD9D9D9')
                self.assertEqual(sheet[f'{TITLE_COL}2'].border.left.style, 'thin')
                self.assertEqual(sheet.freeze_panes, 'A2')
                self.assertEqual(sheet.auto_filter.ref, f'A1:{LAST_COL}6')
                self.assertEqual(sheet[f'{TITLE_COL}5'].data_type, 's')
                self.assertIsNone(sheet[f'{TITLE_COL}5'].fill.patternType)
                self.assertIsNone(sheet[f'{TITLE_COL}6'].fill.patternType)
                workbook.close()
            page_rows = [r for e in events if e['type'] == 'page' for r in e['rows']]
            self.assertTrue(all('강조 근거' in r for r in page_rows))
            self.assertTrue(any('과업: 연구' in r['강조 근거'] for r in page_rows))
            self.assertEqual(
                next(r for r in page_rows if r['공고명'] == 'AI 제외')['강조 근거'],
                '',
            )
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
            self.assertFalse(any(e['type'] == 'collected' for e in events))
            self.assertEqual(events[-1]['failures'], ['nara'])

    def test_original_excel_layouts_keep_highlight_and_exclusion(self):
        nara_items = [
            {'bidNtceNo': 'R26BK1', 'bidNtceOrd': '000', 'bidNtceNm': '일반 시설 용역',
             'srvceDivNm': '일반용역', 'intrbidYn': 'N', 'ntceKindNm': '등록공고',
             'ntceInsttNm': '조달청', 'dminsttNm': '수요처',
             'bidNtceDt': '2026-10-27 15:43:35', 'bidClseDt': '2026-11-17 10:00:00'},
            {'bidNtceNo': 'R26BK2', 'bidNtceOrd': '001', 'bidNtceNm': 'AI 연구',
             'srvceDivNm': '기술용역', 'ntceKindNm': '재공고', 'ntceInsttNm': '기관',
             'bidNtceDt': '2026-10-27 09:00:00', 'bidClseDt': ''},
            {'bidNtceNo': 'R26BK3', 'bidNtceOrd': '000', 'bidNtceNm': 'AI 제외 연구',
             'srvceDivNm': '일반용역', 'intrbidYn': 'N', 'ntceKindNm': '등록공고'},
        ]
        d2b_items = [
            {'bidNm': '정수기 임차 용역', 'pblancSe': '긴급공고', 'pblancNo': 'LKH0072', 'pblancOdr': '1',
             'pblancDate': '20261027', 'demandYear': '2026', 'orntCode': 'LKH', 'dcsNo': '43625',
             'ornt': '제39보병사단', 'bidPartcptRegistClosDt': '202611021100',
             'biddocPresentnClosDt': '202611031000', 'cntrctMth': '제한경쟁', 'bidStle': '전자입찰',
             'bsicExpt': '53460000'},
            {'bidNm': 'AI 연구', 'pblancSe': '재공고', 'pblancNo': 'SFC0158', 'pblancOdr': '2',
             'pblancDate': '20261027', 'demandYear': '2026', 'orntCode': 'SFC', 'dcsNo': '42471'},
        ]
        keywords, rules = ['AI'], {}
        nara_rows = ui_runner.mark_highlight_suppressed(
            [ui_runner.nara.convert(item, '나라장터') for item in nara_items], ['제외'], ['공고명'])
        d2b_rows = [ui_runner.d2b.convert(item, '조달청') for item in d2b_items]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            settings = {'output_directory': 'output', 'highlight_keywords': ['사용안함'], 'highlight_rules': rules}
            with patch.object(ui_runner, 'BASE_DIR', root), \
                 patch.object(ui_runner, 'load_settings', return_value=settings):
                nara_path = ui_runner.save_ui_results('nara', nara_rows, '2026-10-27', '2026-10-27',
                                                      keywords, original=True)
                d2b_path = ui_runner.save_ui_results('d2b', d2b_rows, '2026-10-27', '2026-10-27',
                                                     keywords, original=True)
            self.assertEqual(nara_path.name, '나라장터_원본_20261027_20261027.xlsx')
            self.assertEqual(d2b_path.name, 'D2B_원본_20261027_20261027.xlsx')

            workbook = load_workbook(nara_path)
            sheet = workbook.active
            self.assertEqual(sheet.title, '입찰공고')
            self.assertEqual(sheet['A1'].value, '입찰공고목록')
            self.assertTrue(sheet['A3'].value.startswith('생성일 : '))
            self.assertEqual([c.value for c in sheet[5]],
                             ['업무구분', '업무여부', '구분', '입찰공고번호', '공고명', '공고기관',
                              '수요기관', '게시일시(입찰마감일시)'])
            self.assertEqual(sheet['A5'].fill.fgColor.rgb, 'FFC0C0C0')
            # 강조 공고가 먼저 오고, 제외 조건에 걸린 공고는 남되 강조하지 않습니다.
            self.assertEqual([c.value for c in sheet['E'][5:]], ['AI 연구', '일반 시설 용역', 'AI 제외 연구'])
            self.assertEqual([c.value for c in sheet[6]],
                             ['기술용역', '-', '재공고', 'R26BK2-001', 'AI 연구', '기관', None, '2026/10/27 09:00(-)'])
            self.assertEqual(sheet['H7'].value, '2026/10/27 15:43(2026/11/17 10:00)')
            self.assertEqual(sheet['B7'].value, '내자')
            self.assertEqual(sheet['E6'].fill.fgColor.rgb, 'FFFFFF00')
            self.assertIsNone(sheet['E7'].fill.patternType)
            self.assertIsNone(sheet['E8'].fill.patternType)
            self.assertEqual(sheet.auto_filter.ref, 'A5:H8')
            workbook.close()

            workbook = load_workbook(d2b_path)
            sheet = workbook.active
            self.assertEqual(sheet.title, '데이터')
            self.assertEqual(sheet['A1'].value, '순번')
            self.assertEqual(sheet['O1'].value, '기초예가')
            self.assertEqual([c.value for c in sheet[2]],
                             ['1', '경쟁입찰', '재공고', '2026-10-27', '재공고 SFC0158-2', '2026SFC42471',
                              '42471', 'AI 연구', None, None, None, None, None, None, None])
            self.assertEqual([c.value for c in sheet[3]],
                             ['2', '경쟁입찰', '긴급공고', '2026-10-27', '긴급공고 LKH0072-1', '2026LKH43625',
                              '43625', '정수기 임차 용역', '제39보병사단', None, '2026-11-02 11:00',
                              '2026-11-03 10:00', '제한경쟁', '전자입찰', '53,460,000'])
            self.assertEqual(sheet['H2'].fill.fgColor.rgb, 'FFFFFF00')
            self.assertIsNone(sheet['H3'].fill.patternType)
            self.assertEqual(sheet['A1'].font.name, 'Arial')
            workbook.close()

if __name__=='__main__':
    unittest.main()
