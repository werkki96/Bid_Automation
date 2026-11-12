# 파일 역할: UI 금액 표시와 저장하지 않은 강조 키워드의 검색 적용을 검증합니다.
# 마지막 업데이트: 2026-11-02
# 작업자: SMI 장지수

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from exporter import EXCEL_COLUMNS
from ui_helpers import format_price
import ui_runner


class PresentationTests(unittest.TestCase):
    def test_price_format(self):
        for value, expected in [(153670610, '153,670,610'), ('1,234,567', '1,234,567'),
                                ('1234.50', '1,234.50'), (0, '0'), (None, ''), ('', ''),
                                ('미정', '미정'), ('NaN', 'NaN')]:
            with self.subTest(value=value):
                self.assertEqual(format_price(value), expected)

    def test_unsaved_highlight_keywords_apply_to_screen_and_excel(self):
        with tempfile.TemporaryDirectory() as tmp:
            settings = {'highlight_keywords': ['기존단어'], 'highlight_rules': {}}
            events = []
            def collect(key, first, last, settings, on_page):
                rows = [{'공고명': '새주제 과제'}]
                on_page(rows, 1, 1, 1)
                return rows
            with patch.object(ui_runner, 'BASE_DIR', Path(tmp)), \
                 patch.object(ui_runner, 'load_settings', return_value=settings), \
                 patch.object(ui_runner, 'api_keys', return_value={'nara': 'test'}), \
                 patch.object(ui_runner.nara, 'collect', side_effect=collect):
                ui_runner.run_ui('2026-10-26', '2026-10-27', '', events.append,
                                 ('nara',), [], ['새주제'])
                row = next(e for e in events if e['type'] == 'page')['rows'][0]
                # 저장 버튼도 검색 때 사용한 UI 강조 키워드를 그대로 전달합니다.
                file = ui_runner.save_ui_results('nara', [row], '2026-10-26', '2026-10-27', ['새주제'])
            self.assertEqual(settings['highlight_keywords'], ['기존단어'])
            self.assertEqual(row['강조 근거'], '지정 키워드: 새주제')
            workbook = load_workbook(file)
            self.assertEqual(workbook.active[f"{get_column_letter(EXCEL_COLUMNS.index('공고명') + 1)}2"].fill.fgColor.rgb, 'FFFFFF00')
            workbook.close()
