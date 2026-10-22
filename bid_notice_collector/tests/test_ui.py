import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch
import ui_runner
from http_client import collect_pages

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
            self.assertTrue((root/'output'/'나라장터_20261012_20261014.csv').exists())
            self.assertEqual(events[-1],{'type':'done','failures':[]})

if __name__=='__main__':
    unittest.main()
