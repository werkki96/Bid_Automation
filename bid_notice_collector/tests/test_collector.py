import csv
from datetime import datetime, date, timedelta, timezone
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from config import COLUMNS, read_env, resolve_date_range, api_keys, load_settings
from collectors import nara, d2b
from exporter import filter_keywords, save_csv
from main import run

SETTINGS = {
    "date_range": {"start_date": None, "end_date": None},
    "page_size": 1,
    "nara": {"display_name": "나라장터", "industry_code": "1169", "exclude_closed": "Y"},
    "d2b": {"display_name": "조달청"},
    "include_keywords": [],
    "exclude_keywords": [],
    "keyword_search_columns": ["공고명"],
}


class CollectorTests(unittest.TestCase):
    def test_date_kst_and_override(self):
        # UTC Oct 16 16:00 is Oct 17 01:00 KST; default Oct 16~17.
        now = datetime(2026, 10, 16, 16, 0, tzinfo=timezone.utc)
        self.assertEqual(resolve_date_range(SETTINGS, now=now),
                         (date(2026, 10, 16), date(2026, 10, 17)))
        self.assertEqual(resolve_date_range(SETTINGS, "2026-10-12", "2026-10-14"),
                         (date(2026, 10, 12), date(2026, 10, 14)))
        custom = {**SETTINGS, "date_range": {"start_date": "2026-10-12", "end_date": "2026-10-14"}}
        self.assertEqual(resolve_date_range(custom, now=now),
                         (date(2026, 10, 12), date(2026, 10, 14)))

    def test_invalid_range_validation(self):
        for first, last in [("2026-10-15", None), (None, "2026-10-16"),
                            ("2026-10-18", "2026-10-16"), ("2026-02-30", "2026-03-01")]:
            with self.subTest(start=first, end=last):
                custom = {**SETTINGS, "date_range": {"start_date": first, "end_date": last}}
                with self.assertRaises(ValueError):
                    resolve_date_range(custom)
        with self.assertRaises(ValueError):
            resolve_date_range(SETTINGS, explicit_start="2026-10-12")

    def test_read_env_does_not_need_dotenv_dependency(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / ".env"
            path.write_text("# comment\nNARA_API_KEY='encoded%2Btoken'\nD2B_API_KEY=abc\n", encoding="utf-8")
            values = read_env(path)
            with patch.dict("os.environ", {}, clear=True):
                self.assertEqual(api_keys(values), {"nara": "encoded+token", "d2b": "abc"})

    def test_nara_params_pagination_and_mapping(self):
        params_seen = []
        def fake_fetch(url, params):
            self.assertEqual(url, nara.API_URL)
            params_seen.append(dict(params))
            number = params["pageNo"]
            obj = {"response": {
                "header": {"resultCode": "00", "resultMsg": "정상"},
                "body": {
                    "totalCount": 2,
                    "items": {"item": [{
                        "bidNtceNo": f"R26BK{number}", "bidNtceOrd": "000",
                        "bidNtceDt": "2026-10-15 12:00:00",
                        "bidNtceNm": "학술 연구 용역", "presmptPrce": "1000000",
                        "ntceInsttNm": "A", "dminsttNm": "B",
                        "bidNtceDtlUrl": "https://example.org/detail",
                    }]},
                },
            }}
            return json.dumps(obj).encode("utf-8")

        with patch.object(nara, "fetch_bytes", side_effect=fake_fetch):
            records = nara.collect("test-token", date(2026, 10, 12), date(2026, 10, 14), SETTINGS)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["공고번호"], "R26BK1-000")
        self.assertEqual(records[1]["상세보기 주소"], "https://example.org/detail")
        self.assertEqual(params_seen[0]["inqryBgnDt"], "202610120000")
        self.assertEqual(params_seen[0]["inqryEndDt"], "202610142359")
        self.assertEqual(params_seen[0]["indstrytyCd"], "1169")
        self.assertEqual(params_seen[0]["bidClseExcpYn"], "Y")
        self.assertEqual(params_seen[1]["pageNo"], 2)

    def test_d2b_params_and_unavailable_fields(self):
        observed = []
        def fake_fetch(url, params):
            observed.append(dict(params))
            return b'''<response><header><resultCode>00</resultCode></header>
            <body><items><item><pblancNo>D200</pblancNo><pblancOdr>001</pblancOdr>
            <bidNm>Project X</bidNm><pblancDate>20261015</pblancDate>
            <ornt>Agency</ornt><pblancSe>Notice</pblancSe>
            <bsicExpt>500000</bsicExpt><bidPartcptRegistClosDt>20261019</bidPartcptRegistClosDt>
            </item></items><totalCount>1</totalCount></body></response>'''
        with patch.object(d2b, "fetch_bytes", side_effect=fake_fetch):
            records = d2b.collect("test-token", date(2026, 10, 12), date(2026, 10, 14), SETTINGS)
        self.assertEqual(observed[0]["anmtDateBegin"], "20261012")
        self.assertEqual(observed[0]["anmtDateEnd"], "20261014")
        self.assertEqual(records[0]["공고번호"], "D200-001")
        self.assertEqual(records[0]["추정가격(원)"], "")
        self.assertEqual(records[0]["수요기관"], "")
        self.assertEqual(records[0]["상세보기 주소"], d2b.D2B_LIST_URL)

    def test_keyword_filter_and_csv_header(self):
        rows = [
            {"공고명": "공공 연구", "공고위치": "나라장터"},
            {"공고명": "청소 용역", "공고위치": "조달청"},
        ]
        kept, removed = filter_keywords(rows, ["청소"], ["공고명"])
        self.assertEqual((len(kept), removed), (1, 1))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "test.csv"
            save_csv(path, kept)
            with path.open(encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                self.assertEqual(reader.fieldnames, COLUMNS)
                self.assertEqual(list(reader)[0]["공고명"], "공공 연구")
            save_csv(path, [])
            with path.open(encoding="utf-8-sig", newline="") as f:
                self.assertEqual(len(list(csv.reader(f))), 1)

    def test_include_keywords_or_and_exclude_precedence(self):
        rows = [
            {"공고명": "AI 연구 개발"},
            {"공고명": "시장 조사 사업"},
            {"공고명": "청소 용역"},
            {"공고명": "연구 청소"},
        ]
        kept, removed = filter_keywords(
            rows, ["청소"], ["공고명"], include_keywords=["연구", "조사"]
        )
        self.assertEqual([row["공고명"] for row in kept], ["AI 연구 개발", "시장 조사 사업"])
        self.assertEqual(removed, 2)
        kept, removed = filter_keywords(rows, [], ["공고명"], include_keywords=[])
        self.assertEqual((len(kept), removed), (4, 0))
        kept, removed = filter_keywords(rows, [], ["공고명"], include_keywords=["  ai  "])
        self.assertEqual([row["공고명"] for row in kept], ["AI 연구 개발"])

    def test_include_keywords_settings_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "settings.json"
            invalid = {**SETTINGS, "include_keywords": "연구"}
            path.write_text(json.dumps(invalid), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "include_keywords"):
                load_settings(path)

    def test_main_calls_both_collectors_and_writes_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            settings = dict(SETTINGS)
            settings["include_keywords"] = ["학술", "과제"]
            settings["exclude_keywords"] = ["제외"]
            settings["output_directory"] = "output"
            fake_nara = [{"공고명": "학술 연구", "공고위치": "나라장터"},
                         {"공고명": "제외할 공고", "공고위치": "나라장터"}]
            fake_d2b = [{"공고명": "과제 공고", "공고위치": "조달청"}]
            with patch("main.BASE_DIR", tmp_dir), \
                 patch("main.load_settings", return_value=settings), \
                 patch("main.api_keys", return_value={"nara": "test1", "d2b": "test2"}), \
                 patch("main.nara.collect", return_value=fake_nara) as nc, \
                 patch("main.d2b.collect", return_value=fake_d2b) as dc:
                self.assertEqual(run("2026-10-12", "2026-10-14"), 0)
                nc.assert_called_once_with("test1", date(2026, 10, 12), date(2026, 10, 14), settings)
                dc.assert_called_once_with("test2", date(2026, 10, 12), date(2026, 10, 14), settings)
            nara_file = tmp_dir / "output" / "나라장터_20261012_20261014.csv"
            d2b_file = tmp_dir / "output" / "D2B_20261012_20261014.csv"
            self.assertTrue(nara_file.is_file())
            self.assertTrue(d2b_file.is_file())
            with nara_file.open(encoding="utf-8-sig", newline="") as f:
                self.assertEqual(len(list(csv.DictReader(f))), 1)


if __name__ == "__main__":
    unittest.main()
