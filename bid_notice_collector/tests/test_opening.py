# 파일 역할: 나라장터 개찰결과 조회(기간 분할·응답 변환·검색)와 평가비율 컬럼을 실제 API 호출 없이 검증합니다.
# 마지막 업데이트: 2026-11-11
# 작업자: SMI 장지수

from datetime import date
import json
import unittest
from unittest.mock import patch

from collectors import d2b, nara, nara_opening


def g2b_page(items, total=None):
    return json.dumps({"response": {
        "header": {"resultCode": "00", "resultMsg": "정상"},
        "body": {"totalCount": len(items) if total is None else total, "items": items},
    }}).encode("utf-8")


class EvaluationRatioTests(unittest.TestCase):
    def test_nara_ratio_from_api_fields(self):
        row = nara.convert({"bidNtceNo": "R1", "bidNtceOrd": "000",
                            "techAbltEvlRt": "90", "bidPrceEvlRt": "10"}, "나라장터")
        self.assertEqual(row["평가비율(기술:가격)"], "90:10")
        self.assertEqual(nara.convert({"bidNtceNo": "R2"}, "나라장터")["평가비율(기술:가격)"], "")

    def test_d2b_ratio_points_to_notice_document(self):
        self.assertEqual(d2b.convert({"cntrctMth": "협상에의한계약(전자)"}, "D2B")["평가비율(기술:가격)"], "공고문 확인")
        self.assertEqual(d2b.convert({"cntrctMth": "제한경쟁"}, "D2B")["평가비율(기술:가격)"], "")

    def test_range_error_object_is_reported(self):
        raw = json.dumps({"nkoneps.com.response.ResponseError": {
            "header": {"resultCode": "07", "resultMsg": "입력범위값 초과 에러"}}}).encode("utf-8")
        with self.assertRaisesRegex(RuntimeError, "07 / 입력범위값 초과"):
            nara.parse_page(raw)


class OpeningResultTests(unittest.TestCase):
    def test_date_chunks_respect_seven_day_limit(self):
        chunks = list(nara_opening.date_chunks(date(2026, 11, 1), date(2026, 11, 15)))
        self.assertEqual(chunks, [
            (date(2026, 11, 1), date(2026, 11, 7)),
            (date(2026, 11, 8), date(2026, 11, 14)),
            (date(2026, 11, 15), date(2026, 11, 15)),
        ])
        self.assertEqual(list(nara_opening.date_chunks(date(2026, 11, 3), date(2026, 11, 3))),
                         [(date(2026, 11, 3), date(2026, 11, 3))])

    def test_convert_first_bidder_and_failed_opening(self):
        opened = nara_opening.convert({
            "bidNtceNo": "R26BK1", "bidNtceOrd": "000", "bidNtceNm": "AI 연구",
            "opengDt": "2026-11-10 09:00:00", "prtcptCnum": "4", "progrsDivCdNm": "개찰완료",
            "opengRsltNtcCntnts": "개찰완료",
            "opengCorpInfo": "연구소^1234567890^홍길동^89900000^96.928",
        })
        self.assertEqual(opened["공고번호"], "R26BK1-000")
        self.assertEqual(opened["개찰일시"], "2026-11-10 09:00")
        self.assertEqual((opened["1순위 업체"], opened["투찰금액(원)"], opened["투찰률(%)"]),
                         ("연구소", "89,900,000", "96.928"))
        self.assertEqual(opened["비고"], "")
        self.assertIn("bidPbancNo=R26BK1&bidPbancOrd=000", opened["상세보기 주소"])
        # 협상계약은 투찰률이 0이고, 유찰은 업체 정보가 비어 있습니다.
        self.assertEqual(nara_opening.first_bidder("업체^1^대표^36287900^0"), ("업체", "36,287,900", ""))
        self.assertEqual(nara_opening.first_bidder("업체^1^대표"), ("업체", "", ""))
        failed = nara_opening.convert({"bidNtceNo": "R2", "bidNtceOrd": "000", "progrsDivCdNm": "유찰",
                                       "opengRsltNtcCntnts": "단일응찰", "opengCorpInfo": ""})
        self.assertEqual((failed["1순위 업체"], failed["비고"], failed["참가업체수"]), ("", "단일응찰", "0"))

    def test_matches_or_search(self):
        row = {"공고명": "국방 AI 연구", "공고기관": "국방부", "수요기관": "", "1순위 업체": "연구소"}
        self.assertTrue(nara_opening.matches(row, ""))
        self.assertTrue(nara_opening.matches(row, "없음, ai"))
        self.assertTrue(nara_opening.matches(row, "연구소"))
        self.assertFalse(nara_opening.matches(row, "청소, 경비"))

    def test_collect_splits_range_and_sorts_latest_first(self):
        seen = []

        def fake_fetch(url, params):
            seen.append(dict(params))
            day = params["inqryBgnDt"][:8]
            return g2b_page([{"bidNtceNo": f"N{day}", "bidNtceOrd": "000",
                              "opengDt": f"{day[:4]}-{day[4:6]}-{day[6:]} 10:00:00"}])

        pages = []
        with patch.object(nara_opening, "fetch_bytes", side_effect=fake_fetch):
            rows = nara_opening.collect("key", date(2026, 11, 1), date(2026, 11, 10), "1169",
                                        on_page=lambda r, done, total: pages.append((len(r), done, total)))
        self.assertEqual([(p["inqryBgnDt"], p["inqryEndDt"]) for p in seen],
                         [("202611010000", "202611072359"), ("202611080000", "202611102359")])
        self.assertTrue(all(p["inqryDiv"] == "2" and p["indstrytyCd"] == "1169" for p in seen))
        self.assertEqual([row["공고번호"] for row in rows], ["N20261108-000", "N20261101-000"])
        self.assertEqual(pages, [(1, 1, 2), (1, 2, 2)])

        seen.clear()
        with patch.object(nara_opening, "fetch_bytes", side_effect=fake_fetch):
            nara_opening.collect("key", date(2026, 11, 1), date(2026, 11, 1), None)
        self.assertNotIn("indstrytyCd", seen[0])
        with self.assertRaises(ValueError):
            nara_opening.collect("key", date(2026, 11, 2), date(2026, 11, 1))


if __name__ == "__main__":
    unittest.main()
