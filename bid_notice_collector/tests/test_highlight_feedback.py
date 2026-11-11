# 파일 역할: 사용자가 지정한 강조·제외 판정 기록이 현재 settings.json에서도 모두 맞는지 검증합니다.
# 마지막 업데이트: 2026-11-11
# 작업자: SMI 장지수

import unittest

import highlight_feedback


class HighlightFeedbackTests(unittest.TestCase):
    def test_cases_are_well_formed(self):
        cases = highlight_feedback.load_cases()
        self.assertTrue(cases)
        for case in cases:
            with self.subTest(title=case.get("title")):
                self.assertIn(case["expected"], ("강조", "제외"))
                self.assertTrue(case["title"].strip())

    def test_current_settings_match_every_recorded_decision(self):
        # 규칙을 조정하다 예전에 맞춰 둔 판정이 다시 틀어지면 여기서 잡힙니다.
        wrong = [(case["title"], case["expected"], reasons)
                 for case, reasons in highlight_feedback.failures()]
        self.assertEqual(wrong, [])


if __name__ == "__main__":
    unittest.main()
