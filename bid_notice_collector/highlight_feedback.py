# 파일 역할: highlight_feedback.json의 강조·제외 판정 기록이 현재 설정에서도 그대로 맞는지 검사합니다.
# 마지막 업데이트: 2026-11-11
# 작업자: SMI 장지수

"""Run `python highlight_feedback.py` to list recorded cases that the current settings get wrong."""
import json
import sys

from config import BASE_DIR, load_settings
from exporter import mark_highlight_suppressed
from highlighter import highlight_reasons

FEEDBACK_PATH = BASE_DIR / "highlight_feedback.json"


def load_cases(path=FEEDBACK_PATH):
    return json.loads(path.read_text(encoding="utf-8"))["cases"]


def highlight_result(title, issuer, settings):
    """Return the highlight reasons the app would show for this title (empty = not highlighted)."""
    # 화면·Excel과 같은 순서로 제외 키워드 억제를 먼저 보고, 그다음 강조 규칙을 적용합니다.
    record = mark_highlight_suppressed(
        [{"공고명": title}], settings.get("exclude_keywords", []), ["공고명"],
        settings.get("exclude_exceptions", []),
    )[0]
    if record["_highlight_suppressed"]:
        return []
    return highlight_reasons(title, settings.get("highlight_keywords", []),
                             settings.get("highlight_rules", {}), issuer)


def failures(settings=None, cases=None):
    """Return (case, reasons) pairs whose current result differs from the recorded decision."""
    settings = settings or load_settings()
    wrong = []
    for case in cases if cases is not None else load_cases():
        reasons = highlight_result(case["title"], case.get("issuer"), settings)
        if bool(reasons) != (case["expected"] == "강조"):
            wrong.append((case, reasons))
    return wrong


def main():
    cases = load_cases()
    wrong = failures(cases=cases)
    for case, reasons in wrong:
        actual = "강조(" + "; ".join(reasons) + ")" if reasons else "강조 안 됨"
        print(f"[불일치] {case['title']}\n  기록: {case['expected']} / 현재: {actual}")
    print(f"판정 기록 {len(cases)}건 중 {len(cases) - len(wrong)}건 일치, {len(wrong)}건 불일치")
    return 1 if wrong else 0


if __name__ == "__main__":
    sys.exit(main())
