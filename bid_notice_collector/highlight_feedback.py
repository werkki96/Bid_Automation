# 파일 역할: highlight_feedback.json의 강조·제외 판정 기록이 현재 설정에서도 그대로 맞는지 검사합니다.
# 마지막 업데이트: 2026-11-11
# 작업자: SMI 장지수

"""Run `python highlight_feedback.py [--reference 과제목록.xlsx]` to list cases the current settings get wrong."""
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


def load_reference(path):
    """Read (사업명, 발주처) pairs from a reference project-list Excel; duplicates removed."""
    from openpyxl import load_workbook
    workbook = load_workbook(path, read_only=True, data_only=True)
    projects, seen = [], set()
    for sheet in workbook.worksheets:
        header = None
        for row in sheet.iter_rows(values_only=True):
            # 시트마다 제목 줄이 있으므로 '번호'로 시작하는 줄을 머리글로 씁니다.
            if row and row[0] == "번호":
                header = list(row)
                continue
            if header and "사업명" in header and row[header.index("사업명")]:
                title = str(row[header.index("사업명")]).strip()
                issuer = str(row[header.index("발주처")] or "").strip() if "발주처" in header else ""
                if title not in seen:
                    seen.add(title)
                    projects.append((title, issuer))
    return projects


def reference_report(path, settings=None):
    """Print how many reference projects the current rules would highlight, and list the misses."""
    settings = settings or load_settings()
    projects = load_reference(path)
    missed = [(t, i) for t, i in projects if not highlight_result(t, i, settings)]
    for title, issuer in missed:
        print(f"[미강조 과제] {issuer} | {title}")
    print(f"기준 과제 {len(projects)}건 중 {len(projects) - len(missed)}건 강조, {len(missed)}건 미강조")


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    cases = load_cases()
    wrong = failures(cases=cases)
    for case, reasons in wrong:
        actual = "강조(" + "; ".join(reasons) + ")" if reasons else "강조 안 됨"
        print(f"[불일치] {case['title']}\n  기록: {case['expected']} / 현재: {actual}")
    print(f"판정 기록 {len(cases)}건 중 {len(cases) - len(wrong)}건 일치, {len(wrong)}건 불일치")
    # 예: python highlight_feedback.py --reference 과제목록.xlsx
    if "--reference" in argv and argv.index("--reference") + 1 < len(argv):
        reference_report(argv[argv.index("--reference") + 1])
    return 1 if wrong else 0


if __name__ == "__main__":
    sys.exit(main())
