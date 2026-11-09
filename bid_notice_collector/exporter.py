# 파일 역할: 포함·제외 키워드 필터링, 강조 공고 우선 정렬, Excel 서식 저장 및 CSV 저장을 담당합니다.
# 마지막 업데이트: 2026-11-02
# 작업자: SMI 장지수

"""Keyword filtering, CSV output, and highlighted Excel output."""

import csv
from datetime import datetime
import os
from pathlib import Path
import tempfile
from config import COLUMNS, KST
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from highlighter import highlight_reasons, normalize

EXCEL_COLUMNS = COLUMNS + ["강조 근거"]

# 각 사이트에서 내려받는 '입찰공고목록' Excel의 열 순서·너비·정렬입니다.
ORIGINAL_LAYOUTS = {
    "nara": {
        "sheet": "입찰공고",
        "font": "맑은 고딕",
        "size": 9,
        "title": "입찰공고목록",
        "header_row": 5,
        "title_column": "공고명",
        "columns": [
            ("업무구분", 20, "center"), ("업무여부", 20, "center"), ("구분", 20, "center"),
            ("입찰공고번호", 20, "center"), ("공고명", 30, "left"), ("공고기관", 20, "left"),
            ("수요기관", 20, "left"), ("게시일시(입찰마감일시)", 30, "center"),
        ],
    },
    "d2b": {
        "sheet": "데이터",
        "font": "Arial",
        "size": 10,
        "title": None,
        "header_row": 1,
        "title_column": "입찰건명",
        "columns": [
            ("순번", 5.33, None), ("입찰구분", 9.11, None), ("공고구분", 9.11, None),
            ("공고일자", 11, None), ("G2B공고번호-차수", 12.22, None), ("통합참조번호", 15.11, None),
            ("판단번호", 11.33, None), ("입찰건명", 44.66, None), ("발주기관", 15.22, None),
            ("사전심사 마감일시", 10.22, None), ("입찰참가등록 마감일시", 22.11, None),
            ("입찰서/견적서제출 마감일시", 16.44, None), ("계약방법", 9.11, None),
            ("입찰형태", 9.11, None), ("기초예가", 11.78, None),
        ],
    },
}


def rank_highlighted(records, highlight_keywords=None, highlight_rules=None):
    """Return (record, reasons) pairs with highlighted notices first."""
    ranked = []
    for record in records:
        suppressed = bool(record.get("_highlight_suppressed"))
        reasons = [] if suppressed else highlight_reasons(
            record.get("공고명"), highlight_keywords, highlight_rules,
            record.get("공고기관(발주기관)"),
        )
        ranked.append((record, reasons))
    # 안정 정렬: 강조 공고를 위로 모으되 원래 수집 순서는 유지합니다.
    ranked.sort(key=lambda item: not item[1])
    return ranked


def _replace_workbook(workbook, target):
    # 완성된 임시 파일만 교체하여 저장 실패 시 기존 결과 파일을 보호합니다.
    fd, temporary = tempfile.mkstemp(prefix=".tmp_", suffix=".xlsx", dir=target.parent)
    os.close(fd)
    try:
        workbook.save(temporary)
        os.replace(temporary, target)
    finally:
        workbook.close()
        if os.path.exists(temporary):
            os.unlink(temporary)


def save_original_excel(path, source, records, highlight_keywords=None, highlight_rules=None,
                        created=None):
    """Save records in the source website's own Excel layout, keeping highlight rules."""
    layout = ORIGINAL_LAYOUTS[source]
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = layout["sheet"]
    font = Font(name=layout["font"], size=layout["size"])
    side = Side(style="thin", color="FF000000")
    border = Border(left=side, right=side, top=side, bottom=side)
    header_fill = PatternFill(fill_type="solid", fgColor="FFC0C0C0")
    yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    if layout["title"]:
        sheet["A1"] = layout["title"]
        sheet["A1"].font = Font(name=layout["font"], size=14, bold=True)
        sheet.row_dimensions[1].height = 21
        for row, text in ((2, "생성자 : -"), (3, f"생성일 : {created or datetime.now(KST):%Y/%m/%d}")):
            sheet.cell(row, 1, text).font = font
            sheet.cell(row, 1).alignment = Alignment(vertical="center")
    # 나라장터 원본은 모든 칸을 가운데·줄바꿈, D2B 원본은 아래쪽 정렬입니다.
    wrap = bool(layout["title"])
    vertical = "center" if wrap else "bottom"
    header = layout["header_row"]
    for column, (name, width, horizontal) in enumerate(layout["columns"], 1):
        cell = sheet.cell(header, column, name)
        cell.font, cell.fill, cell.border = font, header_fill, border
        cell.alignment = Alignment(horizontal="center" if wrap else None, vertical=vertical, wrap_text=wrap)
        sheet.column_dimensions[get_column_letter(column)].width = width
    ranked = rank_highlighted(records, highlight_keywords, highlight_rules)
    for number, (record, reasons) in enumerate(ranked, 1):
        values = dict(record.get("_원본") or {}, 순번=str(number))
        row = header + number
        for column, (name, _, horizontal) in enumerate(layout["columns"], 1):
            cell = sheet.cell(row, column, str(values.get(name) or ""))
            # API text must remain literal text, including values starting with '='.
            cell.data_type = "s"
            cell.font, cell.border = font, border
            cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)
            if name == layout["title_column"] and reasons:
                cell.fill = yellow
    last = get_column_letter(len(layout["columns"]))
    sheet.auto_filter.ref = f"A{header}:{last}{header + len(ranked)}"
    _replace_workbook(workbook, target)


def save_excel(path, records, highlight_keywords=None, highlight_rules=None):
    """Save records with highlighted notices first and matching title cells highlighted."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "입찰공고"
    sheet.append(EXCEL_COLUMNS)
    header_fill = PatternFill(fill_type="solid", fgColor="FFD9D9D9")
    yellow = PatternFill(fill_type="solid", fgColor="FFFFFF00")
    cell_border = Border(
        left=Side(style="thin", color="FF000000"),
        right=Side(style="thin", color="FF000000"),
        top=Side(style="thin", color="FF000000"),
        bottom=Side(style="thin", color="FF000000"),
    )
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.border = cell_border
    for record, reasons in rank_highlighted(records, highlight_keywords, highlight_rules):
        row = sheet.max_row + 1
        for column, name in enumerate(EXCEL_COLUMNS, 1):
            value = "; ".join(reasons) if name == "강조 근거" else record.get(name)
            cell = sheet.cell(row, column, "" if value is None else str(value))
            # API text must remain literal text, including values starting with '='.
            cell.data_type = "s"
            cell.border = cell_border
            if name == "공고명" and reasons:
                cell.fill = yellow
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    # 컬럼 추가로 위치가 바뀌어도 같은 열에 너비가 적용되도록 이름으로 찾습니다.
    sheet.column_dimensions[get_column_letter(EXCEL_COLUMNS.index("공고명") + 1)].width = 70
    sheet.column_dimensions[get_column_letter(EXCEL_COLUMNS.index("강조 근거") + 1)].width = 80
    _replace_workbook(workbook, target)


def filter_keywords(records, keywords, search_columns, include_keywords=None):
    """Include if ANY inclusion term matches; exclude matching rows from the result.

    An empty include list allows all records through the inclusion phase.
    Matching ignores case and surrounding whitespace.
    """
    # 포함은 OR, 제외도 OR이며 둘 다 해당하면 제외합니다. 강조 규칙과 달리 내부 공백은 유지합니다.
    includes = [term.strip().casefold() for term in (include_keywords or []) if term.strip()]
    terms = [term.strip().casefold() for term in keywords if term.strip()]
    kept = []
    removed = 0
    for record in records:
        haystacks = [str(record.get(name, "") or "").casefold() for name in search_columns]
        if includes and not any(term in text for term in includes for text in haystacks):
            removed += 1
            continue
        if any(term in text for term in terms for text in haystacks):
            removed += 1
        else:
            kept.append(record)
    return kept, removed


def matches_exception(record, exceptions, search_columns):
    # 띄어쓰기와 관계없이 비교합니다. 예: '법령 정비'와 '법령정비'
    haystacks = [normalize(record.get(name, "")) for name in search_columns]
    return any(
        normalize(term) and normalize(term) in text
        for term in exceptions or []
        for text in haystacks
    )


def mark_highlight_suppressed(records, keywords, search_columns, exceptions=None):
    """Mark rows whose highlight should be suppressed without removing them.

    A row that matches an exception term keeps its highlight even if it matches an exclusion.
    """
    return [
        dict(
            record,
            _highlight_suppressed=matches_exclusion(record, keywords, search_columns)
            and not matches_exception(record, exceptions, search_columns),
        )
        for record in records
    ]


def matches_exclusion(record, keywords, search_columns):
    haystacks = [str(record.get(name, "") or "").casefold() for name in search_columns]
    for entry in keywords:
        terms = [entry] if isinstance(entry, str) else entry
        groups = [
            [alternative.strip().casefold() for alternative in term.split("|") if alternative.strip()]
            for term in terms
            if isinstance(term, str) and term.strip()
        ]
        if groups and all(
            any(alternative in text for alternative in group for text in haystacks)
            for group in groups
        ):
            return True
    return False


def safe_csv_cell(value):
    """Prevent spreadsheet formula execution from untrusted API text fields."""
    cell = str(value) if value is not None else ""
    if cell.lstrip().startswith(("=", "+", "-", "@")) or cell.startswith(("\t", "\r", "\n")):
        return "'" + cell
    return cell


def save_csv(path, records):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    # Atomic replacement: failed runs do not leave partially written CSV files.
    fd, tempfile_name = tempfile.mkstemp(prefix=".tmp_", suffix=".csv", dir=target.parent)
    try:
        with os.fdopen(fd, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=COLUMNS, extrasaction="ignore")
            writer.writeheader()
            for record in records:
                writer.writerow({column: safe_csv_cell(record.get(column, "")) for column in COLUMNS})
        os.replace(tempfile_name, target)
    finally:
        if os.path.exists(tempfile_name):
            os.unlink(tempfile_name)
