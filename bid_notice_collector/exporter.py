"""Keyword filtering and Excel-compatible UTF-8 BOM CSV output."""

import csv
import os
from pathlib import Path
import tempfile
from config import COLUMNS


def filter_keywords(records, keywords, search_columns, include_keywords=None):
    """Include if ANY inclusion term matches; then exclude if ANY exclusion term matches.

    An empty include list allows all records through the inclusion phase.
    Matching ignores case and surrounding whitespace.
    """
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
