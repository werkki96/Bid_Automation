# 파일 역할: 프로그램 진입점. 날짜·인증키를 읽고 두 API 수집, 키워드 필터링, Excel 저장을 순서대로 실행합니다.
# 마지막 업데이트: 2026-10-23
# 작업자: SMI 장지수

"""Entry point: python main.py"""

import argparse
from datetime import datetime
from pathlib import Path
import sys

from config import BASE_DIR, api_keys, load_settings, resolve_date_range
from collectors import nara, d2b
from exporter import filter_keywords, mark_excluded, save_excel


def run(start_date=None, end_date=None):
    settings = load_settings()
    start, end = resolve_date_range(settings, start_date, end_date)
    range_token = f"{start:%Y%m%d}_{end:%Y%m%d}"
    keys = api_keys()
    output_dir = BASE_DIR / settings.get("output_directory", "output")
    # Prevent output path escaping the package directory through settings.
    if not output_dir.resolve().is_relative_to(BASE_DIR.resolve()):
        raise ValueError("output_directory는 프로젝트 내부 디렉터리여야 합니다.")

    print(f"조회 공고일자: {start.isoformat()} ~ {end.isoformat()} (양끝 날짜 포함, KST 기준)")
    print("키워드 제외 대상 컬럼:", ", ".join(settings.get("keyword_search_columns", ["공고명"])))
    print("포함 키워드 (OR):", settings.get("include_keywords", []))
    print("제외 키워드 수:", len(settings.get("exclude_keywords", [])))

    jobs = [
        ("나라장터", "nara", nara.collect, f"나라장터_{range_token}.xlsx"),
        ("D2B", "d2b", d2b.collect, f"D2B_{range_token}.xlsx"),
    ]
    errors = []
    for title, key_name, collector, filename in jobs:
        print(f"\n[{title}] 수집 시작")
        if not keys[key_name]:
            print(f"  실패: {key_name.upper()}_API_KEY를 .env에 입력하세요.")
            errors.append(title)
            continue
        try:
            records = collector(keys[key_name], start, end, settings)
            # 제외 키워드 공고도 저장하되, 결과에서 붉은색으로 표시합니다.
            kept, removed = filter_keywords(
                records,
                [],
                settings.get("keyword_search_columns", ["공고명"]),
                include_keywords=settings.get("include_keywords", []),
            )
            kept = mark_excluded(
                kept,
                settings.get("exclude_keywords", []),
                settings.get("keyword_search_columns", ["공고명"]),
            )
            destination = output_dir / filename
            save_excel(destination, kept, settings.get("highlight_keywords", []),
                       settings.get("highlight_rules", {}))
            print(f"  완료: API {len(records)}건 / 포함조건 제외 {removed}건 / 저장 {len(kept)}건")
            print(f"  결과 파일: {destination}")
        except Exception as exc:
            # One source failing does not stop the other.
            print(f"  실패: {type(exc).__name__}: {exc}")
            errors.append(title)
    # 종료코드: 정상 0, 수집 실패 1. CLI 설정 오류는 아래 진입점에서 2로 처리합니다.
    if errors:
        print(f"\n실패한 API: {', '.join(errors)}. 정상 API의 저장 파일은 유지됩니다.")
        return 1
    print("\n두 API의 수집 및 저장이 완료되었습니다.")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="나라장터 및 D2B 기간별 입찰공고 수집")
    parser.add_argument("--start-date", metavar="YYYY-MM-DD", help="조회 시작일 (종료일과 함께 지정)")
    parser.add_argument("--end-date", metavar="YYYY-MM-DD", help="조회 종료일 (시작일과 함께 지정)")
    parser.add_argument("--date", metavar="YYYY-MM-DD", help="이전 버전 호환: 지정한 날짜 하루만 조회")
    args = parser.parse_args()
    try:
        if args.date is not None:
            if args.start_date is not None or args.end_date is not None:
                parser.error("--date와 --start-date/--end-date를 동시에 사용할 수 없습니다.")
            sys.exit(run(args.date, args.date))
        sys.exit(run(args.start_date, args.end_date))
    except (ValueError, FileNotFoundError, KeyError) as exc:
        print(f"설정 오류: {exc}", file=sys.stderr)
        sys.exit(2)
