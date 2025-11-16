#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import hwp_utils

DEFAULT_KEYWORDS = ("계약", "번호", "스쿨넷", "추정금액", "이중화", "Aggregation")


def _read_bytes(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except FileNotFoundError:  # pragma: no cover - surfaced to CLI
        print(f"[!] 파일을 찾을 수 없습니다: {path}", file=sys.stderr)
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HWP 텍스트 추출 상태를 CLI에서 점검합니다.")
    parser.add_argument("file", type=Path, help="검증할 HWP 파일 경로")
    parser.add_argument(
        "--keywords",
        nargs="*",
        default=list(DEFAULT_KEYWORDS),
        help="필수로 포함되어야 하는 키워드 목록",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="추출한 텍스트를 저장할 .txt 경로",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="요약 정보를 JSON으로 출력",
    )
    parser.add_argument(
        "--fail-missing",
        action="store_true",
        help="키워드가 모두 발견되지 않으면 종료 코드를 2로 반환",
    )
    args = parser.parse_args(argv)

    file_bytes = _read_bytes(args.file)
    text, status = hwp_utils.convert_hwp_local_to_text(file_bytes)
    if not text:
        print(f"[!] 텍스트 추출 실패 — {status}", file=sys.stderr)
        return 1

    keywords = args.keywords or []
    stats = hwp_utils.collect_text_statistics(text, keywords)
    stats["status"] = status
    stats["file"] = str(args.file)

    if args.output:
        args.output.write_text(text, encoding="utf-8")
        stats["output"] = str(args.output)

    if args.json:
        print(json.dumps(stats, ensure_ascii=False, indent=2))
    else:
        print(f"파일: {args.file}")
        print(f"상태: {status}")
        print(f"총 글자 수: {stats['length']:,}")
        print(f"한글 글자 수: {stats['korean_characters']:,}")
        print(f"한글 포함 여부: {'예' if stats['contains_korean'] else '아니오'}")
        print(f"포함된 키워드: {', '.join(stats['found_keywords']) or '-'}")
        print(f"누락된 키워드: {', '.join(stats['missing_keywords']) or '-'}")
        if args.output:
            print(f"텍스트 저장 경로: {args.output}")

    if args.fail_missing and stats["missing_keywords"]:
        return 2

    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry
    raise SystemExit(main())
