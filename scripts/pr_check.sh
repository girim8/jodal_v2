#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

SAMPLE_HWP="${ROOT_DIR}/공고서_국가_일반_국내_유선통신서비스_1458+0036_공동이행_하도급사전승인_장기계속계약_차등점수제_수요기관평가.hwp"
KEYWORDS=(계약 번호 스쿨넷 추정금액 이중화 Aggregation)

echo "[1/3] python -m compileall"
python -m compileall app.py hwp_utils.py scripts/hwp_inspector.py

echo "[2/3] pytest"
pytest

if [[ -f "$SAMPLE_HWP" ]]; then
  echo "[3/3] scripts/hwp_inspector.py on sample"
  python scripts/hwp_inspector.py "$SAMPLE_HWP" --keywords "${KEYWORDS[@]}" --json --fail-missing
else
  echo "[3/3] 샘플 HWP 파일을 찾지 못해 hwp_inspector 단계는 생략합니다." >&2
fi
