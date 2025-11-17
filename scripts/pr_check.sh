#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

SAMPLE_HWP="${ROOT_DIR}/공고서_국가_일반_국내_유선통신서비스_1458+0036_공동이행_하도급사전승인_장기계속계약_차등점수제_수요기관평가.hwp"
KEYWORDS=(계약 번호 금액)

echo "[1/3] python -m compileall"
python -m compileall app.py hwp_utils.py scripts/hwp_inspector.py

echo "[2/3] pytest"
pytest

if [[ -f "$SAMPLE_HWP" ]]; then
  echo "[3/3] scripts/hwp_inspector.py on sample"
  TMP_JSON="$(mktemp)"
  python scripts/hwp_inspector.py "$SAMPLE_HWP" --keywords "${KEYWORDS[@]}" --json | tee "$TMP_JSON"
  python - "$TMP_JSON" <<'PY'
import json
import sys

path = sys.argv[1]
with open(path, "r", encoding="utf-8") as fh:
    stats = json.load(fh)

found = set(stats.get("found_keywords", []))
missing = set(stats.get("missing_keywords", []))
required = {"계약", "번호"}
if not required.issubset(found):
    raise SystemExit("필수 키워드(계약, 번호)가 모두 포함되지 않았습니다.")
if "금액" in found:
    raise SystemExit("'금액' 키워드가 발견되면 안 됩니다.")
if "금액" not in missing:
    raise SystemExit("'금액' 키워드가 누락 상태로 보고되지 않았습니다.")

print("샘플 키워드 상태 검증 완료 — 계약 O / 번호 O / 금액 X")
PY
  rm -f "$TMP_JSON"
else
  echo "[3/3] 샘플 HWP 파일을 찾지 못해 hwp_inspector 단계는 생략합니다." >&2
fi
