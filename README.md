# jodal_v2

조달입찰 분석 시스템(Streamlit)과 HWP 텍스트 추출 도구 모음입니다. 아래 절차를 따르면 로컬/Streamlit Cloud 어디서든 동일한 구성이 가능합니다.

## 1. 환경 준비
1. Python 3.11 설치 (Streamlit Cloud는 `runtime.txt`로 3.11.9를 명시합니다.)
2. 의존성 설치
   ```bash
   pip install -r requirements.txt
   ```
   - 문서 변환에 실제로 사용되는 라이브러리만 남겼기 때문에(예: PyPDF2, reportlab, Pillow, olefile) `pip install` 시간이 이전 대비 크게 단축되었습니다.
3. (Streamlit Cloud) `packages.txt`를 **완전히 제거**했으므로 apt 단계 자체가 실행되지 않습니다. 이제 "Your app is in the oven" 상태가 apt 설치 때문에 길어질 일이 없습니다.

> 📁 **폰트 직접 번들링도 가능** — `assets/fonts/NanumGothic.ttf` 또는 `assets/fonts/NanumGothic-Regular.ttf`를 추가하면 ReportLab이 우선 사용합니다. 폰트를 추가하지 않아도 새로 추가된 CID 폰트(Adobe-Korea1 기반)로 한국어가 정상 표기됩니다.

## 2. 앱 실행
```bash
streamlit run app.py
```
- 사이드바에서 Excel(`filtered` 시트) 업로드 후 필터/차트를 확인합니다.
- "내고객 분석하기" 메뉴에서는 HWP/HWPX/문서 업로드 → GPT 보고서/챗봇까지 포함됩니다.

## 3. HWP 추출 품질 점검
새로 추가된 CLI로 현장에서 바로 검증할 수 있습니다.
```bash
python scripts/hwp_inspector.py 공고서.hwp \
  --keywords 계약 번호 스쿨넷 추정금액 이중화 Aggregation \
  --output 공고서.txt --json --fail-missing
```
- `--json` : 요약을 JSON으로 확인해 배포 파이프라인에서 기계적으로 검증할 수 있습니다.
- `--fail-missing` : 필수 키워드가 모두 나오지 않으면 종료 코드 2로 실패 처리합니다.
- 추출 텍스트는 UTF-8로 저장되며, `hwp_utils.collect_text_statistics`를 통해 한글 글자 수·키워드 유무를 바로 확인합니다.

## 4. 테스트
```bash
pytest
```
- OLE BodyText 파서를 이용한 단위 테스트와(필수 키워드 `계약`, `번호`) 샘플 HWP 파일이 있을 경우의 통합 테스트를 모두 커버합니다.

## 5. Troubleshooting
- Streamlit Cloud에서 "Your app is in the oven"이 오래 지속된다면, `pip install` 단계에서 설치되는 패키지 수가 크게 줄어든 최신 `requirements.txt`가 반영됐는지와, 저장소 루트에 **`packages.txt`가 없는지**(apt 단계 완전 생략)를 먼저 확인하세요.
- HWP 추출이 의심될 때는 `scripts/hwp_inspector.py`로 동일 파일을 로컬에서 확인한 뒤 `tests/test_hwp_conversion.py`의 skip 조건을 만족하도록 샘플 파일을 배치하면 됩니다.

## 6. PR 체크리스트 한 번에 실행하기
CI 대신 로컬에서 PR 제출 전 동일한 절차를 돌려보고 싶다면 아래 스크립트를 실행하세요.

```bash
./scripts/pr_check.sh
```

- `python -m compileall`, `pytest`를 순차 실행합니다.
- 저장소 루트에 샘플 HWP가 있으면 `scripts/hwp_inspector.py --fail-missing`까지 자동 수행해 키워드 검증도 끝냅니다.
