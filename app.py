# =============================
# ✅ 서비스구분 컬럼 생성
# =============================
classification_rules = {
    '통신': '전용회선', '회선': '전용회선', '전송': '전용회선', '망': '전용회선',
    '인터넷': '인터넷', '콜': '전화', '문자': 'SMS', '고객센터': '전화',
    'C그룹': '전화', '전용회선': '전용회선', '단말기': 'NSI',
    '스마트기기': 'NSI', '스마트 기기': 'NSI', 'LTE': '무선', '5G': '무선', '무선': '무선',
    '대표번호': '전화', 'IDC': 'IDC', 'CDN': 'IDC', '스쿨넷': '전용회선',
    '클라우드': 'IDC', '와이파이': '인터넷', '백업': 'IDC', 'IoT': '무선',
    '메시지': '문자', '메세지': '문자', 'Contact': '전화', 'cloud': 'IDC',
    '디도스': '보안', '보안': '보안', '관제': '보안', '재난': '보안',
    '유지보수': '유지보수',
    '안심알리미': 'NSI',
    '안심 알리미': 'NSI',
    '전기공사': '유지보수',
    '스토리지': 'NSI',
    '음식물': 'NSI',
    '소액': 'NSI',
    '통화': '전화',
    '위협': '전화',
    '전화기': '전화',
    '모바일행정전화': '전화',
    '휴대폰': '무선',
    'LED': 'NSI',
    '조명': 'NSI',
    '태블릿': 'NSI',
    '네트워크': '전용회선',
    '스마트단말': 'NSI',
    '운영대행': '유지보수',
    '모바일': '무선',
    'AI': 'AI',
    '인공지능': 'AI',
    '빅데이터': 'AI',
    '구내전화': '전화', 'IPTV': '미디어', 'CCTV': 'CCTV'
}


def add_service_category(df: pd.DataFrame) -> pd.DataFrame:
    if "서비스구분" in df.columns:
        df = df.copy()
        _ = df.pop("서비스구분")

    df["서비스구분"] = "미분류"  # 맨 뒤에 생성

    if "입찰공고명" not in df.columns:
        return df

    rule_items = sorted(classification_rules.items(), key=lambda x: len(x[0]), reverse=True)

    def classify_title(title: str) -> str:
        t = "" if pd.isna(title) else str(title)
        tl = t.lower()
        for k, label in rule_items:
            if (k in t) or (k.lower() in tl):
                return label
        return "미분류"

    df["서비스구분"] = df["입찰공고명"].apply(classify_title)
    return df


# =============================
# 첨부 링크 매트릭스
# =============================
CSS_COMPACT = """
<style>
.attch-wrap { display:flex; flex-direction:column; gap:14px; background:#eef6ff; padding:8px; border-radius:12px; }
.attch-card { border:1px solid #cfe1ff; border-radius:12px; padding:12px 14px; background:#f4f9ff; }
.attch-title { font-weight:700; margin-bottom:8px; font-size:13px; line-height:1.4; word-break:break-word; color:#0b2e5b; }
.attch-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:10px; }
.attch-box { border:1px solid #cfe1ff; border-radius:10px; overflow:hidden; background:#ffffff; }
.attch-box-header { background:#0d6efd; color:#fff; font-weight:700; font-size:11px; padding:6px 8px; display:flex; align-items:center; justify-content:space-between; }
.badge { background:rgba(255,255,255,0.2); color:#fff; padding:0 6px; border-radius:999px; font-size:10px; }
.attch-box-body { padding:8px; font-size:12px; line-height:1.45; word-break:break-word; color:#0b2447; }
.attch-box-body a { color:#0b5ed7; text-decoration:none; }
.attch-box-body a:hover { text-decoration:underline; }
.attch-box-body details summary { cursor:pointer; font-weight:600; list-style:none; outline:none; color:#0b2447; }
.attch-box-body details summary::-webkit-details-marker { display:none; }
.attch-box-body details summary:after { content:"▼"; font-size:10px; margin-left:6px; color:#0b2447; }
</style>
"""


def _is_url(val: str) -> bool:
    s = str(val).strip()
    return s.startswith("http://") or s.startswith("https://")


def _filename_from_url(url: str) -> str:
    try:
        path = urlparse(url).path
        if not path:
            return url
        return unquote(path.split("/")[-1]) or url
    except Exception:
        return url


def build_attachment_matrix(df_like: pd.DataFrame, title_col: str) -> pd.DataFrame:
    if title_col not in df_like.columns:
        return pd.DataFrame(columns=[title_col, "본공고링크", "제안요청서", "공고서", "과업지시서", "규격서", "기타"])
    buckets = {}

    def add_link(title, category, name, url):
        if title not in buckets:
            buckets[title] = {k: {} for k in ["본공고링크", "제안요청서", "공고서", "과업지시서", "규격서", "기타"]}
        if url not in buckets[title][category]:
            buckets[title][category][url] = name

    n_cols = df_like.shape[1]
    for _, row in df_like.iterrows():
        title = str(row.get(title_col, ""))
        if not title:
            continue
        for j in range(1, n_cols):
            url_col = df_like.columns[j]
            name_col = df_like.columns[j - 1]
            url_val = row.get(url_col, None)
            name_val = row.get(name_col, None)
            if pd.isna(url_val):
                continue
            raw = str(url_val).strip()
            if _is_url(raw):
                urls = [raw]
            else:
                toks = [u.strip() for u in raw.replace("\n", ";").split(";")]
                urls = [u for u in toks if _is_url(u)]
                if not urls:
                    continue
            name_base = "" if pd.isna(name_val) else str(name_val).strip()
            name_tokens = [n.strip() for n in (name_base.replace("\n", ";") if name_base else "").split(";")]
            for k, u in enumerate(urls):
                disp_name = name_tokens[k] if k < len(name_tokens) and name_tokens[k] else (name_base or _filename_from_url(u))
                low = (disp_name or "").lower() + " " + _filename_from_url(u).lower()

                if ("제안요청서" in low) or ("rfp" in low):
                    add_link(title, "제안요청서", disp_name, u)
                elif ("공고서" in low) or ("공고문" in low):
                    add_link(title, "공고서", disp_name, u)
                elif "과업지시서" in low:
                    add_link(title, "과업지시서", disp_name, u)
                elif ("규격서" in low) or ("spec" in low):
                    add_link(title, "규격서", disp_name, u)
                else:
                    add_link(title, "기타", disp_name, u)

    def join_html(d):
        if not d:
            return ""
        return " | ".join([f"<a href='{url}' target='_blank' rel='nofollow noopener'>{name}</a>" for url, name in d.items()])

    rows = []
    for title, catmap in buckets.items():
        rows.append(
            {
                title_col: title,
                "본공고링크": join_html(catmap["본공고링크"]),
                "제안요청서": join_html(catmap["제안요청서"]),
                "공고서": join_html(catmap["공고서"]),
                "과업지시서": join_html(catmap["과업지시서"]),
                "규격서": join_html(catmap["규격서"]),
                "기타": join_html(catmap["기타"]),
            }
        )
    return pd.DataFrame(rows).sort_values(by=[title_col]).reset_index(drop=True)


def render_attachment_cards_html(df_links: pd.DataFrame, title_col: str) -> str:
    cat_cols = ["본공고링크", "제안요청서", "공고서", "과업지시서", "규격서", "기타"]
    present_cols = [c for c in cat_cols if c in df_links.columns]
    if title_col not in df_links.columns:
        return "<p>표시할 데이터가 없습니다.</p>"
    html = [CSS_COMPACT, '<div class="attch-wrap">']
    for _, r in df_links.iterrows():
        title = str(r.get(title_col, "") or "")
        html.append('<div class="attch-card">')
        html.append(f'<div class="attch-title">{title}</div>')
        html.append('<div class="attch-grid">')
        for col in present_cols:
            raw = str(r.get(col, "") or "").strip()
            if not raw:
                continue
            parts = [p.strip() for p in raw.split("|") if p.strip()]
            count = len(parts)
            if count <= 3:
                body_html = raw
            else:
                head = " | ".join(parts[:3])
                tail = " | ".join(parts[3:])
                body_html = head + f'<details style="margin-top:6px;"><summary>더보기 ({count-3})</summary>{tail}</details>'
            html.append('<div class="attch-box">')
            html.append(f'<div class="attch-box-header">{col} <span class="badge">{count}</span></div>')
            html.append(f'<div class="attch-box-body">{body_html}</div>')
            html.append('</div>')
        html.append('</div></div>')
    html.append('</div>')
    return "\n".join(html)


# =============================
# 벤더 정규화/색상
# =============================
VENDOR_COLOR_MAP = {
    "엘지유플러스": "#FF1493",
    "케이티": "#FF0000",
    "에스케이브로드밴드": "#FFD700",
    "에스케이텔레콤": "#1E90FF",
}
OTHER_SEQ = ["#2E8B57", "#6B8E23", "#556B2F", "#8B4513", "#A0522D", "#CD853F", "#228B22", "#006400"]


def normalize_vendor(name: str) -> str:
    s = str(name) if pd.notna(name) else ""
    if "엘지유플러스" in s or "LG유플러스" in s or "LG U" in s.upper():
        return "엘지유플러스"
    if s.startswith("케이티") or " KT" in s or s == "KT" or "주식회사 케이티" in s:
        return "케이티"
    if "브로드밴드" in s or "SK브로드밴드" in s:
        return "에스케이브로드밴드"
    if "텔레콤" in s or "SK텔레콤" in s:
        return "에스케이텔레콤"
    return s or "기타"


# =============================
# 로그인 게이트 & 사이드바
# =============================
INFO_BOX = "사번/생년월일은 사내 배포용으로만 사용됩니다."


def login_gate():
    st.title("🔐 로그인")
    emp = st.text_input("사번", value="", placeholder="예: 9999")
    dob = st.text_input("생년월일(YYMMDD)", value="", placeholder="예: 990101", type="password")
    users = _get_auth_users_from_secrets()
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("로그인", type="primary", use_container_width=True):
            ok = False
            if emp == "2855" and dob == "910518":
                ok = True
                st.session_state["role"] = "admin"
            elif any((str(u.get("emp")) == emp and str(u.get("dob")) == dob) for u in users):
                ok = True
                st.session_state["role"] = "user"
            if ok:
                st.session_state["authed"] = True
                st.success("로그인 성공")
                st.rerun()
            else:
                st.error("인증 실패. 사번/생년월일을 확인하세요.")
    with col2:
        st.info(INFO_BOX)


def render_sidebar_base():
    st.sidebar.title("📂 데이터 업로드")

    up = st.sidebar.file_uploader(
        "filtered 시트가 포함된 병합 엑셀 업로드 (.xlsx)",
        type=["xlsx"],
        key="uploaded_file"
    )
    if up is not None:
        st.session_state["uploaded_file_obj"] = up

    st.sidebar.radio("# 📋 메뉴 선택", ["조달입찰결과현황", "내고객 분석하기"], key="menu")

    with st.sidebar.expander("🔑 Gemini API Key", expanded=True):
        if _get_gemini_key_from_secrets():
            st.success("st.secrets에서 Gemini 키를 불러왔습니다. (권장)")
        key_in = st.text_input(
            "사이드바에서 키 입력(선택) — st.secrets가 우선 적용됩니다.",
            type="password",
            placeholder="AIza...."
        )
        if st.button("키 적용", use_container_width=True):
            if key_in and key_in.strip().startswith("AIza"):
                st.session_state["GEMINI_API_KEY"] = key_in.strip()
                st.success("세션에 Gemini 키가 적용되었습니다.")
            else:
                st.warning("유효한 Gemini 키를 입력하세요 (AIza...).")

    if _get_gemini_key():
        st.sidebar.success("Gemini 사용 가능")
    else:
        st.sidebar.warning("Gemini 비활성 — st.secrets.GEMINI_API_KEY 설정 필요")

    if _cloudconvert_supported():
        st.sidebar.success("CloudConvert 사용 가능")
    else:
        st.sidebar.warning("CloudConvert 비활성 — st.secrets.CLOUDCONVERT_API_KEY 설정 필요")
    
    # [삭제됨] Gemini 추가 요구사항 입력창 제거


def render_sidebar_filters(df: pd.DataFrame):
    st.sidebar.markdown("---")
    st.sidebar.subheader("🧰 필터")

    if "서비스구분" in df.columns:
        options = sorted([str(x) for x in df["서비스구분"].dropna().unique()])
        defaults = [x for x in SERVICE_DEFAULT if x in options]
        st.sidebar.multiselect(
            "서비스구분 선택",
            options=options,
            default=defaults,
            key="svc_filter_ms",
        )

    st.sidebar.subheader("🔍 부가 필터")
    st.sidebar.checkbox("(필터)낙찰자선정여부 = 'Y' 만 보기", value=True, key="only_winner")

    if "대표업체" in df.columns:
        company_list = sorted(df["대표업체"].dropna().unique())
        st.sidebar.multiselect("대표업체 필터 (복수 가능)", company_list, key="selected_companies")

    demand_col_sidebar = "수요기관명" if "수요기관명" in df.columns else ("수요기관" if "수요기관" in df.columns else None)
    if demand_col_sidebar:
        org_list = sorted(df[demand_col_sidebar].dropna().unique())
        st.sidebar.multiselect(f"{demand_col_sidebar} 필터 (복수 가능)", org_list, key="selected_orgs")

    st.sidebar.subheader("📆 공고게시일자 필터")
    if "공고게시일자_date" in df.columns:
        df["_tmp_date"] = pd.to_datetime(df["공고게시일자_date"], errors="coerce")
    else:
        df["_tmp_date"] = pd.NaT

    df["_tmp_year"] = df["_tmp_date"].dt.year
    year_list = sorted([int(x) for x in df["_tmp_year"].dropna().unique()])
    st.sidebar.multiselect("연도 선택 (복수 가능)", year_list, default=[], key="selected_years")

    df["_tmp_month"] = df["_tmp_date"].dt.month
    st.sidebar.multiselect("월 선택 (복수 가능)", list(range(1, 13)), default=[], key="selected_months")


# ===== 진입 가드 =====
if not st.session_state.get("authed", False):
    login_gate()
    st.stop()

render_sidebar_base()

# =============================
# 업로드/데이터 로드
# =============================
uploaded_file = st.session_state.get("uploaded_file_obj")
if not uploaded_file:
    st.title("📊 조달입찰 분석 시스템")
    st.caption("좌측 사이드바에서 'filtered' 시트를 포함한 엑셀 파일을 업로드하세요.")
    st.stop()

try:
    df = pd.read_excel(uploaded_file, sheet_name="filtered", engine="openpyxl")
except Exception as e:
    st.error(f"엑셀 로드 실패: {e}")
    st.stop()

df = add_service_category(df)
df_original = df.copy()

render_sidebar_filters(df_original)

# =============================
# 사이드바 필터 값 읽기 & 적용
# =============================
service_selected = st.session_state.get("svc_filter_ms", [])
only_winner = st.session_state.get("only_winner", True)
selected_companies = st.session_state.get("selected_companies", [])
selected_orgs = st.session_state.get("selected_orgs", [])
selected_years = st.session_state.get("selected_years", [])
selected_months = st.session_state.get("selected_months", [])

demand_col_sidebar = "수요기관명" if "수요기관명" in df.columns else ("수요기관" if "수요기관" in df.columns else None)

df_filtered = df.copy()
if "공고게시일자_date" in df_filtered.columns:
    df_filtered["공고게시일자_date"] = pd.to_datetime(df_filtered["공고게시일자_date"], errors="coerce")
else:
    df_filtered["공고게시일자_date"] = pd.NaT

df_filtered["year"] = df_filtered["공고게시일자_date"].dt.year
df_filtered["month"] = df_filtered["공고게시일자_date"].dt.month

if selected_years:
    df_filtered = df_filtered[df_filtered["year"].isin(selected_years)]
if selected_months:
    df_filtered = df_filtered[df_filtered["month"].isin(selected_months)]
if only_winner and "낙찰자선정여부" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["낙찰자선정여부"] == "Y"]
if selected_companies and "대표업체" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["대표업체"].isin(selected_companies)]
if selected_orgs and demand_col_sidebar:
    df_filtered = df_filtered[df_filtered[demand_col_sidebar].isin(selected_orgs)]
if service_selected and "서비스구분" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["서비스구분"].astype(str).isin(service_selected)]


# =============================
# 기본 분석(차트)
# =============================
def render_basic_analysis_charts(base_df: pd.DataFrame):
    def pick_unit(max_val: float):
        if max_val >= 1_0000_0000_0000:
            return ("조원", 1_0000_0000_0000)
        elif max_val >= 100_000_000:
            return ("억원", 100_000_000)
        elif max_val >= 1_000_000:
            return ("백만원", 1_000_000)
        else:
            return ("원", 1)

    def apply_unit(values: pd.Series, mode: str = "자동"):
        unit_map = {"원": ("원", 1), "백만원": ("백만원", 1_000_000), "억원": ("억원", 100_000_000), "조원": ("조원", 1_0000_0000_0000)}
        if mode == "자동":
            u, f = pick_unit(values.max() if len(values) else 0)
            return values / f, u
        else:
            u, f = unit_map.get(mode, ("원", 1))
            return values / f, u

    st.markdown("## 📊 기본 통계 분석")
    st.caption("※ 이하 모든 차트는 **낙찰자선정여부 == 'Y'** 기준입니다.")

    if "낙찰자선정여부" not in base_df.columns:
        st.warning("컬럼 '낙찰자선정여부'를 찾을 수 없습니다.")
        return
    dwin = base_df[base_df["낙찰자선정여부"] == "Y"].copy()
    if dwin.empty:
        st.warning("낙찰(Y) 데이터가 없습니다.")
        return

    for col in ["투찰금액", "배정예산금액", "투찰율"]:
        if col in dwin.columns:
            dwin[col] = pd.to_numeric(dwin[col], errors="coerce")

    if "대표업체" in dwin.columns:
        dwin["대표업체_표시"] = dwin["대표업체"].map(normalize_vendor)
    else:
        dwin["대표업체_표시"] = "기타"

    st.markdown("### 1) 대표업체별 분포")
    unit_choice = st.selectbox("파이차트(투찰금액 합계) 표기 단위", ["자동", "원", "백만원", "억원", "조원"], index=0)
    col_pie1, col_pie2 = st.columns(2)

    with col_pie1:
        if "투찰금액" in dwin.columns:
            sum_by_company = dwin.groupby("대표업체_표시")["투찰금액"].sum().reset_index().sort_values("투찰금액", ascending=False)
            scaled_vals, unit_label = apply_unit(sum_by_company["투찰금액"].fillna(0), unit_choice)
            sum_by_company["표시금액"] = scaled_vals
            fig1 = px.pie(
                sum_by_company,
                names="대표업체_표시",
                values="표시금액",
                title=f"대표업체별 투찰금액 합계 — 단위: {unit_label}",
                color="대표업체_표시",
                color_discrete_map=VENDOR_COLOR_MAP,
                color_discrete_sequence=OTHER_SEQ,
            )
            st.plotly_chart(fig1, use_container_width=True)

    with col_pie2:
        cnt_by_company = dwin["대표업체_표시"].value_counts().reset_index()
        cnt_by_company.columns = ["대표업체_표시", "건수"]
        fig2 = px.pie(
            cnt_by_company,
            names="대표업체_표시",
            values="건수",
            title="대표업체별 낙찰 건수",
            color="대표업체_표시",
            color_discrete_map=VENDOR_COLOR_MAP,
            color_discrete_sequence=OTHER_SEQ,
        )
        st.plotly_chart(fig2, use_container_width=True)


# =============================
# LLM 분석용 텍스트 추출 (Smart Fallback 적용)
# =============================
TEXT_EXTS = {".txt", ".csv", ".md", ".log"}
DIRECT_PDF_EXTS = {".pdf"}
BINARY_EXTS = {".hwp", ".hwpx", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx"}


def extract_text_combo_gemini_first(uploaded_files):
    combined_texts, convert_logs, generated_pdfs = [], [], []

    for idx, f in enumerate(uploaded_files):
        name = f.name
        data = f.read()
        ext = (os.path.splitext(name)[1] or "").lower()

        # 무료 티어 429 방지용 지연
        if idx > 0:
            time.sleep(2.0)
        
        # ✅ 반환값 2개(텍스트, 모델) 받기
        gem_txt, used_model = gemini_try_extract_text_from_file(data, name)
        
        if gem_txt:
            convert_logs.append(f"🤖 {name}: Gemini[{used_model}] 추출 성공 ({len(gem_txt)}자)")
            combined_texts.append(f"\n\n===== [{name} | Gemini-{used_model}] =====\n{gem_txt}\n")
            continue
        else:
            convert_logs.append(f"🤖 {name}: Gemini 추출 실패 → 폴백 진행")

        if ext in {".hwp", ".hwpx"}:
            try:
                txt, fmt = convert_to_text(data, name)
                convert_logs.append(f"📄 {name}: 로컬 {fmt} 텍스트 추출 성공 ({len(txt)} chars)")
                combined_texts.append(f"\n\n===== [{name} | 로컬 {fmt} 추출] =====\n{_redact_secrets(txt)}\n")
                continue
            except Exception as e:
                convert_logs.append(f"📄 {name}: 로컬 HWP/HWPX 추출 실패 ({e}) → 다음 단계")

        if ext in TEXT_EXTS:
            for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
                try:
                    txt = data.decode(enc)
                    break
                except Exception:
                    continue
            else:
                txt = data.decode("utf-8", errors="ignore")

            convert_logs.append(f"🗒️ {name}: 로컬 텍스트 로드 완료")
            combined_texts.append(f"\n\n===== [{name}] =====\n{_redact_secrets(txt)}\n")
            continue

        if ext in DIRECT_PDF_EXTS:
            txt = extract_text_from_pdf_bytes(data)
            convert_logs.append(f"✅ {name}: 로컬 PDF 텍스트 추출 {len(txt)} chars")
            combined_texts.append(f"\n\n===== [{name}] =====\n{_redact_secrets(txt)}\n")
            continue

        if ext in BINARY_EXTS:
            pdf_bytes, dbg = cloudconvert_convert_to_pdf(data, name)
            if pdf_bytes:
                generated_pdfs.append((os.path.splitext(name)[0] + ".pdf", pdf_bytes))
                txt = extract_text_from_pdf_bytes(pdf_bytes)
                convert_logs.append(f"✅ {name} → CloudConvert PDF 성공 ({dbg}), 텍스트 {len(txt)} chars")
                combined_texts.append(f"\n\n===== [{name} → CloudConvert PDF] =====\n{_redact_secrets(txt)}\n")
            else:
                convert_logs.append(f"🛑 {name}: CloudConvert 실패 ({dbg})")
            continue

        convert_logs.append(f"ℹ️ {name}: 미지원 형식(패스)")

    return "\n".join(combined_texts).strip(), convert_logs, generated_pdfs


# =============================
# 메뉴
# =============================
menu_val = st.session_state.get("menu")

if menu_val == "조달입찰결과현황":
    st.title("📑 조달입찰결과현황")
    dl_buf = BytesIO()
    df_filtered.to_excel(dl_buf, index=False, engine="openpyxl")
    dl_buf.seek(0)
    st.download_button(
        label="📥 필터링된 데이터 다운로드 (Excel)",
        data=dl_buf,
        file_name=f"filtered_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.data_editor(df_filtered, use_container_width=True, key="result_editor", height=520)
    with st.expander("📊 기본 통계 분석(차트) 열기", expanded=False):
        render_basic_analysis_charts(df_filtered)

elif menu_val == "내고객 분석하기":
    st.title("🧑‍💼 내고객 분석하기")
    st.info("ℹ️ 이 메뉴는 사이드바 필터와 무관하게 **전체 원본 데이터**를 대상으로 검색합니다.")

    demand_col = None
    for col in ["수요기관명", "수요기관", "기관명"]:
        if col in df_original.columns:
            demand_col = col
            break
    if not demand_col:
        st.error("⚠️ 수요기관 관련 컬럼을 찾을 수 없습니다.")
        st.stop()
    st.success(f"✅ 검색 대상 컬럼: **{demand_col}**")

    customer_input = st.text_input(f"고객사명을 입력하세요 ({demand_col} 기준, 쉼표로 복수 입력 가능)", help="예) 조달청, 국방부")

    with st.expander(f"📋 전체 {demand_col} 목록 보기 (검색 참고용)"):
        unique_orgs = sorted(df_original[demand_col].dropna().unique())
        st.write(f"총 {len(unique_orgs)}개 기관")
        search_org = st.text_input("기관명 검색", key="search_org_in_my")
        view_orgs = [o for o in unique_orgs if (search_org in str(o))] if search_org else unique_orgs
        st.write(view_orgs[:120])

    if customer_input:
        customers = [c.strip() for c in customer_input.split(",") if c.strip()]
        if customers:
            result = df_original[df_original[demand_col].isin(customers)]
            st.subheader(f"📊 검색 결과: {len(result)}건")
            if not result.empty:
                rb = BytesIO()
                result.to_excel(rb, index=False, engine="openpyxl")
                rb.seek(0)
                st.download_button(
                    label="📥 결과 데이터 다운로드 (Excel)",
                    data=rb,
                    file_name=f"{'_'.join(customers)}_이력_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.data_editor(result, use_container_width=True, key="customer_editor", height=520)

                # ===== 첨부 링크 매트릭스 =====
                st.markdown("---")
                st.subheader("🔗 입찰공고명 기준으로 URL을 분류합니다.")
                st.caption("(본공고링크/제안요청서/공고서/과업지시서/규격서/기타, URL 중복 제거)")
                title_col = next((c for c in ["입찰공고명", "공고명"] if c in result.columns), None)
                if title_col:
                    attach_df = build_attachment_matrix(result, title_col)
                    if not attach_df.empty:
                        use_compact = st.toggle("🔀 그룹형(Compact) 보기", value=True)
                        if use_compact:
                            st.markdown(render_attachment_cards_html(attach_df, title_col), unsafe_allow_html=True)
                        else:
                            st.dataframe(
                                attach_df.applymap(
                                    lambda x: '' if pd.isna(x) else re.sub(r"<[^>]+>", "", str(x))
                                )
                            )

                # ===== Gemini 분석 =====
                st.markdown("---")
                st.subheader("🤖 Gemini 분석")
                # [삭제됨] 캡션 삭제 요청 반영됨

                src_files = st.file_uploader(
                    "분석할 파일 업로드 (여러 개 가능)",
                    type=["pdf", "hwp", "hwpx", "doc", "docx", "ppt", "pptx", "xls", "xlsx", "txt", "csv", "md", "log"],
                    accept_multiple_files=True,
                    key="src_files_uploader",
                )

                if st.button("🧠 Gemini 분석 보고서 생성", type="primary", use_container_width=True):
                    if not src_files:
                        st.warning("먼저 분석할 파일을 업로드하세요.")
                    else:
                        with st.spinner("Gemini가 업로드된 자료로 보고서를 작성 중..."):
                            combined_text, logs, generated_pdfs = extract_text_combo_gemini_first(src_files)

                            st.session_state["gpt_convert_logs"] = logs

                            if not combined_text.strip():
                                st.error("업로드된 파일에서 텍스트를 추출하지 못했습니다.")
                            else:
                                prompt = f"""
다음은 조달/입찰 관련 문서들의 텍스트입니다.
핵심 요구사항, 기술/가격 평가 비율, 계약조건, 월과 일을 포함한 정확한 일정(입찰 마감/계약기간),
공동수급/하도급/긴급공고 여부, 주요 장비/스펙/구간,
배정예산/추정가격/예가 등을 표와 불릿으로 요약하세요.

[문서 통합 텍스트]
{combined_text[:180000]}
""".strip()
                                try:
                                    # ✅ 반환값 2개 받기
                                    report, used_model = call_gemini(
                                        [
                                            {"role": "system", "content": "당신은 SK브로드밴드 망설계/조달 제안 컨설턴트입니다."},
                                            {"role": "user", "content": prompt},
                                        ],
                                        model="gemini-2.0-flash",
                                        max_tokens=2000,
                                        temperature=0.4,
                                    )

                                    st.session_state["gpt_report_md"] = report
                                    st.session_state["generated_src_pdfs"] = generated_pdfs

                                    # ✅ 화면에 사용된 모델 표시
                                    st.success(f"보고서 생성이 완료되었습니다. (사용된 모델: **{used_model}**)")

                                except Exception as e:
                                    st.error(f"보고서 생성 중 오류: {e}")

                convert_logs_ss = st.session_state.get("gpt_convert_logs", [])
                if convert_logs_ss:
                    st.write("### 변환/추출 로그")
                    for line in convert_logs_ss:
                        st.write("- " + line)

                report_md = st.session_state.get("gpt_report_md")
                generated_pdfs = st.session_state.get("generated_src_pdfs", [])

                if report_md:
                    st.markdown("### 📝 Gemini 분석 보고서")
                    st.markdown(report_md)

                    base_fname = f"{'_'.join(customers)}_Gemini분석_{datetime.now().strftime('%Y%m%d_%H%M')}"

                    st.download_button(
                        "📥 보고서 다운로드 (.md)",
                        data=report_md.encode("utf-8"),
                        file_name=f"{base_fname}.md",
                        mime="text/markdown",
                        use_container_width=True
                    )

                    pdf_bytes, dbg = markdown_to_pdf_korean(report_md, title="Gemini 분석 보고서")
                    if pdf_bytes:
                        st.download_button(
                            "📥 보고서 다운로드 (.pdf)",
                            data=pdf_bytes,
                            file_name=f"{base_fname}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.caption(f"PDF 생성 상태: {dbg}")

                    if generated_pdfs:
                        st.markdown("---")
                        st.markdown("### 🗂️ CloudConvert로 변환된 PDF 내려받기")
                        for i, (fname, pbytes) in enumerate(generated_pdfs):
                            st.download_button(
                                label=f"📥 {fname}",
                                data=pbytes,
                                file_name=fname,
                                mime="application/pdf",
                                key=f"dl_ccpdf_{i}",
                                use_container_width=True,
                            )

                # ===== 컨텍스트 챗봇 =====
                st.markdown("---")
                st.subheader("💬 보고서/테이블 참조 챗봇")
                question = st.chat_input("질문을 입력하세요")
                if question:
                    st.session_state.setdefault("chat_messages", [])
                    st.session_state["chat_messages"].append({"role": "user", "content": question})

                    ctx_df = result.head(200).copy()
                    df_sample_csv = ctx_df.to_csv(index=False)[:20000]
                    report_ctx = st.session_state.get("gpt_report_md") or "(아직 보고서 없음)"

                    q_prompt = f"""
[요약 보고서]
{report_ctx}

[표 데이터 일부 CSV]
{df_sample_csv}

질문: {question}
컨텍스트에 근거해 한국어로 간결하게 답하세요. 표/불릿 적극 활용.
""".strip()

                    try:
                        # ✅ 반환값 2개 받기
                        ans, used_model = call_gemini(
                            [
                                {"role": "system", "content": "당신은 조달/통신 제안 분석 챗봇입니다. 컨텍스트 기반으로만 답하세요."},
                                {"role": "user", "content": q_prompt},
                            ],
                            model="gemini-2.0-flash",
                            max_tokens=1200,
                            temperature=0.2,
                        )
                        # ✅ 답변 끝에 모델명 붙여주기
                        final_ans = f"{ans}\n\n_(Generated by **{used_model}**)_"
                        st.session_state["chat_messages"].append({"role": "assistant", "content": final_ans})
                    except Exception as e:
                        st.session_state["chat_messages"].append({"role": "assistant", "content": f"오류: {e}"})

                for m in st.session_state.get("chat_messages", []):
                    st.chat_message("user" if m["role"] == "user" else "assistant").markdown(m["content"])

# === EOF ===
