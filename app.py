# ... (기존 사이드바 코드 아래에 추가) ...
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🛠️ 디버깅 도구")
    
    if st.sidebar.button("내 API로 쓸 수 있는 모델 확인하기"):
        try:
            import google.generativeai as genai
            
            # 1. 키 가져오기 (입력값 우선 -> 없으면 Secrets)
            chk_key = st.session_state.get("user_input_gemini_key", "").strip()
            if not chk_key:
                chk_key = _get_gemini_key_from_secrets()
                
            if not chk_key:
                st.sidebar.error("API 키가 없습니다.")
            else:
                # 2. 모델 조회
                genai.configure(api_key=chk_key)
                models = genai.list_models()
                
                valid_models = []
                for m in models:
                    if 'generateContent' in m.supported_generation_methods:
                        valid_models.append(m.name)
                
                # 3. 결과 출력
                st.sidebar.success("조회 성공!")
                st.sidebar.code("\n".join(valid_models))
                
        except Exception as e:
            st.sidebar.error(f"조회 실패: {e}")
