import streamlit as st

# CSS 스타일 정의
# Streamlit 앱에 사용자 정의 CSS를 주입하여 디자인을 커스터마이징합니다.
# Streamlit의 내부 DOM 구조에 의존하는 부분이 있으므로, Streamlit 버전 업데이트 시
# 일부 CSS 셀렉터는 변경될 수 있음을 유의해주세요.
custom_css = """
<style>
    /* 기본 폰트 설정 (Google Noto Sans KR 폰트 임포트) */
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700&display=swap');
    
    /* Streamlit 앱의 전체적인 배경 및 폰트 설정 */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Noto Sans KR', sans-serif;
        margin: 0;
        padding: 0;
        background-color: #f0f2f5; /* 전체 앱 배경색 */
        color: #333; /* 기본 텍스트 색상 */
    }

    /* Streamlit 메인 컨테이너 폭 조절 및 그림자, 모서리 둥글게 */
    [data-testid="stAppViewContainer"] {
        max-width: 800px; /* 컨테이너 최대 너비 */
        margin: auto; /* 페이지 중앙 정렬 */
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1); /* 그림자 효과 */
        border-radius: 8px; /* 모서리 둥글게 */
        overflow: hidden; /* 자식 요소가 컨테이너를 벗어나지 않도록 숨김 */
        background-color: #fff; /* 메인 컨테이너 배경색을 흰색으로 설정 */
    }

    /* 상단 디자인 BAR 스타일 */
    /* Streamlit의 st.container를 사용하여 디자인 바를 만듭니다. */
    .top-design-bar {
        background-color: #2c3e50; /* 어두운 파란색/회색 */
        color: #fff;
        padding: 15px 20px;
        text-align: center;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        /* 고정(sticky) 기능은 제거하고 디자인적인 분리만 강조 */
        margin-left: -1rem; /* Streamlit 기본 좌우 마진 상쇄 */
        margin-right: -1rem; /* Streamlit 기본 좌우 마진 상쇄 */
        width: calc(100% + 2rem); /* Streamlit 기본 좌우 마진 상쇄 */
    }
    .top-design-bar h1 {
        color: #fff; /* 제목 텍스트 색상 흰색 */
        margin: 0;
        font-size: 1.5em;
        font-weight: 700;
    }

    /* 하단 디자인 BAR 스타일 */
    .bottom-design-bar {
        background-color: #2ecc71; /* 초록색 */
        color: #fff;
        padding: 15px;
        text-align: center;
        border-bottom-left-radius: 8px;
        border-bottom-right-radius: 8px;
        box-shadow: 0 -2px 5px rgba(0, 0, 0, 0.1);
        /* 고정(sticky) 기능은 제거하고 디자인적인 분리만 강조 */
        margin-left: -1rem; /* Streamlit 기본 좌우 마진 상쇄 */
        margin-right: -1rem; /* Streamlit 기본 좌우 마진 상쇄 */
        width: calc(100% + 2rem); /* Streamlit 기본 좌우 마진 상쇄 */
    }
    
    /* 대본 입력 방식 선택 섹션 */
    .input-method-selection-box {
        background-color: #e0f2f7; /* 연한 파란색 배경 */
        padding: 10px 15px;
        border-radius: 8px;
        margin-bottom: 20px;
        text-align: center;
        display: flex; /* Flexbox를 사용하여 아이콘과 텍스트 정렬 */
        justify-content: center; /* 가로 중앙 정렬 */
        align-items: center; /* 세로 중앙 정렬 */
        gap: 8px; /* 아이콘과 텍스트 사이 간격 */
        font-weight: 700;
        color: #2c3e50; /* 텍스트 색상 */
        font-size: 1.1em; /* 요청하신 크기 조절 (더 작게) */
    }
    .input-method-selection-box .icon {
        font-size: 1.2em; /* 아이콘 크기 조절 */
    }

    /* Streamlit 탭 위젯 커스터마이징 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        border-bottom: 1px solid #ddd;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #fff;
        border-radius: 4px 4px 0px 0px;
        padding: 10px 15px;
        font-weight: 500;
        color: #555;
    }
    /* 활성화된 탭 스타일 */
    .stTabs [aria-selected="true"] { 
        border-bottom: 2px solid #3498db !important; 
        color: #3498db !important; 
        font-weight: 700;
        background-color: #fff;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #f5f5f5;
    }

    /* Streamlit 파일 업로더 커스터마이징 */
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #a0d8f0;
        border-radius: 8px;
        background-color: #f7fcfe;
        padding: 30px 20px;
        height: 180px; /* 높이 고정 */
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        position: relative; /* 자식 요소 절대 위치 지정을 위해 */
    }
    /* Streamlit 파일 업로더의 기본 텍스트와 아이콘 숨기기 */
    [data-testid="stFileUploaderDropzoneInstructions"] > div > span {
        display: none; 
    }
    [data-testid="stFileUploaderDropzoneInstructions"] > div > small {
        display: none; 
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: none; /* 드롭존 지시사항 전체 숨기기 */
    }
    
    /* Browse files 버튼 스타일 조정 */
    [data-testid="stFileUploaderBrowseButton"] > button {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 0.9em;
        font-weight: 600;
        transition: background-color 0.3s ease;
        position: absolute; /* 드롭존 내에서 절대 위치 지정 */
        bottom: 20px;
        right: 20px;
    }
    [data-testid="stFileUploaderBrowseButton"] > button:hover {
        background-color: #2980b9;
    }

    /* Expander (Word 파일 업로드 시 문제가 발생하나요?) */
    .stExpander {
        border: 1px solid #eee;
        border-radius: 8px;
        background-color: #f9f9f9;
        margin-top: 20px;
    }
    .stExpander > div > div > details > summary {
        color: #666;
        font-size: 0.9em;
        padding: 10px 15px;
        outline: none;
    }
    .stExpander > div > div > details > summary:hover {
        background-color: #f0f0f0;
        border-radius: 8px;
    }
    .stExpander > div > div > details > summary::marker {
        content: '';
    }
    .stExpander > div > div > details > summary::before {
        content: '▼';
        font-size: 0.8em;
        margin-right: 5px;
        transition: transform 0.2s;
    }
    .stExpander > div > div > details[open] > summary::before {
        transform: rotate(180deg);
    }
    .stExpander > div > div > details > div {
        padding: 5px 15px 10px;
        border-top: 1px dashed #eee;
        font-size: 0.85em;
        color: #777;
    }

    /* PPT 자동 생성 시작 버튼 */
    .stButton > button {
        background-color: #2ecc71;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 1.2em;
        font-weight: 700;
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        transition: background-color 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #27ae60;
    }

    /* 반응형 디자인 */
    @media (max-width: 768px) {
        [data-testid="stAppViewContainer"] {
            border-radius: 0;
            box-shadow: none;
        }
        .top-design-bar, .bottom-design-bar { /* 변경된 클래스 이름 사용 */
            border-radius: 0;
        }
    }
</style>
"""

# Streamlit 앱에 사용자 정의 CSS 주입
st.markdown(custom_css, unsafe_allow_html=True)

# --- Streamlit 앱 UI 구성 시작 ---

# 상단 디자인 BAR
# st.container를 사용하여 디자인적인 BAR를 만듭니다.
with st.container():
    st.markdown('<div class="top-design-bar">', unsafe_allow_html=True)
    st.markdown("<h1>촬영 대본 PPT 자동 생성 AI (KoSimCSE)</h1>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# 메인 콘텐츠 영역은 Streamlit의 기본 레이아웃을 따르며,
# [data-testid="stAppViewContainer"]에 지정된 배경색으로 흰색 바탕을 유지합니다.

# 대본 입력 방식 선택 섹션 (더 작게, 이모지 반영)
st.markdown('<div class="input-method-selection-box"><span class="icon">📁</span> 대본 입력 방식 선택</div>', unsafe_allow_html=True)

# 탭 메뉴 구성 (st.tabs 위젯 사용)
tab1, tab2 = st.tabs(["Word 파일 업로드", "텍스트 직접 입력"])

with tab1:
    st.write("Word 파일 (.docx)을 업로드해주세요.")

    # 파일 업로더 위젯
    # 기본 라벨은 숨기고 (label_visibility="collapsed"), 커스텀 텍스트를 마크다운으로 삽입
    uploaded_file = st.file_uploader(
        "Upload your DOCX file here", # 이 텍스트는 내부적으로 사용되지만, CSS로 숨김.
        type=["docx"], # 허용되는 파일 형식
        accept_multiple_files=False, # 단일 파일만 허용
        label_visibility="collapsed" # 기본 라벨 숨기기
    )
    
    # 드래그 앤 드롭 영역 내 커스텀 텍스트 및 아이콘 (CSS로 위치 조정)
    # 이 부분은 st.file_uploader의 위에 띄워지는 형태입니다.
    st.markdown("""
        <div style="text-align: center; margin-top: -160px; pointer-events: none; position: relative; z-index: 1;">
            <i class="fas fa-cloud-upload-alt" style="font-size: 3em; color: #3498db; margin-bottom: 5px;"></i>
            <p style="margin:0; font-size: 1.1em; color: #666;">Drag and drop file here</p>
            <p style="margin:0; font-size: 0.85em; color: #888; margin-top: 5px;">Limit 200MB per file • DOCX</p>
        </div>
    """, unsafe_allow_html=True)
    # `pointer-events: none`은 마크다운 오버레이가 파일 업로더 클릭을 방해하지 않도록 합니다.
    # `margin-top`과 `z-index`는 텍스트와 아이콘이 파일 업로더 위에 적절히 표시되도록 조절합니다.

    if uploaded_file is not None:
        st.success(f"파일 '{uploaded_file.name}'이(가) 업로드되었습니다.")
        # 여기에 업로드된 파일을 처리하는 로직을 추가합니다.
        # 예: bytes_data = uploaded_file.getvalue()
        # st.write(bytes_data)

    # 문제 해결 드롭다운 (st.expander 위젯 사용)
    with st.expander("🙁 Word 파일 업로드 시 문제가 발생하나요?"):
        st.write("문제가 발생할 경우 다음 사항을 확인해주세요:")
        st.markdown("- 파일 형식이 `.docx`인지 확인해주세요.")
        st.markdown("- 파일 크기가 200MB를 초과하지 않는지 확인해주세요.")
        st.markdown("- 네트워크 연결이 안정적인지 확인해주세요.")
        st.markdown("- 다른 이름으로 저장 후 다시 시도해보세요.")

with tab2:
    st.text_area(
        "대본을 직접 입력하세요.",
        height=200,
        placeholder="여기에 대본을 입력해주세요...",
        help="여기에 입력된 텍스트로 PPT 대본이 생성됩니다."
    )

# 하단 디자인 BAR
with st.container():
    st.markdown('<div class="bottom-design-bar">', unsafe_allow_html=True)
    if st.button("🚀 PPT 자동 생성 시작"):
        # 버튼 클릭 시 실행될 로직
        st.success("PPT 생성 중입니다... 잠시만 기다려 주세요.")
        # 여기에 PPT 생성 및 다운로드 로직을 추가합니다.
    st.markdown('</div>', unsafe_allow_html=True)