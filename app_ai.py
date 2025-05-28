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

    /* Streamlit 헤더 영역 스타일 (상단 바 역할) */
    /* Streamlit 버전업에 따라 data-testid 값은 변경될 수 있습니다. */
    [data-testid="stHeader"] {
        background-color: #2c3e50; /* 어두운 파란색/회색 */
        color: #fff;
        padding: 15px 20px;
        text-align: center;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        position: sticky; /* 스크롤 시 상단에 고정 */
        top: 0; /* 상단에 붙임 */
        z-index: 999; /* 다른 요소 위에 표시되도록 */
        /* Streamlit 기본 마진 상쇄 및 너비 조절 */
        margin-left: -1rem; 
        margin-right: -1rem;
        width: calc(100% + 2rem);
    }
    /* 상단 바 제목 (Streamlit의 기본 제목 스타일 오버라이드) */
    [data-testid="stHeader"] h1 {
        color: #fff;
        margin: 0;
        font-size: 1.5em;
        font-weight: 700;
    }

    /* 하단 바 스타일 (Streamlit의 버튼 컨테이너 활용) */
    /* st.button이 포함될 컨테이너를 타겟팅합니다. */
    .bottom-bar {
        background-color: #2ecc71; /* 초록색 */
        color: #fff;
        padding: 15px;
        text-align: center;
        border-bottom-left-radius: 8px;
        border-bottom-right-radius: 8px;
        box-shadow: 0 -2px 5px rgba(0, 0, 0, 0.1);
        position: sticky; /* 스크롤 시 하단에 고정 */
        bottom: 0; /* 하단에 붙임 */
        z-index: 999;
        /* Streamlit 기본 마진 상쇄 및 너비 조절 */
        margin-left: -1rem;
        margin-right: -1rem;
        width: calc(100% + 2rem);
    }
    
    /* Streamlit 메인 콘텐츠 영역 (기본 패딩을 활용) */
    /* 이 부분은 Streamlit이 자동으로 패딩을 추가하므로, 별도의 컨테이너를 만들지 않고
       css로 전체 앱 컨테이너의 배경색을 흰색으로 설정하여 흰색 바탕을 유지합니다. */
    .st-emotion-cache-1c7y2vl { /* 메인 콘텐츠를 감싸는 Streamlit 내부 div - 셀렉터 변경될 수 있음 */
        padding: 20px; /* 내부 여백 */
        background-color: #fff; /* 메인 콘텐츠 배경색 */
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
    /* st.tabs는 내부적으로 Shadow DOM을 사용하므로, 외부 CSS로 모든 것을 제어하기 어렵습니다.
       아래는 가능한 범위 내에서 기본 스타일을 조정합니다. */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px; /* 탭 사이 간격 제거 */
        border-bottom: 1px solid #ddd; /* 탭 목록 하단 테두리 */
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
        border-bottom: 2px solid #3498db !important; /* 파란색 밑줄 (Streamlit 기본 스타일 오버라이드) */
        color: #3498db !important; /* 활성화된 탭 텍스트 색상 파란색 */
        font-weight: 700;
        background-color: #fff;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background-color: #f5f5f5;
    }

    /* Streamlit 파일 업로더 커스터마이징 */
    /* st.file_uploader의 드롭존(Dropzone) 스타일 */
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #a0d8f0; /* 연한 파란색 점선 테두리 */
        border-radius: 8px;
        background-color: #f7fcfe; /* 아주 연한 파란색 배경 */
        padding: 30px 20px; /* 내부 패딩 */
        height: 180px; /* 높이 고정 (원하는 높이로 조절) */
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    /* 파일 업로더의 기본 안내 텍스트 숨기기 */
    [data-testid="stFileUploaderDropzoneInstructions"] > div > span {
        display: none; 
    }
    /* 파일 업로더의 기본 제한 텍스트 숨기기 */
    [data-testid="stFileUploaderDropzoneInstructions"] > div > small {
        display: none; 
    }
    /* 파일 업로더의 "Browse files" 버튼 숨기기 (원한다면) */
    /* [data-testid="stFileUploaderBrowseButton"] {
        display: none;
    } */
    /* 드래그 앤 드롭 아이콘 커스터마이징을 위한 stFileUploaderDropzoneTarget */
    [data-testid="stFileUploaderDropzoneTarget"] {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 100%;
        width: 100%;
        position: relative; /* 자식 요소 절대 위치 지정을 위해 */
    }
    /* 자체적으로 아이콘과 텍스트 추가 (st.markdown으로) */
    /* 기존 browse files 버튼 위치 조절 */
    [data-testid="stFileUploaderBrowseButton"] {
        position: absolute;
        bottom: 20px;
        right: 20px;
    }
    [data-testid="stFileUploaderBrowseButton"] > button {
        background-color: #3498db; /* 파란색 */
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 0.9em;
        font-weight: 600;
        transition: background-color 0.3s ease;
    }
    [data-testid="stFileUploaderBrowseButton"] > button:hover {
        background-color: #2980b9; /* 더 어두운 파란색 */
    }


    /* 문제 해결 Expander (st.expander) 스타일 */
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
        outline: none; /* 클릭 시 기본 외곽선 제거 */
    }
    .stExpander > div > div > details > summary:hover {
        background-color: #f0f0f0;
        border-radius: 8px;
    }
    .stExpander > div > div > details > summary::marker { /* 기본 드롭다운 마커 제거 */
        content: '';
    }
    .stExpander > div > div > details > summary::before { /* 사용자 정의 화살표 */
        content: '▼';
        font-size: 0.8em;
        margin-right: 5px;
        transition: transform 0.2s;
    }
    .stExpander > div > div > details[open] > summary::before {
        transform: rotate(180deg); /* 열렸을 때 화살표 회전 */
    }
    .stExpander > div > div > details > div { /* Expander 내부 콘텐츠 */
        padding: 5px 15px 10px;
        border-top: 1px dashed #eee; /* 내용 위 점선 구분선 */
        font-size: 0.85em;
        color: #777;
    }

    /* PPT 자동 생성 시작 버튼 */
    .stButton > button {
        background-color: #2ecc71; /* 초록색 */
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        cursor: pointer;
        font-size: 1.2em;
        font-weight: 700;
        width: 100%; /* 버튼 너비 100% */
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px; /* 아이콘과 텍스트 사이 간격 */
        transition: background-color 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #27ae60; /* 호버 시 더 어두운 초록색 */
    }

    /* 반응형 디자인 (선택 사항: 화면 크기가 작아질 때 조절) */
    @media (max-width: 768px) {
        [data-testid="stAppViewContainer"] {
            border-radius: 0; /* 모바일에서 전체 화면 사용 */
            box-shadow: none;
        }

        [data-testid="stHeader"], .bottom-bar {
            border-radius: 0; /* 모바일에서 바도 둥근 모서리 제거 */
        }
    }
</style>
"""

# Streamlit 앱에 사용자 정의 CSS 주입
st.markdown(custom_css, unsafe_allow_html=True)

# --- Streamlit 앱 UI 구성 시작 ---

# 상단 바 (st.markdown을 사용하여 HTML h1 태그 삽입)
# st.header나 st.title을 사용하면 Streamlit 기본 스타일이 적용되어 CSS 오버라이딩이 더 어려울 수 있습니다.
# 여기서는 CSS가 적용되는 [data-testid="stHeader"]를 활용합니다.
st.markdown("<h1 style='display: none;'>촬영 대본 PPT 자동 생성 AI (KoSimCSE)</h1>", unsafe_allow_html=True)
# 실제 텍스트는 CSS의 [data-testid="stHeader"] h1에 의해 표시됩니다.

# 대본 입력 방식 선택 섹션
st.markdown('<div class="input-method-selection-box"><span class="icon">📁</span> 대본 입력 방식 선택</div>', unsafe_allow_html=True)

# 탭 메뉴 구성 (st.tabs 위젯 사용)
tab1, tab2 = st.tabs(["Word 파일 업로드", "텍스트 직접 입력"])

with tab1:
    st.write("Word 파일 (.docx)을 업로드해주세요.")

    # 파일 업로더 위젯
    # 기본 라벨은 숨기고 (label_visibility="collapsed"), 커스텀 텍스트를 마크다운으로 삽입
    uploaded_file = st.file_uploader(
        "파일을 드래그 앤 드롭하거나 찾아보세요.", # 이 텍스트는 st.file_uploader의 드롭존에 기본적으로 표시됩니다.
        type=["docx"], # 허용되는 파일 형식
        accept_multiple_files=False, # 단일 파일만 허용
        label_visibility="collapsed" # 기본 라벨 숨기기
    )
    
    # 드래그 앤 드롭 영역 내 커스텀 텍스트 및 아이콘 (CSS로 위치 조정)
    st.markdown("""
        <div style="text-align: center; margin-top: -160px; pointer-events: none; position: relative; z-index: 1;">
            <i class="fas fa-cloud-upload-alt" style="font-size: 3em; color: #3498db; margin-bottom: 5px;"></i>
            <p style="margin:0; font-size: 1.1em; color: #666;">Drag and drop file here</p>
        </div>
        <div style="text-align: center; font-size: 0.85em; color: #888; margin-top: 10px; position: relative; z-index: 1;">
            Limit 200MB per file • DOCX
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
    # st.info("여기에 입력된 텍스트로 PPT 대본이 생성됩니다.") # help 속성으로 대체 가능

# 하단 바 (st.markdown을 사용하여 HTML div를 만들고 그 안에 버튼 배치)
# 버튼은 st.button을 사용하여 Streamlit의 기능적인 버튼을 유지합니다.
with st.container(): # 하단 바 영역을 위한 컨테이너
    st.markdown('<div class="bottom-bar">', unsafe_allow_html=True) # 하단 바 CSS 클래스 적용
    if st.button("🚀 PPT 자동 생성 시작"):
        # 버튼 클릭 시 실행될 로직
        st.success("PPT 생성 중입니다... 잠시만 기다려 주세요.")
        # 여기에 PPT 생성 및 다운로드 로직을 추가합니다.
    st.markdown('</div>', unsafe_allow_html=True)