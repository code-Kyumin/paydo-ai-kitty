# Paydo AI PPT 생성기 with KoSimCSE 적용 및 오류 수정

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
import docx
from io import BytesIO
from sentence_transformers import SentenceTransformer, util

# Streamlit 세팅
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
# st.title("🎬 AI PPT 생성기 (KoSimCSE)") # 이 라인은 더 이상 사용하지 않습니다.

# 모델 로딩 (한 번만)
@st.cache_resource
def load_model():
    return SentenceTransformer("jhgan/ko-sbert-nli")

model = load_model()

# Word 파일 텍스트 추출
def extract_text_from_word(uploaded_file):
    try:
        file_bytes = BytesIO(uploaded_file.read())
        doc = docx.Document(file_bytes)
        return [p.text for p in doc.paragraphs if p.text.strip()]
    except Exception as e:
        st.error(f"Word 파일 처리 오류: {e}")
        return None

# 텍스트 줄 수 계산
def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
        else:
            lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True))
    return lines

# 문장 분할
def smart_sentence_split(text):
    paragraphs = text.split('\n')
    sentences = []
    for paragraph in paragraphs:
        # 서술어 단독 분리 방지를 위해 문장 끝 마침표 기준이 아닌, 약간 넓게 split
        temp_sentences = re.split(r'(?<=[^\d][.!?])\s+(?=[\"\'\uAC00-\uD7A3])', paragraph)
        sentences.extend([s.strip() for s in temp_sentences if s.strip()])
    return sentences

# 슬라이드 분할 with 유사도 + 짧은 문장 병합 개선
def split_text_into_slides_with_similarity(text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt, model, similarity_threshold=0.85):
    slides, split_flags, slide_number = [], [], 1
    current_text, current_lines, needs_check = "", 0, False

    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)
        if not sentences:
            continue

        embeddings = model.encode(sentences)

        i = 0
        while i < len(sentences):
            sentence = sentences[i]
            sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

            # 다음 문장과 병합을 시도 (너무 짧은 문장 방지)
            if sentence_lines <= 2 and i + 1 < len(sentences):
                next_sentence = sentences[i + 1]
                merged = sentence + " " + next_sentence
                merged_lines = calculate_text_lines(merged, max_chars_per_line_ppt)
                if merged_lines <= max_lines_per_slide:
                    sentence = merged
                    sentence_lines = merged_lines
                    i += 1  # 추가로 하나 더 소비

            if sentence_lines > max_lines_per_slide:
                wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line_ppt, break_long_words=True)
                temp_text, temp_lines = "", 0
                for line in wrapped_lines:
                    line_lines = calculate_text_lines(line, max_chars_per_line_ppt)
                    if temp_lines + line_lines <= max_lines_per_slide:
                        temp_text += line + "\n"
                        temp_lines += line_lines
                    else:
                        slides.append(temp_text.strip())
                        split_flags.append(True)
                        slide_number += 1
                        temp_text = line + "\n"
                        temp_lines = line_lines
                if temp_text:
                    slides.append(temp_text.strip())
                    split_flags.append(True)
                    slide_number += 1
                current_text, current_lines = "", 0
                i += 1
                continue

            if current_lines + sentence_lines <= max_lines_per_slide:
                current_text += sentence + "\n"
                current_lines += sentence_lines
            else:
                slides.append(current_text.strip())
                split_flags.append(needs_check)
                slide_number += 1
                current_text = sentence + "\n"
                current_lines = sentence_lines
                needs_check = False
            i += 1

    if current_text:
        slides.append(current_text.strip())
        split_flags.append(needs_check)

    return slides, split_flags

def create_ppt(slide_texts, split_flags, max_chars_per_line_in_ppt=18, font_size=54):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    total_slides = len(slide_texts)

    for i, text in enumerate(slide_texts):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_text_to_slide(slide, text, font_size, PP_ALIGN.CENTER, max_chars_per_line_in_ppt)
        if split_flags[i]:
            add_check_needed_shape(slide)
        if i == total_slides - 1:
            add_end_mark(slide)
    return prs

def add_text_to_slide(slide, text, font_size, alignment, max_chars_per_line):
    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    text_frame.word_wrap = True

    wrapped_lines = textwrap.wrap(text, width=max_chars_per_line, break_long_words=True)
    for line in wrapped_lines:
        p = text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.name = 'Noto Color Emoji'
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = alignment
        p.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

    text_frame.auto_size = None

def add_check_needed_shape(slide):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(2.5), Inches(0.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    p = shape.text_frame.paragraphs[0]
    p.text = "확인 필요!"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    shape.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER

def add_end_mark(slide):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10), Inches(6), Inches(2), Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    p = shape.text_frame.paragraphs[0]
    p.text = "끝"
    p.font.size = Pt(36)
    p.font.color.rgb = RGBColor(255, 255, 255)
    shape.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    p.alignment = PP_ALIGN.CENTER


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
        /* 하단 고정 바 때문에 메인 컨테이너 하단에 패딩 추가 */
        padding-bottom: 90px; /* 하단 고정 바의 높이(padding 15+15+버튼 높이 고려)에 맞춰 조절 */
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
        font-size: 1.5em; /* 이 부분은 Python 코드의 인라인 스타일이 우선합니다. */
        font-weight: 700;
    }

    /* 고정된 하단 바 스타일 (새로 추가) */
    .fixed-bottom-bar { 
        background-color: #A2D9CE; /* 옅은 녹색으로 변경 (연두색으로 보이도록) */
        padding: 15px 20px;
        text-align: center;
        box-shadow: 0 -2px 5px rgba(0, 0, 0, 0.1);
        position: fixed; /* 뷰포트 하단에 고정 */
        bottom: 0; /* 하단에 붙임 */
        left: 50%; /* 왼쪽 50% 이동 */
        transform: translateX(-50%); /* 자신의 너비의 절반만큼 왼쪽으로 이동하여 중앙 정렬 */
        width: 100%; /* 너비 100% */
        max-width: 800px; /* 메인 컨테이너와 동일한 최대 너비 적용 */
        z-index: 1000; /* 다른 요소 위에 표시되도록 가장 높은 z-index 부여 */
        display: flex; /* 내부 버튼을 중앙 정렬하기 위한 flexbox */
        justify-content: center; /* 버튼을 중앙에 정렬 */
        align-items: center;
        box-sizing: border-box; /* padding이 width에 포함되도록 */
        border-bottom-left-radius: 8px; /* 메인 컨테이너와 일치하도록 */
        border-bottom-right-radius: 8px; /* 메인 컨테이너와 일치하도록 */
    }

    /* 고정된 하단 바 안에 있는 Streamlit 버튼 컨테이너 (.stButton) */
    .fixed-bottom-bar .stButton {
        width: auto; /* flex 컨테이너 내에서 콘텐츠 크기에 맞게 너비 조절 */
        display: flex; /* 내부 버튼을 가운데 정렬하기 위해 flexbox 적용 */
        justify-content: center; /* 이 stButton 내부의 실제 버튼을 가운데 정렬 */
        margin: 0; /* Streamlit 기본 마진 상쇄 (필요 시) */
    }

    /* 고정된 하단 바 안에 있는 실제 버튼 (button 태그) 스타일 */
    .fixed-bottom-bar .stButton > button { 
        background-color: #2ecc71; /* 초록색 (기존과 동일하게 유지) */
        color: white;
        border: none;
        padding: 12px 25px; /* 패딩 증가로 버튼 크기 키우기 */
        border-radius: 8px; /* 더 둥글게 */
        cursor: pointer;
        font-size: 1.3em; /* 폰트 크기 키우기 */
        font-weight: 700;
        width: auto; /* 버튼 콘텐츠 크기에 맞게 너비 조절 */
        max-width: 400px; /* 최대 너비 제한 (너무 길어지는 것을 방지) */
        display: flex; /* flexbox 사용 */
        align-items: center;
        justify-content: center;
        gap: 10px;
        transition: background-color 0.3s ease;
    }
    .fixed-bottom-bar .stButton > button:hover {
        background-color: #27ae60; /* 호버 시 더 어두운 초록색 */
    }
    
    /* 기존의 전역 .stButton > button 스타일은 삭제하거나 주석 처리 */
    /*
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
    */


    /* Streamlit 메인 콘텐츠 영역 (기본 패딩을 활용) */
    /* 이 부분은 Streamlit이 자동으로 패딩을 추가하므로, 별도의 컨테이너를 만들지 않고
       css로 전체 앱 컨테이너의 배경색을 흰색으로 설정하여 흰색 바탕을 유지합니다. */
    /*
    .st-emotion-cache-1c7y2vl { // 메인 콘텐츠를 감싸는 Streamlit 내부 div - 셀렉터 변경될 수 있음
        padding: 20px; // 내부 여백
        background-color: #fff; // 메인 콘텐츠 배경색
    }
    */
    /* 위 주석 처리된 부분 대신 [data-testid="stAppViewContainer"]에 padding-bottom을 추가하여
       하단 고정 바가 콘텐츠를 가리지 않도록 했습니다. */


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

    /* 아래 버튼 스타일은 .fixed-bottom-bar 안에 있는 버튼에만 적용됩니다. */
    /*
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
    */

    /* 반응형 디자인 (선택 사항: 화면 크기가 작아질 때 조절) */
    @media (max-width: 768px) {
        [data-testid="stAppViewContainer"] {
            border-radius: 0; /* 모바일에서 전체 화면 사용 */
            box-shadow: none;
        }

        [data-testid="stHeader"], .fixed-bottom-bar {
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
# 폰트 크기를 인라인 스타일로 직접 지정 (CSS보다 우선순위가 높음)
st.markdown("""
    <div class="top-design-bar">
        <h1 style='color: #fff; margin: 0; 
                   font-size: 0.4em !important; /* !important를 인라인에 추가 */
                   font-weight: 700; text-align: center; 
                   display: flex; align-items: center; justify-content: center; gap: 10px;'>
            🎬 촬영 대본 PPT 자동 생성 AI (KoSimCSE)
        </h1>
    </div>
""", unsafe_allow_html=True)


# 대본 입력 방식 선택 섹션
st.markdown('<div class="input-method-selection-box"><span class="icon">📁</span> 대본 입력 방식 선택</div>', unsafe_allow_html=True)

# 탭 메뉴 구성 (st.tabs 위젯 사용)
tab1, tab2 = st.tabs(["Word 파일 업로드", "텍스트 직접 입력"])

with tab1:
    st.write("Word 파일 (.docx)을 업로드해주세요.")

    # 파일 업로더 위젯
    # 기본 라벨은 숨기고 (label_visibility="collapsed"), 커스텀 텍스트를 마크다운으로 삽입
    uploaded_file_tab1 = st.file_uploader( # 변수명 통일 (uploaded_file_tab1)
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

    if uploaded_file_tab1 is not None: # 변수명 통일
        st.success(f"파일 '{uploaded_file_tab1.name}'이(가) 업로드되었습니다.")
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
    text_input_tab2 = st.text_area( # 변수명 통일 (text_input_tab2)
        "대본을 직접 입력하세요.",
        height=200,
        placeholder="여기에 대본을 입력해주세요...",
        help="여기에 입력된 텍스트로 PPT 대본이 생성됩니다."
    )
    # st.info("여기에 입력된 텍스트로 PPT 대본이 생성됩니다.") # help 속성으로 대체 가능

# UI 입력 (기존 하단 UI 입력 슬라이더 부분)
# 이 부분은 페이지 하단에 배치됩니다.
st.markdown("---") # 구분선 추가
st.subheader("⚙️ PPT 생성 옵션")
st.write("생성될 PPT의 세부 옵션을 설정할 수 있습니다.")

max_lines = st.slider("슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.slider("한 줄당 최대 글자 수", 10, 100, 18)
font_size = st.slider("폰트 크기", 10, 60, 54)
sim_threshold = st.slider("문맥 유사도 기준", 0.0, 1.0, 0.85, step=0.05)


# 고정된 하단 바 (새롭게 추가)
st.markdown('<div class="fixed-bottom-bar">', unsafe_allow_html=True) 
if st.button("🚀 PPT 자동 생성 시작"): # 이 버튼이 div 안에 들어갑니다.
    paragraphs = []
    target_file = None
    target_text_input = ""

    if uploaded_file_tab1 is not None:
        paragraphs = extract_text_from_word(uploaded_file_tab1)
    elif text_input_tab2.strip():
        paragraphs = [p.strip() for p in text_input_tab2.split("\n\n") if p.strip()]
    else:
        st.warning("PPT 생성을 위해 Word 파일을 업로드하거나 대본을 직접 입력해주세요.")
        st.stop()

    if not paragraphs:
        st.error("유효한 텍스트가 없습니다.")
        st.stop()

    with st.spinner("PPT 생성 중..."):
        slides, flags = split_text_into_slides_with_similarity(
            paragraphs, max_lines, max_chars, model, similarity_threshold=sim_threshold
        )
        ppt = create_ppt(slides, flags, max_chars, font_size)

        if ppt:
            ppt_io = io.BytesIO()
            ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button(
                label="📥 PPT 다운로드",
                data=ppt_io,
                file_name="paydo_script_ai.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.success(f"총 {len(slides)}개의 슬라이드가 생성되었습니다.")
            if any(flags):
                flagged = [i+1 for i, f in enumerate(flags) if f]
                st.warning(f"⚠️ 확인이 필요한 슬라이드: {flagged}")
st.markdown('</div>', unsafe_allow_html=True)