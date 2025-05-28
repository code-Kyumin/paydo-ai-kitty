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
from sentence_transformers import SentenceTransformer

# 기본 스타일 설정
st.set_page_config(page_title="촬영 대본 PPT 자동 생성 AI", layout="centered")
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        font-family: 'Segoe UI', sans-serif;
    }
    h1.title-style {
        font-size: 1.8rem;
        color: #222;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    .section {
        background-color: #f9f9f9;
        padding: 1rem 1.2rem;
        border-radius: 0.5rem;
        border: 1px solid #ddd;
        margin-bottom: 1rem;
    }
    .stSlider > div {
        padding-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# 제목 및 안내
st.markdown('<h1 class="title-style">🎬 촬영 대본 PPT 자동 생성 AI (KoSimCSE)</h1>', unsafe_allow_html=True)
st.markdown("""
<div class="section">
    📢 Word 파일 업로드 오류 시, **파일명을 영문으로 변경한 후 업로드**해 주세요. 
    한글 파일명은 시스템 호환성 문제로 인해 오류가 발생할 수 있습니다.
</div>
""", unsafe_allow_html=True)

# 모델 로딩
@st.cache_resource
def load_model():
    return SentenceTransformer("jhgan/ko-sbert-nli")
model = load_model()

# 사이드바 슬라이드 설정
st.sidebar.markdown("#### ⚙️ 슬라이드 설정")
max_lines = st.sidebar.slider("📏 슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.sidebar.slider("🔠 한 줄당 최대 글자 수", 10, 100, 18)
font_size = st.sidebar.slider("🔡 폰트 크기", 10, 60, 54)
sim_threshold = st.sidebar.slider("🧠 문맥 유사도 기준", 0.0, 1.0, 0.85, step=0.05)

# 입력 구역
st.markdown("""
<div class="section">
    <h4 style='margin-bottom:0.8rem'>📤 Word 파일 업로드 또는 텍스트 직접 입력</h4>
""", unsafe_allow_html=True)
uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
st.markdown("<div style='margin-top: 0.5rem'></div>", unsafe_allow_html=True)
st.markdown("✍️ 또는 아래 입력란에 직접 텍스트를 작성하세요:")
text_input = st.text_area("", height=300)
st.markdown("</div>", unsafe_allow_html=True)

# 텍스트 처리 함수들
def extract_text_from_word(uploaded_file):
    try:
        file_bytes = BytesIO(uploaded_file.read())
        doc = docx.Document(file_bytes)
        return [p.text for p in doc.paragraphs if p.text.strip()]
    except Exception:
        raise

def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    paragraphs = text.split('\n')
    for paragraph in paragraphs:
        if not paragraph:
            lines += 1
        else:
            lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True))
    return lines

def smart_sentence_split(text):
    paragraphs = text.split('\n')
    sentences = []
    for paragraph in paragraphs:
        temp_sentences = re.split(r'(?<=[^\d][.!?])\s+(?=[\"\'\uAC00-\uD7A3])', paragraph)
        sentences.extend([s.strip() for s in temp_sentences if s.strip()])
    return sentences

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

            if sentence_lines <= 2 and i + 1 < len(sentences):
                next_sentence = sentences[i + 1]
                merged = sentence + " " + next_sentence
                merged_lines = calculate_text_lines(merged, max_chars_per_line_ppt)
                if merged_lines <= max_lines_per_slide:
                    sentence = merged
                    sentence_lines = merged_lines
                    i += 1

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

# 실행 버튼
st.markdown("<div style='text-align:center; margin-top:1.5rem'>", unsafe_allow_html=True)
if st.button("🚀 PPT 자동 생성 시작", use_container_width=True):
    paragraphs = []
    if uploaded_file:
        try:
            paragraphs = extract_text_from_word(uploaded_file)
        except Exception:
            st.error("❌ Word 파일을 처리하는 중 오류가 발생했습니다.\n\n📌 **파일명을 영문으로 변경한 뒤 다시 업로드해 주세요.**")
            st.stop()
    elif text_input.strip():
        paragraphs = [p.strip() for p in text_input.split("\n\n") if p.strip()]
    else:
        st.warning("📎 Word 파일을 업로드하거나 텍스트를 입력해 주세요.")
        st.stop()

    if not paragraphs:
        st.error("❗ 유효한 텍스트가 없습니다.")
        st.stop()

    with st.spinner("🛠️ PPT 생성 중입니다..."):
        slides, flags = split_text_into_slides_with_similarity(
            paragraphs, max_lines, max_chars, model, similarity_threshold=sim_threshold
        )
        ppt = create_ppt(slides, flags, max_chars, font_size)

        if ppt:
            ppt_io = io.BytesIO()
            ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("📥 PPT 다운로드", ppt_io, "paydo_script_ai.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            st.success(f"✅ 총 {len(slides)}개의 슬라이드가 생성되었습니다.")
            if any(flags):
                flagged = [i+1 for i, f in enumerate(flags) if f]
                st.warning(f"⚠️ 확인이 필요한 슬라이드 번호: {flagged}")
st.markdown("</div>", unsafe_allow_html=True)
