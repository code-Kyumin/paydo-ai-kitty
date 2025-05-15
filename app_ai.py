
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_AUTO_SIZE

import io
import re
import textwrap
import docx
from io import BytesIO
from sentence_transformers import SentenceTransformer, util
import kss
import logging
from PIL import Image

# Streamlit 설정
st.set_page_config(page_title="AI PPT 생성기", layout="wide")
st.title("🎬 AI PPT 생성기 (KoSimCSE + KSS)")

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@st.cache_resource
def load_model():
    logging.info("Loading KoSimCSE model...")
    model = SentenceTransformer("jhgan/ko-sbert-nli")
    logging.info("KoSimCSE model loaded.")
    return model

model = load_model()

def extract_text_from_word(uploaded_file):
    try:
        file_bytes = BytesIO(uploaded_file.read())
        doc = docx.Document(file_bytes)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return paragraphs
    except Exception as e:
        st.error(f"Word 파일 처리 오류: {e}")
        return []

def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    for paragraph in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=False, replace_whitespace=False)
        lines += len(wrapped_lines) or 1
    return lines

def smart_sentence_split(text):
    try:
        return kss.split_sentences(text)
    except Exception:
        return [s.strip() for s in re.split(r'[.!?]\s+', text) if s.strip()]

def is_incomplete(sentence):
    sentence = sentence.strip()
    if not sentence or len(sentence) < 10:
        return True
    if sentence.endswith(('은','는','이','가','을','를','에','으로','고','와','과', '며', '는데', '지만', '거나', '든지', '든지간에', '든가')):
        return True
    if re.match(r'^(그리고|하지만|그러나|또한|그래서|즉|또|그러면|그런데)$', sentence):
        return True
    if not sentence.endswith(('.', '!', '?', '다', '요', '죠', '까', '나', '시오')) and len(sentence) < 15:
        return True
    return False

def merge_sentences(sentences):
    merged, buffer = [], ""
    for i, sentence in enumerate(sentences):
        sentence = sentence.strip()
        if not sentence:
            continue
        if buffer:
            current = buffer + " " + sentence
            if len(current) > 200:
                merged.append(buffer)
                buffer = sentence
            else:
                buffer = current
            if not is_incomplete(buffer) or i == len(sentences) - 1:
                merged.append(buffer)
                buffer = ""
        else:
            buffer = sentence if is_incomplete(sentence) and i < len(sentences) - 1 else ""
            if not buffer:
                merged.append(sentence)
    if buffer:
        merged.append(buffer)
    return merged

def split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, threshold=0.85):
    slides, current_text, current_lines = [], "", 0
    all_sentences = [s for p in paragraphs for s in smart_sentence_split(p)]
    merged_sentences = merge_sentences(all_sentences)
    if not merged_sentences:
        return [""], [False]
    embeddings = model.encode(merged_sentences)
    for i, sentence in enumerate(merged_sentences):
        sentence_lines = calculate_text_lines(sentence, max_chars)
        if sentence_lines > max_lines:
            if current_text:
                slides.append(current_text.strip())
                current_text, current_lines = "", 0
            parts = textwrap.wrap(sentence, width=max_chars * max_lines)
            slides.extend(part.strip() for part in parts)
            continue
        similar = True
        if current_text and i > 0 and i < len(embeddings):
            sim = util.cos_sim(embeddings[i-1], embeddings[i])[0][0]
            if sim < threshold:
                similar = False
        if current_lines + sentence_lines <= max_lines and similar:
            current_text = f"{current_text}\n{sentence}" if current_text else sentence
            current_lines += sentence_lines
        else:
            if current_text:
                slides.append(current_text.strip())
            current_text, current_lines = sentence, sentence_lines
    if current_text:
        slides.append(current_text.strip())
    return slides, [False] * len(slides)

def create_ppt(slides, flags, max_chars, font_size):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for i, (slide_text, flag) in enumerate(zip(slides, flags)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(0.75), Inches(0.75), prs.slide_width - Inches(1.5), prs.slide_height - Inches(1.5))
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        for line in textwrap.wrap(slide_text, width=max_chars):
            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(font_size)
            p.font.bold = True
            p.alignment = PP_ALIGN.LEFT
        if flag:
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(1.5), Inches(0.3))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
            tf = shape.text_frame
            tf.text = "확인 필요"
            tf.paragraphs[0].font.size = Pt(10)
            tf.paragraphs[0].font.bold = True
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    return prs

# --- Streamlit UI ---

uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=200)

st.sidebar.header("⚙️ PPT 설정")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.sidebar.slider("한 줄당 최대 글자 수", 10, 100, 18)
font_size = st.sidebar.slider("폰트 크기", 10, 60, 54)
sim_threshold = st.sidebar.slider("문맥 유사도 기준", 0.0, 1.0, 0.85, step=0.05)

if st.button("✨ PPT 생성"):
    if uploaded_file or text_input:
        paragraphs = extract_text_from_word(uploaded_file) if uploaded_file else [p.strip() for p in text_input.split("\n\n") if p.strip()]
        if not paragraphs:
            st.error("입력된 텍스트가 없습니다.")
            st.stop()
        with st.spinner("PPT 생성 중..."):
            slides, flags = split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, threshold=sim_threshold)
            ppt = create_ppt(slides, flags, max_chars, font_size)
            ppt_io = BytesIO()
            ppt.save(ppt_io)
            ppt_io.seek(0)
            st.download_button("📥 PPT 다운로드", ppt_io, "paydo_script_ai.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            st.success(f"총 {len(slides)}개의 슬라이드가 생성되었습니다.")
            if any(flags):
                flagged = [i+1 for i, f in enumerate(flags) if f]
                st.warning(f"⚠️ 확인이 필요한 슬라이드: {flagged}")
    else:
        st.info("Word 파일을 업로드하거나 텍스트를 입력하세요.")
