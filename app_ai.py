import streamlit as st
from pptx import Presentation
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt  # MSO_VERTICAL_ANCHOR 제거
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR # MSO_VERTICAL_ANCHOR 추가
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io
import re
import textwrap
import docx
from sentence_transformers import SentenceTransformer, util
import kss
from io import BytesIO
import logging

# Streamlit 설정
st.set_page_config(page_title="Paydo AI PPT", layout="centered")
st.title("🎬 AI PPT 생성기 (KoSimCSE + KSS 의미 단위 분할)")

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

@st.cache_resource
def load_model():
    """KoSimCSE 모델 로드 (캐싱)"""
    return SentenceTransformer("jhgan/ko-sbert-nli")

model = load_model()

def extract_text_from_word(uploaded_file):
    """Word 파일에서 텍스트 추출"""
    try:
        file_bytes = BytesIO(uploaded_file.read())
        doc = docx.Document(file_bytes)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        logging.debug(f"Word paragraphs extracted: {len(paragraphs)} paragraphs")
        return paragraphs
    except FileNotFoundError:
        st.error("오류: Word 파일을 찾을 수 없습니다.")
        return []
    except docx.exceptions.PackageNotFoundError:
        st.error("오류: Word 파일이 유효하지 않습니다.")
        return []
    except Exception as e:
        st.error(f"오류: Word 파일 처리 중 오류 발생: {e}")
        logging.error(f"Word 파일 처리 오류 상세: {e}")
        return []

def calculate_text_lines(text, max_chars_per_line):
    """텍스트 줄 수 계산"""
    lines = 0
    for paragraph in text.split('\n'):
        lines += len(textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=True)) or 1
    return lines

def smart_sentence_split(text):
    """KSS를 이용한 문장 분리"""
    return kss.split_sentences(text)

def is_incomplete(sentence):
    """불완전한 문장 여부 확인"""
    return sentence.endswith(('은','는','이','가','을','를','에','으로','고','와','과')) or len(sentence) < 8 or re.match(r'^(그리고|하지만|그러나|또한|그래서|즉|또|그러면|그런데)$', sentence.strip())

def merge_sentences(sentences):
    """불완전한 문장 병합"""
    merged = []
    buffer = ""
    for sentence in sentences:
        if buffer:
            buffer += " " + sentence
            if not is_incomplete(sentence):
                merged.append(buffer.strip())
                buffer = ""
        else:
            if is_incomplete(sentence):
                buffer = sentence
            else:
                merged.append(sentence)
    if buffer:
        merged.append(buffer.strip())
    return merged

def split_text_into_slides_with_similarity(text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt, model, similarity_threshold=0.85):
    """의미 단위 및 문맥 유사도 기반 슬라이드 분할"""

    slides, split_flags = [], []
    current_text = ""
    current_lines = 0
    needs_check = False

    for paragraph in text_paragraphs:
        sentences = smart_sentence_split(paragraph)
        merged_sentences = merge_sentences(sentences)
        embeddings = model.encode(merged_sentences)

        for i, sentence in enumerate(merged_sentences):
            sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

            # 짧은 문장 다음 문장과 병합
            if sentence_lines <= 2 and i + 1 < len(merged_sentences) and calculate_text_lines(merged_sentences[i] + " " + merged_sentences[i+1], max_chars_per_line_ppt) <= max_lines_per_slide:
                sentence = merged_sentences[i] + " " + merged_sentences[i+1]
                sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)
                i += 1

            if sentence_lines > max_lines_per_slide:
                wrapped_lines = textwrap.wrap(sentence, max_chars_per_line_ppt, break_long_words=True)
                temp_text = ""
                temp_lines = 0
                for line in wrapped_lines:
                    line_lines = calculate_text_lines(line, max_chars_per_line_ppt)
                    if temp_lines + line_lines <= max_lines_per_slide:
                        temp_text += line + "\n"
                        temp_lines += line_lines
                    else:
                        slides.append(temp_text.strip())
                        split_flags.append(True)
                        temp_text = line + "\n"
                        temp_lines = line_lines
                if temp_text:
                    slides.append(temp_text.strip())
                    split_flags.append(True)
                current_text = ""
                current_lines = 0
            elif current_lines + sentence_lines <= max_lines_per_slide:
                # 유사도 검사 추가 (첫 문장 제외)
                if current_text and i > 0 and util.cos_sim(embeddings[i-1], embeddings[i])[0][0] < similarity_threshold:
                    slides.append(current_text.strip())
                    split_flags.append(needs_check)
                    current_text = sentence + "\n"
                    current_lines = sentence_lines
                    needs_check = False
                else:
                    current_text += sentence + "\n"
                    current_lines += sentence_lines
            else:
                slides.append(current_text.strip())
                split_flags.append(needs_check)
                current_text = sentence + "\n"
                current_lines = sentence_lines
                needs_check = False
        if current_text:
            slides.append(current_text.strip())
            split_flags.append(needs_check)
    return slides, split_flags

def create_ppt(slides, split_flags, max_chars_per_line_in_ppt, font_size):
    """PPT 생성"""

    prs = Presentation()
    for i, text in enumerate(slides):
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            textbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(6.2))
            text_frame = textbox.text_frame
            text_frame.clear()
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
            text_frame.word_wrap = True
            for line in textwrap.wrap(text, width=max_chars_per_line_in_ppt, break_long_words=True):
                p = text_frame.add_paragraph()
                p.text = line
                p.font.size = Pt(font_size)
                p.font.name = 'Noto Color Emoji'
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER
        except Exception as e:
            st.error(f"슬라이드 생성 오류: {e}")
            logging.error(f"슬라이드 생성 오류 상세: {e}")
            return None
        if split_flags[i]:
            # 확인 필요 슬라이드 표시
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(2.5), Inches(0.5))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
            shape.text_frame.text = "확인 필요!"
            shape.text_frame.paragraphs[0].font.size = Pt(18)
            shape.text_frame.paragraphs[0].font.bold = True
            shape.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return prs

# UI
uploaded_file = st.file_uploader("📄 Word 파일 업로드", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력:", height=300)
max_lines = st.slider("슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.slider("한 줄당 최대 글자 수", 10, 100, 18)
font_size = st.slider("폰트 크기", 10, 60, 54)
sim_threshold = st.slider("문맥 유사도 기준", 0.0, 1.0, 0.85, step=0.05)

if st.button("🚀 PPT 생성"):
    paragraphs = []
    if uploaded_file:
        paragraphs = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        paragraphs = [p.strip() for p in text_input.split("\n\n") if p.strip()]
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    if not paragraphs:
        st.error("유효한 텍스트가 없습니다.")
        st.stop()

    with st.spinner("PPT 생성 중..."):
        try:
            slides, flags = split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, sim_threshold)
            ppt = create_ppt(slides, flags, max_chars, font_size)

            if ppt:
                ppt_bytes = BytesIO()
                ppt.save(ppt_bytes)
                ppt_bytes.seek(0)
                st.download_button("📥 PPT 다운로드", ppt_bytes, "paydo_script_ai.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                st.success(f"총 {len(slides)}개의 슬라이드가 생성되었습니다.")
                if any(flags):
                    flagged = [i+1 for i, f in enumerate(flags) if f]
                    st.warning(f"⚠️ 확인이 필요한 슬라이드: {flagged}")
        except Exception as e:
            st.error(f"PPT 생성 오류: {e}")
            logging.error(f"PPT 생성 오류 상세: {e}")