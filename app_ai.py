import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.slide import PP_SLIDE_LAYOUT  # Import slide layout enum
from pptx.enum.enum import MSO_TRANSITION_TYPE  # Import transition type enum
from pptx.enum.shapes import MSO_SHAPE_TYPE  # Import shape type enum
from pptx.util import Cm

import io
import re
import textwrap
import docx
from io import BytesIO
from sentence_transformers import SentenceTransformer, util
import kss
import logging
from PIL import Image  # Import PIL for image handling

# Streamlit 설정
st.set_page_config(page_title="AI PPT 생성기", layout="wide")
st.title("🎬 AI PPT 생성기 (KoSimCSE + KSS)")

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 모델 로딩 (캐싱)
@st.cache_resource
def load_model():
    logging.info("Loading KoSimCSE model...")
    model = SentenceTransformer("jhgan/ko-sbert-nli")
    logging.info("KoSimCSE model loaded.")
    return model

model = load_model()

# --- Helper Functions ---

def extract_text_from_word(uploaded_file):
    """Word 파일에서 텍스트 추출"""
    try:
        file_bytes = BytesIO(uploaded_file.read())
        doc = docx.Document(file_bytes)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        logging.debug(f"Extracted {len(paragraphs)} paragraphs from Word file.")
        return paragraphs
    except Exception as e:
        logging.error(f"Error extracting text from Word: {e}", exc_info=True)
        st.error(f"Error processing Word file. Please check the file and try again.")
        return []

def calculate_text_lines(text, max_chars_per_line):
    """텍스트 줄 수 계산"""
    lines = 0
    for paragraph in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, 
                                      break_long_words=False, replace_whitespace=False)
        lines += len(wrapped_lines) or 1
    return lines

def smart_sentence_split(text):
    """KSS를 이용한 문장 분리"""
    try:
        sentences = kss.split_sentences(text)
        logging.debug(f"Split text into {len(sentences)} sentences.")
        return sentences
    except Exception as e:
        logging.error(f"Error splitting sentences with KSS: {e}", exc_info=True)
        # Fallback to a simpler split if KSS fails
        return [s.strip() for s in re.split(r'[.!?]\s+', text) if s.strip()]

def is_incomplete(sentence):
    """불완전한 문장 여부 확인"""
    sentence = sentence.strip()
    if not sentence:
        return False
    if len(sentence) < 10:
        return True
    if sentence.endswith(('은','는','이','가','을','를','에','으로','고','와','과', '며', '는데', '지만', '거나', '든지', '든지간에', '든가')):
        return True
    if re.match(r'^(그리고|하지만|그러나|또한|그래서|즉|또|그러면|그런데)$', sentence):
        return True
    if not sentence.endswith(('.', '!', '?', '다', '요', '죠', '까', '나', '시오')) and len(sentence) < 15:
         return True
    return False

def merge_sentences(sentences):
    """불완전한 문장 병합"""
    merged = []
    buffer = ""
    for i, sentence in enumerate(sentences):
        sentence = sentence.strip()
        if not sentence:
            continue

        if buffer:
            current_candidate = buffer + " " + sentence
            if len(current_candidate) > 200:
                merged.append(buffer)
                buffer = sentence
            else:
                buffer = current_candidate

            if not is_incomplete(buffer) or i == len(sentences) - 1:
                merged.append(buffer)
                buffer = ""
        else:
            if is_incomplete(sentence) and i < len(sentences) - 1:
                buffer = sentence
            else:
                merged.append(sentence)
                buffer = ""
    
    if buffer:
        merged.append(buffer)
    return merged

def split_text_into_slides_with_similarity(text_paragraphs, max_lines_per_slide, 
                                           max_chars_per_line_ppt, model, 
                                           similarity_threshold=0.85):
    """의미 단위 및 문맥 유사도 기반 슬라이드 분할"""
    slides = []
    current_slide_text = ""
    current_slide_lines = 0

    all_sentences = []
    for paragraph in text_paragraphs:
        sentences_from_para = smart_sentence_split(paragraph)
        all_sentences.extend(sentences_from_para)
    
    if not all_sentences:
        return [""], [False]

    merged_sentences = merge_sentences(all_sentences)
    
    if not merged_sentences:
        return [""], [False]

    embeddings = model.encode(merged_sentences)

    for i, sentence in enumerate(merged_sentences):
        sentence = sentence.strip()
        if not sentence:
            continue

        sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

        if sentence_lines > max_lines_per_slide:
            if current_slide_text:
                slides.append(current_slide_text.strip())
                current_slide_text = ""
                current_slide_lines = 0
            
            wrapped_sentence_parts = textwrap.wrap(sentence, 
                                                   width=max_chars_per_line_ppt * max_lines_per_slide,
                                                   break_long_words=False, 
                                                   replace_whitespace=False, 
                                                   drop_whitespace=False)
            for part in wrapped_sentence_parts:
                slides.append(part.strip())
            continue

        can_add_to_current_slide = (current_slide_lines + sentence_lines <= max_lines_per_slide)
        
        is_similar_enough = True
        if current_slide_text and i > 0 and (i-1) < len(embeddings) and i < len(embeddings) :
            similarity = util.cos_sim(embeddings[i-1], embeddings[i])[0][0]
            if similarity < similarity_threshold:
                is_similar_enough = False

        if can_add_to_current_slide and is_similar_enough:
            if current_slide_text:
                current_slide_text += "\n" + sentence
            else:
                current_slide_text = sentence
            current_slide_lines += sentence_lines
        else:
            if current_slide_text:
                slides.append(current_slide_text.strip())
            
            current_slide_text = sentence
            current_slide_lines = sentence_lines

    if current_slide_text:
        slides.append(current_slide_text.strip())

    split_flags = [False] * len(slides)
    if not slides:
        return [""], [False]
        
    return slides, split_flags

def create_ppt(slides_data, max_chars_per_line_in_ppt, font_size_pt,
               font_name_to_use, template_path=None, transition_type=None):
    """PPT 생성 (템플릿, 전환 효과, 글꼴 설정)"""

    if template_path:
        try:
            prs = Presentation(template_path)
        except Exception as e:
            logging.error(f"Error loading template '{template_path}': {e}. Using default template.", exc_info=True)
            prs = Presentation()
    else:
        prs = Presentation()

    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for i, slide_data in enumerate(slides_data):
        try:
            if template_path and len(prs.slides) > i:
                slide = prs.slides[i]  # Use existing slide from template
            else:
                slide_layout_index = 6  # 빈 레이아웃 (인덱스는 환경에 따라 다를 수 있음)
                slide_layout = prs.slide_layouts[slide_layout_index]
                slide = prs.slides.add_slide(slide_layout)

            # Clear existing content (shapes) on the slide
            for shape in list(slide.shapes):
                slide.shapes.element.remove(shape.element)

            text_content = slide_data['text']
            image_path = slide_data.get('image', None)  # Get image path from slide data

            if text_content:
                add_text_to_slide(slide, text_content, font_size_pt, font_name_to_use, 
                                 max_chars_per_line_in_ppt)

            if image_path:
                try:
                    add_image_to_slide(slide, image_path)
                except Exception as img_err:
                    logging.error(f"Error adding image to slide {i + 1}: {img_err}", exc_info=True)
                    st.error(f"Error adding image to slide {i + 1}. Please check the image path.")

            if transition_type:
                set_slide_transition(slide, transition_type)

            if slide_data.get('is_flagged', False):
                add_check_needed_shape(slide)
        
        except Exception as slide_err:
            logging.error(f"Error creating slide {i + 1}: {slide_err}", exc_info=True)
            st.error(f"Error creating slide {i + 1}. Please check the input data.")

    return prs

def add_text_to_slide(slide, text, font_size_pt, font_name_to_use, max_chars_per_line):
    """슬라이드에 텍스트 추가"""

    margin_horizontal = Inches(0.75)
    margin_vertical = Inches(0.75)
    left = margin_horizontal
    top = margin_vertical
    width = prs.slide_width - (2 * margin_horizontal)
    height = prs.slide_height - (2 * margin_vertical)

    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.clear()
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
    text_frame.word_wrap = True

    text_frame.margin_left = Inches(0.1)
    text_frame.margin_right = Inches(0.1)
    text_frame.margin_top = Inches(0.1)
    text_frame.margin_bottom = Inches(0.1)

    wrapped_lines = textwrap.wrap(text, width=max_chars_per_line, 
                                  break_long_words=False, replace_whitespace=False, 
                                  drop_whitespace=True)

    for line_text in wrapped_lines:
        p = text_frame.add_paragraph()
        p.text = line_text
        p.font.size = Pt(font_size_pt)
        try:
            p.font.name = font_name_to_use
        except Exception as font_e:
            logging.warning(f"Font '{font_name_to_use}' not found: {font_e}. Using default font.")
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT

def add_image_to_slide(slide, image_path):
    """슬라이드에 이미지 추가"""

    img = Image.open(image_path)
    img_width, img_height = img.size

    left = Inches(1)
    top = Inches(1)
    width = Inches(5)  # 적절한 기본값
    height = Inches(3)

    slide.shapes.add_picture(image_path, left, top, width=width, height=height)

def set_slide_transition(slide, transition_type):
    """슬라이드 전환 효과 설정"""

    try:
        transition = slide.transition
        transition.type = transition_type
        transition.duration = 1  # Duration in seconds
    except Exception as e:
        logging.error(f"Error setting transition: {e}", exc_info=True)
        st.warning(f"Transition effect '{transition_type}' could not be applied.")

def add_check_needed_shape(slide):
    """확인 필요 표시 추가"""

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), 
                                  Inches(1.5), Inches(0.3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
    tf = shape.text_frame
    tf.text = "확인 필요"
    tf.paragraphs[0].font.size = Pt(10)
    tf.paragraphs[0].font.bold = True
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# --- Streamlit UI ---

# Input
uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=200)

# Sidebar Options
st.sidebar.header("⚙️ PPT 설정")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.sidebar.slider("한 줄당 최대 글자 수", 10, 100, 18)
font_size = st.sidebar.slider("폰트 크기", 10, 60, 54)
font_name = st.sidebar.text_input("폰트 이름", "맑은 고딕")  # Font selection
sim_threshold = st.sidebar.slider("문맥 유사도 기준", 0.0, 1.0, 0.85, step=0.05)
template_option = st.sidebar.selectbox("PPT 템플릿 선택", 
                                      ["기본", "템플릿 1", "템플릿 2", "사용자 지정"], 
                                      index=0)
template_path = None
if template_option == "사용자 지정":
    template_path = st.sidebar.file_uploader("템플릿 파일 업로드 (.pptx)", type=["pptx"])
    if template_path:
        template_path = BytesIO(template_path.read())  # Read the file into BytesIO

transition_option = st.sidebar.selectbox("슬라이드 전환 효과",
                                         [None, "None", "Fade", "Push", "Wipe", "Split"],
                                         index=0)
transition_type = None
if transition_option:
    transition_type = getattr(MSO_TRANSITION_TYPE, transition_option.upper(), None)
    if not transition_type:
        st.sidebar.warning("Invalid transition effect selected.")

# Main Process
if st.button("✨ PPT 생성"):
    if uploaded_file or text_input:
        try:
            paragraphs = extract_text_from_word(uploaded_file) if uploaded_file else [p.strip() for p in text_input.split("\n\n") if p.strip()]
            if not paragraphs:
                st.error("입력된 텍스트가 없습니다.")
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
                    st.download_button("📥 PPT 다운로드", ppt_io, "paydo_script_ai.pptx",
                                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                    st.success(f"총 {len(slides)}개의 슬라이드가 생성되었습니다.")
                    if any(flags):
                        flagged = [i+1 for i, f in enumerate(flags) if f]
                        st.warning(f"⚠️ 확인이 필요한 슬라이드: {flagged}")
                else:
                    st.error("PPT 생성 실패.")

        except Exception as e:
            st.error(f"오류 발생: {e}")

    else:
        st.info("Word 파일을 업로드하거나 텍스트를 입력하세요.")