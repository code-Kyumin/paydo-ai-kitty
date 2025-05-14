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
from sentence_transformers import SentenceTransformer, util
import kss
from io import BytesIO
import logging

# Streamlit 설정
st.set_page_config(page_title="Paydo AI PPT", layout="wide")
st.title("🎬 AI PPT 생성기 (KoSimCSE + KSS 의미 단위 분할)")

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

@st.cache_resource
def load_model():
    """KoSimCSE 모델 로드 (캐싱)"""
    logging.info("Loading SentenceTransformer model...")
    model = SentenceTransformer("jhgan/ko-sbert-nli")
    logging.info("SentenceTransformer model loaded.")
    return model

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
        logging.error(f"Word 파일 처리 오류 상세: {e}", exc_info=True)
        return []

def calculate_text_lines(text, max_chars_per_line):
    """텍스트 줄 수 계산"""
    lines = 0
    for paragraph in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=False, replace_whitespace=False)
        lines += len(wrapped_lines) or 1
    return lines

def smart_sentence_split(text):
    """KSS를 이용한 문장 분리"""
    try:
        return kss.split_sentences(text)
    except Exception as e:
        logging.error(f"KSS sentence splitting error: {e}", exc_info=True)
        return [s.strip() for s in text.split('.') if s.strip()]


def is_incomplete(sentence):
    """불완전한 문장 여부 확인 (서술어 잘림 방지 강화)"""
    sentence_stripped = sentence.strip()
    if len(sentence_stripped) < 10:
        return True
    if sentence_stripped.endswith(('은','는','이','가','을','를','에','으로','고','와','과', '며', '는데', '지만', '거나', '든지', '든지간에', '든가')):
        return True
    if re.match(r'^(그리고|하지만|그러나|또한|그래서|즉|또|그러면|그런데)$', sentence_stripped):
        return True
    if not sentence_stripped.endswith(('.', '!', '?', '다', '요', '죠', '까', '나', '시오')) and len(sentence_stripped) < 15:
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

def split_text_into_slides_with_similarity(text_paragraphs, max_lines_per_slide, max_chars_per_line_ppt, model, similarity_threshold=0.85):
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
            
            wrapped_sentence_parts = textwrap.wrap(sentence, width=max_chars_per_line_ppt * max_lines_per_slide,
                                                   break_long_words=False, replace_whitespace=False, drop_whitespace=False)
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


def create_ppt(slides, split_flags, max_chars_per_line_in_ppt, font_size_pt):
    """PPT 생성 (16:9 비율 및 폰트 수정)"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # --- 폰트 이름 '맑은 고딕'으로 변경 ---
    # Streamlit Cloud 환경에 '맑은 고딕'이 없을 경우, 기본 폰트로 대체될 수 있습니다.
    font_name_to_use = '맑은 고딕'
    # --- 폰트 이름 변경 끝 ---


    for i, slide_text_content in enumerate(slides):
        try:
            slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(slide_layout)

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
            
            wrapped_lines = textwrap.wrap(slide_text_content, 
                                          width=max_chars_per_line_in_ppt, 
                                          break_long_words=False,
                                          replace_whitespace=False,
                                          drop_whitespace=True)

            for line_text in wrapped_lines:
                p = text_frame.add_paragraph()
                p.text = line_text
                p.font.size = Pt(font_size_pt)
                try:
                    p.font.name = font_name_to_use
                except Exception as font_e:
                    logging.warning(f"Font '{font_name_to_use}' not found or could not be applied: {font_e}. Using default font.")
                p.font.bold = True
                p.alignment = PP_ALIGN.LEFT

            if i < len(split_flags) and split_flags[i]:
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(1.5), Inches(0.3))
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
                tf = shape.text_frame
                tf.text = "확인 필요"
                tf.paragraphs[0].font.size = Pt(10)
                tf.paragraphs[0].font.bold = True
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        except Exception as e:
            st.error(f"슬라이드 {i+1} 생성 중 오류 발생: {e}")
            logging.error(f"슬라이드 {i+1} 생성 오류 상세: {e}", exc_info=True)
    return prs

# --- Streamlit UI 부분 ---
uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=250)

# 슬라이드 옵션 (기본값 및 범위 조정)
st.sidebar.header("⚙️ 슬라이드 옵션")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수 (예상)", 3, 15, 5)
max_chars = st.sidebar.slider("한 줄당 최대 글자 수 (예상)", 20, 80, 35)
# --- 폰트 크기 슬라이더 설정 (기본값 54, 최대 65) ---
font_size = st.sidebar.slider(
    "폰트 크기 (Pt)",
    min_value=18,
    max_value=65,
    value=54,
    step=1
)
# --- 폰트 크기 슬라이더 설정 끝 ---
sim_threshold = st.sidebar.slider("문장 병합 유사도 기준 (낮을수록 많이 병합)", 0.5, 0.95, 0.75, step=0.05)

if st.button("🚀 PPT 생성"):
    paragraphs = []
    if uploaded_file:
        st.write(f"'{uploaded_file.name}' 파일 처리 중...")
        paragraphs = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        st.write("입력된 텍스트 처리 중...")
        paragraphs = [p.strip() for p in re.split(r'\n\s*\n', text_input) if p.strip()]
    else:
        st.warning("Word 파일을 업로드하거나 텍스트를 입력하세요.")
        st.stop()

    if not paragraphs:
        st.error("유효한 텍스트가 없습니다. Word 파일 내용 또는 입력된 텍스트를 확인해주세요.")
        st.stop()

    logging.info(f"입력된 문단 수: {len(paragraphs)}")
    if paragraphs:
         logging.debug(f"첫 번째 문단 내용 (일부): {paragraphs[0][:100]}")

    with st.spinner("AI가 열심히 PPT를 만들고 있어요... 잠시만 기다려주세요! ☕️"):
        try:
            logging.info("Splitting text into slides...")
            slides_content, slide_flags = split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, sim_threshold)
            
            if not slides_content or (len(slides_content) == 1 and not slides_content[0].strip()):
                st.error("슬라이드로 변환할 내용이 생성되지 않았습니다. 입력 텍스트나 분할 로직을 확인해주세요.")
                st.stop()

            logging.info(f"생성될 슬라이드 수: {len(slides_content)}")
            
            logging.info("Creating PPT...")
            ppt = create_ppt(slides_content, slide_flags, max_chars, font_size)

            if ppt:
                ppt_bytes = BytesIO()
                ppt.save(ppt_bytes)
                ppt_bytes.seek(0)
                
                st.success(f"🎉 와우! 총 {len(slides_content)}개의 슬라이드가 포함된 PPT가 완성되었습니다!")
                
                from datetime import datetime
                now = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_filename = f"paydo_script_ai_{now}.pptx"

                st.download_button(
                    label="📥 PPT 다운로드 (16:9)",
                    data=ppt_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
                if any(slide_flags):
                    flagged_indices = [i + 1 for i, flag in enumerate(slide_flags) if flag]
                    st.warning(f"⚠️ 다음 슬라이드 번호를 확인해주세요: {flagged_indices}")
            else:
                st.error("PPT 생성에 실패했습니다. 로그를 확인해주세요.")

        except Exception as e:
            st.error(f"PPT 생성 과정 중 심각한 오류 발생: {e}")
            logging.error(f"PPT 생성 전체 프로세스 오류 상세: {e}", exc_info=True)