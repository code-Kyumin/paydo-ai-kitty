import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt  # Cm은 현재 사용되지 않아 제거
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
# MSO_SHAPE_TYPE, MSO_THEME_COLOR_INDEX, MSO_AUTO_SIZE는 현재 사용되지 않아 제거

import io
import re
import textwrap
import docx  # python-docx 라이브러리
from io import BytesIO
from sentence_transformers import SentenceTransformer, util
import kss
import logging
import time
# from PIL import Image  # 현재 코드에서 PIL Image 직접 사용 안 함 (필요시 추가)
import math  # ceil 함수 사용 시 필요 (현재 코드에서는 직접 사용 안 함)
import torch  # torch 임포트 추가 (주의: requirements.txt와 버전 일치시켜야 함)

# --- Streamlit 페이지 설정 ---
st.set_page_config(page_title="AI 촬영 대본 PPT 생성기", layout="wide")
st.title("🎬 AI 촬영 대본 PPT 생성기")
st.caption("텍스트를 입력하면 촬영 대본 형식의 PPT를 자동으로 생성해주는 도구입니다.")

# --- 로깅 설정 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- 모델 로드 ---
@st.cache_resource  # 리소스 캐싱 (모델 로드에 적합)
def load_sbert_model():
    logger.info("SentenceTransformer 모델 로드를 시작합니다...")
    try:
        model = SentenceTransformer("jhgan/ko-sbert-nli")  # 모델명 확인 필요 (최신 모델로 변경 고려)
        logger.info("SentenceTransformer 모델 로드 완료.")
        return model
    except Exception as e:
        logger.error(f"SentenceTransformer 모델 로드 중 오류 발생: {e}", exc_info=True)
        st.error(f"모델 로드 중 오류가 발생했습니다: {e}")
        return None  # 모델 로드 실패 시 None 반환

model = load_sbert_model()

# --- Word 파일에서 텍스트 추출 ---
def extract_text_from_word(uploaded_file):
    logger.info(f"'{uploaded_file.name}'에서 텍스트 추출 시작")
    try:
        doc = docx.Document(uploaded_file)
        full_text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        logger.info(f"'{uploaded_file.name}'에서 텍스트 추출 완료, 총 {len(full_text)} 문자")
        return full_text
    except docx.opc.exceptions.PackageNotFoundError:
        logger.error(f"'{uploaded_file.name}'은 유효한 .docx 파일이 아닙니다.")
        st.error(f"올바른 .docx 파일을 업로드해주세요.")
        return None
    except Exception as e:
        logger.error(f"'{uploaded_file.name}' 처리 중 예외 발생: {e}", exc_info=True)
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
        return None

# --- 문장 분리 함수 ---
def smart_sentence_split(text):
    try:
        return kss.split_sentences(text)
    except Exception as e:
        logger.warning(f"kss 문장 분리 실패, 기본 분리 시도: {e}")
        return re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s', text)

# --- 유사도 기반 슬라이드 분할 함수 ---
def split_text_into_slides_with_similarity(
    full_text, model, max_lines_per_slide, max_chars_per_line, similarity_threshold, update_progress_callback
):
    sentences = smart_sentence_split(full_text)
    
    # 모델 로드 실패 시 빈 리스트 반환하여 PPT 생성 중단
    if model is None:
        return [], []
    
    embeddings = model.encode(sentences)
    slides = []
    current_slide_text = ""
    current_slide_lines = 0
    review_flags = []  # 각 슬라이드의 검토 필요 여부 저장

    for i, (sentence, embedding) in enumerate(zip(sentences, embeddings)):
        num_lines = math.ceil(len(sentence) / max_chars_per_line)
        
        # 슬라이드 최대 줄 수 초과 시 강제 분할
        if num_lines > max_lines_per_slide:
            logger.warning(f"문장 길이로 인해 슬라이드 강제 분할: {sentence}")
            
            # 현재 슬라이드 내용 추가 (있을 경우)
            if current_slide_text:
                slides.append(current_slide_text.strip())
                review_flags.append(False)  # 이전 슬라이드는 검토 불필요
            
            # 긴 문장 분할하여 새 슬라이드 생성
            wrapped_lines = textwrap.wrap(sentence, width=max_chars_per_line, replace_whitespace=False)
            slides.extend(wrapped_lines)
            review_flags.extend([True] * len(wrapped_lines))  # 분할된 슬라이드는 검토 필요
            
            current_slide_text = ""
            current_slide_lines = 0
            continue

        # 유사도 검사 및 슬라이드 병합 로직 (이전과 동일)
        if current_slide_lines + num_lines <= max_lines_per_slide:
            if current_slide_text:
                similarity = util.cos_sim(embeddings[i-1], embedding).item()
                if similarity >= similarity_threshold:
                    current_slide_text += " " + sentence
                    current_slide_lines += num_lines
                else:
                    slides.append(current_slide_text.strip())
                    review_flags.append(False)  # 이전 슬라이드는 검토 불필요
                    current_slide_text = sentence
                    current_slide_lines = num_lines
            else:
                current_slide_text = sentence
                current_slide_lines = num_lines
        else:
            slides.append(current_slide_text.strip())
            review_flags.append(False)  # 이전 슬라이드는 검토 불필요
            current_slide_text = sentence
            current_slide_lines = num_lines

        update_progress_callback((i + 1) / len(sentences), f"슬라이드 분할 중 ({i + 1}/{len(sentences)})")

    if current_slide_text:
        slides.append(current_slide_text.strip())
        review_flags.append(False)  # 마지막 슬라이드 검토 불필요

    return slides, review_flags

# --- PPT 생성 함수 ---
def create_presentation(slides_content, review_flags, max_chars_per_line, font_size, update_progress_callback):
    presentation = Presentation()
    slide_width_inch = 16  # 슬라이드 가로 크기 (인치)
    slide_height_inch = 9   # 슬라이드 세로 크기 (인치)
    presentation.slide_width = Inches(slide_width_inch)
    presentation.slide_height = Inches(slide_height_inch)

    for i, (text, flag) in enumerate(zip(slides_content, review_flags)):
        try:
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])  # 빈 슬라이드 레이아웃
            
            # 텍스트 박스 추가 및 설정 (크기 조정)
            left = Inches(1)
            top = Inches(1)
            width = Inches(slide_width_inch - 2)
            height = Inches(slide_height_inch - 2)
            text_box = slide.shapes.add_textbox(left, top, width, height)
            text_frame = text_box.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # 상단 정렬
            
            # 텍스트 추가 및 스타일 설정
            p = text_frame.paragraphs[0]
            p.text = textwrap.fill(text, width=max_chars_per_line, replace_whitespace=False)
            p.font.size = Pt(font_size)
            p.font.bold = True
            p.alignment = PP_ALIGN.CENTER

            # "확인 필요" 도형 추가
            if flag:
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(2.0), Inches(0.5))
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
                tf = shape.text_frame
                tf.text = "확인 필요"
                tf.paragraphs[0].font.size = Pt(16)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER

            update_progress_callback((i + 1) / len(slides_content), f"PPT 슬라이드 생성 중 ({i + 1}/{len(slides_content)})")

        except Exception as e:
            logger.error(f"슬라이드 생성 중 오류 발생 (슬라이드 {i+1}): {e}", exc_info=True)
            update_progress_callback(0, f"오류 발생: {e}")  # 오류 발생 시 진행률 초기화
            return None  # 오류 발생 시 None 반환

    return presentation

# --- Streamlit UI 업데이트 함수 ---
def update_ui_progress(progress_value, progress_text):
    st.session_state.progress_bar.progress(progress_value, text=progress_text)

# --- 메인 로직 ---
uploaded_file = st.file_uploader("🎬 촬영 대본(.docx) 파일을 업로드하세요", type=["docx"])
text_input = st.text_area("또는 텍스트를 직접 입력하세요", height=200)

max_lines_option = st.sidebar.slider("슬라이드 당 최대 줄 수", 2, 15, 6)
max_chars_option = st.sidebar.slider("한 줄당 최대 글자 수", 20, 100, 40)
font_size_option = st.sidebar.slider("글자 크기", 10, 60, 24)
similarity_threshold_option = st.sidebar.slider("문장 유사도 병합 기준", 0.0, 1.0, 0.8, step=0.05)

if "progress_bar" not in st.session_state:
    st.session_state.progress_bar = st.empty()  # 초기화

if st.button("✨ PPT 생성 ✨"):
    if model is None:
        st.error("모델 로드에 실패했습니다. 잠시 후 다시 시도해주세요.")
    elif uploaded_file or text_input:
        try:
            if uploaded_file:
                full_text = extract_text_from_word(uploaded_file)
            else:
                full_text = text_input
            
            if full_text is None:
                st.error("Word 파일 처리 또는 텍스트 입력 과정에서 오류가 발생했습니다.")
                st.stop()
            
            st.session_state.progress_bar = st.progress(0, text="PPT 생성 준비 중...")  # 진행률 표시기 초기화
            
            generated_slides_content, review_flags = split_text_into_slides_with_similarity(
                full_text, model, max_lines_option, max_chars_option, similarity_threshold_option, update_ui_progress
            )
            
            # 모델 로드 실패 시 여기서도 체크 (split_text_into_slides_with_similarity 내부에서도 체크)
            if model is None:
                st.error("모델 로드에 실패했습니다. PPT 생성을 중단합니다.")
                st.stop()
            
            if not generated_slides_content:
                st.error("유효한 슬라이드 내용을 생성하지 못했습니다. 입력 텍스트를 확인해주세요.")
                st.stop()
            
            presentation_object = create_presentation(
                generated_slides_content, review_flags, 
                max_chars_option, font_size_option, update_ui_progress
            )
            
            if presentation_object is None:
                st.error("PPT 생성 중 오류가 발생했습니다. 로그를 확인해주세요.")
                st.stop()

            ppt_file_stream = BytesIO()
            presentation_object.save(ppt_file_stream)
            ppt_file_stream.seek(0)
            
            update_ui_progress(1.0, "PPT 생성 완료!")
            
            st.download_button(
                label="⬇️ 생성된 PPT 다운로드",
                data=ppt_file_stream,
                file_name="generated_presentation_ai.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_button_key"
            )
            st.success(f"🎉 PPT 생성이 완료되었습니다! 총 {len(generated_slides_content)}개의 슬라이드가 만들어졌습니다.")
            
            flagged_slide_indices = [idx + 1 for idx, flag_val in enumerate(review_flags) if flag_val]
            if flagged_slide_indices:
                st.warning(f"⚠️ 다음 슬라이드는 내용이 길거나 구성상 강제 분할되었을 수 있으니 확인해주세요: {', '.join(map(str, flagged_slide_indices))}")

        except Exception as e:
            logger.error("PPT 생성 과정에서 심각한 오류 발생:", exc_info=True)
            st.error(f"PPT 생성 중 오류가 발생했습니다: {e}. 로그를 확인해주세요.")
            update_ui_progress(0, f"오류 발생: {e}")  # 오류 시 진행률 초기화

# --- 앱 하단 정보 ---
st.markdown("""
<br><br><br>
---
**AI PPT 생성기**는 텍스트를 입력하면 자동으로 촬영 대본 형식의 PPT를 생성해주는 도구입니다.

**주요 기능:**

* **텍스트 기반 PPT 생성:** 입력된 텍스트를 분석하여 슬라이드를 자동으로 구성합니다.
* **다양한 설정 옵션:** 슬라이드 당 최대 줄 수, 한 줄당 최대 글자 수, 글자 크기, 문장 유사도 병합 기준 등을 설정할 수 있습니다.
* **강력한 자연어 처리 모델:** KoSimCSE 모델을 사용하여 문맥을 이해하고 자연스러운 슬라이드를 생성합니다.

**문의사항:**

* """)