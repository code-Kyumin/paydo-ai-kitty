import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt # Cm은 현재 사용되지 않아 제거
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
# MSO_SHAPE_TYPE, MSO_THEME_COLOR_INDEX, MSO_AUTO_SIZE는 현재 사용되지 않아 제거

import io
import re
import textwrap
import docx # python-docx 라이브러리
from io import BytesIO
from sentence_transformers import SentenceTransformer, util
import kss
import logging
import time
# from PIL import Image # 현재 코드에서 PIL Image 직접 사용 안 함 (필요시 추가)
import math # ceil 함수 사용 시 필요 (현재 코드에서는 직접 사용 안 함)

# --- Streamlit 페이지 설정 ---
st.set_page_config(page_title="AI 촬영 대본 PPT 생성기", layout="wide")
st.title("🎬 AI 촬영 대본 PPT 생성기")
st.caption("텍스트를 입력하면 촬영 대본 형식의 PPT를 자동으로 생성합니다.")

# --- 로깅 설정 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- 모델 로드 ---
@st.cache_resource # 리소스 캐싱 (모델 로드에 적합)
def load_sbert_model():
    logger.info("SentenceTransformer 모델 로드를 시작합니다...")
    try:
        model = SentenceTransformer("jhgan/ko-sbert-nli") # 모델명 확인 필요
        logger.info("SentenceTransformer 모델 로드가 완료되었습니다.")
        return model
    except Exception as e:
        logger.error(f"SentenceTransformer 모델 로드 중 오류 발생: {e}", exc_info=True)
        st.error(f"모델 로드에 실패했습니다: {e}. 인터넷 연결 및 모델명을 확인해주세요.")
        return None

sbert_model = load_sbert_model()
if sbert_model is None:
    st.error("모델 초기화 실패로 인해 앱을 실행할 수 없습니다. 관리자에게 문의하거나 잠시 후 다시 시도해주세요.")
    st.stop()


# --- 텍스트 처리 함수 ---
def extract_text_from_word(uploaded_file_obj):
    logger.info(f"'{uploaded_file_obj.name}' 파일에서 텍스트 추출을 시작합니다.")
    try:
        uploaded_file_obj.seek(0) # 파일 포인터 초기화
        doc = docx.Document(BytesIO(uploaded_file_obj.read()))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        if not paragraphs:
            logger.warning(f"'{uploaded_file_obj.name}' 파일에 추출할 텍스트 내용이 없습니다.")
            st.warning("업로드된 Word 파일에 내용이 없거나 읽을 수 있는 텍스트가 없습니다.")
        return paragraphs
    except docx.opc.exceptions.PackageNotFoundError:
        logger.error(f"'{uploaded_file_obj.name}' 파일은 유효한 .docx 형식이 아닙니다.")
        st.error("올바른 .docx 파일을 업로드해주세요. (.doc 파일은 지원하지 않습니다)")
        return []
    except Exception as e:
        logger.error(f"'{uploaded_file_obj.name}' 파일 처리 중 예상치 못한 오류 발생: {e}", exc_info=True)
        st.error(f"Word 파일 처리 중 오류가 발생했습니다: {e}")
        return []

def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    if not text: return 0
    for paragraph_part in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph_part, width=max_chars_per_line, break_long_words=True, replace_whitespace=False)
        lines += len(wrapped_lines) if wrapped_lines else 1
    return max(1, lines) # 최소 1줄로 처리

def smart_sentence_split_kss(text_block):
    try:
        sentences = kss.split_sentences(text_block)
        return sentences
    except Exception as e:
        logger.warning(f"KSS 문장 분리 실패 (기본 분리 사용): {e}")
        # KSS 실패 시 간단한 구두점 기반 분리 (정규표현식 개선 가능)
        sentences = re.split(r'(?<=[.!?])\s+', text_block.strip())
        return [s.strip() for s in sentences if s.strip()]

def is_short_non_sentence(sentence_text, min_char_len=5):
    sentence_text = sentence_text.strip()
    if not sentence_text: return False
    common_sentence_endings = ('.', '!', '?', '다', '요', '죠', '까', '시오', '습니다', '합니다')
    if len(sentence_text) < min_char_len and not sentence_text.endswith(common_sentence_endings):
        return True
    return False

def is_incomplete_sentence(sentence_text):
    sentence_text = sentence_text.strip()
    if not sentence_text: return False
    # 어미 기반 불완전성 판단 (좀 더 보수적으로)
    incomplete_endings = ('은', '는', '이', '가', '을', '를', '에', '으로', '고', '와', '과', '며', '는데', '지만')
    # 길이가 짧으면서 특정 조사/어미로 끝나면 불완전으로 간주
    if len(sentence_text) < 15 and sentence_text.endswith(incomplete_endings):
        return True
    return False

def merge_script_sentences(sentences_list, max_segment_chars=250):
    merged_segments = []
    current_segment_buffer = ""
    for i, sentence in enumerate(sentences_list):
        sentence = sentence.strip()
        if not sentence: continue

        if is_short_non_sentence(sentence): # 매우 짧은 비문장성 텍스트는 가능하면 단독 처리
            if current_segment_buffer:
                merged_segments.append(current_segment_buffer)
                current_segment_buffer = ""
            merged_segments.append(sentence)
            continue

        if not current_segment_buffer:
            current_segment_buffer = sentence
        else:
            potential_segment = current_segment_buffer + " " + sentence
            if len(potential_segment) <= max_segment_chars:
                current_segment_buffer = potential_segment
            else: # 너무 길어지면 현재 버퍼를 확정하고 새 버퍼 시작
                merged_segments.append(current_segment_buffer)
                current_segment_buffer = sentence
        
        # 문맥상 완전하거나, 마지막 문장이거나, 다음 문장과 합치면 너무 길어질 것 같으면 현재 버퍼 확정
        if not is_incomplete_sentence(current_segment_buffer) or \
           i == len(sentences_list) - 1 or \
           (i + 1 < len(sentences_list) and len(current_segment_buffer + " " + sentences_list[i+1].strip()) > max_segment_chars) :
            if current_segment_buffer:
                merged_segments.append(current_segment_buffer)
                current_segment_buffer = ""

    if current_segment_buffer: # 남은 버퍼 추가
        merged_segments.append(current_segment_buffer)
    
    return [seg for seg in merged_segments if seg]


def split_text_into_slides(text_paragraphs, max_lines_per_slide, max_chars_per_line, sentence_model, similarity_threshold, progress_callback_func):
    logger.info("슬라이드 분할 로직을 시작합니다.")
    final_slides_text = []
    final_split_flags = []

    # 1. 문단에서 모든 문장 추출 및 기본 병합
    progress_callback_func(0.05, "텍스트 전처리 중...")
    all_raw_sentences = [s for para in text_paragraphs for s in smart_sentence_split_kss(para)]
    meaningful_segments = merge_script_sentences(all_raw_sentences)
    
    if not meaningful_segments:
        logger.warning("분할할 유의미한 텍스트 세그먼트가 없습니다.")
        return [""], [False]

    progress_callback_func(0.15, "문장 임베딩 생성 중...")
    try:
        segment_embeddings = sentence_model.encode(meaningful_segments, show_progress_bar=False) # UI에 이미 진행바 있음
    except Exception as e:
        logger.error(f"문장 임베딩 생성 중 오류: {e}", exc_info=True)
        st.error(f"문장 분석 중 오류가 발생했습니다: {e}")
        return [""], [False] # 또는 적절한 오류 처리된 슬라이드 반환

    current_slide_buffer = ""
    current_slide_line_count = 0
    last_added_segment_embedding = None

    # 2. 의미적 유사도 기반 슬라이드 분할
    for i, segment in enumerate(meaningful_segments):
        progress_callback_func(0.15 + (0.5 * (i / len(meaningful_segments))), f"슬라이드 내용 구성 중 ({i+1}/{len(meaningful_segments)})...")
        
        segment_line_count = calculate_text_lines(segment, max_chars_per_line)

        # 한 세그먼트 자체가 최대 줄 수를 넘는 경우 (긴급 분할)
        if segment_line_count > max_lines_per_slide:
            if current_slide_buffer: # 이전 버퍼가 있으면 먼저 슬라이드로
                final_slides_text.append(current_slide_buffer.strip())
                final_split_flags.append(False)
                current_slide_buffer, current_slide_line_count = "", 0
                last_added_segment_embedding = None
            
            # 긴 세그먼트를 줄 단위로 나누어 슬라이드 생성
            wrapped_lines = textwrap.wrap(segment, width=max_chars_per_line, break_long_words=True, replace_whitespace=False)
            temp_long_segment_slide = ""
            temp_long_segment_lines = 0
            for line in wrapped_lines:
                if temp_long_segment_lines + 1 <= max_lines_per_slide:
                    temp_long_segment_slide += line + "\n"
                    temp_long_segment_lines += 1
                else:
                    final_slides_text.append(temp_long_segment_slide.strip())
                    final_split_flags.append(True) # 강제 분할 플래그
                    temp_long_segment_slide = line + "\n"
                    temp_long_segment_lines = 1
            if temp_long_segment_slide: # 남은 부분
                final_slides_text.append(temp_long_segment_slide.strip())
                final_split_flags.append(True)
            last_added_segment_embedding = segment_embeddings[i] # 이 긴 세그먼트의 임베딩 사용
            continue

        # 일반적인 경우: 현재 슬라이드에 추가할지, 새 슬라이드로 시작할지 결정
        should_start_new_slide = False
        if not current_slide_buffer: # 첫 세그먼트
            should_start_new_slide = False
        elif current_slide_line_count + segment_line_count > max_lines_per_slide: # 공간 부족
            should_start_new_slide = True
        elif last_added_segment_embedding is not None: # 유사도 체크
            similarity_score = util.cos_sim(last_added_segment_embedding, segment_embeddings[i])[0][0].item()
            if similarity_score < similarity_threshold:
                should_start_new_slide = True
        
        if should_start_new_slide and current_slide_buffer:
            final_slides_text.append(current_slide_buffer.strip())
            final_split_flags.append(False)
            current_slide_buffer = segment
            current_slide_line_count = segment_line_count
        else:
            current_slide_buffer = f"{current_slide_buffer}\n{segment}" if current_slide_buffer else segment
            current_slide_line_count += segment_line_count
        
        last_added_segment_embedding = segment_embeddings[i]

    if current_slide_buffer: # 마지막 남은 버퍼 슬라이드로 추가
        final_slides_text.append(current_slide_buffer.strip())
        final_split_flags.append(False)

    # 3. 짧은 슬라이드 병합 시도 (후처리)
    processed_slides_text = []
    processed_split_flags = []
    skip_next_slide_index = -1

    for i in range(len(final_slides_text)):
        progress_callback_func(0.65 + (0.15 * (i / len(final_slides_text))), f"슬라이드 최적화 중 ({i+1}/{len(final_slides_text)})...")
        if i <= skip_next_slide_index: continue

        current_text = final_slides_text[i]
        current_flag = final_split_flags[i]
        current_lines = calculate_text_lines(current_text, max_chars_per_line)

        if current_lines <= 2 and i + 1 < len(final_slides_text): # 2줄 이하이고 다음 슬라이드가 있다면
            next_text = final_slides_text[i+1]
            next_flag = final_split_flags[i+1]
            combined_text = current_text + "\n" + next_text
            combined_lines = calculate_text_lines(combined_text, max_chars_per_line)

            if combined_lines <= max_lines_per_slide: # 합쳐도 최대 줄 수를 넘지 않으면
                processed_slides_text.append(combined_text)
                processed_split_flags.append(current_flag or next_flag) # 둘 중 하나라도 True면 True
                skip_next_slide_index = i + 1 # 다음 슬라이드는 건너뜀
                logger.info(f"슬라이드 {i+1}과 {i+2}를 병합했습니다.")
                continue
        
        processed_slides_text.append(current_text)
        processed_split_flags.append(current_flag)
        
    if not processed_slides_text: # 모든 슬라이드가 비어있는 극단적인 경우 방지
         logger.warning("최적화 후 슬라이드 내용이 없습니다. 기본 빈 슬라이드를 반환합니다.")
         return [""], [False]

    logger.info(f"총 {len(processed_slides_text)}개의 슬라이드로 분할 완료.")
    return processed_slides_text, processed_split_flags


def create_presentation_from_slides(slides_data, slide_flags, chars_per_line, text_font_size, progress_callback_func):
    logger.info("PPT 생성을 시작합니다.")
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 ratio
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6] # 내용 없는 레이아웃 사용

    total_slides_to_create = len(slides_data)
    # 아래 라인의 변수명에서 공백을 제거합니다.
    for i, (slide_content, is_flagged_for_review) in enumerate(zip(slides_data, slide_flags)): # <--- 수정됨
        progress_callback_func(0.8 + (0.2 * (i / total_slides_to_create)), f"슬라이드 {i+1}/{total_slides_to_create} 생성 중...")
        
        slide = prs.slides.add_slide(blank_slide_layout)

        # 텍스트 영역 설정
        left_margin, top_margin = Inches(0.5), Inches(0.5)
        width = prs.slide_width - (left_margin * 2)
        height = prs.slide_height - (top_margin * 2) # 하단 여백 고려
        
        textbox = slide.shapes.add_textbox(left_margin, Inches(0.7), width, prs.slide_height - Inches(1.5)) # 상단 여백 약간 더 줌
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE # 기본 중앙 정렬

        cleaned_content = slide_content.strip()
        for line_text in cleaned_content.split('\n'): # 이미 \n으로 줄바꿈된 텍스트 사용
            p = text_frame.add_paragraph()
            p.text = line_text
            p.font.size = Pt(text_font_size)
            p.font.bold = True
            p.font.name = '맑은 고딕' # 기본 폰트
            p.alignment = PP_ALIGN.CENTER

        # "확인 필요" 도형 (요구사항 2)
        if is_flagged_for_review: # <--- 여기 변수명도 동일하게 수정됨
            flag_shape_width, flag_shape_height = Inches(2.2), Inches(0.6)
            flag_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), flag_shape_width, flag_shape_height)
            flag_shape.fill.solid()
            flag_shape.fill.fore_color.rgb = RGBColor(255, 255, 0) # 노란색
            
            flag_tf = flag_shape.text_frame
            flag_tf.text = "⚠️ 확인 필요"
            flag_p = flag_tf.paragraphs[0]
            flag_p.font.size = Pt(20)
            flag_p.font.name = '맑은 고딕'
            flag_p.font.bold = True
            flag_p.font.color.rgb = RGBColor(0, 0, 0) # 검은색 글자
            flag_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            flag_p.alignment = PP_ALIGN.CENTER

        # 페이지 번호 (요구사항 1) - 우측 하단
        pn_box_width, pn_box_height = Inches(0.8), Inches(0.3)
        pn_left = prs.slide_width - pn_box_width - Inches(0.3) # 우측 여백
        pn_top = prs.slide_height - pn_box_height - Inches(0.2) # 하단 여백
        
        pn_shape = slide.shapes.add_textbox(pn_left, pn_top, pn_box_width, pn_box_height)
        pn_tf = pn_shape.text_frame
        pn_tf.text = f"{i+1}/{total_slides_to_create}"
        pn_p = pn_tf.paragraphs[0]
        pn_p.font.size = Pt(10)
        pn_p.font.name = '맑은 고딕'
        pn_p.alignment = PP_ALIGN.RIGHT

        # "끝" 표시 (요구사항 9) - 마지막 슬라이드, 페이지 번호 왼쪽
        if i == total_slides_to_create - 1:
            end_mark_diameter = Inches(0.7)
            end_mark_left = pn_left - end_mark_diameter - Inches(0.1) # 페이지 번호 왼쪽에
            end_mark_top = pn_top + (pn_box_height / 2) - (end_mark_diameter / 2) # 페이지 번호와 수직 중앙 정렬 시도

            if end_mark_left < Inches(0.2) : end_mark_left = Inches(0.2) # 너무 왼쪽으로 가지 않도록

            end_mark_shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, end_mark_left, end_mark_top, end_mark_diameter, end_mark_diameter
            )
            end_mark_shape.fill.solid()
            end_mark_shape.fill.fore_color.rgb = RGBColor(255, 0, 0) # 빨간색
            
            end_mark_tf = end_mark_shape.text_frame
            end_mark_tf.text = "끝"
            end_mark_p = end_mark_tf.paragraphs[0]
            end_mark_p.font.size = Pt(20) # 원 크기 고려하여 조정
            end_mark_p.font.name = '맑은 고딕'
            end_mark_p.font.bold = True
            end_mark_p.font.color.rgb = RGBColor(255, 255, 255) # 흰색
            end_mark_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            end_mark_p.alignment = PP_ALIGN.CENTER
            
    logger.info(f"PPT 생성 완료. 총 {total_slides_to_create} 슬라이드.")
    return prs

# --- Streamlit UI 구성 ---
uploaded_word_file = st.file_uploader("1. Word 파일 업로드 (.docx):", type=["docx"], key="file_uploader_key")
raw_text_input = st.text_area("또는 2. 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=200, key="text_area_key")

st.sidebar.header("⚙️ PPT 생성 옵션")
max_lines_option = st.sidebar.slider("슬라이드당 최대 줄 수:", 1, 10, 4, key="max_lines_key")
max_chars_option = st.sidebar.slider("한 줄당 글자 수 (참고용):", 10, 100, 35, key="max_chars_key") # 이름 및 기본값 변경
font_size_option = st.sidebar.slider("본문 폰트 크기 (Pt):", 10, 70, 44, key="font_size_key") # 기본값 변경
similarity_threshold_option = st.sidebar.slider("문맥 유사도 기준 (낮을수록 잘 나눔):", 0.50, 1.00, 0.70, step=0.01, key="similarity_key") # 기본값 및 step 변경

if st.button("🚀 PPT 생성 실행!", key="generate_button_key", type="primary"):
    final_paragraphs = []
    if uploaded_word_file:
        final_paragraphs = extract_text_from_word(uploaded_word_file)
    elif raw_text_input:
        final_paragraphs = [p.strip() for p in raw_text_input.split("\n\n") if p.strip()]
    
    if not final_paragraphs:
        st.warning("PPT를 생성할 텍스트 내용이 없습니다. 파일을 업로드하거나 텍스트를 입력해주세요.")
    else:
        progress_bar_ui = st.progress(0)
        status_text_ui = st.empty()
        start_process_time = time.time()

        def update_ui_progress(progress_value, message_text):
            current_elapsed_time = time.time() - start_process_time
            # 예상 남은 시간 계산 (간단한 방식, 정확도 낮을 수 있음)
            estimated_remaining_time = int((current_elapsed_time / progress_value) * (1 - progress_value)) if progress_value > 0.01 and progress_value < 1.0 else 0
            
            status_text_ui.text(f"{message_text} - {int(progress_value*100)}% (예상 남은 시간: {estimated_remaining_time}초)")
            progress_bar_ui.progress(min(progress_value, 1.0))

        try:
            update_ui_progress(0.01, "PPT 생성 준비 중...")
            
            generated_slides_content, review_flags = split_text_into_slides(
                final_paragraphs, max_lines_option, max_chars_option, 
                sbert_model, similarity_threshold_option, update_ui_progress
            )
            
            if not generated_slides_content or (len(generated_slides_content) == 1 and not generated_slides_content[0]):
                 st.error("슬라이드로 변환할 내용을 찾지 못했습니다. 입력 텍스트를 확인해주세요.")
            else:
                presentation_object = create_presentation_from_slides(
                    generated_slides_content, review_flags, 
                    max_chars_option, font_size_option, update_ui_progress
                )
                
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
            update_ui_progress(0, f"오류 발생: {e}") # 오류 시 진행률 초기화

# --- 앱 하단 정보 ---
st.markdown("---")
st.markdown("AI 기반 촬영 대본 PPT 자동 생성 도구")
st.markdown(f"현재시간: {time.strftime('%Y-%m-%d %H:%M:%S %Z')}") # [2025-05-12] 이 부분은 수정되지 않도록 해줘.