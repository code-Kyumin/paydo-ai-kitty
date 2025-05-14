import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR # MSO_VERTICAL_ANCHOR 위치 변경
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
st.set_page_config(page_title="Paydo AI PPT", layout="wide") # layout="wide"로 변경하여 16:9 비율에 더 적합하게
st.title("🎬 AI PPT 생성기 (KoSimCSE + KSS 의미 단위 분할)")

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s') # DEBUG에서 INFO로 변경 (선택 사항)

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
        # break_long_words=False 로 변경하여 단어 중간 잘림 최소화 시도 (한글의 경우 효과 미미할 수 있음)
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=False, replace_whitespace=False)
        lines += len(wrapped_lines) or 1 # 빈 줄도 1줄로 처리
    return lines

def smart_sentence_split(text):
    """KSS를 이용한 문장 분리"""
    try:
        return kss.split_sentences(text)
    except Exception as e:
        logging.error(f"KSS sentence splitting error: {e}", exc_info=True)
        # KSS 오류 시 기본적인 문장 분리 (마침표 기준) 또는 원본 반환
        return [s.strip() for s in text.split('.') if s.strip()]


def is_incomplete(sentence):
    """불완전한 문장 여부 확인 (서술어 잘림 방지 강화)"""
    sentence_stripped = sentence.strip()
    # 매우 짧은 문장 (10자 미만, 사용자 설정 가능)
    if len(sentence_stripped) < 10: # 최소 문장 길이를 늘려 짧은 문장 병합 유도
        return True
    # 특정 조사/어미/연결어로 끝나는 경우
    # 예: '다.', '요.', '죠.', '~이다.' 등으로 끝나지 않는 경우 불완전으로 간주할 수 있도록 조건 추가
    if sentence_stripped.endswith(('은','는','이','가','을','를','에','으로','고','와','과', '며', '는데', '지만', '거나', '든지', '든지간에', '든가')):
        return True
    # 문장이 특정 접속 부사로만 이루어진 경우
    if re.match(r'^(그리고|하지만|그러나|또한|그래서|즉|또|그러면|그런데)$', sentence_stripped):
        return True
    # 문장 부호 없이 명사형으로 끝나는 매우 짧은 어구 (추가적인 품사 분석 없이 단순 길이로 제한)
    # 예: "중요한 것은" (O), "중요한 것" (X, 너무 짧으면 불완전 간주)
    # 이 부분은 더 정교한 NLP 분석이 필요할 수 있으나, 우선은 길이와 끝 글자로 판단
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
            # 버퍼와 현재 문장을 합쳤을 때 너무 길어지면 버퍼를 먼저 추가
            # (max_chars_per_line_ppt 값을 여기서는 알 수 없으므로, 일반적인 문장 길이로 제한)
            if len(current_candidate) > 200: # 임의의 최대 문장 길이, 필요시 조정
                merged.append(buffer)
                buffer = sentence
            else:
                buffer = current_candidate

            if not is_incomplete(buffer) or i == len(sentences) - 1: # 마지막 문장이거나, 합쳐진 문장이 완전하면
                merged.append(buffer)
                buffer = ""
        else:
            if is_incomplete(sentence) and i < len(sentences) - 1: # 마지막 문장이 아니면서 불완전하면 버퍼에 저장
                buffer = sentence
            else: # 완전한 문장이거나, 불완전해도 마지막 문장이면 그냥 추가
                merged.append(sentence)
                buffer = "" # 버퍼 초기화
    
    if buffer: # 루프 후 버퍼에 남은 내용 처리
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
        return [""], [False] # 빈 텍스트 처리

    merged_sentences = merge_sentences(all_sentences)
    
    if not merged_sentences:
        return [""], [False]

    embeddings = model.encode(merged_sentences)

    for i, sentence in enumerate(merged_sentences):
        sentence = sentence.strip()
        if not sentence:
            continue

        sentence_lines = calculate_text_lines(sentence, max_chars_per_line_ppt)

        # 한 문장이 슬라이드 최대 줄 수를 넘는 경우 (매우 긴 문장)
        if sentence_lines > max_lines_per_slide:
            if current_slide_text: # 이전까지의 내용을 먼저 슬라이드로 만듦
                slides.append(current_slide_text.strip())
                current_slide_text = ""
                current_slide_lines = 0
            
            # 긴 문장을 여러 슬라이드로 분할
            wrapped_sentence_parts = textwrap.wrap(sentence, width=max_chars_per_line_ppt * max_lines_per_slide, # 슬라이드당 총 글자수 기준으로 분할 시도
                                                   break_long_words=False, replace_whitespace=False, drop_whitespace=False)
            for part in wrapped_sentence_parts:
                slides.append(part.strip())
            continue

        # 현재 슬라이드에 추가할 수 있는지 확인
        can_add_to_current_slide = (current_slide_lines + sentence_lines <= max_lines_per_slide)
        
        # 유사도 검사 (첫 문장이 아니고, 현재 슬라이드에 내용이 있을 때)
        is_similar_enough = True
        if current_slide_text and i > 0 and (i-1) < len(embeddings) and i < len(embeddings) : # Ensure valid indices for embeddings
            # 이전 문장과의 유사도가 아닌, 현재 슬라이드의 마지막 문장과 다음 문장의 유사도
            # 이를 위해서는 현재 슬라이드의 마지막 문장을 알아야 함. 단순화를 위해 이전 문장과의 유사도 사용
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
            # 슬라이드 나누기
            if current_slide_text: # 기존 내용이 있으면 슬라이드로 추가
                slides.append(current_slide_text.strip())
            
            current_slide_text = sentence # 새 슬라이드 시작
            current_slide_lines = sentence_lines

    # 마지막 남은 텍스트 추가
    if current_slide_text:
        slides.append(current_slide_text.strip())

    # split_flags는 현재 로직에서 명시적으로 사용되지 않으므로, 모두 False로 반환하거나 관련 로직 추가 필요
    # 여기서는 모든 슬라이드가 정상적으로 분리되었다고 가정
    split_flags = [False] * len(slides)
    if not slides: # 아무 슬라이드도 생성되지 않은 경우
        return [""], [False]
        
    return slides, split_flags


def create_ppt(slides, split_flags, max_chars_per_line_in_ppt, font_size_pt):
    """PPT 생성 (16:9 비율 및 폰트 수정)"""
    prs = Presentation()
    prs.slide_width = Inches(13.333) # 16:9 너비 (1280px / 96dpi)
    prs.slide_height = Inches(7.5)   # 16:9 높이 (720px / 96dpi)

    # 한글 표시가 원활한 폰트 (Streamlit Cloud 환경에 따라 사용 가능 여부 확인 필요)
    # 사용 가능한 폰트가 없을 경우, 시스템 기본 폰트가 사용됨
    # font_name = '맑은 고딕' # Windows 기본
    font_name = 'NanumGothic' # 나눔고딕 (Streamlit Cloud에 없을 수 있음)
    # font_name = 'Noto Sans KR' # 웹폰트지만, 시스템에 설치되어 있어야 함
    # 특정 폰트 지정 대신 None 또는 기본값으로 두면 환경에 따라 자동 선택됨.
    # 안전하게는 폰트 이름을 지정하지 않거나, Streamlit Cloud의 기본 제공 폰트 확인 필요.
    # 여기서는 예시로 '맑은 고딕'을 사용하되, 주석 처리하여 사용자가 선택하도록 함.
    # 실제 적용 시에는 아래 p.font.name = font_name 부분에서 주석 해제 또는 변경
    # font_name_to_use = '맑은 고딕'
    font_name_to_use = 'Arial' # Arial은 대부분의 환경에서 사용 가능 (테스트용)
                               # 한글이 제대로 나오는지 확인 필요. 'Malgun Gothic' 등 시도.

    for i, slide_text_content in enumerate(slides):
        try:
            # 빈 슬라이드 레이아웃 사용 (인덱스 5 또는 6 - 환경에 따라 다를 수 있음, 6이 Blank인 경우가 많음)
            slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(slide_layout)

            # 텍스트 상자 위치 및 크기 (16:9 슬라이드에 맞게 조정)
            # 상하좌우 여백을 고려하여 설정
            margin_horizontal = Inches(0.75)
            margin_vertical = Inches(0.75)
            left = margin_horizontal
            top = margin_vertical
            width = prs.slide_width - (2 * margin_horizontal)
            height = prs.slide_height - (2 * margin_vertical)

            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            text_frame.clear() # 이전 텍스트 제거
            text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP # 상단 정렬
            text_frame.word_wrap = True # 자동 줄 바꿈 활성화

            # 텍스트 프레임 내부 여백 (선택 사항)
            text_frame.margin_left = Inches(0.1)
            text_frame.margin_right = Inches(0.1)
            text_frame.margin_top = Inches(0.1)
            text_frame.margin_bottom = Inches(0.1)
            
            # textwrap.wrap을 사용하여 슬라이드 텍스트를 다시 한 번 줄바꿈 처리
            # 이는 max_chars_per_line_in_ppt 기준으로 텍스트를 나누기 위함
            wrapped_lines = textwrap.wrap(slide_text_content, 
                                          width=max_chars_per_line_in_ppt, 
                                          break_long_words=False, # 단어 단위 줄바꿈 (한글에서는 어절 단위)
                                          replace_whitespace=False, # 공백 유지
                                          drop_whitespace=True) # 양 끝 공백 제거

            for line_text in wrapped_lines:
                p = text_frame.add_paragraph()
                p.text = line_text
                p.font.size = Pt(font_size_pt)
                try:
                    p.font.name = font_name_to_use # 폰트 이름 적용
                except Exception as font_e:
                    logging.warning(f"Font '{font_name_to_use}' not found or could not be applied: {font_e}. Using default font.")
                    # 폰트 적용 실패 시 기본 폰트 사용
                p.font.bold = True
                p.alignment = PP_ALIGN.LEFT # 중앙 정렬 대신 왼쪽 정렬로 변경 (가독성 고려)

            # 확인 필요 슬라이드 표시 로직 (필요하다면 유지)
            if i < len(split_flags) and split_flags[i]: # split_flags 길이 체크
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(1.5), Inches(0.3))
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor(255, 255, 0) # 노란색
                tf = shape.text_frame
                tf.text = "확인 필요"
                tf.paragraphs[0].font.size = Pt(10)
                tf.paragraphs[0].font.bold = True
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        except Exception as e:
            st.error(f"슬라이드 {i+1} 생성 중 오류 발생: {e}")
            logging.error(f"슬라이드 {i+1} 생성 오류 상세: {e}", exc_info=True)
            # 오류 발생 시 빈 PPT라도 반환할지, None을 반환할지 결정
            # return None # 여기서 None을 반환하면 전체 PPT 생성이 중단됨
    return prs

# --- Streamlit UI 부분 ---
uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=250)

# 슬라이드 옵션 (기본값 및 범위 조정)
st.sidebar.header("⚙️ 슬라이드 옵션")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수 (예상)", 3, 15, 5) # 기본값 및 범위 변경
max_chars = st.sidebar.slider("한 줄당 최대 글자 수 (예상)", 20, 80, 35) # 기본값 및 범위 변경
font_size = st.sidebar.slider("폰트 크기 (Pt)", 18, 48, 28) # 기본값 및 범위 변경
sim_threshold = st.sidebar.slider("문장 병합 유사도 기준 (낮을수록 많이 병합)", 0.5, 0.95, 0.75, step=0.05) # 기본값 및 설명 변경

if st.button("🚀 PPT 생성"):
    paragraphs = []
    if uploaded_file:
        st.write(f"'{uploaded_file.name}' 파일 처리 중...")
        paragraphs = extract_text_from_word(uploaded_file)
    elif text_input.strip():
        st.write("입력된 텍스트 처리 중...")
        # 여러 빈 줄을 하나의 문단 구분으로 처리
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
                
                # 다운로드 파일 이름에 현재 날짜/시간 포함 (선택 사항)
                from datetime import datetime
                now = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_filename = f"paydo_script_ai_{now}.pptx"

                st.download_button(
                    label="📥 PPT 다운로드 (16:9)",
                    data=ppt_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
                # 확인 필요한 슬라이드가 있다면 메시지 표시
                if any(slide_flags): # slide_flags가 사용된다면
                    flagged_indices = [i + 1 for i, flag in enumerate(slide_flags) if flag]
                    st.warning(f"⚠️ 다음 슬라이드 번호를 확인해주세요: {flagged_indices}")
            else:
                st.error("PPT 생성에 실패했습니다. 로그를 확인해주세요.")

        except Exception as e:
            st.error(f"PPT 생성 과정 중 심각한 오류 발생: {e}")
            logging.error(f"PPT 생성 전체 프로세스 오류 상세: {e}", exc_info=True)