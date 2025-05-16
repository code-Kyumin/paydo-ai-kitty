import sys
import asyncio

# 🧩 Python 3.12에서 Streamlit event loop 오류 우회
try:
    asyncio.get_running_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

# 🧩 PyTorch 내부 torch._classes 오류 회피
# import types
# sys.modules['torch._classes'] = types.SimpleNamespace()

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
import kss
import logging
import time

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

# Word 문서에서 텍스트 추출
from docx import Document
from docx.opc.exceptions import PackageNotFoundError

def extract_text_from_word(uploaded_file):
    try:
        # 파일 포인터를 처음으로 되돌림
        uploaded_file.seek(0)

        # 바이트 스트림으로 읽기
        file_bytes = BytesIO(uploaded_file.read())

        # docx 문서 로딩
        doc = Document(file_bytes)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return paragraphs

    except PackageNotFoundError:
        st.error("❌ 이 파일은 .docx 형식이 아닙니다. .docx 파일만 업로드해 주세요.")
        return []

    except Exception as e:
        st.error(f"❌ Word 파일 처리 중 오류가 발생했습니다: {e}")
        return []


# 텍스트 줄 수 계산
def calculate_text_lines(text, max_chars_per_line):
    lines = 0
    if not text:
        return 0
    for paragraph in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=False)
        lines += len(wrapped_lines) if wrapped_lines else 1
    return lines if lines > 0 else 1

# 문장 분할 (kss 또는 백업 정규식)
def smart_sentence_split(text):
    try:
        return kss.split_sentences(text)
    except Exception:
        return [s.strip() for s in re.split(r'[.!?]\s+', text) if s.strip()]
# 문장이 아닌 것으로 간주되는 짧은 문장 판단
def is_potentially_non_sentence(sentence_text, min_length=5):
    sentence_text = sentence_text.strip()
    if not sentence_text:
        return False
    if len(sentence_text) < min_length and not sentence_text.endswith(('.', '!', '?', '다', '요', '죠', '까', '나', '시오')):
        return True
    return False

# 불완전한 문장 판단 (어미 기반)
def is_incomplete(sentence):
    sentence = sentence.strip()
    if not sentence:
        return False
    incomplete_endings = ('은', '는', '이', '가', '을', '를', '에', '으로', '고', '와', '과', 
                          '며', '는데', '지만', '거나', '든지', '든지간에', '든가', '고,', '며,', '는데,')
    return sentence.endswith(incomplete_endings)

# 문장 병합 (요구사항 3, 4 반영)
def merge_sentences(sentences, max_chars_per_sentence_segment=200):
    merged = []
    buffer = ""
    for i, sentence in enumerate(sentences):
        sentence = sentence.strip()
        if not sentence:
            continue
        if is_potentially_non_sentence(sentence):
            if buffer:
                merged.append(buffer)
                buffer = ""
            merged.append(sentence)
            continue

        if buffer:
            current_candidate = buffer + " " + sentence
            if len(current_candidate) > max_chars_per_sentence_segment:
                merged.append(buffer)
                buffer = sentence
            elif not is_incomplete(buffer) and is_incomplete(sentence) and i < len(sentences) - 1:
                buffer = current_candidate
            elif not is_incomplete(buffer):
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
def split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, threshold=0.85, progress_callback=None):
    slides = []
    split_flags = []

    all_sentences_original = [s for p in paragraphs for s in smart_sentence_split(p)]
    merged_sentences = merge_sentences(all_sentences_original)

    if not merged_sentences:
        return [""], [False]

    if progress_callback:
        progress_callback(0.1, "문장 임베딩 중...")

    embeddings = model.encode(merged_sentences)
    current_text = ""
    current_lines = 0
    last_sentence_embedding = None

    for i, sentence in enumerate(merged_sentences):
        if progress_callback:
            progress_callback(0.1 + (0.5 * (i / len(merged_sentences))), f"슬라이드 분할 중 ({i+1}/{len(merged_sentences)})...")

        sentence_actual_lines = calculate_text_lines(sentence, max_chars)

        # 슬라이드당 최대 줄 수 초과하는 긴 문장 처리
        if sentence_actual_lines > max_lines:
            if current_text:
                slides.append(current_text.strip())
                split_flags.append(False)
                current_text, current_lines = "", 0
                last_sentence_embedding = None

            wrapped_sentence_lines = textwrap.wrap(sentence, width=max_chars, break_long_words=False)
            temp_slide_text = ""
            temp_slide_lines = 0
            for line_text in wrapped_sentence_lines:
                if temp_slide_lines + 1 <= max_lines:
                    temp_slide_text += line_text + "\n"
                    temp_slide_lines += 1
                else:
                    slides.append(temp_slide_text.strip())
                    split_flags.append(True)
                    temp_slide_text = line_text + "\n"
                    temp_slide_lines = 1
            if temp_slide_text:
                slides.append(temp_slide_text.strip())
                split_flags.append(True)
            last_sentence_embedding = embeddings[i]
            continue

        # 일반적인 경우
        if current_lines + sentence_actual_lines <= max_lines:
            similar_to_previous = True
            if current_text and last_sentence_embedding is not None and i < len(embeddings):
                prev_emb = embeddings[i-1] if i > 0 else None
                if prev_emb is not None:
                    sim = util.cos_sim(prev_emb, embeddings[i])[0][0].item()
                    if sim < threshold:
                        similar_to_previous = False
            if similar_to_previous:
                current_text = f"{current_text}\n{sentence}" if current_text else sentence
                current_lines += sentence_actual_lines
                last_sentence_embedding = embeddings[i]
            else:
                if current_text:
                    slides.append(current_text.strip())
                    split_flags.append(False)
                current_text = sentence
                current_lines = sentence_actual_lines
                last_sentence_embedding = embeddings[i]
        else:
            if current_text:
                slides.append(current_text.strip())
                split_flags.append(False)
            current_text = sentence
            current_lines = sentence_actual_lines
            last_sentence_embedding = embeddings[i]

    if current_text:
        slides.append(current_text.strip())
        split_flags.append(False)

    # 짧은 슬라이드 병합 (2줄 이하)
    final_slides = []
    final_flags = []
    skip_next = False
    for i in range(len(slides)):
        if progress_callback:
            progress_callback(0.6 + (0.2 * (i / len(slides))), f"짧은 슬라이드 병합 중 ({i+1}/{len(slides)})...")
        if skip_next:
            skip_next = False
            continue
        current_slide_text = slides[i]
        current_slide_lines = calculate_text_lines(current_slide_text, max_chars)
        if current_slide_lines <= 2:
            if i + 1 < len(slides):
                next_slide_text = slides[i + 1]
                combined_text = current_slide_text + "\n" + next_slide_text
                combined_lines = calculate_text_lines(combined_text, max_chars)
                if combined_lines <= max_lines:
                    final_slides.append(combined_text)
                    final_flags.append(split_flags[i] or split_flags[i+1])
                    skip_next = True
                    continue
        final_slides.append(current_slide_text)
        final_flags.append(split_flags[i])

    if not final_slides:
        return [""], [False]

    return final_slides, final_flags
def create_ppt(slides_content, flags, max_chars_per_line, font_size_pt, progress_callback=None):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6]

    for i, (slide_text, is_flagged) in enumerate(zip(slides_content, flags)):
        if progress_callback:
            progress_callback(0.8 + (0.2 * (i / len(slides_content))), f"PPT 슬라이드 생성 중 ({i+1}/{len(slides_content)})...")

        slide = prs.slides.add_slide(blank_slide_layout)

        left = Inches(0.75)
        top = Inches(0.5)
        width = prs.slide_width - Inches(1.5)
        height = prs.slide_height - Inches(1.0)

        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

        for line in slide_text.strip().split('\n'):
            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(font_size_pt)
            p.font.bold = True
            p.font.name = '맑은 고딕'
            p.alignment = PP_ALIGN.CENTER

        # ⚠️ 확인 필요 도형
        if is_flagged:
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), Inches(2.2), Inches(0.6))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 0)
            tf = shape.text_frame
            tf.text = "⚠️ 확인 필요"
            p_flag = tf.paragraphs[0]
            p_flag.font.size = Pt(20)
            p_flag.font.name = '맑은 고딕'
            p_flag.font.bold = True
            p_flag.font.color.rgb = RGBColor(0, 0, 0)
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            p_flag.alignment = PP_ALIGN.CENTER

        # 📄 페이지 번호
        pn_left = prs.slide_width - Inches(1.0)
        pn_top = prs.slide_height - Inches(0.5)
        page_number_shape = slide.shapes.add_textbox(pn_left, pn_top, Inches(0.8), Inches(0.3))
        pn_tf = page_number_shape.text_frame
        pn_tf.text = f"{i+1}/{len(slides_content)}"
        p_pn = pn_tf.paragraphs[0]
        p_pn.font.size = Pt(10)
        p_pn.font.name = '맑은 고딕'
        p_pn.alignment = PP_ALIGN.RIGHT

        # 🔴 마지막 슬라이드에 "끝" 표시
        if i == len(slides_content) - 1:
            end_shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, pn_left - Inches(0.9), pn_top, Inches(0.8), Inches(0.8))
            end_shape.fill.solid()
            end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
            end_tf = end_shape.text_frame
            end_tf.text = "끝"
            p_end = end_tf.paragraphs[0]
            p_end.font.size = Pt(40)
            p_end.font.name = '맑은 고딕'
            p_end.font.bold = True
            p_end.font.color.rgb = RGBColor(255, 255, 255)
            end_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            p_end.alignment = PP_ALIGN.CENTER

    return prs

# --- Streamlit UI ---

uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=200)

st.sidebar.header("⚙️ PPT 설정")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수", 1, 10, 4)
max_chars = st.sidebar.slider("한 줄당 최대 글자 수", 10, 100, 30)
font_size = st.sidebar.slider("본문 폰트 크기 (Pt)", 10, 70, 48)
sim_threshold = st.sidebar.slider("문맥 유사도 기준", 0.5, 1.0, 0.75, step=0.01)

if st.button("✨ PPT 생성"):
    if uploaded_file or text_input:
        paragraphs_raw = extract_text_from_word(uploaded_file) if uploaded_file else [
            p.strip() for p in text_input.split("\n\n") if p.strip()
        ]
        if not paragraphs_raw:
            st.error("입력된 텍스트가 없습니다.")
            st.stop()

        progress_bar = st.progress(0)
        status_text = st.empty()
        start_time = time.time()

        def update_progress(value, message):
            elapsed = time.time() - start_time
            eta = int((1.0 - value) * elapsed / value) if value > 0 else 0
            status_text.text(f"{message} ⏳ {int(value*100)}% (예상 {eta}초 남음)")
            progress_bar.progress(min(value, 1.0))

        try:
            with st.spinner("PPT 생성 중..."):
                update_progress(0.05, "텍스트 분할 중...")
                slides, flags = split_text_into_slides_with_similarity(
                    paragraphs_raw, max_lines, max_chars, model, threshold=sim_threshold, progress_callback=update_progress
                )
                update_progress(0.8, "PPT 슬라이드 생성 중...")
                ppt = create_ppt(slides, flags, max_chars, font_size, progress_callback=update_progress)

                ppt_io = BytesIO()
                ppt.save(ppt_io)
                ppt_io.seek(0)
                update_progress(1.0, "완료!")

                st.download_button(
                    label="📥 PPT 다운로드",
                    data=ppt_io,
                    file_name="paydo_script_ai_generated.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success(f"총 {len(slides)}개의 슬라이드가 생성되었습니다.")
                if any(flags):
                    flagged_indices = [i+1 for i, f in enumerate(flags) if f]
                    st.warning(f"⚠️ 일부 슬라이드는 내용이 길어 강제로 분할되었습니다: {flagged_indices}")
        except Exception as e:
            st.error(f"오류 발생: {e}")
            logging.exception(e)
    else:
        st.info("Word 파일을 업로드하거나 텍스트를 입력하세요.")
