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
    if not text:
        return 0
    for paragraph in text.split('\n'):
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, break_long_words=False)
        lines += len(wrapped_lines) if wrapped_lines else 1
    return lines if lines > 0 else 1

def smart_sentence_split(text):
    try:
        sentences = kss.split_sentences(text)
        return sentences
    except Exception:
        return [s.strip() for s in re.split(r'[.!?]\s+', text) if s.strip()]

def is_potentially_non_sentence(sentence_text, min_length=5):
    sentence_text = sentence_text.strip()
    if not sentence_text:
        return False
    if len(sentence_text) < min_length and not sentence_text.endswith(('.', '!', '?', '다', '요', '죠', '까', '나', '시오')):
        return True
    return False

def is_incomplete(sentence):
    sentence = sentence.strip()
    if not sentence:
        return False
    incomplete_endings = ('은','는','이','가','을','를','에','으로','고','와','과', '며', '는데', '지만', '거나', '든지', '든지간에', '든가', '고,', '며,', '는데,')
    if sentence.endswith(incomplete_endings):
        return True
    return False

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
