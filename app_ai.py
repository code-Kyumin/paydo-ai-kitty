
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
from PIL import Image  # Import PIL for image handling

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

# 나머지 함수들과 Streamlit UI 코드는 기존과 동일하게 유지
