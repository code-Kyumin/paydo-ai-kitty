import streamlit as st
import utils  # utils.py를 가져옵니다
from sentence_transformers import SentenceTransformer
import logging

# Streamlit 설정
st.set_page_config(page_title="AI PPT 생성기", layout="wide")
st.title("🎬 촬영 대본 자동 생성 프로그램")

# 로깅 설정 (기존 로깅 설정 유지)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 모델 로드 (캐싱하여 성능 향상)
@st.cache_resource
def load_model():
    logging.info("Loading SentenceTransformer model...")
    model = SentenceTransformer("jhgan/ko-sbert-nli")  # KoSimCSE 모델 사용
    logging.info("SentenceTransformer model loaded.")
    return model

model = load_model()

# 입력 방식 선택
input_type = st.radio("입력 방식 선택", ["텍스트 입력", "파일 업로드"])

if input_type == "텍스트 입력":
    script_text = st.text_area("강의/교육 대본을 입력하세요:", height=200)
else:
    uploaded_file = st.file_uploader("파일을 업로드하세요 (.txt, .docx)", type=["txt", "docx"])
    if uploaded_file:
        script_text = utils.read_script_file(uploaded_file)  # utils에서 파일 읽기 함수 사용
    else:
        script_text = ""

if st.button("대본 분석 및 PPT 생성"):
    if script_text:
        # 핵심 로직 실행
        slides_data = utils.process_script(script_text, model)  # utils에서 전체 처리 함수 사용

        # 사용자 수정 기능 (Streamlit 위젯 활용)
        st.subheader("생성된 슬라이드 (수정 가능)")
        modified_slides = []
        for i, slide_data in enumerate(slides_data):
            modified_text = st.text_area(f"슬라이드 {i + 1}", slide_data["text"], height=150)
            modified_slides.append({"text": modified_text, "flags": slide_data["flags"]})

        if st.button("PPT 생성"):
            utils.create_ppt(modified_slides)  # 수정된 슬라이드로 PPT 생성
            st.success("PPT가 성공적으로 생성되었습니다!")
    else:
        st.warning("대본을 입력해주세요.")