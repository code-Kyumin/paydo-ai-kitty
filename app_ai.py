import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
# from pptx.enum.dml import MSO_THEME_COLOR_INDEX # MSO_THEME_COLOR_INDEX는 사용되지 않아 주석 처리
# from pptx.enum.text import MSO_AUTO_SIZE # MSO_AUTO_SIZE는 사용되지 않아 주석 처리
app_ai


def split_text_into_slides_with_similarity(paragraphs, max_lines, max_chars, model, threshold=0.85, progress_callback=None):
    slides = []
    split_flags = []  # True면 강제 분할된 슬라이드 (확인 필요)

    # 1. 모든 문장 추출 및 병합
    all_sentences_original = [s for p_idx, p in enumerate(paragraphs) for s_idx, s in enumerate(smart_sentence_split(p))]
    # st.write("원본 문장:", all_sentences_original) # 디버깅용
    merged_sentences = merge_sentences(all_sentences_original)
    # st.write("병합된 문장:", merged_sentences) # 디버깅용

    if not merged_sentences:
        return [""], [False]

    if progress_callback:
        progress_callback(0.1, "문장 임베딩 중...")

    embeddings = model.encode(merged_sentences)

    current_text = ""
    current_lines = 0
    last_sentence_embedding = None

    # 2. 슬라이드 분할 로직
    for i, sentence in enumerate(merged_sentences):
        if progress_callback:
            progress_callback(0.1 + (0.5 * (i / len(merged_sentences))), f"슬라이드 분할 중 ({i+1}/{len(merged_sentences)})...")

        sentence_actual_lines = calculate_text_lines(sentence, max_chars)

        # 요구사항 7 & 6: 한 문장 자체가 max_lines를 초과하는 경우
        if sentence_actual_lines > max_lines:
            if current_text: # 이전에 쌓인 텍스트가 있으면 먼저 슬라이드로 만듦
                slides.append(current_text.strip())
                split_flags.append(False)
                current_text, current_lines = "", 0
                last_sentence_embedding = None

            # 긴 문장을 max_chars에 맞춰 여러 줄로 나눔
            wrapped_sentence_lines = textwrap.wrap(sentence, width=max_chars, break_long_words=True, replace_whitespace=False)
            
            temp_slide_text = ""
            temp_slide_lines = 0
            for line_idx, line_text in enumerate(wrapped_sentence_lines):
                if temp_slide_lines + 1 <= max_lines:
                    temp_slide_text += (line_text + "\n")
                    temp_slide_lines += 1
                else: # 현재 슬라이드가 꽉 차면 추가하고 새 슬라이드 시작
                    slides.append(temp_slide_text.strip())
                    split_flags.append(True) # 강제 분할 플래그
                    temp_slide_text = line_text + "\n"
                    temp_slide_lines = 1
            
            if temp_slide_text: # 남은 부분 추가
                slides.append(temp_slide_text.strip())
                split_flags.append(True) # 강제 분할 플래그
            # 이 경우, 다음 문장으로 넘어가므로 current_text 등은 이미 초기화됨
            last_sentence_embedding = embeddings[i] # 이 긴 문장의 임베딩을 다음 문장과의 유사도 비교를 위해 저장
            continue


        # 일반적인 슬라이드 추가 로직
        # 현재 문장을 추가했을 때 최대 줄 수를 넘는지 확인
        if current_lines + sentence_actual_lines <= max_lines:
            similar_to_previous = True
            if current_text and last_sentence_embedding is not None and i < len(embeddings): # current_text가 있고, 이전 임베딩이 있어야 비교 가능
                # 현재 문장(embeddings[i])과 이전 문장(last_sentence_embedding)의 유사도
                # 여기서 last_sentence_embedding은 이전 슬라이드의 마지막 문장이 아니라, 바로 직전 문장의 임베딩이어야 함
                # 따라서, last_sentence_embedding은 이전 for 루프의 embeddings[i-1]을 사용해야함.
                # 다만, 위에서 긴 문장 처리 후 last_sentence_embedding = embeddings[i]로 설정했으므로,
                # 일반적인 경우에는 embeddings[i-1]을 사용하도록 수정.
                prev_emb_to_compare = embeddings[i-1] if i > 0 else None
                if prev_emb_to_compare is not None:
                    sim = util.cos_sim(prev_emb_to_compare, embeddings[i])[0][0].item()
                    if sim < threshold:
                        similar_to_previous = False
                else: # 첫 문장이거나, 이전 문장이 없으면 (예: 긴 문장 바로 다음) 유사도 체크 안 함
                    similar_to_previous = True


            if similar_to_previous:
                current_text = f"{current_text}\n{sentence}" if current_text else sentence
                current_lines += sentence_actual_lines
                last_sentence_embedding = embeddings[i] # 현재 문장의 임베딩을 저장
            else: # 유사도 낮으면 새 슬라이드
                if current_text:
                    slides.append(current_text.strip())
                    split_flags.append(False)
                current_text = sentence
                current_lines = sentence_actual_lines
                last_sentence_embedding = embeddings[i]
        else: # 현재 슬라이드에 공간 부족, 새 슬라이드 시작
            if current_text:
                slides.append(current_text.strip())
                split_flags.append(False)
            current_text = sentence
            current_lines = sentence_actual_lines
            last_sentence_embedding = embeddings[i]

    if current_text: # 마지막 남은 텍스트 추가
        slides.append(current_text.strip())
        split_flags.append(False)

    # 요구사항 5: 문장이 2줄 이하일 경우 앞/뒤 문장과 슬라이드 합치기 시도
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
            # 뒷 슬라이드와 합치기 시도 (주로 이 경우를 먼저 고려)
            if i + 1 < len(slides):
                next_slide_text = slides[i+1]
                combined_text = current_slide_text + "\n" + next_slide_text
                combined_lines = calculate_text_lines(combined_text, max_chars)
                if combined_lines <= max_lines:
                    final_slides.append(combined_text)
                    final_flags.append(split_flags[i] or split_flags[i+1]) # 하나라도 True면 True
                    skip_next = True # 다음 슬라이드는 건너뜀
                    continue
            
            # (옵션) 앞 슬라이드와 합치기 시도 (이미 final_slides에 추가된 것과)
            # 이 로직은 복잡해질 수 있어, 일단 뒷 슬라이드와의 병합을 우선
            # if final_slides:
            #    prev_slide_text = final_slides[-1]
            #    combined_text_with_prev = prev_slide_text + "\n" + current_slide_text
            #    combined_lines_with_prev = calculate_text_lines(combined_text_with_prev, max_chars)
            #    if calculate_text_lines(final_slides[-1], max_chars) <=2 and combined_lines_with_prev <= max_lines : # 이전것도 짧다면
            #       final_slides[-1] = combined_text_with_prev
            #       final_flags[-1] = final_flags[-1] or split_flags[i]
            #       continue # 현재 슬라이드는 추가 안함

        final_slides.append(current_slide_text)
        final_flags.append(split_flags[i])
    
    if not final_slides: # 병합 등으로 모든 슬라이드가 사라진 극단적인 경우
        return [""], [False]

    return final_slides, final_flags


def create_ppt(slides_content, flags, max_chars_per_line, font_size_pt, progress_callback=None):
    prs = Presentation()
    prs.slide_width = Inches(13.333) # 16:9 비율
    prs.slide_height = Inches(7.5)
    
    blank_slide_layout = prs.slide_layouts[6]  # 내용 없는 슬라이드 레이아웃

    for i, (slide_text, is_flagged) in enumerate(zip(slides_content, flags)):
        if progress_callback:
            progress_callback(0.8 + (0.2 * (i / len(slides_content))), f"PPT 슬라이드 생성 중 ({i+1}/{len(slides_content)})...")

        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 텍스트 박스 (슬라이드 중앙보다 약간 위쪽으로)
        left = Inches(0.75)
        top = Inches(0.5) # 기존 0.75에서 더 위로
        width = prs.slide_width - Inches(1.5)
        height = prs.slide_height - Inches(1.0) # 페이지 번호, 끝 표시 공간 확보 위해 약간 줄임

        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE # 중앙 정렬로 변경 (또는 TOP)
        # text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE # 내용에 맞게 자동 크기 조절 (필요시)

        # 줄바꿈 단위로 텍스트 추가
        cleaned_slide_text = slide_text.strip() # 앞뒤 공백 제거
        for line_content in cleaned_slide_text.split('\n'): # 이미 \n으로 구분된 내용을 사용
            p = text_frame.add_paragraph()
            # textwrap.wrap을 여기서 한 번 더 쓰는 것은 split_text_into_slides_with_similarity에서 \n으로 구분한 것과 충돌 가능성
            # 이미 적절히 \n으로 나뉘어 있다고 가정하고, 그 줄을 그대로 넣음
            # 만약 한 줄이 max_chars_per_line을 넘는다면, 이 부분에서 다시 textwrap을 적용해야 하나,
            # split_text_into_slides_with_similarity 에서 \n으로 나눌 때 max_chars를 고려해야 함.
            # 현재 calculate_text_lines 와 연동하여 \n으로 구분된 텍스트를 가정.
            p.text = line_content 
            p.font.size = Pt(font_size_pt)
            p.font.bold = True
            p.font.name = '맑은 고딕' # 기본 폰트 지정 (맑은 고딕이 없다면 시스템 기본 폰트)
            p.alignment = PP_ALIGN.CENTER

        # 요구사항 2: "확인 필요" 도형 수정
        if is_flagged:
            shape_width = Inches(2.2) # 약간 넓게
            shape_height = Inches(0.6) # 약간 높게
            # 위치는 좌상단 유지 또는 필요시 조정
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(0.2), shape_width, shape_height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 255, 0) # 노란색 배경
            
            tf_flag = shape.text_frame
            tf_flag.text = "⚠️ 확인 필요" # 특수기호 추가
            p_flag = tf_flag.paragraphs[0]
            p_flag.font.size = Pt(20)      # 폰트 크기 20pt
            p_flag.font.name = '맑은 고딕'   # 폰트 맑은 고딕
            p_flag.font.bold = True
            p_flag.font.color.rgb = RGBColor(0, 0, 0) # 검은색 글자
            tf_flag.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            p_flag.alignment = PP_ALIGN.CENTER

        # 요구사항 1: 페이지 번호 슬라이드 우측 하단 배치
        pn_left = prs.slide_width - Inches(1.0) # 우측에서 1인치 떨어진 곳
        pn_top = prs.slide_height - Inches(0.5) # 하단에서 0.5인치 떨어진 곳
        pn_width = Inches(0.8)
        pn_height = Inches(0.3)
        
        page_number_shape = slide.shapes.add_textbox(pn_left, pn_top, pn_width, pn_height)
        pn_tf = page_number_shape.text_frame
        pn_tf.text = f"{i+1}/{len(slides_content)}"
        p_pn = pn_tf.paragraphs[0]
        p_pn.font.size = Pt(10)
        p_pn.font.name = '맑은 고딕'
        p_pn.alignment = PP_ALIGN.RIGHT

        # 요구사항 9: "끝" 표시 도형 (마지막 슬라이드에만)
        if i == len(slides_content) - 1:
            end_shape_diameter = Inches(0.8) # 원의 지름
            end_shape_left = prs.slide_width - end_shape_diameter - Inches(0.2) # 페이지번호와 겹치지 않게 조정
            end_shape_top = prs.slide_height - end_shape_diameter - Inches(0.6) # 페이지번호 위에 위치하도록 조정 (또는 다른 위치)

            # 페이지 번호 위치를 고려하여 "끝" 표시 위치 조정 (페이지 번호보다 약간 위 또는 왼쪽)
            # 예시: 페이지 번호 박스(pn_left, pn_top)의 왼쪽에 위치
            end_shape_left = pn_left - end_shape_diameter - Inches(0.1) 
            end_shape_top = pn_top # 페이지 번호와 같은 높이로 하되, 더 왼쪽에

            if end_shape_left < 0 : end_shape_left = Inches(0.1) # 너무 왼쪽으로 가면 조정


            end_shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, # 원형 도형
                end_shape_left, 
                end_shape_top,
                end_shape_diameter, 
                end_shape_diameter
            )
            end_shape.fill.solid()
            end_shape.fill.fore_color.rgb = RGBColor(255, 0, 0) # 빨간색 배경
            
            end_tf = end_shape.text_frame
            end_tf.text = "끝"
            p_end = end_tf.paragraphs[0]
            p_end.font.size = Pt(20) # 40pt는 원 안에 너무 클 수 있어 20pt로 조정. 필요시 폰트 크기 재조정.
            p_end.font.name = '맑은 고딕'
            p_end.font.bold = True
            p_end.font.color.rgb = RGBColor(255, 255, 255) # 흰색 글자
            end_tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            p_end.alignment = PP_ALIGN.CENTER
            
    return prs

# --- Streamlit UI ---

uploaded_file = st.file_uploader("📄 Word 파일 업로드 (.docx)", type=["docx"])
text_input = st.text_area("또는 텍스트 직접 입력 (문단은 빈 줄로 구분):", height=200, key="text_input_main")

st.sidebar.header("⚙️ PPT 설정")
max_lines = st.sidebar.slider("슬라이드당 최대 줄 수", 1, 10, 4, key="max_lines_slider")
max_chars = st.sidebar.slider("한 줄당 최대 글자 수 (참고용)", 10, 100, 30, key="max_chars_slider") # 이름 변경 (실제 줄바꿈은 \n 기준)
font_size = st.sidebar.slider("본문 폰트 크기 (Pt)", 10, 70, 48, key="font_size_slider") # 기본 폰트 크기 조정
sim_threshold = st.sidebar.slider("문맥 유사도 기준 (낮을수록 잘 나눔)", 0.5, 1.0, 0.75, step=0.01, key="sim_threshold_slider") # step 조정 및 설명 변경

if st.button("✨ PPT 생성", key="generate_ppt_button"):
    if uploaded_file or text_input:
        paragraphs_raw = extract_text_from_word(uploaded_file) if uploaded_file else [p.strip() for p in text_input.split("\n\n") if p.strip()]
        
        if not paragraphs_raw:
            st.error("입력된 텍스트가 없습니다.")
            st.stop()

        # 요구사항 8: 진행률 표시 UI
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        def update_progress(value, message):
            progress_bar.progress(min(value, 1.0)) # 1.0을 넘지 않도록
            status_text.text(message)

        update_progress(0, "준비 중...")

        try:
            with st.spinner("PPT 생성 중... 이 작업은 몇 분 정도 소요될 수 있습니다."): # 좀 더 친절한 메시지
                # 1. 슬라이드 내용 분할
                update_progress(0.05, "텍스트 분할 및 분석 시작...")
                # split_text_into_slides_with_similarity 함수가 progress_callback을 받도록 수정됨
                slides, flags = split_text_into_slides_with_similarity(paragraphs_raw, max_lines, max_chars, model, threshold=sim_threshold, progress_callback=update_progress)
                
                # 2. PPT 파일 생성
                update_progress(0.8, "PPT 파일 생성 시작...")
                # create_ppt 함수가 progress_callback을 받도록 수정됨
                ppt = create_ppt(slides, flags, max_chars, font_size, progress_callback=update_progress)
                
                update_progress(1.0, "PPT 생성 완료!")

                ppt_io = BytesIO()
                ppt.save(ppt_io)
                ppt_io.seek(0)
                
                st.download_button(
                    label="📥 PPT 다운로드",
                    data=ppt_io,
                    file_name="paydo_script_ai_generated.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.success(f"총 {len(slides)}개의 슬라이드가 포함된 PPT가 생성되었습니다.")
                
                if any(flags):
                    flagged_indices = [i + 1 for i, flag_val in enumerate(flags) if flag_val]
                    st.warning(f"⚠️ 일부 슬라이드({len(flagged_indices)}개)가 내용이 길거나 구성상 강제 분할되었습니다. 확인이 필요한 슬라이드 번호: {flagged_indices}")

        except Exception as e:
            st.error(f"PPT 생성 중 오류 발생: {e}")
            logging.error(f"PPT 생성 실패: {e}", exc_info=True)
            update_progress(0, f"오류 발생: {e}") # 오류 발생 시 진행률 초기화 및 메시지
            
    else:
        st.info("Word 파일을 업로드하거나 텍스트를 직접 입력한 후 'PPT 생성' 버튼을 클릭하세요.")