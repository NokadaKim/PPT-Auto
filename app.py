"""
PPT Auto Generator - Streamlit App
- Manual Input (기본) / AI Auto Generate (선택)
- 콘텐츠를 한 슬라이드 내에 자동 레이아웃 배치
- 이미지 자동 생성 체크박스
"""
import streamlit as st
import os
import tempfile

from generator.ppt_builder import PPTBuilder
from generator.image_fetcher import ImageFetcher
from generator.content_formatter import format_content_for_slide, extract_image_keywords

# ─── Unsplash API Key ───
UNSPLASH_KEY = "vsqPu708Y6eSVNwZRsM0du7xwV2r_Gqw3MGeByiM0rs"

# ─── AI 콘텐츠 생성 (g4f 사용) ───
def generate_outline_with_ai(topic, num_chapters):
    """g4f를 사용하여 PPT 아웃라인 생성"""
    try:
        from g4f.client import Client
        client = Client()

        prompt = f"""'{topic}' 주제로 PPT 발표 자료의 아웃라인을 만들어주세요.
총 {num_chapters}개의 챕터(슬라이드)로 구성해주세요.
각 챕터마다 제목과 핵심 내용 3-5줄을 작성해주세요.

형식:
## 챕터1: [제목]
- 내용1
- 내용2
- 내용3

## 챕터2: [제목]
- 내용1
...

한국어로 작성해주세요."""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI 생성 오류: {e}"


def parse_outline(outline_text):
    """AI가 생성한 아웃라인을 파싱하여 챕터 리스트로 변환"""
    chapters = []
    current_title = ""
    current_content = []

    for line in outline_text.split('\n'):
        line = line.strip()
        if line.startswith('## ') or line.startswith('# '):
            if current_title:
                chapters.append({"title": current_title, "content": "\n".join(current_content)})
            current_title = line.lstrip('#').strip()
            # "챕터1: " 등의 접두사 제거
            if ':' in current_title:
                current_title = current_title.split(':', 1)[1].strip()
            elif '：' in current_title:
                current_title = current_title.split('：', 1)[1].strip()
            current_content = []
        elif line.startswith('- ') or line.startswith('• ') or line.startswith('* '):
            current_content.append(line.lstrip('-•* ').strip())
        elif line:
            current_content.append(line)

    if current_title:
        chapters.append({"title": current_title, "content": "\n".join(current_content)})

    return chapters


# ═══════════════════════════════════════════════
# Streamlit UI
# ═══════════════════════════════════════════════
st.set_page_config(page_title="PPT Auto Generator", page_icon="📊", layout="wide")

st.title("📊 PPT Auto Generator")
st.caption("콘텐츠를 입력하면 자동 레이아웃으로 PPT를 생성합니다")

# ─── 사이드바 설정 ───
with st.sidebar:
    st.header("⚙️ 설정")

    # 모드 선택 (Manual이 기본)
    mode = st.radio("모드 선택", ["Manual Input", "AI Auto Generate"], index=0)

    st.divider()

    # 디자인 테마
    theme = st.selectbox("디자인 테마", [
        "ir_book",
        "dark_modern", "light_clean", "nature_green", "corporate_blue",
        "warm_orange", "purple_creative", "minimal_gray", "tech_dark"
    ], index=0)

    st.divider()

    # 엔딩 슬라이드
    ending_text = st.text_input("엔딩 슬라이드 텍스트", value="Thank You")
    ending_subtext = st.text_input("엔딩 부제", value="")

# ─── 메인 영역 ───
st.divider()

# 공통: 프레젠테이션 제목
ppt_title = st.text_input("📌 프레젠테이션 제목", placeholder="예: 2026 AI 트렌드 분석")
ppt_subtitle = st.text_input("📝 부제목 (선택)", placeholder="예: 기술 동향과 시장 전망")

st.divider()

# ═══════════════════════════════════════════════
# MODE 1: Manual Input (기본)
# ═══════════════════════════════════════════════
if mode == "Manual Input":
    st.subheader("✏️ 슬라이드 직접 입력")

    # 슬라이드 개수
    num_slides = st.number_input("슬라이드 개수", min_value=1, max_value=20, value=3)

    slides_data = []

    for i in range(num_slides):
        with st.expander(f"📄 슬라이드 {i+1}", expanded=(i == 0)):
            slide_title = st.text_input(
                f"제목", key=f"title_{i}",
                placeholder="슬라이드 제목을 입력하세요"
            )
            slide_content = st.text_area(
                f"내용", key=f"content_{i}",
                placeholder="내용을 입력하세요.\n줄바꿈으로 항목을 구분합니다.\n\n예:\n시장 규모: 2026년 500억 달러 예상\n성장률: 연평균 15.3%\n주요 동인: AI 기술 발전, 자동화 수요 증가",
                height=180
            )

            # ★ 이미지 자동 생성 체크박스
            auto_image = st.checkbox(
                "🖼️ 이미지 자동 생성",
                key=f"img_{i}",
                help="체크하면 콘텐츠에 맞는 이미지를 자동으로 검색하여 삽입합니다"
            )

            slide_label = st.text_input(
                "라벨 (선택)", key=f"label_{i}",
                placeholder="예: Chapter 1, 개요, 분석 등"
            )

            slides_data.append({
                "title": slide_title,
                "content": slide_content,
                "auto_image": auto_image,
                "label": slide_label
            })

    st.divider()

    # 생성 버튼
    if st.button("🚀 PPT 생성", type="primary", use_container_width=True):
        if not ppt_title:
            st.error("프레젠테이션 제목을 입력하세요!")
        elif not any(s["title"] or s["content"] for s in slides_data):
            st.error("최소 1개 슬라이드의 내용을 입력하세요!")
        else:
            with st.spinner("PPT 생성 중..."):
                builder = PPTBuilder(theme=theme)
                fetcher = ImageFetcher(unsplash_key=UNSPLASH_KEY)

                # 표지
                builder.add_title_slide(ppt_title, ppt_subtitle)

                progress = st.progress(0)

                for idx, sd in enumerate(slides_data):
                    if not sd["title"] and not sd["content"]:
                        continue

                    title = sd["title"] or f"슬라이드 {idx+1}"
                    content = sd["content"] or ""
                    label = sd["label"] or ""

                    # 콘텐츠 구조 분석 (한 페이지 내 레이아웃 결정)
                    formatted = format_content_for_slide(title, content)

                    # 이미지 자동 검색
                    img_path = None
                    if sd["auto_image"] and content:
                        st.text(f"  🖼️ 슬라이드 {idx+1} 이미지 검색 중...")
                        keywords = extract_image_keywords(title, content)
                        img_path = fetcher.search_and_download(
                            keywords,
                            filename=f"slide_{idx+1}.jpg"
                        )

                    # 자동 레이아웃으로 슬라이드 생성
                    builder.add_formatted_slide(formatted, image_path=img_path, slide_label=label)

                    progress.progress((idx + 1) / len(slides_data))

                # 엔딩 슬라이드
                if ending_text:
                    builder.add_ending_slide(ending_text, ending_subtext)

                # 저장
                output_path = os.path.join(tempfile.gettempdir(), f"{ppt_title}.pptx")
                builder.save(output_path)

                st.success("✅ PPT 생성 완료!")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 PPT 다운로드",
                        data=f.read(),
                        file_name=f"{ppt_title}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

# ═══════════════════════════════════════════════
# MODE 2: AI Auto Generate
# ═══════════════════════════════════════════════
elif mode == "AI Auto Generate":
    st.subheader("🤖 AI 자동 생성")

    topic = st.text_input("🎯 주제 입력", placeholder="예: 2026년 글로벌 AI 시장 동향")
    num_chapters = st.slider("챕터 수", min_value=2, max_value=10, value=5)

    # ★ 이미지 자동 생성 (전체 적용)
    auto_image_all = st.checkbox(
        "🖼️ 모든 슬라이드에 이미지 자동 생성",
        value=False,
        help="체크하면 모든 슬라이드에 관련 이미지를 자동으로 삽입합니다"
    )

    if st.button("🤖 AI로 아웃라인 생성", use_container_width=True):
        if not topic:
            st.error("주제를 입력하세요!")
        else:
            with st.spinner("AI가 아웃라인을 생성하고 있습니다..."):
                outline_text = generate_outline_with_ai(topic, num_chapters)
                st.session_state["outline_text"] = outline_text
                st.session_state["outline_chapters"] = parse_outline(outline_text)

    # 아웃라인 편집
    if "outline_text" in st.session_state:
        st.divider()
        st.subheader("📋 생성된 아웃라인 (편집 가능)")

        edited_outline = st.text_area(
            "아웃라인",
            value=st.session_state["outline_text"],
            height=400
        )

        if st.button("🔄 아웃라인 다시 파싱"):
            st.session_state["outline_chapters"] = parse_outline(edited_outline)
            st.success("파싱 완료!")

        # 파싱 결과 미리보기
        if "outline_chapters" in st.session_state:
            chapters = st.session_state["outline_chapters"]
            st.write(f"총 {len(chapters)}개 챕터 감지됨:")
            for idx, ch in enumerate(chapters):
                st.write(f"  **{idx+1}. {ch['title']}** — {len(ch['content'].split(chr(10)))}줄")

            st.divider()

            if st.button("🚀 PPT 생성", type="primary", use_container_width=True):
                if not ppt_title:
                    ppt_title = topic

                with st.spinner("PPT 생성 중..."):
                    builder = PPTBuilder(theme=theme)
                    fetcher = ImageFetcher(unsplash_key=UNSPLASH_KEY)

                    builder.add_title_slide(ppt_title, ppt_subtitle)

                    progress = st.progress(0)

                    for idx, ch in enumerate(chapters):
                        title = ch["title"]
                        content = ch["content"]
                        label = f"Chapter {idx+1}"

                        # 자동 레이아웃 분석
                        formatted = format_content_for_slide(title, content)

                        # 이미지
                        img_path = None
                        if auto_image_all and content:
                            st.text(f"  🖼️ 챕터 {idx+1} 이미지 검색 중...")
                            keywords = extract_image_keywords(title, content)
                            img_path = fetcher.search_and_download(
                                keywords,
                                filename=f"ch_{idx+1}.jpg"
                            )

                        builder.add_formatted_slide(formatted, image_path=img_path, slide_label=label)

                        progress.progress((idx + 1) / len(chapters))

                    if ending_text:
                        builder.add_ending_slide(ending_text, ending_subtext)

                    output_path = os.path.join(tempfile.gettempdir(), f"{ppt_title}.pptx")
                    builder.save(output_path)

                    st.success("✅ PPT 생성 완료!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 PPT 다운로드",
                            data=f.read(),
                            file_name=f"{ppt_title}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
