"""
PPT Auto Generator - Streamlit App
- 챕터 → 슬라이드 계층 구조
- 챕터별 간지 페이지 + 본문 좌상단 "01 챕터제목"
- 콘텐츠 섹션별 분리 (빈 줄 기준) → 별도 텍스트박스
- Manual Input (기본) / AI Auto Generate (선택)
"""
import streamlit as st
import os
import tempfile

from generator.ppt_builder import PPTBuilder
from generator.image_fetcher import ImageFetcher
from generator.content_formatter import format_content_for_slide

# ─── Unsplash API Key ───
UNSPLASH_KEY = "vsqPu708Y6eSVNwZRsM0du7xwV2r_Gqw3MGeByiM0rs"


# ─── AI 콘텐츠 생성 ───
def generate_outline_with_ai(topic, num_chapters):
    try:
        from g4f.client import Client
        client = Client()
        prompt = f"""'{topic}' 주제로 PPT 발표 자료의 아웃라인을 만들어주세요.
총 {num_chapters}개의 챕터로 구성하고, 각 챕터에 2-3개의 슬라이드를 포함해주세요.

형식:
## 챕터1: [챕터 제목]
### 슬라이드1: [슬라이드 제목]
- 내용1
- 내용2

### 슬라이드2: [슬라이드 제목]
- 내용1
- 내용2

## 챕터2: [챕터 제목]
### 슬라이드1: [슬라이드 제목]
...

한국어로 작성해주세요."""
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI 생성 오류: {e}"


def parse_chapter_outline(outline_text):
    """챕터-슬라이드 구조로 파싱"""
    chapters = []
    current_chapter = None
    current_slide = None

    for line in outline_text.split('\n'):
        line = line.strip()
        if line.startswith('## ') or (line.startswith('# ') and not line.startswith('### ')):
            if current_chapter:
                if current_slide:
                    current_chapter["slides"].append(current_slide)
                chapters.append(current_chapter)
            title = line.lstrip('#').strip()
            if ':' in title:
                title = title.split(':', 1)[1].strip()
            elif '：' in title:
                title = title.split('：', 1)[1].strip()
            current_chapter = {"title": title, "slides": []}
            current_slide = None
        elif line.startswith('### '):
            if current_slide and current_chapter:
                current_chapter["slides"].append(current_slide)
            title = line.lstrip('#').strip()
            if ':' in title:
                title = title.split(':', 1)[1].strip()
            elif '：' in title:
                title = title.split('：', 1)[1].strip()
            current_slide = {"title": title, "content": ""}
        elif line.startswith(('- ', '• ', '* ')) and current_slide is not None:
            content_line = line.lstrip('-•* ').strip()
            if current_slide["content"]:
                current_slide["content"] += "\n"
            current_slide["content"] += content_line
        elif line and current_slide is not None:
            if current_slide["content"]:
                current_slide["content"] += "\n"
            current_slide["content"] += line

    if current_chapter:
        if current_slide:
            current_chapter["slides"].append(current_slide)
        chapters.append(current_chapter)

    return chapters


# ═══════════════════════════════════════════════
# Streamlit UI
# ═══════════════════════════════════════════════
st.set_page_config(page_title="PPT Auto Generator", page_icon="📊", layout="wide")

st.title("📊 PPT Auto Generator")
st.caption("챕터/슬라이드 구조로 전문적인 PPT를 자동 생성합니다")

# ─── 사이드바 설정 ───
with st.sidebar:
    st.header("⚙️ 설정")
    mode = st.radio("모드 선택", ["Manual Input", "AI Auto Generate"], index=0)
    st.divider()
    theme = st.selectbox("디자인 테마", [
        "ir_pro", "corporate_blue", "light_clean", "dark_modern"
    ], index=0)
    st.divider()
    ending_text = st.text_input("엔딩 슬라이드 텍스트", value="Thank You")
    ending_subtext = st.text_input("엔딩 부제", value="")

# ─── 메인 영역 ───
st.divider()
ppt_title = st.text_input("📌 프레젠테이션 제목", placeholder="예: 2026 AI 트렌드 분석")
ppt_subtitle = st.text_input("📝 부제목 (선택)", placeholder="예: 기술 동향과 시장 전망")
st.divider()


# ═══════════════════════════════════════════════
# MODE 1: Manual Input
# ═══════════════════════════════════════════════
if mode == "Manual Input":
    st.subheader("✏️ 챕터/슬라이드 직접 입력")

    num_chapters = st.number_input("챕터 수", min_value=1, max_value=10, value=2)

    # 챕터별 데이터 저장
    chapters_data = []

    for ch_idx in range(num_chapters):
        ch_num = ch_idx + 1
        with st.expander(f"📁 챕터 {ch_num}", expanded=(ch_idx == 0)):
            ch_title = st.text_input(
                f"챕터 {ch_num} 제목 (간지 페이지에 표시)",
                key=f"ch_title_{ch_idx}",
                placeholder=f"예: 챕터 {ch_num} 제목을 입력하세요"
            )

            num_slides = st.number_input(
                f"챕터 {ch_num} 슬라이드 수",
                min_value=1, max_value=10, value=2,
                key=f"ch_slides_{ch_idx}"
            )

            slides = []
            for sl_idx in range(num_slides):
                st.markdown(f"---")
                st.markdown(f"**슬라이드 {sl_idx + 1}**")

                sl_title = st.text_input(
                    "슬라이드 제목",
                    key=f"sl_title_{ch_idx}_{sl_idx}",
                    placeholder="슬라이드 제목을 입력하세요"
                )

                sl_content = st.text_area(
                    "콘텐츠 (빈 줄로 섹션 구분)",
                    key=f"sl_content_{ch_idx}_{sl_idx}",
                    placeholder="""예시 (빈 줄로 섹션 구분):

Environmental
· 전체 온실가스의 16% 농축산업에서 발생
*60~70%  → 메탄(CH₄)
*20%  → 아산화질소(N₂O)

Ethical
· 공장식 축산으로 인한 문제발생
· 가축전염병, 비위생적 환경

Resource
*햄버거 패티 한 장에 소요되는 자원량
· 3kg의 곡물과 사료
· 200L의 식수/관개수""",
                    height=200
                )

                auto_image = st.checkbox(
                    "🖼️ 이미지 자동 생성",
                    key=f"img_{ch_idx}_{sl_idx}",
                    help="체크하면 콘텐츠에 맞는 이미지를 자동으로 삽입합니다"
                )

                slides.append({
                    "title": sl_title,
                    "content": sl_content,
                    "auto_image": auto_image
                })

            chapters_data.append({
                "title": ch_title,
                "slides": slides
            })

    st.divider()

    # ─── 생성 버튼 ───
    if st.button("🚀 PPT 생성", type="primary", use_container_width=True):
        if not ppt_title:
            st.error("프레젠테이션 제목을 입력하세요!")
        elif not any(ch["title"] for ch in chapters_data):
            st.error("최소 1개 챕터의 제목을 입력하세요!")
        else:
            with st.spinner("PPT 생성 중..."):
                builder = PPTBuilder(theme=theme)
                fetcher = ImageFetcher(unsplash_key=UNSPLASH_KEY)

                # 1) 표지
                builder.add_title_slide(ppt_title, ppt_subtitle)

                # 2) INDEX (목차)
                ch_titles = [ch["title"] or f"Chapter {i+1}"
                             for i, ch in enumerate(chapters_data)]
                builder.add_index_slide(ch_titles)

                # 슬라이드 총 수 계산 (진행률용)
                total_slides = sum(len(ch["slides"]) for ch in chapters_data)
                done = 0
                progress = st.progress(0)

                # 3) 챕터별 처리
                for ch_idx, ch in enumerate(chapters_data):
                    ch_num = ch_idx + 1
                    ch_title = ch["title"] or f"Chapter {ch_num}"

                    # 간지 페이지
                    builder.add_chapter_title_slide(ch_num, ch_title)

                    # 슬라이드
                    for sl_idx, sl in enumerate(ch["slides"]):
                        if not sl["title"] and not sl["content"]:
                            done += 1
                            progress.progress(done / max(total_slides, 1))
                            continue

                        slide_title = sl["title"] or f"슬라이드 {sl_idx + 1}"
                        content = sl["content"] or ""

                        # 콘텐츠 분석 (섹션 분리)
                        formatted = format_content_for_slide(slide_title, content)

                        # 이미지
                        img_path = None
                        if sl["auto_image"] and content:
                            st.text(f"  🖼️ Ch{ch_num}-Slide{sl_idx+1} 이미지 검색 중...")
                            img_path = fetcher.search_and_download(
                                title=slide_title,
                                content=content,
                                filename=f"ch{ch_num}_sl{sl_idx+1}.jpg"
                            )

                        # 슬라이드 생성 (챕터 정보 포함)
                        builder.add_formatted_slide(
                            formatted,
                            chapter_num=ch_num,
                            chapter_title=ch_title,
                            image_path=img_path
                        )

                        done += 1
                        progress.progress(done / max(total_slides, 1))

                # 4) 엔딩
                if ending_text:
                    builder.add_ending_slide(ending_text, ending_subtext)

                # 5) 저장 & 다운로드
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
    num_ai_chapters = st.slider("챕터 수", min_value=2, max_value=10, value=3)

    auto_image_all = st.checkbox(
        "🖼️ 모든 슬라이드에 이미지 자동 생성",
        value=False
    )

    if st.button("🤖 AI로 아웃라인 생성", use_container_width=True):
        if not topic:
            st.error("주제를 입력하세요!")
        else:
            with st.spinner("AI가 아웃라인을 생성하고 있습니다..."):
                outline_text = generate_outline_with_ai(topic, num_ai_chapters)
                st.session_state["outline_text"] = outline_text
                st.session_state["ai_chapters"] = parse_chapter_outline(outline_text)

    if "outline_text" in st.session_state:
        st.divider()
        st.subheader("📋 생성된 아웃라인 (편집 가능)")

        edited_outline = st.text_area(
            "아웃라인",
            value=st.session_state["outline_text"],
            height=400
        )

        if st.button("🔄 아웃라인 다시 파싱"):
            st.session_state["ai_chapters"] = parse_chapter_outline(edited_outline)
            st.success("파싱 완료!")

        if "ai_chapters" in st.session_state:
            chapters = st.session_state["ai_chapters"]
            st.write(f"총 {len(chapters)}개 챕터:")
            for ci, ch in enumerate(chapters):
                st.write(f"  **{ci+1}. {ch['title']}** — {len(ch['slides'])}개 슬라이드")

            st.divider()

            if st.button("🚀 PPT 생성", type="primary", use_container_width=True):
                if not ppt_title:
                    ppt_title = topic

                with st.spinner("PPT 생성 중..."):
                    builder = PPTBuilder(theme=theme)
                    fetcher = ImageFetcher(unsplash_key=UNSPLASH_KEY)

                    builder.add_title_slide(ppt_title, ppt_subtitle)

                    ch_titles = [ch["title"] for ch in chapters]
                    builder.add_index_slide(ch_titles)

                    total_slides = sum(len(ch["slides"]) for ch in chapters)
                    done = 0
                    progress = st.progress(0)

                    for ch_idx, ch in enumerate(chapters):
                        ch_num = ch_idx + 1
                        ch_title = ch["title"]

                        builder.add_chapter_title_slide(ch_num, ch_title)

                        for sl_idx, sl in enumerate(ch["slides"]):
                            slide_title = sl["title"]
                            content = sl["content"]

                            formatted = format_content_for_slide(slide_title, content)

                            img_path = None
                            if auto_image_all and content:
                                st.text(f"  🖼️ Ch{ch_num}-Slide{sl_idx+1} 이미지 검색 중...")
                                img_path = fetcher.search_and_download(
                                    title=slide_title,
                                    content=content,
                                    filename=f"ai_ch{ch_num}_sl{sl_idx+1}.jpg"
                                )

                            builder.add_formatted_slide(
                                formatted,
                                chapter_num=ch_num,
                                chapter_title=ch_title,
                                image_path=img_path
                            )

                            done += 1
                            progress.progress(done / max(total_slides, 1))

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
