"""
generator/design_intelligence.py
키워드 → 톤앤매너 분석 → 최적 레이아웃 매칭
"""

from g4f.client import Client
import json
import re

client = Client()

LAYOUT_TYPES = {
    "content": {
        "name": "Standard Content",
        "desc": "일반 텍스트 + 불릿 포인트",
        "best_for": ["설명", "일반내용"]
    },
    "bento_grid": {
        "name": "Bento Grid",
        "desc": "카드형 격자 레이아웃",
        "best_for": ["비교", "기능소개", "특징나열", "4개이상_항목"]
    },
    "stat_cards": {
        "name": "Statistics Cards",
        "desc": "숫자/데이터 강조 카드형",
        "best_for": ["통계", "수치", "성과", "KPI", "시장규모"]
    },
    "timeline": {
        "name": "Timeline",
        "desc": "시간순 흐름/단계 표현",
        "best_for": ["역사", "로드맵", "과정", "단계", "프로세스", "연도"]
    },
    "colorbox": {
        "name": "Color Comparison Boxes",
        "desc": "두 가지 컬러 박스로 대비/비교",
        "best_for": ["비교", "장단점", "AS-IS/TO-BE", "전후", "vs"]
    },
    "quote_slide": {
        "name": "Quote / Callout",
        "desc": "인용문이나 핵심 메시지 강조",
        "best_for": ["인용", "핵심메시지", "결론", "비전"]
    },
    "big_number": {
        "name": "Big Number Focus",
        "desc": "하나의 핵심 숫자를 크게 강조",
        "best_for": ["핵심지표", "임팩트수치", "한줄요약", "시장규모"]
    },
    "icon_grid": {
        "name": "Icon Grid",
        "desc": "아이콘 + 짧은 설명 그리드",
        "best_for": ["핵심가치", "서비스목록", "기능목록"]
    },
}

TONE_PRESETS = {
    "corporate_formal": {
        "name": "Corporate Formal",
        "colors": {"primary": "#1A1A2E", "secondary": "#16213E", "accent": "#E64A19"},
        "style": "clean, minimal, structured",
        "corner_radius": 0.03,
        "spacing": "wide",
        "reference_urls": [
            "https://dribbble.com/search/corporate-presentation",
            "https://www.behance.net/search/projects/corporate%20presentation%20minimal"
        ]
    },
    "tech_modern": {
        "name": "Tech Modern",
        "colors": {"primary": "#0D1117", "secondary": "#161B22", "accent": "#58A6FF"},
        "style": "dark mode, gradient accents, bento grid",
        "corner_radius": 0.08,
        "spacing": "medium",
        "reference_urls": [
            "https://dribbble.com/search/tech-presentation-dark",
            "https://www.behance.net/search/projects/tech%20presentation%20dark%20mode"
        ]
    },
    "creative_bold": {
        "name": "Creative Bold",
        "colors": {"primary": "#6C2BD9", "secondary": "#F43F5E", "accent": "#FBBF24"},
        "style": "vibrant, asymmetric, oversized typography",
        "corner_radius": 0.12,
        "spacing": "medium",
        "reference_urls": [
            "https://dribbble.com/search/creative-presentation-bold",
            "https://www.behance.net/search/projects/creative%20bold%20presentation"
        ]
    },
    "eco_nature": {
        "name": "Eco / Nature",
        "colors": {"primary": "#064E3B", "secondary": "#059669", "accent": "#D97706"},
        "style": "earthy tones, soft curves, organic feel",
        "corner_radius": 0.10,
        "spacing": "wide",
        "reference_urls": [
            "https://dribbble.com/search/eco-presentation",
            "https://www.behance.net/search/projects/sustainability%20presentation"
        ]
    },
    "medical_clean": {
        "name": "Medical Clean",
        "colors": {"primary": "#0E4DA4", "secondary": "#3B82F6", "accent": "#10B981"},
        "style": "clean, trustworthy, blue tones",
        "corner_radius": 0.06,
        "spacing": "wide",
        "reference_urls": [
            "https://dribbble.com/search/medical-presentation",
            "https://www.behance.net/search/projects/medical%20health%20presentation"
        ]
    },
    "edu_friendly": {
        "name": "Education Friendly",
        "colors": {"primary": "#7C3AED", "secondary": "#8B5CF6", "accent": "#F59E0B"},
        "style": "warm, approachable, playful accents",
        "corner_radius": 0.10,
        "spacing": "medium",
        "reference_urls": [
            "https://dribbble.com/search/education-presentation",
            "https://www.behance.net/search/projects/education%20presentation"
        ]
    },
    "finance_trust": {
        "name": "Finance Trust",
        "colors": {"primary": "#1E293B", "secondary": "#334155", "accent": "#22C55E"},
        "style": "professional, data-driven, dark navy",
        "corner_radius": 0.04,
        "spacing": "wide",
        "reference_urls": [
            "https://dribbble.com/search/finance-presentation",
            "https://www.behance.net/search/projects/finance%20presentation"
        ]
    },
    "minimal_elegant": {
        "name": "Minimal Elegant",
        "colors": {"primary": "#18181B", "secondary": "#3F3F46", "accent": "#EF4444"},
        "style": "lots of whitespace, subtle accents",
        "corner_radius": 0.02,
        "spacing": "extra-wide",
        "reference_urls": [
            "https://dribbble.com/tags/minimal-presentation",
            "https://www.behance.net/search/projects/minimal%20presentation"
        ]
    },
}


def analyze_tone_and_manner(title: str) -> dict:
    prompt = f"""You are a presentation design consultant.
Analyze this title and choose the best tone.

Title: "{title}"

Choose ONE from: {json.dumps(list(TONE_PRESETS.keys()))}

Respond ONLY in this JSON format:
{{"tone": "key_name", "reason": "brief reason in Korean"}}"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group())
            tone_key = result.get("tone", "minimal_elegant")
            if tone_key in TONE_PRESETS:
                return {
                    "tone_key": tone_key,
                    "preset": TONE_PRESETS[tone_key],
                    "reason": result.get("reason", ""),
                }
    except Exception as e:
        print(f"Tone analysis error: {e}")

    return {
        "tone_key": "minimal_elegant",
        "preset": TONE_PRESETS["minimal_elegant"],
        "reason": "기본 스타일 적용",
    }


def analyze_content_layout(slide_title: str, content_points: list) -> str:
    layout_info = json.dumps(
        {k: v["best_for"] for k, v in LAYOUT_TYPES.items()},
        ensure_ascii=False
    )
    points_text = json.dumps(content_points, ensure_ascii=False)

    prompt = f"""Choose the best slide layout for this content.

Title: "{slide_title}"
Points: {points_text}

Layouts: {layout_info}

Rules:
- If 4+ short items → bento_grid or icon_grid
- If numbers/percentages/money → stat_cards or big_number
- If comparing two things → colorbox
- If process/steps/years → timeline
- If quote or key message → quote_slide
- Otherwise → content

Respond with ONLY the layout key, nothing else."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
        )
        layout_key = response.choices[0].message.content.strip().strip('"').strip("'").lower()
        if layout_key in LAYOUT_TYPES:
            return layout_key
    except Exception as e:
        print(f"Layout analysis error: {e}")

    return "content"


def restructure_content_for_layout(slide_title: str, original_points: list, layout_type: str) -> dict:
    if layout_type == "content":
        return {"type": "content", "title": slide_title, "points": original_points}

    points_text = "\n".join(original_points)

    format_instructions = {
        "bento_grid": '{"type":"bento_grid","title":"...","items":[{"title":"short title","desc":"description"},...]  (3-6 items)}',
        "stat_cards": '{"type":"stat_cards","title":"...","stats":[{"number":"85%","label":"description"},...]  (2-4 items)}',
        "timeline": '{"type":"timeline","title":"...","steps":[{"year":"2024","title":"short","desc":"detail"},...]  (3-6 steps)}',
        "colorbox": '{"type":"colorbox","title":"...","box1_title":"Left Title","box1_items":["item1","item2"],"box2_title":"Right Title","box2_items":["item1","item2"]}',
        "quote_slide": '{"type":"quote_slide","title":"...","quote":"the quote text","author":"optional author"}',
        "big_number": '{"type":"big_number","title":"...","number":"$25B","description":"main description","subtitle":"optional detail"}',
        "icon_grid": '{"type":"icon_grid","title":"...","items":[{"title":"Feature","desc":"short description"},...]  (4-6 items)}',
    }

    fmt = format_instructions.get(layout_type, format_instructions["bento_grid"])

    prompt = f"""Restructure this slide content into the specified layout format.
Keep the original language (Korean or English). Keep all key information.

Slide Title: "{slide_title}"
Original Content:
{points_text}

Target Layout: {layout_type}
Required JSON format: {fmt}

Respond with ONLY valid JSON, no other text."""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group())
            result["type"] = layout_type
            if "title" not in result:
                result["title"] = slide_title
            return result
    except Exception as e:
        print(f"Content restructure error: {e}")

    return {"type": "content", "title": slide_title, "points": original_points}
