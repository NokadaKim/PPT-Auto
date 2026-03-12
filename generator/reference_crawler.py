"""
reference_crawler.py
Behance, Dribbble, Slidesgo에서 레퍼런스 이미지 수집 + AI 분석
"""

import requests
import os
import json
import hashlib
from g4f.client import Client

client = Client()
CACHE_DIR = "design_cache"
os.makedirs(CACHE_DIR, exist_ok=True)


def search_reference_images(keyword: str, tone: str, num_results=5) -> list:
    """
    Google 이미지 검색으로 PPT 디자인 레퍼런스 수집
    (Behance, Dribbble, Slidesgo 등에서 필터링)
    """
    search_queries = [
        f"{keyword} presentation slide design {tone}",
        f"{keyword} PPT layout design minimal professional",
        f"{keyword} slide deck design Behance Dribbble",
    ]

    results = []
    for query in search_queries:
        try:
            # SerpAPI 또는 무료 이미지 검색 API 사용
            # 여기서는 간단한 구현 예시
            api_url = "https://www.googleapis.com/customsearch/v1"
            params = {
                "q": query,
                "searchType": "image",
                "num": num_results,
                "key": os.getenv("GOOGLE_CSE_KEY", ""),
                "cx": os.getenv("GOOGLE_CSE_ID", ""),
                "imgSize": "large",
                "safe": "active"
            }

            if params["key"]:
                resp = requests.get(api_url, params=params, timeout=10)
                if resp.status_code == 200:
                    data = resp.json()
                    for item in data.get("items", []):
                        results.append({
                            "url": item["link"],
                            "title": item.get("title", ""),
                            "source": item.get("displayLink", ""),
                            "thumbnail": item.get("image", {}).get("thumbnailLink", "")
                        })
        except Exception as e:
            print(f"Reference search error: {e}")

    return results[:num_results]


def analyze_reference_design(image_url: str) -> dict:
    """
    AI로 레퍼런스 이미지의 디자인 요소를 분석
    - 컬러 팔레트
    - 레이아웃 구조
    - 타이포그래피 스타일
    - 여백/간격 패턴
    """
    # 캐시 체크
    cache_key = hashlib.md5(image_url.encode()).hexdigest()
    cache_file = os.path.join(CACHE_DIR, f"{cache_key}.json")

    if os.path.exists(cache_file):
        with open(cache_file, "r", encoding="utf-8") as f:
            return json.load(f)

    prompt = f"""
Analyze this presentation slide design image and extract design elements.
Image URL: {image_url}

Respond in this exact JSON format:
{{
    "color_palette": {{
        "primary": "#hex",
        "secondary": "#hex",
        "accent": "#hex",
        "background": "#hex",
        "text": "#hex"
    }},
    "layout_structure": {{
        "type": "bento_grid|split|full_image|cards|timeline|centered",
        "columns": 1-4,
        "has_sidebar": true/false,
        "content_alignment": "left|center|right",
        "image_position": "left|right|top|bottom|background|none"
    }},
    "typography": {{
        "title_style": "bold|light|uppercase|mixed",
        "title_size_ratio": "large|medium|small",
        "body_style": "regular|light",
        "line_spacing": "tight|normal|wide"
    }},
    "spacing": {{
        "margins": "narrow|medium|wide",
        "element_gap": "tight|medium|loose",
        "whitespace_ratio": "low|medium|high"
    }},
    "visual_elements": {{
        "has_icons": true/false,
        "has_shapes": true/false,
        "corner_style": "sharp|rounded|pill",
        "shadow_depth": "none|subtle|strong",
        "gradient": true/false
    }},
    "overall_mood": "description in Korean"
}}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.choices[0].message.content.strip()
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            result = json.loads(json_match.group())
            # 캐시 저장
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            return result
    except Exception as e:
        print(f"Design analysis error: {e}")

    return {}


import re

def get_design_recommendation(title: str, tone_key: str) -> dict:
    """
    키워드 + 톤에 맞는 디자인 추천을 종합적으로 생성
    """
    # 내장 디자인 규칙 (레퍼런스 사이트 분석 결과를 미리 학습한 것)
    DESIGN_RULES = {
        "tech_modern": {
            "bg_style": "dark_gradient",
            "card_style": "glass_morphism",
            "accent_usage": "neon_highlight",
            "preferred_layouts": ["bento_grid", "stat_cards", "split_two_column"],
            "reference_sites": [
                "https://dribbble.com/search/tech-presentation-dark",
                "https://www.behance.net/search/projects/tech%20presentation%20dark%20mode"
            ]
        },
        "corporate_formal": {
            "bg_style": "white_clean",
            "card_style": "flat_subtle_shadow",
            "accent_usage": "underline_accent",
            "preferred_layouts": ["split_two_column", "color_boxes", "stat_cards"],
            "reference_sites": [
                "https://dribbble.com/search/corporate-presentation",
                "https://www.behance.net/search/projects/corporate%20presentation%20minimal"
            ]
        },
        "creative_bold": {
            "bg_style": "vibrant_gradient",
            "card_style": "bold_rounded",
            "accent_usage": "color_blocks",
            "preferred_layouts": ["bento_grid", "full_image_overlay", "big_number"],
            "reference_sites": [
                "https://dribbble.com/search/creative-presentation-bold",
                "https://www.behance.net/search/projects/creative%20bold%20presentation"
            ]
        },
        "eco_nature": {
            "bg_style": "warm_earth",
            "card_style": "organic_rounded",
            "accent_usage": "nature_accent",
            "preferred_layouts": ["split_two_column", "icon_grid", "timeline"],
            "reference_sites": [
                "https://dribbble.com/search/eco-presentation",
                "https://www.behance.net/search/projects/eco%20sustainability%20presentation"
            ]
        },
        "minimal_elegant": {
            "bg_style": "white_clean",
            "card_style": "borderless",
            "accent_usage": "single_color_pop",
            "preferred_layouts": ["split_two_column", "quote_slide", "big_number"],
            "reference_sites": [
                "https://dribbble.com/tags/minimal-presentation",
                "https://www.behance.net/search/projects/minimal%20presentation"
            ]
        },
        "medical_clean": {
            "bg_style": "light_blue_white",
            "card_style": "clean_card",
            "accent_usage": "trust_blue_green",
            "preferred_layouts": ["stat_cards", "split_two_column", "icon_grid"],
            "reference_sites": [
                "https://dribbble.com/search/medical-presentation",
                "https://www.behance.net/search/projects/medical%20health%20presentation"
            ]
        },
        "edu_friendly": {
            "bg_style": "warm_light",
            "card_style": "playful_rounded",
            "accent_usage": "warm_highlight",
            "preferred_layouts": ["icon_grid", "bento_grid", "timeline"],
            "reference_sites": [
                "https://dribbble.com/search/education-presentation",
                "https://www.behance.net/search/projects/education%20presentation%20friendly"
            ]
        },
        "finance_trust": {
            "bg_style": "dark_navy",
            "card_style": "structured_card",
            "accent_usage": "green_highlight",
            "preferred_layouts": ["stat_cards", "big_number", "color_boxes"],
            "reference_sites": [
                "https://dribbble.com/search/finance-presentation",
                "https://www.behance.net/search/projects/finance%20investment%20presentation"
            ]
        }
    }

    rules = DESIGN_RULES.get(tone_key, DESIGN_RULES["minimal_elegant"])
    return {
        "tone_key": tone_key,
        "design_rules": rules,
        "reference_urls": rules["reference_sites"]
    }
