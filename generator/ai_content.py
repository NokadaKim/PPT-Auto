from g4f.client import Client
import json


class AIContentGenerator:
    def __init__(self):
        self.client = Client()

    def extract_json(self, text):
        text = text.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            text = "\n".join(lines[1:])
            if text.endswith("```"):
                text = text[:-3].strip()
            elif "```" in text:
                text = text[:text.rfind("```")].strip()

        start = text.find("{")
        if start == -1:
            return None

        depth = 0
        for i in range(start, len(text)):
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    return text[start:i + 1]
        return None

    def generate_outline(self, topic, num_chapters=3, slides_per_chapter=2, language="ko"):
        if language == "ko":
            lang_instruction = """한국어를 기본으로 작성하되, 아래 규칙을 따르세요:
- 전문 용어, 브랜드명, 기술명은 영문 그대로 사용 (예: AI, Marketing, ROI, KPI, Digital Transformation)
- 챕터 제목에 영문 키워드를 자연스럽게 포함 (예: "디지털 마케팅 Trend", "AI Technology 활용")
- 통계 수치와 퍼센트는 숫자로 표기
- 자연스러운 한영 혼용체 사용"""
        else:
            lang_instruction = "Write all content in English."

        prompt = f"""You are a professional presentation designer. Create a detailed PPT outline.

Topic: {topic}

Structure: {num_chapters} chapters, {slides_per_chapter} slides per chapter.

{lang_instruction}

Rules:
- Each slide must have 3-5 bullet points (each 1-2 sentences)
- Important keywords should be marked with * at the beginning
- Each chapter should have a clear focus
- Professional and persuasive tone
- If comparing two things, use "colorbox" as slide type

Respond ONLY with valid JSON. No extra text before or after the JSON:
{{
    "title": "Presentation title",
    "subtitle": "Subtitle",
    "chapters": [
        {{
            "chapter_title": "Chapter 1 Title",
            "slides": [
                {{
                    "title": "Slide Title",
                    "type": "content",
                    "bullet_points": [
                        "First point",
                        "*Highlighted important point",
                        "Third point"
                    ]
                }}
            ]
        }}
    ]
}}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
            )

            txt = response.choices[0].message.content.strip()
            json_str = self.extract_json(txt)
            if json_str is None:
                return None, "AI 응답에서 JSON을 찾을 수 없습니다."

            outline = json.loads(json_str)
            return outline, None

        except json.JSONDecodeError as e:
            return None, f"AI 응답이 올바른 JSON이 아닙니다: {e}"
        except Exception as e:
            return None, f"오류 발생: {e}"
