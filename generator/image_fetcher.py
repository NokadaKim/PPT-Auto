"""
image_fetcher.py
- Unsplash + Pexels 이중 검색
- 한국어 키워드 영어 자동 변환
"""
import requests
import os
import re
import hashlib


# 한국어 → 영어 키워드 매핑 (자주 쓰이는 비즈니스/기술 용어)
KO_EN_MAP = {
    "시장": "market", "규모": "scale", "성장": "growth",
    "기술": "technology", "인공지능": "artificial intelligence",
    "자동화": "automation", "데이터": "data", "분석": "analysis",
    "투자": "investment", "매출": "revenue", "수익": "profit",
    "전략": "strategy", "혁신": "innovation", "개발": "development",
    "산업": "industry", "제조": "manufacturing", "유통": "distribution",
    "금융": "finance", "의료": "medical", "건강": "health",
    "교육": "education", "환경": "environment", "에너지": "energy",
    "건설": "construction", "부동산": "real estate", "물류": "logistics",
    "마케팅": "marketing", "브랜드": "brand", "고객": "customer",
    "서비스": "service", "플랫폼": "platform", "클라우드": "cloud",
    "보안": "security", "네트워크": "network", "모바일": "mobile",
    "로봇": "robot", "반도체": "semiconductor", "배터리": "battery",
    "바이오": "biotechnology", "제약": "pharmaceutical",
    "식품": "food", "농업": "agriculture", "수산": "fishery",
    "관광": "tourism", "문화": "culture", "스포츠": "sports",
    "패션": "fashion", "뷰티": "beauty", "게임": "gaming",
    "미디어": "media", "콘텐츠": "content", "광고": "advertising",
    "경영": "management", "리더십": "leadership", "팀워크": "teamwork",
    "회의": "meeting", "발표": "presentation", "보고서": "report",
    "차트": "chart", "그래프": "graph", "통계": "statistics",
    "트렌드": "trend", "전망": "forecast", "비전": "vision",
    "목표": "goal", "성과": "achievement", "비교": "comparison",
    "문제": "problem", "해결": "solution", "도전": "challenge",
    "기회": "opportunity", "경쟁": "competition", "협력": "cooperation",
    "글로벌": "global", "세계": "world", "미래": "future",
    "디지털": "digital", "스마트": "smart", "친환경": "eco friendly",
    "지속가능": "sustainable", "ESG": "ESG",
}


class ImageFetcher:
    def __init__(self, unsplash_key=None, pexels_key=None, save_dir="images"):
        self.unsplash_key = unsplash_key or ""
        self.pexels_key = pexels_key or ""
        self.save_dir = save_dir
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir)

    def _translate_keywords(self, query):
        """한국어 키워드를 영어로 변환"""
        words = re.findall(r'[가-힣a-zA-Z0-9]+', query)
        en_words = []
        for w in words:
            if re.match(r'^[a-zA-Z0-9]+$', w):
                en_words.append(w)
            elif w in KO_EN_MAP:
                en_words.append(KO_EN_MAP[w])
            else:
                # 부분 매칭 시도
                for ko, en in KO_EN_MAP.items():
                    if ko in w:
                        en_words.append(en)
                        break

        if not en_words:
            en_words = ["business", "technology"]

        return " ".join(en_words[:3])

    def search_and_download(self, query, filename=None):
        """이미지 검색 후 다운로드. Unsplash 우선, 실패 시 Pexels"""
        # 영어 키워드로 변환
        search_query = self._translate_keywords(query)

        # 1차: Unsplash
        if self.unsplash_key:
            path = self._search_unsplash(search_query, filename)
            if path:
                return path

        # 2차: Pexels (키 없어도 시도)
        path = self._search_pexels(search_query, filename)
        if path:
            return path

        return None

    def _search_unsplash(self, query, filename=None):
        try:
            url = "https://api.unsplash.com/search/photos"
            params = {
                "query": query,
                "per_page": 1,
                "orientation": "landscape",
                "client_id": self.unsplash_key
            }
            resp = requests.get(url, params=params, timeout=10)
            if resp.status_code != 200:
                return None

            data = resp.json()
            results = data.get("results", [])
            if not results:
                return None

            img_url = results[0]["urls"].get("regular", results[0]["urls"]["small"])

            if not filename:
                filename = hashlib.md5(query.encode()).hexdigest()[:12] + ".jpg"
            filepath = os.path.join(self.save_dir, filename)

            img_resp = requests.get(img_url, timeout=15)
            if img_resp.status_code == 200:
                with open(filepath, "wb") as f:
                    f.write(img_resp.content)
                return filepath
        except Exception as e:
            print(f"Unsplash error: {e}")
        return None

    def _search_pexels(self, query, filename=None):
        """Pexels 무료 API로 이미지 검색 (백업)"""
        try:
            api_key = self.pexels_key or "563492ad6f91700001000001a1b2c3d4e5f6a7b8c9d0e1f2"
            url = "https://api.pexels.com/v1/search"
            headers = {"Authorization": api_key}
            params = {"query": query, "per_page": 1, "orientation": "landscape"}

            resp = requests.get(url, headers=headers, params=params, timeout=10)
            if resp.status_code != 200:
                return None

            data = resp.json()
            photos = data.get("photos", [])
            if not photos:
                return None

            img_url = photos[0]["src"].get("large", photos[0]["src"]["medium"])

            if not filename:
                filename = hashlib.md5(query.encode()).hexdigest()[:12] + ".jpg"
            filepath = os.path.join(self.save_dir, filename)

            img_resp = requests.get(img_url, timeout=15)
            if img_resp.status_code == 200:
                with open(filepath, "wb") as f:
                    f.write(img_resp.content)
                return filepath
        except Exception as e:
            print(f"Pexels error: {e}")
        return None
