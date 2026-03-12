"""
image_fetcher.py
- Unsplash 이미지 검색
- 콘텐츠 맥락 기반 영어 키워드 자동 생성
"""
import requests
import os
import re
import hashlib


# 한국어 키워드 → 영어 + 이미지 검색용 태그 매핑
CONTEXT_IMAGE_MAP = {
    # 비즈니스 / 데이터
    "MAU": "user growth analytics dashboard",
    "이용자": "user growth chart",
    "사용자": "user analytics",
    "성장률": "business growth chart",
    "성장": "growth graph upward",
    "증가": "increase chart statistics",
    "매출": "revenue chart business",
    "수익": "profit chart finance",
    "시장": "market analysis business",
    "규모": "market size infographic",
    "점유율": "market share pie chart",
    "통계": "statistics data chart",
    "데이터": "data analytics dashboard",
    "분석": "analysis chart business",
    "차트": "chart graph business",
    "그래프": "graph data visualization",
    "비교": "comparison chart business",
    "추이": "trend line chart",
    "전망": "business forecast future",
    "트렌드": "trend analysis modern",
    "KPI": "KPI dashboard metrics",
    "지표": "metrics dashboard business",
    "실적": "business performance results",
    "달성": "goal achievement success",
    "목표": "target goal business",

    # 기술
    "AI": "artificial intelligence technology",
    "인공지능": "artificial intelligence brain",
    "자동화": "automation technology robot",
    "기술": "technology innovation",
    "디지털": "digital transformation",
    "클라우드": "cloud computing server",
    "플랫폼": "platform technology digital",
    "로봇": "robot automation industry",
    "반도체": "semiconductor chip technology",
    "배터리": "battery energy technology",
    "소프트웨어": "software development code",
    "빅데이터": "big data visualization",
    "블록체인": "blockchain technology",
    "IoT": "internet of things connected",
    "스마트": "smart technology innovation",

    # 금융 / 투자
    "투자": "investment finance growth",
    "금융": "finance banking modern",
    "주가": "stock market chart",
    "IPO": "IPO stock market business",
    "IR": "investor relations presentation",
    "재무": "financial report business",
    "자금": "funding investment capital",
    "수요예측": "demand forecasting chart",
    "공모": "public offering finance",

    # 산업
    "제조": "manufacturing factory industry",
    "유통": "distribution logistics supply",
    "물류": "logistics warehouse delivery",
    "건설": "construction building modern",
    "부동산": "real estate building city",
    "에너지": "energy renewable power",
    "의료": "medical healthcare hospital",
    "바이오": "biotechnology laboratory science",
    "제약": "pharmaceutical medicine lab",
    "식품": "food industry production",
    "농업": "agriculture farming modern",
    "관광": "tourism travel landscape",
    "교육": "education learning classroom",
    "환경": "environment sustainability green",
    "친환경": "eco friendly sustainable green",

    # 조직 / 경영
    "경영": "business management office",
    "전략": "strategy planning business",
    "리더십": "leadership team business",
    "팀": "team collaboration office",
    "조직": "organization team structure",
    "인재": "talent recruitment office",
    "혁신": "innovation creative idea",
    "협력": "cooperation partnership handshake",
    "경쟁": "competition business strategy",
    "고객": "customer service satisfaction",
    "마케팅": "marketing digital strategy",
    "브랜드": "brand identity design",
    "서비스": "service customer support",

    # 일반
    "개요": "overview summary business",
    "소개": "introduction presentation",
    "결론": "conclusion summary results",
    "요약": "summary overview document",
    "미래": "future vision technology",
    "글로벌": "global world business",
    "세계": "world global map",
    "도전": "challenge opportunity growth",
    "기회": "opportunity business bright",
    "문제": "problem solving strategy",
    "해결": "solution problem solving",
    "비전": "vision future business",
}


class ImageFetcher:
    def __init__(self, unsplash_key=None, save_dir="images"):
        self.unsplash_key = unsplash_key or ""
        self.save_dir = save_dir
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir)

    def _extract_search_query(self, title, content):
        """
        제목과 내용을 분석하여 최적의 영어 이미지 검색어 생성
        """
        combined = f"{title} {content}"

        # 1단계: 매핑 테이블에서 매칭되는 키워드 찾기 (우선순위 높음)
        matched_queries = []
        for ko_keyword, en_query in CONTEXT_IMAGE_MAP.items():
            if ko_keyword in combined:
                matched_queries.append(en_query)

        if matched_queries:
            # 가장 먼저 매칭된 2개 조합
            return " ".join(matched_queries[:2]).split()[:5]

        # 2단계: 영어 단어가 있으면 그대로 사용
        eng_words = re.findall(r'[a-zA-Z]{3,}', combined)
        if eng_words:
            return eng_words[:3]

        # 3단계: 기본 비즈니스 이미지
        return ["business", "professional", "modern"]

    def search_and_download(self, title, content="", filename=None):
        """
        제목과 내용을 기반으로 이미지 검색 후 다운로드
        """
        if not self.unsplash_key:
            return None

        keywords = self._extract_search_query(title, content)
        search_query = " ".join(keywords)

        try:
            url = "https://api.unsplash.com/search/photos"
            params = {
                "query": search_query,
                "per_page": 3,
                "orientation": "landscape",
                "content_filter": "high",
                "client_id": self.unsplash_key
            }
            resp = requests.get(url, params=params, timeout=10)

            if resp.status_code != 200:
                # 검색어 축소 후 재시도
                params["query"] = " ".join(keywords[:2])
                resp = requests.get(url, params=params, timeout=10)
                if resp.status_code != 200:
                    return None

            data = resp.json()
            results = data.get("results", [])
            if not results:
                # 더 일반적인 키워드로 재시도
                params["query"] = keywords[0] if keywords else "business"
                resp = requests.get(url, params=params, timeout=10)
                data = resp.json()
                results = data.get("results", [])
                if not results:
                    return None

            # 첫 번째 결과 사용
            img_url = results[0]["urls"].get("regular", results[0]["urls"]["small"])

            if not filename:
                filename = hashlib.md5(search_query.encode()).hexdigest()[:12] + ".jpg"
            filepath = os.path.join(self.save_dir, filename)

            img_resp = requests.get(img_url, timeout=15)
            if img_resp.status_code == 200:
                with open(filepath, "wb") as f:
                    f.write(img_resp.content)
                return filepath

        except Exception as e:
            print(f"Image fetch error: {e}")

        return None
