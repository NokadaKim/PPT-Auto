"""
image_fetcher.py
- Unsplash API로 콘텐츠 관련 이미지 자동 검색 및 다운로드
"""
import requests
import os
import re
import hashlib


class ImageFetcher:
    def __init__(self, unsplash_key=None, save_dir="images"):
        self.unsplash_key = unsplash_key or ""
        self.save_dir = save_dir
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir)

    def search_and_download(self, query, filename=None):
        """
        Unsplash에서 query로 이미지 검색 후 다운로드
        반환: 로컬 이미지 경로 또는 None
        """
        if not self.unsplash_key:
            return None

        try:
            # 한국어 키워드 처리: 간단한 영어 변환 시도
            search_query = self._prepare_query(query)

            url = "https://api.unsplash.com/search/photos"
            params = {
                "query": search_query,
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

            # 다운로드
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

    def _prepare_query(self, query):
        """
        검색어 정리 - 너무 길면 핵심 키워드만 추출
        """
        # 한국어 포함 시 영단어 우선 추출
        eng_words = re.findall(r'[a-zA-Z]{3,}', query)
        if eng_words:
            return " ".join(eng_words[:3])

        # 한국어만 있으면 그대로 사용 (Unsplash가 일부 지원)
        words = query.split()
        return " ".join(words[:4])
