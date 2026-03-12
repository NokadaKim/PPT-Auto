"""
content_formatter.py
- 슬라이드 내용을 분석하여 한 페이지 내에서 레이아웃을 결정
- 표(table) 형식 자동 감지
- 내용을 페이지 분리하지 않고, 한 슬라이드 안에서 구조화
"""
import re
from collections import Counter


def detect_table_data(raw_text):
    """
    탭(\t) 또는 여러 칸 공백으로 구분된 표 형식 데이터를 감지
    반환: {"is_table": bool, "headers": list, "rows": list[list]} 또는 None
    """
    lines = [l for l in raw_text.strip().split('\n') if l.strip()]

    if len(lines) < 2:
        return None

    # 구분자 감지: 탭, |, 2개 이상 연속 공백
    def split_row(line):
        line = line.strip()
        # 1) 탭 구분
        if '\t' in line:
            return [c.strip() for c in line.split('\t') if c.strip()]
        # 2) | 구분 (마크다운 테이블)
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c and not re.match(r'^[-:]+$', c)]
            if len(cells) >= 2:
                return cells
        # 3) 2칸 이상 공백 구분
        parts = re.split(r'\s{2,}', line)
        if len(parts) >= 2:
            return [p.strip() for p in parts if p.strip()]
        return None

    # 모든 행 파싱 시도
    parsed_rows = []
    for line in lines:
        # 마크다운 구분선 (---) 건너뛰기
        if re.match(r'^[\s|:-]+$', line):
            continue
        row = split_row(line)
        if row and len(row) >= 2:
            parsed_rows.append(row)
        else:
            # 표가 아닌 줄이 섞여있으면 표로 간주하지 않음
            return None

    if len(parsed_rows) < 2:
        return None

    # 열 개수 일관성 확인 (첫 행 기준 ±1 허용)
    col_count = len(parsed_rows[0])
    consistent = all(abs(len(r) - col_count) <= 1 for r in parsed_rows)

    if not consistent:
        return None

    # 열 개수 통일 (부족하면 빈 문자열 추가)
    max_cols = max(len(r) for r in parsed_rows)
    for r in parsed_rows:
        while len(r) < max_cols:
            r.append("")

    return {
        "is_table": True,
        "headers": parsed_rows[0],
        "rows": parsed_rows[1:]
    }


def analyze_content_structure(raw_text):
    lines = [l.strip() for l in raw_text.strip().split('\n') if l.strip()]

    if not lines:
        return {"layout": "bullets", "sections": [{"type": "text", "content": "내용 없음"}]}

    # ★ 표 감지 우선
    table_data = detect_table_data(raw_text)
    if table_data and table_data["is_table"]:
        return {
            "layout": "table",
            "sections": [{
                "type": "table",
                "headers": table_data["headers"],
                "rows": table_data["rows"]
            }]
        }

    # key:value 패턴 감지
    kv_pattern = re.compile(r'^(.{2,20})\s*[:：\-–]\s*(.+)$')
    kv_count = sum(1 for l in lines if kv_pattern.match(l))
    total = len(lines)

    if kv_count >= 3 and kv_count / total >= 0.5:
        pairs = []
        extras = []
        for l in lines:
            m = kv_pattern.match(l)
            if m:
                pairs.append({"key": m.group(1).strip(), "value": m.group(2).strip()})
            else:
                extras.append(l)
        return {
            "layout": "key_value",
            "sections": [
                {"type": "key_value", "pairs": pairs},
                *([{"type": "text", "content": " ".join(extras)}] if extras else [])
            ]
        }

    # 짧은 강조 문장
    if total <= 3 and all(len(l) > 30 for l in lines):
        return {
            "layout": "highlight",
            "sections": [
                {"type": "highlight", "main": lines[0]},
                *([{"type": "support", "points": lines[1:]}] if len(lines) > 1 else [])
            ]
        }

    # 긴 내용은 two_column
    if total >= 6:
        mid = total // 2
        return {
            "layout": "two_column",
            "sections": [
                {"type": "column", "side": "left", "points": lines[:mid]},
                {"type": "column", "side": "right", "points": lines[mid:]}
            ]
        }

    return {
        "layout": "bullets",
        "sections": [{"type": "bullets", "points": lines}]
    }


def format_content_for_slide(title, raw_text):
    structure = analyze_content_structure(raw_text)
    structure["title"] = title
    structure["original_text"] = raw_text
    return structure


def extract_image_keywords(title, raw_text, lang='ko'):
    """
    제목과 내용에서 이미지 검색용 키워드를 추출
    """
    stop_words_ko = {
        '의', '에', '를', '을', '이', '가', '는', '은', '로', '으로',
        '와', '과', '에서', '까지', '부터', '대한', '위한', '통한',
        '관한', '및', '등', '그', '그리고', '또는', '하는', '있는',
        '없는', '되는', '한', '수', '것', '때', '더', '매우'
    }

    combined = f"{title} {raw_text[:200]}"
    words = re.findall(r'[가-힣a-zA-Z0-9]+', combined)
    keywords = [w for w in words if len(w) > 1 and w not in stop_words_ko]

    freq = Counter(keywords)
    top_keywords = [w for w, _ in freq.most_common(5)]

    eng_keywords = [w for w in top_keywords if re.match(r'^[a-zA-Z]+$', w)]

    if eng_keywords:
        return " ".join(eng_keywords[:3])
    else:
        return title
