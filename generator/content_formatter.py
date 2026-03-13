"""
content_formatter.py
- 콘텐츠를 섹션(빈 줄 기준)으로 분리
- 각 섹션은 별도 텍스트박스로 배치
- 표(table) 형식 자동 감지
"""
import re
from collections import Counter


def detect_table_data(raw_text):
    lines = [l for l in raw_text.strip().split('\n') if l.strip()]
    if len(lines) < 2:
        return None

    def split_row(line):
        line = line.strip()
        if '\t' in line:
            return [c.strip() for c in line.split('\t') if c.strip()]
        if '|' in line:
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c and not re.match(r'^[-:]+$', c)]
            if len(cells) >= 2:
                return cells
        parts = re.split(r'\s{2,}', line)
        if len(parts) >= 2:
            return [p.strip() for p in parts if p.strip()]
        return None

    parsed_rows = []
    for line in lines:
        if re.match(r'^[\s|:-]+$', line):
            continue
        row = split_row(line)
        if row and len(row) >= 2:
            parsed_rows.append(row)
        else:
            return None

    if len(parsed_rows) < 2:
        return None

    col_count = len(parsed_rows[0])
    if not all(abs(len(r) - col_count) <= 1 for r in parsed_rows):
        return None

    max_cols = max(len(r) for r in parsed_rows)
    for r in parsed_rows:
        while len(r) < max_cols:
            r.append("")

    return {"is_table": True, "headers": parsed_rows[0], "rows": parsed_rows[1:]}


def parse_sections(raw_text):
    """
    빈 줄 기준으로 섹션을 분리.
    각 섹션의 첫 줄이 '영문 단어'만 포함하거나 짧으면 라벨로 간주.
    반환: [{"label": "Environmental", "lines": ["· 전체 온실가스...", ...]}, ...]
    """
    blocks = re.split(r'\n\s*\n', raw_text.strip())
    sections = []

    for block in blocks:
        lines = [l.rstrip() for l in block.split('\n') if l.strip()]
        if not lines:
            continue

        first = lines[0].strip()
        # 라벨 판별: 영문만, 또는 짧은 한글 제목 (15자 이하, 불릿/기호 없음)
        is_label = False
        if re.match(r'^[A-Za-z][A-Za-z\s&/\-]+$', first) and len(first) <= 40:
            is_label = True
        elif len(first) <= 15 and not first.startswith(('·', '•', '-', '*', '–')):
            is_label = True

        if is_label and len(lines) > 1:
            sections.append({"label": first, "lines": lines[1:]})
        else:
            sections.append({"label": "", "lines": lines})

    return sections


def format_content_for_slide(title, raw_text):
    """
    콘텐츠를 분석하여 레이아웃 결정.
    - 표가 감지되면 layout="table"
    - 2개 이상 섹션이 감지되면 layout="multi_section"
    - 그 외 기존 로직 (bullets, key_value, highlight, two_column)
    """
    # 표 감지
    table_data = detect_table_data(raw_text)
    if table_data and table_data["is_table"]:
        return {
            "title": title,
            "original_text": raw_text,
            "layout": "table",
            "sections": [{
                "type": "table",
                "headers": table_data["headers"],
                "rows": table_data["rows"]
            }]
        }

    # 섹션 분리
    sections = parse_sections(raw_text)

    if len(sections) >= 2:
        return {
            "title": title,
            "original_text": raw_text,
            "layout": "multi_section",
            "sections": [
                {"type": "section_block", "label": s["label"], "lines": s["lines"]}
                for s in sections
            ]
        }

    # 단일 블록 → 기존 로직
    lines = [l.strip() for l in raw_text.strip().split('\n') if l.strip()]
    if not lines:
        return {
            "title": title, "original_text": raw_text,
            "layout": "bullets",
            "sections": [{"type": "text", "content": "내용 없음"}]
        }

    kv_pattern = re.compile(r'^(.{2,20})\s*[:：\-–]\s*(.+)$')
    kv_count = sum(1 for l in lines if kv_pattern.match(l))
    total = len(lines)

    if kv_count >= 3 and kv_count / total >= 0.5:
        pairs, extras = [], []
        for l in lines:
            m = kv_pattern.match(l)
            if m:
                pairs.append({"key": m.group(1).strip(), "value": m.group(2).strip()})
            else:
                extras.append(l)
        return {
            "title": title, "original_text": raw_text,
            "layout": "key_value",
            "sections": [
                {"type": "key_value", "pairs": pairs},
                *([{"type": "text", "content": " ".join(extras)}] if extras else [])
            ]
        }

    if total <= 3 and all(len(l) > 30 for l in lines):
        return {
            "title": title, "original_text": raw_text,
            "layout": "highlight",
            "sections": [
                {"type": "highlight", "main": lines[0]},
                *([{"type": "support", "points": lines[1:]}] if len(lines) > 1 else [])
            ]
        }

    if total >= 6:
        mid = total // 2
        return {
            "title": title, "original_text": raw_text,
            "layout": "two_column",
            "sections": [
                {"type": "column", "side": "left", "points": lines[:mid]},
                {"type": "column", "side": "right", "points": lines[mid:]}
            ]
        }

    return {
        "title": title, "original_text": raw_text,
        "layout": "bullets",
        "sections": [{"type": "bullets", "points": lines}]
    }
