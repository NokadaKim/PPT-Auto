"""
ppt_builder.py
- 모든 콘텐츠를 한 슬라이드 내에 배치
- 레이아웃별 렌더링 (bullets, two_column, key_value, highlight, table)
- 이미지 삽입 지원
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os
import re


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def is_korean(text):
    return bool(re.search(r'[가-힣]', text))


def get_font_name(text):
    return '에스코어 드림 5 Medium' if is_korean(text) else 'Montserrat SemiBold'


class PPTBuilder:
    def __init__(self, theme="dark_modern", custom_colors=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.slide_count = 0
        self.theme = theme

        themes = {
            "dark_modern": {
                "bg": "#1A1A2E", "title": "#FFFFFF", "text": "#E0E0E0",
                "accent": "#E94560", "accent2": "#0F3460", "card_bg": "#16213E",
                "subtitle": "#B0B0B0", "label_bg": "#E94560",
                "table_header_bg": "#E94560", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#16213E", "table_row_bg2": "#1A1A2E",
                "table_border": "#0F3460"
            },
            "light_clean": {
                "bg": "#FFFFFF", "title": "#2D2D2D", "text": "#4A4A4A",
                "accent": "#4A90D9", "accent2": "#F0F4F8", "card_bg": "#F7F9FC",
                "subtitle": "#7A7A7A", "label_bg": "#4A90D9",
                "table_header_bg": "#4A90D9", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#F7F9FC", "table_row_bg2": "#FFFFFF",
                "table_border": "#D0D8E0"
            },
            "nature_green": {
                "bg": "#F5F7F0", "title": "#2D4A2D", "text": "#3D5A3D",
                "accent": "#6B8F3C", "accent2": "#E8F0DC", "card_bg": "#EDF2E4",
                "subtitle": "#6B8F3C", "label_bg": "#6B8F3C",
                "table_header_bg": "#6B8F3C", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#EDF2E4", "table_row_bg2": "#F5F7F0",
                "table_border": "#C5D4A8"
            },
            "corporate_blue": {
                "bg": "#F0F4F8", "title": "#1A365D", "text": "#2D4A7A",
                "accent": "#3182CE", "accent2": "#EBF4FF", "card_bg": "#FFFFFF",
                "subtitle": "#4A6FA5", "label_bg": "#3182CE",
                "table_header_bg": "#3182CE", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#EBF4FF", "table_row_bg2": "#F0F4F8",
                "table_border": "#B0C4DE"
            },
            "warm_orange": {
                "bg": "#FFF8F0", "title": "#7C3A1A", "text": "#5A3E2A",
                "accent": "#E67E22", "accent2": "#FFF0E0", "card_bg": "#FFFFFF",
                "subtitle": "#B86B2A", "label_bg": "#E67E22",
                "table_header_bg": "#E67E22", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#FFF0E0", "table_row_bg2": "#FFF8F0",
                "table_border": "#F0C8A0"
            },
            "purple_creative": {
                "bg": "#F5F0FF", "title": "#4A1A7A", "text": "#5A3D7A",
                "accent": "#8B5CF6", "accent2": "#EDE5FF", "card_bg": "#FFFFFF",
                "subtitle": "#7C4DFF", "label_bg": "#8B5CF6",
                "table_header_bg": "#8B5CF6", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#EDE5FF", "table_row_bg2": "#F5F0FF",
                "table_border": "#C4B0F0"
            },
            "minimal_gray": {
                "bg": "#FAFAFA", "title": "#333333", "text": "#555555",
                "accent": "#888888", "accent2": "#F0F0F0", "card_bg": "#FFFFFF",
                "subtitle": "#999999", "label_bg": "#666666",
                "table_header_bg": "#666666", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#F5F5F5", "table_row_bg2": "#FAFAFA",
                "table_border": "#DDDDDD"
            },
            "tech_dark": {
                "bg": "#0D1117", "title": "#F0F6FC", "text": "#C9D1D9",
                "accent": "#58A6FF", "accent2": "#161B22", "card_bg": "#161B22",
                "subtitle": "#8B949E", "label_bg": "#1F6FEB",
                "table_header_bg": "#1F6FEB", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#161B22", "table_row_bg2": "#0D1117",
                "table_border": "#30363D"
            },
        }

        if custom_colors and isinstance(custom_colors, dict):
            self.colors = custom_colors
        else:
            self.colors = themes.get(theme, themes["dark_modern"])

    def _set_slide_bg(self, slide):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(self.colors["bg"])

    def _add_page_number(self, slide):
        self.slide_count += 1
        txBox = slide.shapes.add_textbox(Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.3))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = str(self.slide_count)
        p.font.size = Pt(10)
        p.font.color.rgb = hex_to_rgb(self.colors.get("subtitle", "#999999"))
        p.alignment = PP_ALIGN.RIGHT

    def _add_label(self, slide, label_text, left=Inches(0.6), top=Inches(0.4)):
        if not label_text:
            return
        lbl = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(2.2), Inches(0.35)
        )
        lbl.fill.solid()
        lbl.fill.fore_color.rgb = hex_to_rgb(self.colors.get("label_bg", self.colors["accent"]))
        lbl.line.fill.background()
        if hasattr(lbl, 'adjustments') and len(lbl.adjustments) > 0:
            lbl.adjustments[0] = 0.3
        tf = lbl.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = label_text
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = get_font_name(label_text)
        p.alignment = PP_ALIGN.CENTER

    # ─────────────────────────────────────────
    # 표지 슬라이드
    # ─────────────────────────────────────────
    def add_title_slide(self, title, subtitle=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)

        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(3.2), Inches(3), Inches(0.06)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = hex_to_rgb(self.colors["accent"])
        line.line.fill.background()

        txBox = slide.shapes.add_textbox(Inches(1.5), Inches(3.4), Inches(10), Inches(1.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        if subtitle:
            txBox2 = slide.shapes.add_textbox(Inches(1.5), Inches(4.9), Inches(10), Inches(0.8))
            tf2 = txBox2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = subtitle
            p2.font.size = Pt(20)
            p2.font.color.rgb = hex_to_rgb(self.colors.get("subtitle", "#999999"))
            p2.font.name = get_font_name(subtitle)

        self._add_page_number(slide)

    # ─────────────────────────────────────────
    # Bullets 레이아웃
    # ─────────────────────────────────────────
    def add_bullets_slide(self, title, points, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)
        self._add_label(slide, slide_label)

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        content_width = Inches(7.5) if image_path else Inches(11.5)

        txBox2 = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), content_width, Inches(5.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True

        num_points = len(points) if isinstance(points, list) else 1
        if num_points <= 3:
            font_size, spacing = Pt(18), Pt(14)
        elif num_points <= 5:
            font_size, spacing = Pt(16), Pt(10)
        elif num_points <= 8:
            font_size, spacing = Pt(14), Pt(8)
        else:
            font_size, spacing = Pt(12), Pt(6)

        if isinstance(points, list):
            for i, point in enumerate(points):
                p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
                p.text = f"•  {point}"
                p.font.size = font_size
                p.font.color.rgb = hex_to_rgb(self.colors["text"])
                p.font.name = get_font_name(point)
                p.space_after = spacing
        else:
            p = tf2.paragraphs[0]
            p.text = str(points)
            p.font.size = font_size
            p.font.color.rgb = hex_to_rgb(self.colors["text"])
            p.font.name = get_font_name(str(points))

        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, Inches(8.8), Inches(1.8), Inches(4.0), Inches(4.0))
            except Exception:
                pass

        self._add_page_number(slide)

    # ─────────────────────────────────────────
    # Two Column 레이아웃
    # ─────────────────────────────────────────
    def add_two_column_slide(self, title, left_points, right_points, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)
        self._add_label(slide, slide_label)

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        col_width = Inches(5.5)
        col_top = Inches(1.9)

        if image_path and os.path.exists(image_path):
            col_width = Inches(4.0)

        for side, pts, x_offset in [("left", left_points, Inches(0.6)), ("right", right_points, Inches(0.6) + col_width + Inches(0.3))]:
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, x_offset, col_top, col_width, Inches(4.8)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.colors.get("card_bg", self.colors["bg"]))
            card.line.color.rgb = hex_to_rgb(self.colors.get("accent2", "#EEEEEE"))
            card.line.width = Pt(1)
            if hasattr(card, 'adjustments') and len(card.adjustments) > 0:
                card.adjustments[0] = 0.03

            self._fill_column(slide, x_offset + Inches(0.3), col_top + Inches(0.3), col_width - Inches(0.6), pts)

        if image_path and os.path.exists(image_path):
            try:
                img_left = Inches(0.6) + col_width * 2 + Inches(0.6)
                slide.shapes.add_picture(image_path, img_left, col_top, Inches(3.5), Inches(4.8))
            except Exception:
                pass

        self._add_page_number(slide)

    def _fill_column(self, slide, left, top, width, points):
        txBox = slide.shapes.add_textbox(left, top, width, Inches(4.2))
        tf = txBox.text_frame
        tf.word_wrap = True
        num = len(points)
        fsize = Pt(14) if num <= 5 else Pt(12)
        spacing = Pt(8) if num <= 5 else Pt(5)
        for i, pt in enumerate(points):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"•  {pt}"
            p.font.size = fsize
            p.font.color.rgb = hex_to_rgb(self.colors["text"])
            p.font.name = get_font_name(pt)
            p.space_after = spacing

    # ─────────────────────────────────────────
    # Key-Value 레이아웃
    # ─────────────────────────────────────────
    def add_key_value_slide(self, title, pairs, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)
        self._add_label(slide, slide_label)

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        max_cols = 3 if not image_path else 2
        card_area_width = Inches(8.5) if image_path else Inches(12.0)
        card_w = (card_area_width - Inches(0.3) * (max_cols - 1)) / max_cols
        card_h = Inches(2.0)
        start_left = Inches(0.6)
        start_top = Inches(2.0)

        for idx, pair in enumerate(pairs[:6]):
            col = idx % max_cols
            row = idx // max_cols
            c_left = start_left + col * (card_w + Inches(0.3))
            c_top = start_top + row * (card_h + Inches(0.3))

            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, int(c_left), c_top, int(card_w), card_h
            )
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.colors.get("card_bg", self.colors["bg"]))
            card.line.color.rgb = hex_to_rgb(self.colors.get("accent2", "#EEEEEE"))
            card.line.width = Pt(1)
            if hasattr(card, 'adjustments') and len(card.adjustments) > 0:
                card.adjustments[0] = 0.05

            key_text = pair.get("key", "")
            tb_key = slide.shapes.add_textbox(int(c_left) + Inches(0.2), c_top + Inches(0.25), int(card_w) - Inches(0.4), Inches(0.5))
            tf_k = tb_key.text_frame
            tf_k.word_wrap = True
            pk = tf_k.paragraphs[0]
            pk.text = key_text
            pk.font.size = Pt(16)
            pk.font.bold = True
            pk.font.color.rgb = hex_to_rgb(self.colors["accent"])
            pk.font.name = get_font_name(key_text)

            val_text = pair.get("value", "")
            tb_val = slide.shapes.add_textbox(int(c_left) + Inches(0.2), c_top + Inches(0.85), int(card_w) - Inches(0.4), Inches(1.0))
            tf_v = tb_val.text_frame
            tf_v.word_wrap = True
            pv = tf_v.paragraphs[0]
            pv.text = val_text
            pv.font.size = Pt(13)
            pv.font.color.rgb = hex_to_rgb(self.colors["text"])
            pv.font.name = get_font_name(val_text)

        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, Inches(9.3), Inches(2.0), Inches(3.5), Inches(4.3))
            except Exception:
                pass

        self._add_page_number(slide)

    # ─────────────────────────────────────────
    # Highlight 레이아웃
    # ─────────────────────────────────────────
    def add_highlight_slide(self, title, main_text, support_points=None, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)
        self._add_label(slide, slide_label)

        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        highlight_width = Inches(7.5) if image_path else Inches(11.5)
        hl_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(2.0), highlight_width, Inches(2.0)
        )
        hl_box.fill.solid()
        hl_box.fill.fore_color.rgb = hex_to_rgb(self.colors["accent"])
        hl_box.line.fill.background()
        if hasattr(hl_box, 'adjustments') and len(hl_box.adjustments) > 0:
            hl_box.adjustments[0] = 0.05

        tb_hl = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), highlight_width - Inches(0.8), Inches(1.5))
        tf_hl = tb_hl.text_frame
        tf_hl.word_wrap = True
        p_hl = tf_hl.paragraphs[0]
        p_hl.text = main_text
        p_hl.font.size = Pt(22)
        p_hl.font.bold = True
        p_hl.font.color.rgb = RGBColor(255, 255, 255)
        p_hl.font.name = get_font_name(main_text)
        p_hl.alignment = PP_ALIGN.CENTER

        if support_points:
            tb_sp = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), highlight_width, Inches(2.5))
            tf_sp = tb_sp.text_frame
            tf_sp.word_wrap = True
            for i, sp in enumerate(support_points):
                p_sp = tf_sp.paragraphs[0] if i == 0 else tf_sp.add_paragraph()
                p_sp.text = f"•  {sp}"
                p_sp.font.size = Pt(15)
                p_sp.font.color.rgb = hex_to_rgb(self.colors["text"])
                p_sp.font.name = get_font_name(sp)
                p_sp.space_after = Pt(8)

        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, Inches(8.8), Inches(2.0), Inches(4.0), Inches(4.5))
            except Exception:
                pass

        self._add_page_number(slide)

    # ─────────────────────────────────────────
    # ★ 표(Table) 레이아웃 - 신규
    # ─────────────────────────────────────────
    def add_table_slide(self, title, headers, rows, image_path=None, slide_label=""):
        """
        표 형식 데이터를 슬라이드에 테이블로 렌더링
        headers: ["항목", "ChatGPT", "Claude", ...]
        rows: [["추정 MAU", "~8.3억+", "~1,900만", ...], ...]
        """
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)
        self._add_label(slide, slide_label)

        # 제목
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(30)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["title"])
        p.font.name = get_font_name(title)

        # 테이블 크기 계산
        num_rows = len(rows) + 1  # 헤더 포함
        num_cols = len(headers)

        # 이미지가 있으면 테이블 폭 줄임
        table_width = Inches(8.0) if image_path else Inches(12.0)
        table_height = min(Inches(5.0), Inches(0.55) * num_rows)
        table_left = Inches(0.6)
        table_top = Inches(1.9)

        # 테이블 추가
        table_shape = slide.shapes.add_table(
            num_rows, num_cols, table_left, table_top, table_width, table_height
        )
        table = table_shape.table

        # 열 너비 설정: 첫 열은 좁게, 나머지 균등
        first_col_width = Inches(1.8)
        remaining_width = table_width - first_col_width
        other_col_width = remaining_width / (num_cols - 1) if num_cols > 1 else remaining_width

        table.columns[0].width = int(first_col_width)
        for c in range(1, num_cols):
            table.columns[c].width = int(other_col_width)

        # 행 높이 자동 조절
        row_height = int(table_height / num_rows)
        for r in range(num_rows):
            table.rows[r].height = row_height

        # 색상 가져오기
        header_bg = hex_to_rgb(self.colors.get("table_header_bg", self.colors["accent"]))
        header_text_color = hex_to_rgb(self.colors.get("table_header_text", "#FFFFFF"))
        row_bg1 = hex_to_rgb(self.colors.get("table_row_bg1", self.colors.get("card_bg", "#F5F5F5")))
        row_bg2 = hex_to_rgb(self.colors.get("table_row_bg2", self.colors["bg"]))
        border_color = hex_to_rgb(self.colors.get("table_border", "#DDDDDD"))
        text_color = hex_to_rgb(self.colors["text"])

        # 폰트 크기 계산 (행/열 수에 따라 자동 조절)
        if num_rows <= 5 and num_cols <= 4:
            header_font_size = Pt(14)
            cell_font_size = Pt(13)
        elif num_rows <= 8 and num_cols <= 6:
            header_font_size = Pt(12)
            cell_font_size = Pt(11)
        else:
            header_font_size = Pt(10)
            cell_font_size = Pt(9)

        # ─── 헤더 행 ───
        for col_idx, header_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = ""

            # 배경색
            cell_fill = cell.fill
            cell_fill.solid()
            cell_fill.fore_color.rgb = header_bg

            # 텍스트
            para = cell.text_frame.paragraphs[0]
            para.text = header_text
            para.font.size = header_font_size
            para.font.bold = True
            para.font.color.rgb = header_text_color
            para.font.name = get_font_name(header_text)
            para.alignment = PP_ALIGN.CENTER
            cell.text_frame.word_wrap = True
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        # ─── 데이터 행 ───
        for row_idx, row_data in enumerate(rows):
            bg_color = row_bg1 if row_idx % 2 == 0 else row_bg2

            for col_idx in range(num_cols):
                cell = table.cell(row_idx + 1, col_idx)
                cell_text = row_data[col_idx] if col_idx < len(row_data) else ""
                cell.text = ""

                # 배경색 (줄무늬)
                cell_fill = cell.fill
                cell_fill.solid()
                cell_fill.fore_color.rgb = bg_color

                # 텍스트
                para = cell.text_frame.paragraphs[0]
                para.text = cell_text
                para.font.size = cell_font_size
                para.font.color.rgb = text_color
                para.font.name = get_font_name(cell_text)
                para.alignment = PP_ALIGN.CENTER
                cell.text_frame.word_wrap = True
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                # 첫 열은 볼드 + 왼쪽 정렬
                if col_idx == 0:
                    para.font.bold = True
                    para.alignment = PP_ALIGN.LEFT

        # 테이블 테두리 설정
        self._set_table_borders(table, num_rows, num_cols, border_color)

        # 이미지
        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(
                    image_path, Inches(9.0), Inches(1.9), Inches(3.8), Inches(5.0)
                )
            except Exception:
                pass

        self._add_page_number(slide)

    def _set_table_borders(self, table, num_rows, num_cols, border_color):
        """테이블 셀 테두리 설정"""
        from pptx.oxml.ns import qn

        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                cell = table.cell(row_idx, col_idx)
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()

                for border_name in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
                    ln = tcPr.find(qn(border_name))
                    if ln is None:
                        ln = tcPr.makeelement(qn(border_name), {})
                        tcPr.append(ln)

                    ln.set('w', '12700')  # 1pt
                    ln.set('cap', 'flat')
                    ln.set('cmpd', 'sng')

                    solidFill = ln.find(qn('a:solidFill'))
                    if solidFill is None:
                        solidFill = ln.makeelement(qn('a:solidFill'), {})
                        ln.insert(0, solidFill)

                    srgbClr = solidFill.find(qn('a:srgbClr'))
                    if srgbClr is None:
                        srgbClr = solidFill.makeelement(qn('a:srgbClr'), {})
                        solidFill.append(srgbClr)

                    srgbClr.set('val', str(border_color).replace('#', ''))

    # ─────────────────────────────────────────
    # 통합 렌더 함수
    # ─────────────────────────────────────────
    def add_formatted_slide(self, formatted_data, image_path=None, slide_label=""):
        """
        content_formatter.format_content_for_slide()의 반환값을 받아
        한 슬라이드에 모든 내용을 배치
        """
        layout = formatted_data.get("layout", "bullets")
        title = formatted_data.get("title", "")
        sections = formatted_data.get("sections", [])

        if layout == "table":
            headers = []
            rows = []
            for sec in sections:
                if sec["type"] == "table":
                    headers = sec["headers"]
                    rows = sec["rows"]
                    break
            self.add_table_slide(title, headers, rows, image_path, slide_label)

        elif layout == "key_value":
            pairs = []
            for sec in sections:
                if sec["type"] == "key_value":
                    pairs = sec["pairs"]
                    break
            self.add_key_value_slide(title, pairs, image_path, slide_label)

        elif layout == "two_column":
            left_pts, right_pts = [], []
            for sec in sections:
                if sec.get("side") == "left":
                    left_pts = sec.get("points", [])
                elif sec.get("side") == "right":
                    right_pts = sec.get("points", [])
            self.add_two_column_slide(title, left_pts, right_pts, image_path, slide_label)

        elif layout == "highlight":
            main, support = "", []
            for sec in sections:
                if sec["type"] == "highlight":
                    main = sec["main"]
                elif sec["type"] == "support":
                    support = sec.get("points", [])
            self.add_highlight_slide(title, main, support, image_path, slide_label)

        else:
            all_points = []
            for sec in sections:
                if sec.get("type") == "bullets":
                    all_points.extend(sec.get("points", []))
                elif sec.get("type") == "text":
                    all_points.append(sec.get("content", ""))
            self.add_bullets_slide(title, all_points, image_path, slide_label)

    # ─────────────────────────────────────────
    # 엔딩 슬라이드
    # ─────────────────────────────────────────
    def add_ending_slide(self, text="Thank You", subtext=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._set_slide_bg(slide)

        txBox = slide.shapes.add_textbox(Inches(2), Inches(2.5), Inches(9), Inches(1.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.colors["accent"])
        p.font.name = get_font_name(text)
        p.alignment = PP_ALIGN.CENTER

        if subtext:
            txBox2 = slide.shapes.add_textbox(Inches(2), Inches(4.2), Inches(9), Inches(1))
            tf2 = txBox2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = subtext
            p2.font.size = Pt(20)
            p2.font.color.rgb = hex_to_rgb(self.colors.get("subtitle", "#999999"))
            p2.font.name = get_font_name(subtext)
            p2.alignment = PP_ALIGN.CENTER

        self._add_page_number(slide)

    # ─────────────────────────────────────────
    # 저장
    # ─────────────────────────────────────────
    def save(self, filepath):
        directory = os.path.dirname(filepath)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        self.prs.save(filepath)
        return filepath
