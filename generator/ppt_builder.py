"""
ppt_builder.py
- IR Book 레퍼런스 기반 디자인
- 표지/INDEX/엔딩: 배경 이미지 풀스크린 + 좌측 타이틀
- 본문: 좌상단 챕터번호(accent) + 타이틀(무채색)
- 폰트: S-Core Dream 7 / Montserrat ExtraBold (타이틀), S-Core Dream 5 / Montserrat Medium (본문)
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


def get_title_font(text):
    return 'S-Core Dream 7' if is_korean(text) else 'Montserrat ExtraBold'


def get_body_font(text):
    return 'S-Core Dream 5' if is_korean(text) else 'Montserrat Medium'


def set_title_style(para, text, size, color):
    para.text = text
    para.font.size = size
    para.font.color.rgb = color
    para.font.name = get_title_font(text)
    para.font.bold = False
    para.font.italic = False


def set_body_style(para, text, size, color):
    para.text = text
    para.font.size = size
    para.font.color.rgb = color
    para.font.name = get_body_font(text)
    para.font.bold = False
    para.font.italic = False


class PPTBuilder:
    def __init__(self, theme="ir_pro", custom_colors=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.slide_count = 0
        self.chapter_count = 0
        self.theme = theme

        themes = {
            "ir_pro": {
                "bg": "#FFFFFF", "title": "#222222", "text": "#444444",
                "accent": "#1B2A4A", "accent2": "#E8EDF3", "card_bg": "#F4F6F9",
                "subtitle": "#888888", "label_bg": "#1B2A4A",
                "chapter_num": "#1B2A4A", "chapter_title": "#222222",
                "table_header_bg": "#1B2A4A", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#F4F6F9", "table_row_bg2": "#FFFFFF",
                "table_border": "#D0D8E0",
                "cover_overlay": "#1B2A4A", "divider_line": "#1B2A4A",
                "page_num": "#AAAAAA"
            },
            "dark_modern": {
                "bg": "#1A1A2E", "title": "#FFFFFF", "text": "#E0E0E0",
                "accent": "#E94560", "accent2": "#0F3460", "card_bg": "#16213E",
                "subtitle": "#B0B0B0", "label_bg": "#E94560",
                "chapter_num": "#E94560", "chapter_title": "#FFFFFF",
                "table_header_bg": "#E94560", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#16213E", "table_row_bg2": "#1A1A2E",
                "table_border": "#0F3460",
                "cover_overlay": "#1A1A2E", "divider_line": "#E94560",
                "page_num": "#666666"
            },
            "light_clean": {
                "bg": "#FFFFFF", "title": "#2D2D2D", "text": "#4A4A4A",
                "accent": "#4A90D9", "accent2": "#F0F4F8", "card_bg": "#F7F9FC",
                "subtitle": "#7A7A7A", "label_bg": "#4A90D9",
                "chapter_num": "#4A90D9", "chapter_title": "#2D2D2D",
                "table_header_bg": "#4A90D9", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#F7F9FC", "table_row_bg2": "#FFFFFF",
                "table_border": "#D0D8E0",
                "cover_overlay": "#2D2D2D", "divider_line": "#4A90D9",
                "page_num": "#AAAAAA"
            },
            "corporate_blue": {
                "bg": "#F0F4F8", "title": "#1A365D", "text": "#2D4A7A",
                "accent": "#3182CE", "accent2": "#EBF4FF", "card_bg": "#FFFFFF",
                "subtitle": "#4A6FA5", "label_bg": "#3182CE",
                "chapter_num": "#3182CE", "chapter_title": "#1A365D",
                "table_header_bg": "#3182CE", "table_header_text": "#FFFFFF",
                "table_row_bg1": "#EBF4FF", "table_row_bg2": "#F0F4F8",
                "table_border": "#B0C4DE",
                "cover_overlay": "#1A365D", "divider_line": "#3182CE",
                "page_num": "#AAAAAA"
            },
        }

        if custom_colors and isinstance(custom_colors, dict):
            self.colors = custom_colors
        else:
            self.colors = themes.get(theme, themes["ir_pro"])

    def _set_slide_bg(self, slide, color=None):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(color or self.colors["bg"])

    def _add_page_number(self, slide):
        self.slide_count += 1
        txBox = slide.shapes.add_textbox(Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.3))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        set_body_style(p, str(self.slide_count), Pt(10), hex_to_rgb(self.colors.get("page_num", "#AAAAAA")))
        p.alignment = PP_ALIGN.RIGHT

    # ═══════════════════════════════════════
    # 표지 슬라이드 (배경이미지 + 좌측 타이틀)
    # ═══════════════════════════════════════
    def add_title_slide(self, title, subtitle="", bg_image_path=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_image_path and os.path.exists(bg_image_path):
            slide.shapes.add_picture(
                bg_image_path, Inches(0), Inches(0),
                self.prs.slide_width, self.prs.slide_height
            )
            # 좌측 반투명 오버레이
            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                Inches(6.5), self.prs.slide_height
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = hex_to_rgb(self.colors["cover_overlay"])
            from pptx.oxml.ns import qn
            overlay.fill._fill.attrib[qn('a:blipFill')] if False else None
            overlay.line.fill.background()
        else:
            self._set_slide_bg(slide, self.colors["cover_overlay"])

        # 상단 장식선
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.8), Inches(3.5), Inches(0.05)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = hex_to_rgb(self.colors.get("divider_line", "#FFFFFF"))
        line.line.fill.background()

        # 타이틀 (좌측)
        txBox = slide.shapes.add_textbox(Inches(1.2), Inches(3.1), Inches(5.0), Inches(1.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        set_title_style(p, title, Pt(40), RGBColor(255, 255, 255))

        # 부제 (좌측)
        if subtitle:
            txBox2 = slide.shapes.add_textbox(Inches(1.2), Inches(4.9), Inches(5.0), Inches(0.8))
            tf2 = txBox2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            set_body_style(p2, subtitle, Pt(18), RGBColor(200, 210, 225))

        # "INVESTOR RELATIONS" 라벨
        lbl = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(5), Inches(0.5))
        tf_l = lbl.text_frame
        p_l = tf_l.paragraphs[0]
        set_body_style(p_l, "INVESTOR RELATIONS", Pt(14), RGBColor(180, 190, 210))

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # INDEX(목차) 슬라이드
    # ═══════════════════════════════════════
    def add_index_slide(self, chapters, bg_image_path=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_image_path and os.path.exists(bg_image_path):
            slide.shapes.add_picture(
                bg_image_path, Inches(0), Inches(0),
                self.prs.slide_width, self.prs.slide_height
            )
            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                Inches(6.0), self.prs.slide_height
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = hex_to_rgb(self.colors["cover_overlay"])
            overlay.line.fill.background()
        else:
            self._set_slide_bg(slide, self.colors["cover_overlay"])

        # "TABLE OF CONTENTS"
        txBox = slide.shapes.add_textbox(Inches(1.2), Inches(1.0), Inches(4.5), Inches(0.6))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        set_title_style(p, "TABLE OF CONTENTS", Pt(16), RGBColor(180, 190, 210))

        # 목차 항목
        start_top = Inches(2.0)
        for idx, ch_title in enumerate(chapters):
            num = f"{idx + 1:02d}"

            # 번호
            tb_num = slide.shapes.add_textbox(Inches(1.2), start_top + Inches(idx * 0.7), Inches(0.6), Inches(0.5))
            tf_n = tb_num.text_frame
            p_n = tf_n.paragraphs[0]
            set_title_style(p_n, num, Pt(20), hex_to_rgb(self.colors.get("divider_line", "#FFFFFF")))

            # 제목
            tb_title = slide.shapes.add_textbox(Inches(2.0), start_top + Inches(idx * 0.7), Inches(3.5), Inches(0.5))
            tf_t = tb_title.text_frame
            tf_t.word_wrap = True
            p_t = tf_t.paragraphs[0]
            set_body_style(p_t, ch_title, Pt(16), RGBColor(255, 255, 255))

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # 본문 - 챕터 헤더 구조
    # ═══════════════════════════════════════
    def _add_chapter_header(self, slide, title, slide_label=""):
        """좌상단 챕터번호(accent) + 타이틀(무채색)"""
        self._set_slide_bg(slide)

        # 챕터 번호
        self.chapter_count += 1
        num_text = slide_label if slide_label else f"{self.chapter_count:02d}"

        tb_num = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(1.0), Inches(0.5))
        tf_n = tb_num.text_frame
        p_n = tf_n.paragraphs[0]
        set_title_style(p_n, num_text, Pt(14), hex_to_rgb(self.colors["chapter_num"]))

        # 구분선
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.9), Inches(0.5), Inches(0.03)
        )
        line.fill.solid()
        line.fill.fore_color.rgb = hex_to_rgb(self.colors["chapter_num"])
        line.line.fill.background()

        # 타이틀
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.05), Inches(11.5), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        set_title_style(p, title, Pt(26), hex_to_rgb(self.colors["chapter_title"]))

        # 하단 라인
        bottom_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.95), Inches(11.733), Inches(0.015)
        )
        bottom_line.fill.solid()
        bottom_line.fill.fore_color.rgb = hex_to_rgb(self.colors.get("divider_line", "#DDDDDD"))
        bottom_line.line.fill.background()

    # ═══════════════════════════════════════
    # Bullets 레이아웃
    # ═══════════════════════════════════════
    def add_bullets_slide(self, title, points, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._add_chapter_header(slide, title, slide_label)

        content_width = Inches(7.0) if image_path else Inches(11.5)
        txBox = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), content_width, Inches(4.5))
        tf = txBox.text_frame
        tf.word_wrap = True

        num_points = len(points) if isinstance(points, list) else 1
        if num_points <= 3:
            font_size, spacing = Pt(17), Pt(14)
        elif num_points <= 5:
            font_size, spacing = Pt(15), Pt(10)
        elif num_points <= 8:
            font_size, spacing = Pt(13), Pt(8)
        else:
            font_size, spacing = Pt(11), Pt(6)

        if isinstance(points, list):
            for i, point in enumerate(points):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                set_body_style(p, f"•  {point}", font_size, hex_to_rgb(self.colors["text"]))
                p.space_after = spacing
        else:
            p = tf.paragraphs[0]
            set_body_style(p, str(points), font_size, hex_to_rgb(self.colors["text"]))

        if image_path and os.path.exists(image_path):
            try:
                slide.shapes.add_picture(image_path, Inches(8.5), Inches(2.0), Inches(4.3), Inches(4.3))
            except Exception:
                pass

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # Two Column 레이아웃
    # ═══════════════════════════════════════
    def add_two_column_slide(self, title, left_points, right_points, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._add_chapter_header(slide, title, slide_label)

        col_width = Inches(5.5)
        col_top = Inches(2.1)

        for pts, x_offset in [(left_points, Inches(0.8)), (right_points, Inches(0.8) + col_width + Inches(0.3))]:
            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_offset, col_top, col_width, Inches(4.3))
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.colors.get("card_bg", "#F4F6F9"))
            card.line.color.rgb = hex_to_rgb(self.colors.get("accent2", "#E8EDF3"))
            card.line.width = Pt(1)

            txBox = slide.shapes.add_textbox(x_offset + Inches(0.3), col_top + Inches(0.3), col_width - Inches(0.6), Inches(3.7))
            tf = txBox.text_frame
            tf.word_wrap = True
            fsize = Pt(13) if len(pts) <= 5 else Pt(11)
            for i, pt in enumerate(pts):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                set_body_style(p, f"•  {pt}", fsize, hex_to_rgb(self.colors["text"]))
                p.space_after = Pt(7)

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # Key-Value 레이아웃
    # ═══════════════════════════════════════
    def add_key_value_slide(self, title, pairs, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._add_chapter_header(slide, title, slide_label)

        max_cols = 3
        card_area = Inches(12.0)
        card_w = (card_area - Inches(0.3) * (max_cols - 1)) / max_cols
        card_h = Inches(2.0)

        for idx, pair in enumerate(pairs[:6]):
            col = idx % max_cols
            row = idx // max_cols
            c_left = Inches(0.6) + col * (card_w + Inches(0.3))
            c_top = Inches(2.2) + row * (card_h + Inches(0.3))

            card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(c_left), c_top, int(card_w), card_h)
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.colors.get("card_bg", "#F4F6F9"))
            card.line.color.rgb = hex_to_rgb(self.colors.get("accent2", "#E8EDF3"))
            card.line.width = Pt(1)

            # Key
            tb_key = slide.shapes.add_textbox(int(c_left) + Inches(0.2), c_top + Inches(0.2), int(card_w) - Inches(0.4), Inches(0.5))
            pk = tb_key.text_frame.paragraphs[0]
            set_title_style(pk, pair.get("key", ""), Pt(15), hex_to_rgb(self.colors["chapter_num"]))

            # Value
            tb_val = slide.shapes.add_textbox(int(c_left) + Inches(0.2), c_top + Inches(0.8), int(card_w) - Inches(0.4), Inches(1.0))
            tb_val.text_frame.word_wrap = True
            pv = tb_val.text_frame.paragraphs[0]
            set_body_style(pv, pair.get("value", ""), Pt(12), hex_to_rgb(self.colors["text"]))

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # Highlight 레이아웃
    # ═══════════════════════════════════════
    def add_highlight_slide(self, title, main_text, support_points=None, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._add_chapter_header(slide, title, slide_label)

        hl_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.2), Inches(11.5), Inches(2.0)
        )
        hl_box.fill.solid()
        hl_box.fill.fore_color.rgb = hex_to_rgb(self.colors["accent"])
        hl_box.line.fill.background()

        tb_hl = slide.shapes.add_textbox(Inches(1.2), Inches(2.5), Inches(10.7), Inches(1.5))
        tb_hl.text_frame.word_wrap = True
        p_hl = tb_hl.text_frame.paragraphs[0]
        set_title_style(p_hl, main_text, Pt(20), RGBColor(255, 255, 255))
        p_hl.alignment = PP_ALIGN.CENTER

        if support_points:
            tb_sp = slide.shapes.add_textbox(Inches(0.8), Inches(4.5), Inches(11.5), Inches(2.0))
            tb_sp.text_frame.word_wrap = True
            for i, sp in enumerate(support_points):
                p_sp = tb_sp.text_frame.paragraphs[0] if i == 0 else tb_sp.text_frame.add_paragraph()
                set_body_style(p_sp, f"•  {sp}", Pt(14), hex_to_rgb(self.colors["text"]))
                p_sp.space_after = Pt(8)

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # 표(Table) 레이아웃
    # ═══════════════════════════════════════
    def add_table_slide(self, title, headers, rows, image_path=None, slide_label=""):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self._add_chapter_header(slide, title, slide_label)

        num_rows = len(rows) + 1
        num_cols = len(headers)
        table_width = Inches(12.0)
        table_height = min(Inches(4.5), Inches(0.5) * num_rows)

        table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.6), Inches(2.1), table_width, table_height)
        table = table_shape.table

        first_col_w = Inches(1.8)
        other_col_w = (table_width - first_col_w) / (num_cols - 1) if num_cols > 1 else table_width
        table.columns[0].width = int(first_col_w)
        for c in range(1, num_cols):
            table.columns[c].width = int(other_col_w)

        if num_rows <= 5 and num_cols <= 4:
            h_fs, c_fs = Pt(13), Pt(12)
        elif num_rows <= 8 and num_cols <= 6:
            h_fs, c_fs = Pt(11), Pt(10)
        else:
            h_fs, c_fs = Pt(9), Pt(8)

        header_bg = hex_to_rgb(self.colors["table_header_bg"])
        header_tc = hex_to_rgb(self.colors["table_header_text"])
        row_bg1 = hex_to_rgb(self.colors["table_row_bg1"])
        row_bg2 = hex_to_rgb(self.colors["table_row_bg2"])
        text_c = hex_to_rgb(self.colors["text"])

        for ci, ht in enumerate(headers):
            cell = table.cell(0, ci)
            cell.text = ""
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_bg
            para = cell.text_frame.paragraphs[0]
            set_title_style(para, ht, h_fs, header_tc)
            para.alignment = PP_ALIGN.CENTER
            cell.text_frame.word_wrap = True
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        for ri, rd in enumerate(rows):
            bg = row_bg1 if ri % 2 == 0 else row_bg2
            for ci in range(num_cols):
                cell = table.cell(ri + 1, ci)
                ct = rd[ci] if ci < len(rd) else ""
                cell.text = ""
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
                para = cell.text_frame.paragraphs[0]
                set_body_style(para, ct, c_fs, text_c)
                para.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                cell.text_frame.word_wrap = True
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        self._set_table_borders(table, num_rows, num_cols, hex_to_rgb(self.colors["table_border"]))
        self._add_page_number(slide)

    def _set_table_borders(self, table, num_rows, num_cols, border_color):
        from pptx.oxml.ns import qn
        bc = str(border_color).replace('#', '')
        for ri in range(num_rows):
            for ci in range(num_cols):
                tc = table.cell(ri, ci)._tc
                tcPr = tc.get_or_add_tcPr()
                for bn in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
                    ln = tcPr.find(qn(bn))
                    if ln is None:
                        ln = tcPr.makeelement(qn(bn), {})
                        tcPr.append(ln)
                    ln.set('w', '12700')
                    sf = ln.find(qn('a:solidFill'))
                    if sf is None:
                        sf = ln.makeelement(qn('a:solidFill'), {})
                        ln.insert(0, sf)
                    sc = sf.find(qn('a:srgbClr'))
                    if sc is None:
                        sc = sf.makeelement(qn('a:srgbClr'), {})
                        sf.append(sc)
                    sc.set('val', bc)

    # ═══════════════════════════════════════
    # 통합 렌더
    # ═══════════════════════════════════════
    def add_formatted_slide(self, formatted_data, image_path=None, slide_label=""):
        layout = formatted_data.get("layout", "bullets")
        title = formatted_data.get("title", "")
        sections = formatted_data.get("sections", [])

        if layout == "table":
            for sec in sections:
                if sec["type"] == "table":
                    self.add_table_slide(title, sec["headers"], sec["rows"], image_path, slide_label)
                    return
        elif layout == "key_value":
            for sec in sections:
                if sec["type"] == "key_value":
                    self.add_key_value_slide(title, sec["pairs"], image_path, slide_label)
                    return
        elif layout == "two_column":
            lp, rp = [], []
            for sec in sections:
                if sec.get("side") == "left": lp = sec.get("points", [])
                elif sec.get("side") == "right": rp = sec.get("points", [])
            self.add_two_column_slide(title, lp, rp, image_path, slide_label)
            return
        elif layout == "highlight":
            main, sup = "", []
            for sec in sections:
                if sec["type"] == "highlight": main = sec["main"]
                elif sec["type"] == "support": sup = sec.get("points", [])
            self.add_highlight_slide(title, main, sup, image_path, slide_label)
            return

        pts = []
        for sec in sections:
            if sec.get("type") == "bullets": pts.extend(sec.get("points", []))
            elif sec.get("type") == "text": pts.append(sec.get("content", ""))
        self.add_bullets_slide(title, pts, image_path, slide_label)

    # ═══════════════════════════════════════
    # 엔딩 슬라이드 (배경이미지 + 중앙 텍스트)
    # ═══════════════════════════════════════
    def add_ending_slide(self, text="Thank You", subtext="", bg_image_path=None):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_image_path and os.path.exists(bg_image_path):
            slide.shapes.add_picture(
                bg_image_path, Inches(0), Inches(0),
                self.prs.slide_width, self.prs.slide_height
            )
            overlay = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                self.prs.slide_width, self.prs.slide_height
            )
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = hex_to_rgb(self.colors["cover_overlay"])
            overlay.line.fill.background()
        else:
            self._set_slide_bg(slide, self.colors["cover_overlay"])

        txBox = slide.shapes.add_textbox(Inches(2), Inches(2.8), Inches(9), Inches(1.2))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        set_title_style(p, text, Pt(48), RGBColor(255, 255, 255))
        p.alignment = PP_ALIGN.CENTER

        if subtext:
            txBox2 = slide.shapes.add_textbox(Inches(2), Inches(4.2), Inches(9), Inches(0.8))
            tf2 = txBox2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            set_body_style(p2, subtext, Pt(18), RGBColor(200, 210, 225))
            p2.alignment = PP_ALIGN.CENTER

        self._add_page_number(slide)

    # ═══════════════════════════════════════
    # 저장
    # ═══════════════════════════════════════
    def save(self, filepath):
        directory = os.path.dirname(filepath)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        self.prs.save(filepath)
        return filepath
