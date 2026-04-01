"""
ETRI 디자인 PDF 생성기
ReportLab Platypus 기반
"""
from __future__ import annotations
import io
import re
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
pt = 1.0  # 1pt = 1 unit in ReportLab
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Table, TableStyle, Image as RLImage,
    PageBreak, KeepTogether, HRFlowable,
    Flowable,
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.pdfgen import canvas as pdfgen_canvas
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

from . import styles as S
from .docx_parser import (
    Document, Paragraph as DocPara, Table as DocTable,
    Image as DocImage, TableRow, TableCell,
)

EMU_PER_PT = 914400 / 72.0   # 1pt = 12700 EMU  →  914400 / 72


# ── 커스텀 Flowable ────────────────────────────────────────────────

class ChapterMarker(Flowable):
    """챕터명을 페이지 헤더에 전달하는 제로 높이 마커"""
    _CURRENT = ['']

    def __init__(self, chapter_text: str):
        super().__init__()
        self._chapter = chapter_text
        self.width = 0
        self.height = 0

    def wrap(self, aw, ah):
        return 0, 0

    def draw(self):
        ChapterMarker._CURRENT[0] = self._chapter


class PartDividerPage(Flowable):
    """PART N 구분 페이지 (전체 파란 배경)"""

    def __init__(self, part_text: str, chapter_lines: list[str]):
        super().__init__()
        self.part_text = part_text
        self.chapter_lines = chapter_lines
        self.width  = S.PAGE_W
        self.height = S.PAGE_H

    def wrap(self, available_w, available_h):
        return available_w, available_h

    def draw(self):
        c = self.canv
        # 배경
        c.setFillColor(S.BLUE_DIVIDER)
        c.rect(0, 0, S.PAGE_W, S.PAGE_H, fill=1, stroke=0)

        # PART 배지
        badge_w, badge_h = 80 * mm, 9 * mm
        bx = S.PAGE_W / 2 - badge_w / 2
        by = S.PAGE_H * 0.72
        c.setStrokeColor(S.WHITE)
        c.setFillColor(S.BLUE_DIVIDER)
        c.roundRect(bx, by, badge_w, badge_h, radius=4*mm, fill=1, stroke=1)
        c.setFillColor(S.WHITE)
        c.setFont(S.FONT_BOLD, 11)
        c.drawCentredString(S.PAGE_W / 2, by + 2.5 * mm, self.part_text.upper())

        # 구분선
        line_y = by - 8 * mm
        c.setStrokeColor(S.WHITE)
        c.setLineWidth(0.5)
        c.line(S.MARGIN_LEFT, line_y, S.PAGE_W - S.MARGIN_RIGHT, line_y)

        # 챕터 목록
        c.setFont(S.FONT_REGULAR, 10)
        text_y = line_y - 12 * mm
        for line in self.chapter_lines:
            c.drawCentredString(S.PAGE_W / 2, text_y, line)
            text_y -= 9 * mm


class ChapterHeader(Flowable):
    """H1 챕터 헤더 바 (파란 배경에 챕터 번호 + 제목)"""

    def __init__(self, number: str, title: str, available_w: float):
        super().__init__()
        self._number = number
        self._title  = title
        self.width   = available_w
        self.height  = 14 * mm

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        w, h = self.width, self.height

        # 파란 배경 바
        c.setFillColor(S.BLUE_LIGHT)
        c.rect(0, 0, w, h, fill=1, stroke=0)

        # 챕터 번호 (왼쪽, 굵게)
        c.setFillColor(S.WHITE)
        c.setFont(S.FONT_BOLD, S.FS_H1)
        c.drawString(5 * mm, h / 2 - S.FS_H1 * 0.35, self._number)

        # 제목
        c.setFont(S.FONT_BOLD, S.FS_H1)
        num_w = c.stringWidth(self._number + '  ', S.FONT_BOLD, S.FS_H1)
        c.drawString(5 * mm + num_w, h / 2 - S.FS_H1 * 0.35, self._title)


class CalloutBox(Flowable):
    """콜아웃 박스 (왼쪽 파란 선 + 연파랑 배경)"""

    def __init__(self, text: str, available_w: float):
        super().__init__()
        self._text = text
        self.width = available_w
        self._padding = 4 * mm
        self._bar_w   = 3 * mm

    def wrap(self, aw, ah):
        style = ParagraphStyle(
            'cb', fontName=S.FONT_REGULAR, fontSize=S.FS_BODY,
            leading=S.LEADING_BODY,
        )
        p = Paragraph(self._text, style)
        inner_w = self.width - self._bar_w - self._padding * 2
        _, h = p.wrap(inner_w, ah)
        self.height = h + self._padding * 2
        self._p = p
        return self.width, self.height

    def draw(self):
        c = self.canv
        w, h = self.width, self.height
        c.setFillColor(S.BLUE_PALE)
        c.rect(0, 0, w, h, fill=1, stroke=0)
        c.setFillColor(S.BLUE_LIGHT)
        c.rect(0, 0, self._bar_w, h, fill=1, stroke=0)
        self._p.drawOn(c, self._bar_w + self._padding, self._padding)


# ── PDF 생성기 ──────────────────────────────────────────────────────

class ETRIPdfGenerator:

    def __init__(self, output_path: str, doc_meta: dict | None = None):
        self.output_path = output_path
        self.meta = doc_meta or {}
        self._setup_styles()
        self._toc_entries: list[tuple[int, str, int]] = []  # (level, text, pagenum)

    def _setup_styles(self):
        self.para_style = ParagraphStyle(
            'Body',
            fontName=S.FONT_REGULAR,
            fontSize=S.FS_BODY,
            leading=S.LEADING_BODY,
            alignment=TA_JUSTIFY,
            textColor=S.GRAY_DARK,
            spaceAfter=3,
        )
        self.h2_style = ParagraphStyle(
            'H2',
            fontName=S.FONT_BOLD,
            fontSize=S.FS_H2,
            leading=S.LEADING_H2,
            textColor=S.BLUE_PRIMARY,
            spaceBefore=8,
            spaceAfter=4,
        )
        self.h3_style = ParagraphStyle(
            'H3',
            fontName=S.FONT_BOLD,
            fontSize=S.FS_H3,
            leading=S.LEADING_H3,
            textColor=S.BLUE_PRIMARY,
            spaceBefore=6,
            spaceAfter=3,
        )
        self.h4_style = ParagraphStyle(
            'H4',
            fontName=S.FONT_BOLD,
            fontSize=S.FS_BODY,
            leading=S.LEADING_BODY,
            textColor=S.GRAY_DARK,
            spaceBefore=4,
            spaceAfter=2,
        )
        self.caption_style = ParagraphStyle(
            'Caption',
            fontName=S.FONT_REGULAR,
            fontSize=S.FS_SMALL,
            leading=12,
            textColor=S.GRAY_MID,
            alignment=TA_CENTER,
            spaceBefore=2,
            spaceAfter=4,
        )
        self.bullet_style = ParagraphStyle(
            'Bullet',
            fontName=S.FONT_REGULAR,
            fontSize=S.FS_BODY,
            leading=S.LEADING_BODY,
            textColor=S.GRAY_DARK,
            leftIndent=6 * mm,
            bulletIndent=2 * mm,
            spaceAfter=2,
        )
        self.toc1_style = ParagraphStyle(
            'TOC1',
            fontName=S.FONT_BOLD,
            fontSize=S.FS_TOC1,
            leading=15,
            textColor=S.GRAY_DARK,
            spaceBefore=5,
        )
        self.toc2_style = ParagraphStyle(
            'TOC2',
            fontName=S.FONT_REGULAR,
            fontSize=S.FS_TOC2,
            leading=14,
            textColor=S.GRAY_DARK,
            leftIndent=8 * mm,
            spaceBefore=2,
        )
        self.toc3_style = ParagraphStyle(
            'TOC3',
            fontName=S.FONT_REGULAR,
            fontSize=S.FS_TOC3,
            leading=13,
            textColor=S.GRAY_MID,
            leftIndent=16 * mm,
        )

    # ── 페이지 콜백 ─────────────────────────────────────────────────

    def _on_page(self, canv: pdfgen_canvas.Canvas, doc):
        """본문 페이지 헤더/푸터"""
        page_num = doc.page
        if page_num <= 2:
            return

        canv.saveState()

        # ─ 헤더 ─
        header_y = S.PAGE_H - S.MARGIN_TOP + 4 * mm
        chapter = ChapterMarker._CURRENT[0]
        if chapter:
            canv.setFont(S.FONT_REGULAR, 8)
            canv.setFillColor(S.GRAY_MID)
            canv.drawRightString(
                S.PAGE_W - S.MARGIN_RIGHT, header_y, chapter
            )
        canv.setStrokeColor(S.GRAY_LIGHT)
        canv.setLineWidth(0.5)
        canv.line(
            S.MARGIN_LEFT, header_y - 1.5 * mm,
            S.PAGE_W - S.MARGIN_RIGHT, header_y - 1.5 * mm,
        )

        # ─ 푸터 ─
        footer_y = S.MARGIN_BOTTOM - 6 * mm
        canv.setStrokeColor(S.GRAY_LIGHT)
        canv.line(
            S.MARGIN_LEFT, footer_y + 4 * mm,
            S.PAGE_W - S.MARGIN_RIGHT, footer_y + 4 * mm,
        )
        canv.setFont(S.FONT_REGULAR, 8)
        canv.setFillColor(S.GRAY_MID)
        canv.drawCentredString(S.PAGE_W / 2, footer_y, str(page_num - 2))

        canv.restoreState()

    def _on_cover_page(self, canv: pdfgen_canvas.Canvas, doc):
        """표지 페이지 (헤더/푸터 없음)"""
        pass

    # ── 표지 생성 ───────────────────────────────────────────────────

    def _build_cover(self) -> list:
        """표지 페이지 flowables"""
        flowables = []

        # 표지 그리기 (캔버스 직접 조작)
        class CoverPage(Flowable):
            def __init__(self_, meta):
                super().__init__()
                self_.meta = meta
                self_.width  = S.PAGE_W
                self_.height = S.PAGE_H

            def wrap(self_, aw, ah):
                return S.PAGE_W, S.PAGE_H

            def draw(self_):
                c = self_.canv
                w, h = S.PAGE_W, S.PAGE_H

                # 흰 배경
                c.setFillColor(S.WHITE)
                c.rect(0, 0, w, h, fill=1, stroke=0)

                # 상단 파란 장식 바
                c.setFillColor(S.BLUE_PRIMARY)
                c.rect(0, h - 18 * mm, w, 18 * mm, fill=1, stroke=0)

                # ETRI 로고 텍스트
                c.setFillColor(S.WHITE)
                c.setFont(S.FONT_BOLD, 22)
                c.drawString(S.MARGIN_LEFT, h - 13 * mm, 'ETRI')

                # 연도
                year = self_.meta.get('year', '2025')
                c.setFont(S.FONT_REGULAR, 10)
                c.setFillColor(S.WHITE)
                c.drawRightString(w - S.MARGIN_RIGHT, h - 12 * mm, year)

                # 제목 영역 (중앙)
                title     = self_.meta.get('title', '표준체계 및 선도전략')
                subtitle  = self_.meta.get('subtitle', '')
                year_big  = self_.meta.get('year', '2025')

                # 큰 제목
                center_y = h * 0.58
                c.setFont(S.FONT_BOLD, 28)
                c.setFillColor(S.BLUE_PRIMARY)
                c.drawCentredString(w / 2, center_y, title)

                if subtitle:
                    c.setFont(S.FONT_BOLD, 28)
                    c.setFillColor(S.ORANGE)
                    c.drawCentredString(w / 2, center_y - 16 * mm, subtitle)

                # 연도 (크게)
                c.setFont(S.FONT_BOLD, 32)
                c.setFillColor(S.GRAY_DARK)
                c.drawCentredString(w / 2, center_y - 35 * mm, year_big)

                # 하단 구분선
                c.setStrokeColor(S.GRAY_LIGHT)
                c.setLineWidth(0.5)
                c.line(S.MARGIN_LEFT, 35 * mm, w - S.MARGIN_RIGHT, 35 * mm)

                # 날짜
                date_str = self_.meta.get('date', '')
                c.setFont(S.FONT_REGULAR, 9)
                c.setFillColor(S.GRAY_MID)
                c.drawCentredString(w / 2, 28 * mm, date_str)

                # 기관
                org = self_.meta.get('org', 'ICT전략연구소 표준연구본부')
                c.setFont(S.FONT_BOLD, 9)
                c.setFillColor(S.GRAY_DARK)
                c.drawCentredString(w / 2, 18 * mm, org)

                # 하단 파란 띠
                c.setFillColor(S.BLUE_PRIMARY)
                c.rect(0, 0, w, 10 * mm, fill=1, stroke=0)

        flowables.append(CoverPage(self.meta))
        flowables.append(PageBreak())
        return flowables

    # ── 목차 생성 ───────────────────────────────────────────────────

    def _build_toc_header(self) -> list:
        flowables = []
        flowables.append(Spacer(1, 8 * mm))
        title_style = ParagraphStyle(
            'TOCTitle',
            fontName=S.FONT_BOLD,
            fontSize=14,
            leading=20,
            textColor=S.BLUE_PRIMARY,
            spaceBefore=0,
            spaceAfter=6,
        )
        flowables.append(Paragraph('목  차', title_style))
        flowables.append(HRFlowable(
            width=S.CONTENT_W,
            thickness=1.5,
            color=S.BLUE_PRIMARY,
            spaceAfter=6,
        ))
        return flowables

    def _build_toc_entry(self, para: DocPara) -> Flowable | None:
        text = para.text.strip()
        if not text:
            return None

        # DOCX TOC 항목 형식: "1.2.3제목텍스트123" (끝에 페이지 번호 붙음)
        # 끝의 숫자를 페이지 번호로 분리
        m_page = re.match(r'^(.*?)(\d+)$', text)
        if m_page:
            body_part = m_page.group(1).strip()
            page_num  = m_page.group(2)
        else:
            body_part = text
            page_num  = ''

        # 번호와 제목 분리 (예: "1.2.3제목" → "1.2.3  제목")
        m_num = re.match(r'^((?:PART\s+\w+|[IVX]+|\d+(?:\.\d+)*))\s*(.*)', body_part)
        if m_num:
            num   = m_num.group(1).strip()
            title = m_num.group(2).strip()
        else:
            num   = ''
            title = body_part

        # 표시 텍스트 구성
        num_part   = f'{num}  ' if num else ''
        page_part  = f'  {page_num}' if page_num else ''
        # XML 이스케이프
        def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        display = f'{esc(num_part)}{esc(title)}{esc(page_part)}'

        style_map = {
            'TOC1': self.toc1_style,
            'TOC2': self.toc2_style,
            'TOC3': self.toc3_style,
            'TOC4': self.toc3_style,
        }
        style = style_map.get(para.style, self.toc2_style)
        return Paragraph(display, style)

    # ── 본문 flowable 변환 ─────────────────────────────────────────

    def _runs_to_markup(self, para: DocPara) -> str:
        """DocPara runs → ReportLab XML 마크업"""
        parts = []
        for run in para.runs:
            text = run.text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            tags_open = tags_close = ''
            if run.bold:
                tags_open  += f'<b>'
                tags_close = '</b>' + tags_close
            if run.italic:
                tags_open  += '<i>'
                tags_close = '</i>' + tags_close
            if run.color:
                tags_open  += f'<font color="#{run.color}">'
                tags_close = '</font>' + tags_close
            parts.append(f'{tags_open}{text}{tags_close}')
        return ''.join(parts)

    def _convert_paragraph(self, para: DocPara) -> list[Flowable]:
        flowables = []
        text = para.text.strip()
        markup = self._runs_to_markup(para)

        if not markup:
            return []

        style = para.style

        # ── Heading 1 ─
        if style == 'Heading1':
            m = re.match(r'^(\d+)\s*(.*)', text)
            if m:
                num, title = m.group(1), m.group(2).strip()
                chapter_label = f'{num}. {title}'
            else:
                num, title = '', text
                chapter_label = text

            flowables.append(ChapterMarker(chapter_label))
            flowables.append(Spacer(1, 6 * mm))
            flowables.append(ChapterHeader(
                number=num, title=title,
                available_w=S.CONTENT_W,
            ))
            flowables.append(Spacer(1, 4 * mm))
            return flowables

        # ── Heading 2 ─
        if style == 'Heading2':
            m = re.match(r'^([\d.]+)\s*(.*)', text)
            if m:
                num_part  = m.group(1)
                title_part = m.group(2).strip()
                display = f'<font color="#{S.BLUE_PRIMARY.hexval()[2:]}"><b>{num_part}</b></font>  {title_part}'
            else:
                display = f'<b>{markup}</b>'
            flowables.append(Spacer(1, 3 * mm))
            flowables.append(Paragraph(display, self.h2_style))
            return flowables

        # ── Heading 3 ─
        if style == 'Heading3':
            flowables.append(Spacer(1, 2 * mm))
            flowables.append(Paragraph(f'<b>{markup}</b>', self.h3_style))
            return flowables

        # ── Heading 4+ ─
        if style in ('Heading4', 'Heading5'):
            flowables.append(Paragraph(f'<b>{markup}</b>', self.h4_style))
            return flowables

        # ── PART 구분 ─
        if style == 'PART':
            return []  # PartDividerPage 별도 처리

        # ── 캡션 ─
        if style == 'Caption':
            flowables.append(Paragraph(markup, self.caption_style))
            return flowables

        # ── 콜아웃 ─
        if style == 'Callout':
            flowables.append(Spacer(1, 2 * mm))
            flowables.append(CalloutBox(markup, S.CONTENT_W))
            flowables.append(Spacer(1, 2 * mm))
            return flowables

        # ── 불릿/번호 ─
        if para.num_id is not None or style == 'Bullet':
            bullet = '•'
            indent = (para.num_level + 1) * 5 * mm
            style_obj = ParagraphStyle(
                f'Bullet{para.num_level}',
                parent=self.bullet_style,
                leftIndent=indent,
                bulletIndent=indent - 4 * mm,
            )
            flowables.append(Paragraph(f'<bullet>{bullet}</bullet>{markup}', style_obj))
            return flowables

        # ── 일반 본문 ─
        al = {
            'center':  TA_CENTER,
            'right':   TA_RIGHT,
            'left':    TA_LEFT,
            'justify': TA_JUSTIFY,
        }.get(para.alignment, TA_JUSTIFY)
        style_obj = ParagraphStyle(
            'BodyLocal', parent=self.para_style, alignment=al,
        )
        flowables.append(Paragraph(markup, style_obj))
        return flowables

    @staticmethod
    def _optimal_col_widths(
        tbl: 'DocTable', n_cols: int, cell_fs: float, available_w: float
    ) -> list[float]:
        """
        텍스트 자연 너비 기반 컬럼 너비 최적화.
        줄바꿈을 최소화하면서 전체 너비 = available_w 를 채우도록 분배.
        """
        from reportlab.pdfbase import pdfmetrics

        PADDING = 7   # 좌우 패딩 합계 (pt)
        MIN_COL = 12  # 최소 컬럼 너비 (pt)

        # ── 1단계: 각 컬럼의 자연 너비(한 줄에 들어가는 너비) 수집 ──
        natural: list[float] = [0.0] * n_cols   # 최장 줄 너비
        min_word: list[float] = [0.0] * n_cols  # 최장 단어 너비 (절대 최소)

        for row in tbl.rows:
            col_cursor = 0
            for cell in row.cells:
                if col_cursor >= n_cols:
                    break
                span = min(cell.colspan, n_cols - col_cursor)
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if not text:
                        continue
                    for line in text.splitlines():
                        line = line.strip()
                        if not line:
                            continue
                        lw = pdfmetrics.stringWidth(line, S.FONT_REGULAR, cell_fs) + PADDING
                        # colspan을 고려: 넓은 셀은 마지막 컬럼에 주로 기여
                        target_col = min(col_cursor + span - 1, n_cols - 1)
                        natural[target_col] = max(natural[target_col], lw / span)
                        for word in line.split():
                            ww = pdfmetrics.stringWidth(word, S.FONT_REGULAR, cell_fs) + PADDING
                            min_word[col_cursor] = max(min_word[col_cursor], ww)
                col_cursor += span

        # ── 2단계: 최소값 보정 ──
        natural  = [max(n, mw, MIN_COL) for n, mw in zip(natural, min_word)]
        min_word = [max(mw, MIN_COL) for mw in min_word]

        total_natural = sum(natural)

        # ── 3단계: 자연 너비 합이 available_w 이하 → 그대로 사용 ──
        if total_natural <= available_w:
            # 남은 공간을 자연 너비 비율로 균등 확장
            scale = available_w / total_natural
            return [w * scale for w in natural]

        # ── 4단계: 초과 시 → 최소 너비 확보 후 남은 공간 비율 배분 ──
        total_min = sum(min_word)
        if total_min >= available_w:
            # 최소 너비만으로도 넘침 → 균등 분할
            return [available_w / n_cols] * n_cols

        remaining = available_w - total_min
        # 각 컬럼이 자연 너비에서 최소 너비를 뺀 만큼 추가 공간 필요
        extra_need = [max(0.0, n - m) for n, m in zip(natural, min_word)]
        total_extra = sum(extra_need) or 1.0

        return [
            mw + remaining * (en / total_extra)
            for mw, en in zip(min_word, extra_need)
        ]

    # ── 표 디자인 색상 상수 ───────────────────────────────────────────
    _T_HDR_BG    = colors.HexColor('#1565C0')   # 헤더 배경 (진파랑)
    _T_HDR2_BG   = colors.HexColor('#2E75B6')   # 서브헤더 배경 (중간파랑)
    _T_HDR_FG    = colors.white
    _T_ODD_BG    = colors.white                  # 홀수 데이터 행
    _T_EVEN_BG   = colors.HexColor('#EBF3FB')   # 짝수 데이터 행 (연파랑)
    _T_TOTAL_BG  = colors.HexColor('#D0E4F5')   # 합계 행
    _T_MERGE_BG  = colors.HexColor('#C5DDF5')   # rowspan 병합 셀 (첫열)
    _T_BORDER    = colors.HexColor('#BFBFBF')   # 테두리
    _T_HBORDER   = colors.HexColor('#808080')   # 헤더 아래 선

    # 합계/소계 행으로 인식할 키워드
    _TOTAL_KEYS  = {'합계', '계', '소계', '전체', 'total', 'subtotal', 'sum'}

    @staticmethod
    def _is_numeric_cell(text: str) -> bool:
        """숫자/퍼센트/날짜 위주 셀인지 판단"""
        t = text.strip().replace(',', '').replace('%', '').replace('.', '')
        t = t.replace('~', '').replace("'", '').replace('-', '').replace(' ', '')
        return bool(t) and t.isdigit()

    def _convert_table(self, tbl: DocTable) -> list[Flowable]:
        if not tbl.rows:
            return []

        import re as _re

        # ── 실제 그리드 컬럼 수 ──────────────────────────────────────
        n_cols = max(
            (sum(c.colspan for c in row.cells) for row in tbl.rows),
            default=1,
        )
        n_cols = max(n_cols, 1)

        # ── 폰트 크기: 컬럼 수에 따라 ────────────────────────────────
        if n_cols >= 8:
            cell_fs, cell_ld = 7.0, 9
        elif n_cols >= 6:
            cell_fs, cell_ld = 7.5, 10
        elif n_cols >= 4:
            cell_fs, cell_ld = 8.0, 11
        else:
            cell_fs, cell_ld = S.FS_BODY - 0.5, S.LEADING_BODY - 2

        # ── 컬럼 너비 자동 최적화 ────────────────────────────────────
        col_w = self._optimal_col_widths(tbl, n_cols, cell_fs, S.CONTENT_W)

        # ── 셀 높이 제한 ──────────────────────────────────────────────
        MAX_CELL_H       = S.CONTENT_H * 0.45
        max_lines_per_cell = max(3, int(MAX_CELL_H / cell_ld))

        # ── 기본 스타일 명령 ──────────────────────────────────────────
        ts_cmds: list = [
            # 폰트
            ('FONTNAME',      (0, 0), (-1, -1), S.FONT_REGULAR),
            ('FONTSIZE',      (0, 0), (-1, -1), cell_fs),
            ('LEADING',       (0, 0), (-1, -1), cell_ld),
            # 정렬 (기본: 왼쪽/가운데)
            ('ALIGN',         (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            # 패딩
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 6),
            # 외곽 테두리 (약간 굵게)
            ('BOX',       (0, 0), (-1, -1), 0.75, self._T_BORDER),
            # 내부 선
            ('INNERGRID', (0, 0), (-1, -1), 0.5,  self._T_BORDER),
        ]

        # ── 헤더 행 스타일 (0행) ──────────────────────────────────────
        n_header_rows = sum(1 for r in tbl.rows if r.is_header)
        if n_header_rows == 0:
            n_header_rows = 1

        for hr in range(n_header_rows):
            ts_cmds += [
                ('BACKGROUND', (0, hr), (-1, hr), self._T_HDR_BG),
                ('TEXTCOLOR',  (0, hr), (-1, hr), self._T_HDR_FG),
                ('FONTNAME',   (0, hr), (-1, hr), S.FONT_BOLD),
                ('ALIGN',      (0, hr), (-1, hr), 'CENTER'),
            ]
        # 헤더 아래 구분선
        ts_cmds.append(
            ('LINEBELOW', (0, n_header_rows - 1), (-1, n_header_rows - 1),
             1.0, self._T_HBORDER)
        )

        # ── 데이터 행 스타일 ─────────────────────────────────────────
        data_row_idx = 0  # 헤더 제외 카운터 (교대 색상용)
        for r_idx, row in enumerate(tbl.rows):
            if r_idx < n_header_rows:
                continue

            # 합계 행 감지 (첫 셀 텍스트 기준)
            first_text = ''
            if row.cells:
                first_text = ' '.join(
                    p.text.strip() for p in row.cells[0].paragraphs
                ).strip().lower()
            is_total = first_text in self._TOTAL_KEYS

            if is_total:
                row_bg = self._T_TOTAL_BG
                ts_cmds.append(('FONTNAME', (0, r_idx), (-1, r_idx), S.FONT_BOLD))
            else:
                row_bg = self._T_EVEN_BG if data_row_idx % 2 == 1 else self._T_ODD_BG

            ts_cmds.append(('BACKGROUND', (0, r_idx), (-1, r_idx), row_bg))
            data_row_idx += 1

        # ── 데이터 그리드 구성 ───────────────────────────────────────
        from reportlab.pdfbase import pdfmetrics as _pm
        import math as _math

        data: list[list] = []

        for r_idx, row in enumerate(tbl.rows):
            row_data: list = [''] * n_cols
            col_cursor = 0
            is_hdr_row = r_idx < n_header_rows

            for cell in row.cells:
                if col_cursor >= n_cols:
                    break

                # vMerge continuation → 빈 셀 (SPAN은 restart 행에서 이미 등록됨)
                # 반드시 occupied 체크보다 먼저 처리해야 이중 전진(double-advance) 버그 방지
                if cell.vmerge == 'continue':
                    col_cursor += cell.colspan
                    continue

                # 셀 텍스트 구성 및 잘라내기
                lines = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
                eff_col_w = col_w[col_cursor] if col_cursor < len(col_w) else (S.CONTENT_W / n_cols)
                eff_col_w *= cell.colspan
                eff_col_w = max(eff_col_w, 10.0)  # 0 나누기 방지

                # pdfmetrics 기반 정확한 줄 수 추정 (한국어 포함)
                wrapped_count = 0
                trimmed_lines: list[str] = []
                for ln in lines:
                    ln_w = _pm.stringWidth(ln, S.FONT_REGULAR, cell_fs)
                    wrapped = max(1, _math.ceil(ln_w / eff_col_w))
                    if wrapped_count + wrapped > max_lines_per_cell:
                        remaining = max_lines_per_cell - wrapped_count
                        if remaining > 0:
                            ratio = (remaining * eff_col_w) / max(ln_w, 1.0)
                            cutoff = max(1, int(len(ln) * ratio))
                            trimmed_lines.append(ln[:cutoff] + ('…' if len(ln) > cutoff else ''))
                        break
                    trimmed_lines.append(ln)
                    wrapped_count += wrapped

                cell_text = '<br/>'.join(
                    l.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    for l in trimmed_lines
                )

                # 정렬 결정
                plain = ' '.join(trimmed_lines)
                if is_hdr_row or cell.is_header:
                    al = TA_CENTER
                elif cell.align == 'center' or self._is_numeric_cell(plain):
                    al = TA_CENTER
                elif cell.align == 'right':
                    al = TA_RIGHT
                else:
                    al = TA_LEFT

                fn = S.FONT_BOLD if (is_hdr_row or cell.is_header) else S.FONT_REGULAR
                fg = self._T_HDR_FG if (is_hdr_row or cell.is_header) else S.GRAY_DARK

                p_style = ParagraphStyle(
                    f'TC{r_idx}_{col_cursor}',
                    fontName=fn,
                    fontSize=cell_fs,
                    leading=cell_ld,
                    textColor=fg,
                    alignment=al,
                )
                row_data[col_cursor] = Paragraph(cell_text, p_style)

                span_end_c = col_cursor + cell.colspan - 1
                # rowspan을 페이지 높이 기준으로 제한 (너무 큰 rowspan은 LayoutError 유발)
                # 실제 행 높이 ≈ cell_ld * 4 (평균 4줄) + 상하패딩 8pt 로 보수적 추정
                _est_row_h = cell_ld * 4 + 8
                _max_rs = max(2, int(S.CONTENT_H * 0.80 / max(_est_row_h, 1)))
                safe_rowspan = min(cell.rowspan, _max_rs)
                span_end_r = r_idx + safe_rowspan - 1

                # colspan SPAN
                if cell.colspan > 1:
                    ts_cmds.append(('SPAN', (col_cursor, r_idx), (span_end_c, r_idx)))
                    ts_cmds.append(('ALIGN', (col_cursor, r_idx), (span_end_c, r_idx), 'CENTER'))

                # rowspan SPAN (안전한 범위 내에서만 적용)
                if safe_rowspan > 1:
                    ts_cmds.append(('SPAN', (col_cursor, r_idx), (span_end_c, span_end_r)))
                    ts_cmds.append(('VALIGN', (col_cursor, r_idx), (span_end_c, span_end_r), 'MIDDLE'))
                    # 병합 셀 배경 (첫 컬럼 병합 셀은 특별 색상)
                    merge_bg = self._T_MERGE_BG if col_cursor == 0 else self._T_EVEN_BG
                    ts_cmds.append(('BACKGROUND', (col_cursor, r_idx), (span_end_c, span_end_r), merge_bg))

                # 셀별 배경 오버라이드 (DOCX 지정색)
                if cell.bg_color and not (is_hdr_row or cell.is_header):
                    try:
                        ts_cmds.append((
                            'BACKGROUND',
                            (col_cursor, r_idx), (span_end_c, r_idx),
                            colors.HexColor(f'#{cell.bg_color}'),
                        ))
                    except Exception:
                        pass

                col_cursor += cell.colspan

            data.append(row_data)

        # n_rows 확인 (rowspan으로 생긴 빈 행 채우기)
        while len(data) < len(tbl.rows):
            data.append([''] * n_cols)

        try:
            table = Table(
                data, colWidths=col_w,
                repeatRows=n_header_rows,
                splitByRow=1,
            )
            table.setStyle(TableStyle(ts_cmds))
            return [Spacer(1, 3 * mm), table, Spacer(1, 3 * mm)]
        except Exception:
            fallback = []
            for row in tbl.rows:
                line = ' | '.join(
                    ' '.join(p.text for p in cell.paragraphs)
                    for cell in row.cells
                )
                if line.strip():
                    esc = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    fallback.append(Paragraph(esc, self.para_style))
            return fallback

    def _convert_image(self, img: DocImage) -> list[Flowable]:
        if not img.data:
            return []
        buf = io.BytesIO(img.data)
        max_w = S.CONTENT_W
        max_h = S.CONTENT_H * 0.4

        if img.width_emu and img.height_emu:
            w_pt = img.width_emu  / EMU_PER_PT
            h_pt = img.height_emu / EMU_PER_PT
            scale = min(max_w / w_pt, max_h / h_pt, 1.0)
            w_pt *= scale
            h_pt *= scale
        else:
            w_pt = max_w
            h_pt = max_h / 2

        flowables = [
            Spacer(1, 2 * mm),
            RLImage(buf, width=w_pt, height=h_pt),
        ]
        if img.caption:
            cap_text = img.caption.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            flowables.append(Paragraph(cap_text, self.caption_style))
        flowables.append(Spacer(1, 2 * mm))
        return flowables

    # ── 메인 빌드 ──────────────────────────────────────────────────

    def build(self, doc: Document):
        """Document → PDF 파일 생성"""
        doc_template = BaseDocTemplate(
            self.output_path,
            pagesize=portrait((S.PAGE_W, S.PAGE_H)),
            leftMargin=S.MARGIN_LEFT,
            rightMargin=S.MARGIN_RIGHT,
            topMargin=S.MARGIN_TOP,
            bottomMargin=S.MARGIN_BOTTOM,
            title=self.meta.get('title', ''),
            author=self.meta.get('org', ''),
            allowSplitting=1,   # 페이지를 넘는 플로어블 분할 허용
        )

        # 프레임 정의 (패딩 0 → 표/이미지 너비가 CONTENT_W와 정확히 일치)
        content_frame = Frame(
            S.MARGIN_LEFT, S.MARGIN_BOTTOM,
            S.CONTENT_W, S.CONTENT_H,
            id='content',
            leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
        )
        cover_frame = Frame(
            0, 0, S.PAGE_W, S.PAGE_H,
            id='cover',
            leftPadding=0, rightPadding=0,
            topPadding=0, bottomPadding=0,
        )

        # 페이지 템플릿
        cover_tmpl = PageTemplate(
            id='cover', frames=[cover_frame],
            onPage=self._on_cover_page,
        )
        normal_tmpl = PageTemplate(
            id='normal', frames=[content_frame],
            onPage=self._on_page,
        )
        doc_template.addPageTemplates([cover_tmpl, normal_tmpl])

        # ChapterMarker._CURRENT 초기화
        ChapterMarker._CURRENT[0] = ''

        # ── Flowables 구성 ─────────────────────────────────────────
        from reportlab.platypus.doctemplate import NextPageTemplate

        story = []

        # 표지
        story.append(NextPageTemplate('cover'))
        story += self._build_cover()

        # 목차
        story.append(NextPageTemplate('normal'))

        # 본문 아이템 처리
        toc_items   = []
        body_items  = []
        in_toc      = True

        for item in doc.items:
            if isinstance(item, DocPara):
                if item.style in ('TOC1', 'TOC2', 'TOC3', 'TOC4', 'TOC5'):
                    in_toc = True
                    toc_items.append(item)
                    continue
                else:
                    in_toc = False

            body_items.append(item)

        # 목차 섹션
        if toc_items:
            story += self._build_toc_header()
            for ti in toc_items:
                fl = self._build_toc_entry(ti)
                if fl:
                    story.append(fl)
            story.append(PageBreak())

        # PART 구분 페이지 처리: PART 항목 찾기
        part_groups = self._group_by_parts(body_items)

        for part_info, items_in_part in part_groups:
            if part_info:
                part_text, chapter_lines = part_info
                # NextPageTemplate → PageBreak → 내용 순서가 맞아야 함
                story.append(NextPageTemplate('cover'))
                story.append(PageBreak())
                story.append(PartDividerPage(part_text, chapter_lines))
                story.append(NextPageTemplate('normal'))
                story.append(PageBreak())

            for item in items_in_part:
                if isinstance(item, DocPara):
                    fls = self._convert_paragraph(item)
                    story += fls
                elif isinstance(item, DocTable):
                    story += self._convert_table(item)
                elif isinstance(item, DocImage):
                    story += self._convert_image(item)

        doc_template.build(story)
        print(f'PDF 생성 완료: {self.output_path}')

    def _group_by_parts(
        self, items: list
    ) -> list[tuple[tuple | None, list]]:
        """PART 단락을 기준으로 아이템 그룹화"""
        groups: list[tuple[tuple | None, list]] = []
        current_part = None
        current_items: list = []

        # 첫 PART 전에 오는 내용을 위한 초기 그룹
        groups.append((None, current_items))

        for item in items:
            if isinstance(item, DocPara) and item.style == 'PART':
                text = item.text.strip()
                # 다음 PART 아이템 수집 (챕터 라인은 이후 Heading1에서)
                current_items = []
                current_part = (text, [])
                groups.append((current_part, current_items))
            elif isinstance(item, DocPara) and item.style == 'Heading1':
                if current_part:
                    raw = item.text.strip()
                    # "1표준화 추진..." → "01  표준화 추진..."
                    m = re.match(r'^(\d+)\s*(.*)', raw)
                    if m:
                        ch_line = f'{int(m.group(1)):02d}  {m.group(2).strip()}'
                    else:
                        ch_line = raw
                    current_part[1].append(ch_line)
                current_items.append(item)
            else:
                current_items.append(item)

        return groups
