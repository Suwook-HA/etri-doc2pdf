"""
DOCX 파일 파싱 → 콘텐츠 모델 추출
"""
from __future__ import annotations
import zipfile
import io
import os
from dataclasses import dataclass, field
from typing import Any
from lxml import etree

# Word XML 네임스페이스
NS = {
    'w':  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':  'urn:schemas-microsoft-com:vml',
}


# ── 콘텐츠 노드 타입 ────────────────────────────────────────────────

@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False
    color: str | None = None  # hex string e.g. "FF0000"
    size: float | None = None  # half-points


@dataclass
class Paragraph:
    style: str          # 'Normal', 'Heading1' .. 'Heading9', 'PART', 'TOC1'..'TOC3'
    runs: list[TextRun] = field(default_factory=list)
    indent_level: int = 0   # 들여쓰기 수준 (불릿 등)
    is_bullet: bool = False
    num_id: int | None = None
    num_level: int = 0
    alignment: str = 'left'  # left / center / right / justify

    @property
    def text(self) -> str:
        return ''.join(r.text for r in self.runs)


@dataclass
class TableCell:
    paragraphs: list[Paragraph] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    bg_color: str | None = None
    bold: bool = False
    is_header: bool = False
    vmerge: str = ''   # 'restart' | 'continue' | ''
    align: str = ''    # 'center' | 'right' | '' (셀 수준 정렬 힌트)


@dataclass
class TableRow:
    cells: list[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    rows: list[TableRow] = field(default_factory=list)
    col_widths: list[float] = field(default_factory=list)  # relative widths


@dataclass
class Image:
    data: bytes
    ext: str = 'png'
    width_emu: int = 0
    height_emu: int = 0
    caption: str = ''


@dataclass
class Document:
    title: str = ''
    subtitle: str = ''
    year: str = ''
    org: str = ''
    date: str = ''
    cover_image: bytes | None = None
    items: list[Any] = field(default_factory=list)  # Paragraph | Table | Image


# ── 파서 ────────────────────────────────────────────────────────────

class DocxParser:
    def __init__(self, path: str):
        self.path = path
        self.zf = zipfile.ZipFile(path, 'r')
        self._load_rels()
        self._load_numbering()

    def _load_rels(self):
        """이미지 relationship 로드"""
        self.rels: dict[str, str] = {}
        try:
            data = self.zf.read('word/_rels/document.xml.rels')
            root = etree.fromstring(data)
            for rel in root:
                rid  = rel.get('Id', '')
                tgt  = rel.get('Target', '')
                self.rels[rid] = tgt
        except Exception:
            pass

    def _load_numbering(self):
        """번호 매기기 정의 로드"""
        self.num_formats: dict[int, str] = {}   # numId → format type
        try:
            data = self.zf.read('word/numbering.xml')
            root = etree.fromstring(data)
            w = NS['w']
            for num in root.findall(f'{{{w}}}num'):
                nid_el = num.get(f'{{{w}}}numId')
                nid = int(nid_el) if nid_el else 0
                self.num_formats[nid] = 'bullet'
        except Exception:
            pass

    def _get_image_data(self, rid: str) -> tuple[bytes, str] | None:
        tgt = self.rels.get(rid, '')
        if not tgt:
            return None
        path = 'word/' + tgt if not tgt.startswith('/') else tgt.lstrip('/')
        try:
            data = self.zf.read(path)
            ext = os.path.splitext(path)[-1].lstrip('.').lower() or 'png'
            if ext == 'jpeg':
                ext = 'jpg'
            return data, ext
        except Exception:
            return None

    def _parse_run(self, r_el) -> TextRun | None:
        w = NS['w']
        # 텍스트 수집
        texts = []
        for t in r_el.findall(f'{{{w}}}t'):
            if t.text:
                texts.append(t.text)
        text = ''.join(texts)
        if not text:
            return None

        # 서식
        rpr = r_el.find(f'{{{w}}}rPr')
        bold = italic = False
        color = None
        size = None
        if rpr is not None:
            bold   = rpr.find(f'{{{w}}}b') is not None
            italic = rpr.find(f'{{{w}}}i') is not None
            c_el   = rpr.find(f'{{{w}}}color')
            if c_el is not None:
                cv = c_el.get(f'{{{w}}}val', '')
                if cv and cv.upper() != 'AUTO':
                    color = cv
            sz_el = rpr.find(f'{{{w}}}sz')
            if sz_el is not None:
                sv = sz_el.get(f'{{{w}}}val', '')
                if sv:
                    size = int(sv) / 2.0

        return TextRun(text=text, bold=bold, italic=italic, color=color, size=size)

    def _parse_paragraph(self, p_el) -> Paragraph | None:
        w = NS['w']
        ppr = p_el.find(f'{{{w}}}pPr')

        # 스타일
        style_id = ''
        if ppr is not None:
            ps = ppr.find(f'{{{w}}}pStyle')
            if ps is not None:
                style_id = ps.get(f'{{{w}}}val', '')

        # 스타일 매핑
        style = self._map_style(style_id)

        # 정렬
        alignment = 'justify'
        if ppr is not None:
            jc = ppr.find(f'{{{w}}}jc')
            if jc is not None:
                jv = jc.get(f'{{{w}}}val', '')
                if jv in ('center', 'right', 'left'):
                    alignment = jv

        # 번호 매기기
        num_id = None
        num_level = 0
        if ppr is not None:
            num_pr = ppr.find(f'{{{w}}}numPr')
            if num_pr is not None:
                ilvl = num_pr.find(f'{{{w}}}ilvl')
                nid  = num_pr.find(f'{{{w}}}numId')
                if ilvl is not None:
                    num_level = int(ilvl.get(f'{{{w}}}val', '0'))
                if nid is not None:
                    nv = nid.get(f'{{{w}}}val', '0')
                    if nv != '0':
                        num_id = int(nv)

        # 텍스트 런 수집
        runs = []
        for child in p_el:
            tag = child.tag.split('}')[-1]
            if tag == 'r':
                run = self._parse_run(child)
                if run:
                    runs.append(run)
            elif tag == 'hyperlink':
                for r in child.findall(f'{{{w}}}r'):
                    run = self._parse_run(r)
                    if run:
                        runs.append(run)

        return Paragraph(
            style=style,
            runs=runs,
            num_id=num_id,
            num_level=num_level,
            alignment=alignment,
        )

    def _map_style(self, style_id: str) -> str:
        """DOCX 스타일 ID → 내부 스타일명 매핑"""
        mapping = {
            '1':    'Heading1',
            '2':    'Heading2',
            '3':    'Heading3',
            '4':    'Heading4',
            '5':    'Heading5',
            '10':   'TOC1',
            '20':   'TOC2',
            '30':   'TOC3',
            '40':   'TOC4',
            '50':   'TOC5',
            'af6':  'PART',     # PART I, PART II ...
            'ab':   'Normal',
            'a':    'Normal',
            'NF':   'Normal',
            'NO':   'Normal',
            'PL':   'Bullet',
            'TT':   'Caption',
            'ZD':   'Callout',
            'ZGSM': 'Callout',
            'EQ':   'Equation',
        }
        return mapping.get(style_id, 'Normal')

    def _parse_table(self, tbl_el) -> Table:
        w = NS['w']
        rows = []

        # 컬럼 너비
        col_widths = []
        tblpr = tbl_el.find(f'{{{w}}}tblPr')
        tbl_grid = tbl_el.find(f'{{{w}}}tblGrid')
        if tbl_grid is not None:
            for gc in tbl_grid.findall(f'{{{w}}}gridCol'):
                wv = gc.get(f'{{{w}}}w', '1000')
                col_widths.append(int(wv))

        _HEADER_COLORS = {
            '1565C0', '1976D2', '2E75B6', '17375E',
            '003366', '0070C0', '4472C4', '244185',
            '1F3864', '2F5496', '305496',
        }

        for r_idx, tr_el in enumerate(tbl_el.findall(f'{{{w}}}tr')):
            cells = []
            for tc_el in tr_el.findall(f'{{{w}}}tc'):
                tcpr = tc_el.find(f'{{{w}}}tcPr')
                colspan  = 1
                bg_color = None
                vmerge   = ''
                cell_align = ''

                if tcpr is not None:
                    # colspan
                    span_el = tcpr.find(f'{{{w}}}gridSpan')
                    if span_el is not None:
                        colspan = int(span_el.get(f'{{{w}}}val', '1'))

                    # vMerge (rowspan)
                    vm_el = tcpr.find(f'{{{w}}}vMerge')
                    if vm_el is not None:
                        val = vm_el.get(f'{{{w}}}val', '')
                        vmerge = 'restart' if val == 'restart' else 'continue'

                    # 배경색
                    shd = tcpr.find(f'{{{w}}}shd')
                    if shd is not None:
                        fill = shd.get(f'{{{w}}}fill', '')
                        if fill and fill.upper() not in ('AUTO', 'FFFFFF', ''):
                            bg_color = fill

                    # 셀 수준 텍스트 정렬 (tcPr > vAlign)
                    # (수평 정렬은 p 수준에서 처리)

                # 셀 텍스트
                paras = []
                for p_el in tc_el.findall(f'{{{w}}}p'):
                    para = self._parse_paragraph(p_el)
                    if para:
                        paras.append(para)

                # 헤더 행 감지
                is_header = (
                    r_idx == 0
                    or (bg_color and bg_color.upper() in _HEADER_COLORS)
                )

                # 셀 정렬 힌트: 모든 단락이 center면 center
                aligns = [p.alignment for p in paras if p.alignment]
                if aligns and all(a == 'center' for a in aligns):
                    cell_align = 'center'
                elif aligns and all(a == 'right' for a in aligns):
                    cell_align = 'right'

                cells.append(TableCell(
                    paragraphs=paras,
                    colspan=colspan,
                    bg_color=bg_color,
                    is_header=is_header,
                    vmerge=vmerge,
                    align=cell_align,
                ))
            rows.append(TableRow(cells=cells, is_header=r_idx == 0))

        # rowspan 계산: vmerge='restart' 셀의 rowspan을 실제 값으로 설정
        # 그리드 좌표 → 셀 매핑 구성
        grid: list[list[TableCell | None]] = []
        for row in rows:
            grid_row: list[TableCell | None] = []
            col_cursor = 0
            cell_iter = iter(row.cells)
            for cell in row.cells:
                while col_cursor < len(grid_row):
                    col_cursor += 1
                for _ in range(cell.colspan):
                    grid_row.append(cell)
                col_cursor += cell.colspan
            grid.append(grid_row)

        n_grid_cols = max((len(r) for r in grid), default=0)
        for c in range(n_grid_cols):
            r = 0
            while r < len(grid):
                row_g = grid[r]
                if c >= len(row_g):
                    r += 1
                    continue
                cell = row_g[c]
                if cell and cell.vmerge == 'restart':
                    span = 1
                    for r2 in range(r + 1, len(grid)):
                        if c < len(grid[r2]) and grid[r2][c] and grid[r2][c].vmerge == 'continue':
                            span += 1
                        else:
                            break
                    cell.rowspan = span
                r += 1

        return Table(rows=rows, col_widths=col_widths)

    def _parse_image_from_drawing(self, drawing_el) -> Image | None:
        """drawing 요소에서 이미지 추출"""
        a = NS['a']
        pic_ns = NS['pic']
        r_ns = NS['r']

        # blip r:embed 찾기
        blip_els = drawing_el.findall(f'.//{{{a}}}blip')
        if not blip_els:
            return None
        rid = blip_els[0].get(f'{{{r_ns}}}embed', '')
        if not rid:
            return None

        img_data = self._get_image_data(rid)
        if not img_data:
            return None
        data, ext = img_data

        # 크기 (EMU)
        extent_el = drawing_el.find(
            f'.//{{{NS["wp"]}}}extent'
        )
        w_emu = h_emu = 0
        if extent_el is not None:
            w_emu = int(extent_el.get('cx', 0))
            h_emu = int(extent_el.get('cy', 0))

        return Image(data=data, ext=ext, width_emu=w_emu, height_emu=h_emu)

    def parse(self) -> Document:
        doc = Document()
        data = self.zf.read('word/document.xml')
        root = etree.fromstring(data)
        w = NS['w']
        body = root.find(f'{{{w}}}body')
        if body is None:
            return doc

        items = []
        prev_caption = ''

        for child in body:
            tag = child.tag.split('}')[-1]

            if tag == 'p':
                para = self._parse_paragraph(child)
                if para is None:
                    continue

                # 이미지가 문단 안에 있는지 확인
                drawings = child.findall(f'.//{{{w}}}drawing')
                if drawings:
                    for drw in drawings:
                        img = self._parse_image_from_drawing(drw)
                        if img:
                            img.caption = prev_caption
                            items.append(img)
                    prev_caption = ''
                    continue

                text = para.text.strip()
                if not text:
                    continue

                # 캡션 저장
                if para.style == 'Caption':
                    prev_caption = text
                    items.append(para)
                    continue

                # 문서 메타데이터 추출 (표지)
                if not doc.title and para.style in ('Heading1', 'PART'):
                    pass  # 첫 번째 콘텐츠는 본문으로 처리

                items.append(para)

            elif tag == 'tbl':
                tbl = self._parse_table(child)
                items.append(tbl)

        doc.items = items
        return doc

    def close(self):
        self.zf.close()
