"""
ETRI 표준체계 및 선도전략 디자인 스타일 정의
"""
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ── 색상 ─────────────────────────────────────────────────────────────
BLUE_PRIMARY   = colors.HexColor('#1565C0')   # ETRI 주색 (진파랑)
BLUE_LIGHT     = colors.HexColor('#1976D2')   # 섹션 헤더 바
BLUE_PALE      = colors.HexColor('#E3F0FC')   # 콜아웃 배경
BLUE_DIVIDER   = colors.HexColor('#1565C0')   # PART 구분 페이지 배경
ORANGE         = colors.HexColor('#E55B2A')   # 강조 오렌지
GRAY_DARK      = colors.HexColor('#2D2D2D')   # 본문 텍스트
GRAY_MID       = colors.HexColor('#5A5A5A')   # 부제목
GRAY_LIGHT     = colors.HexColor('#AAAAAA')   # 구분선
WHITE          = colors.white
TABLE_HEADER   = colors.HexColor('#1565C0')   # 표 헤더
TABLE_ODD      = colors.HexColor('#F0F6FB')   # 표 홀수행
TABLE_EVEN     = colors.white                  # 표 짝수행
FOOTER_LINE    = colors.HexColor('#CCCCCC')
TOC_DOT        = colors.HexColor('#888888')

# ── 페이지 크기 (B5) ─────────────────────────────────────────────────
PAGE_W = 182 * mm
PAGE_H = 257 * mm

MARGIN_LEFT   = 25 * mm
MARGIN_RIGHT  = 22 * mm
MARGIN_TOP    = 25 * mm
MARGIN_BOTTOM = 20 * mm

CONTENT_W = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_H = PAGE_H - MARGIN_TOP - MARGIN_BOTTOM

# ── 폰트 등록 ──────────────────────────────────────────────────────
FONT_DIR = r'C:\Windows\Fonts'

_FONTS_REGISTERED = False

def register_fonts():
    global _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return
    candidates = {
        'Regular': ['malgun.ttf', 'NanumGothic.ttf'],
        'Bold':    ['malgunbd.ttf', 'NanumGothicBold.ttf'],
    }
    for variant, files in candidates.items():
        for fname in files:
            path = os.path.join(FONT_DIR, fname)
            if os.path.exists(path):
                name = 'Korean' if variant == 'Regular' else 'Korean-Bold'
                pdfmetrics.registerFont(TTFont(name, path))
                break
    _FONTS_REGISTERED = True

register_fonts()

FONT_REGULAR = 'Korean'
FONT_BOLD    = 'Korean-Bold'

# ── 폰트 크기 ──────────────────────────────────────────────────────
FS_H1      = 15    # Chapter 제목
FS_H2      = 12    # 절 제목
FS_H3      = 10.5  # 소절 제목
FS_BODY    = 9.5   # 본문
FS_SMALL   = 8.5   # 캡션, 각주
FS_TOC1    = 10
FS_TOC2    = 9.5
FS_TOC3    = 9

# ── 줄간격 ──────────────────────────────────────────────────────────
LEADING_BODY = 16
LEADING_H1   = 22
LEADING_H2   = 18
LEADING_H3   = 15
