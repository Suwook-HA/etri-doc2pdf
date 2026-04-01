"""
ETRI 표준체계 및 선도전략 DOCX → PDF 변환기
사용법: python main.py <input.docx> [output.pdf] [옵션]
"""
from __future__ import annotations
import argparse
import os
import sys
import re

from .docx_parser import DocxParser
from .pdf_generator import ETRIPdfGenerator


def extract_meta(doc_path: str) -> dict:
    """파일명 및 문서 내용에서 메타데이터 추출"""
    basename = os.path.splitext(os.path.basename(doc_path))[0]

    meta = {
        'title':    '표준체계 및 선도전략',
        'subtitle': '',
        'year':     '2025',
        'date':     '2025.12.',
        'org':      'ICT전략연구소 표준연구본부',
    }

    # 파일명에서 연도 추출 (예: ..._2025_...)
    m = re.search(r'(\d{4})', basename)
    if m:
        meta['year'] = m.group(1)
        meta['date'] = f"{m.group(1)}.12."

    return meta


def main():
    parser = argparse.ArgumentParser(
        description='ETRI DOCX → PDF 출판물 변환기',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python -m etri_doc2pdf 문서.docx
  python -m etri_doc2pdf 문서.docx 출력.pdf
  python -m etri_doc2pdf 문서.docx --title "표준체계" --year 2025
        """,
    )
    parser.add_argument('input',  help='입력 DOCX 파일 경로')
    parser.add_argument('output', nargs='?', help='출력 PDF 파일 경로 (기본: 입력명.pdf)')
    parser.add_argument('--title',    default='', help='문서 제목')
    parser.add_argument('--subtitle', default='', help='부제목')
    parser.add_argument('--year',     default='', help='연도 (예: 2025)')
    parser.add_argument('--date',     default='', help='날짜 (예: 2025.12.)')
    parser.add_argument('--org',      default='', help='발행 기관명')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'오류: 파일을 찾을 수 없습니다: {args.input}', file=sys.stderr)
        sys.exit(1)

    # 출력 경로
    if args.output:
        output_path = args.output
    else:
        base = os.path.splitext(args.input)[0]
        output_path = base + '_ETRI_Design.pdf'

    # 메타데이터
    meta = extract_meta(args.input)
    if args.title:    meta['title']    = args.title
    if args.subtitle: meta['subtitle'] = args.subtitle
    if args.year:     meta['year']     = args.year
    if args.date:     meta['date']     = args.date
    if args.org:      meta['org']      = args.org

    print(f'변환 시작: {args.input}')
    print(f'  제목: {meta["title"]}  연도: {meta["year"]}')

    # 파싱
    parser_obj = DocxParser(args.input)
    try:
        doc = parser_obj.parse()
    finally:
        parser_obj.close()

    print(f'  파싱 완료: {len(doc.items)} 개 항목')

    # PDF 생성
    gen = ETRIPdfGenerator(output_path, meta)
    gen.build(doc)

    print(f'완료: {output_path}')


if __name__ == '__main__':
    main()
