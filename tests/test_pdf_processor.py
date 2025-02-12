import sys
import os
import unittest
import fitz
import pandas as pd

# pdf_processor.py 파일이 있는 현재 디렉토리를 sys.path에 추가합니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_processor import PDFProcessor


##### 테스트 코드 (test_pdf_processor.py 파일로 분리하는 것이 일반적) #####
class TestPDFProcessor(unittest.TestCase):
    """PDFProcessor 클래스에 대한 단위 테스트"""

    def setUp(self):
        """테스트 셋업: PDFProcessor 인스턴스 생성 (dummy_pdf_path 필요)"""
        self.dummy_pdf_path = "dummy_pdf.pdf"  # 실제 테스트 PDF 파일 경로로 변경 필요
        # dummy_pdf.pdf 파일이 존재하지 않으면 생성 (예시)
        if not os.path.exists(self.dummy_pdf_path):
            with fitz.open() as pdf_doc:
                page = pdf_doc.new_page()
                page.insert_text((50, 50), "This is a dummy PDF file for testing.")
                pdf_doc.save(self.dummy_pdf_path)
        self.processor = PDFProcessor(self.dummy_pdf_path)
        self.processor.open_pdf_document() # 문서 열기

    def tearDown(self):
        """테스트 해체: PDFProcessor 객체 정리"""
        if os.path.exists(self.dummy_pdf_path):
            os.remove(self.dummy_pdf_path) # 테스트용 PDF 파일 삭제
        if self.processor.doc:
            self.processor.doc.close() # 문서 객체 닫기


    def test_normalize_text(self):
        """텍스트 정규화 기능 테스트"""
        self.assertEqual(self.processor.normalize_text("  hello   world  "), "helloworld")
        self.assertEqual(self.processor.normalize_text("No space"), "Nospace")
        self.assertEqual(self.processor.normalize_text("  \t\n  "), "") # 공백, 탭, 개행문자 제거 테스트

    def test_is_header_row(self):
        """표 헤더 행 판별 기능 테스트"""
        header_row_long = ["", "보장명", "지급사유", "지급금액", "기타"]
        header_row_short = ["보장명", "지급사유", "지급금액"]
        not_header_row = ["", "내용1", "내용2", "내용3"]

        self.assertTrue(self.processor.is_header_row(header_row_long)) # 긴 헤더 행
        self.assertTrue(self.processor.is_header_row(header_row_short)) # 짧은 헤더 행
        self.assertFalse(self.processor.is_header_row(not_header_row)) # 헤더가 아닌 행
        self.assertTrue(self.processor.is_header_row(["", "  보장명  ", " 지급사유 ", "지급금액 "])) # 공백 포함 헤더 행

    def test_page_has_highlight(self):
        """페이지 하이라이트 감지 기능 테스트"""
        # 테스트 PDF에 하이라이트 annotation 추가 필요 (fitz 또는 다른 PDF 편집 도구 사용)
        # 현재 dummy_pdf.pdf에는 하이라이트 없음
        self.assertFalse(self.processor.page_has_highlight(0)) # 하이라이트 없는 페이지

        # 하이라이트 있는 PDF 생성 및 테스트 필요 (향후 추가)
        highlight_pdf_path = "highlight_pdf.pdf"
        if not os.path.exists(highlight_pdf_path):
             with fitz.open() as pdf_doc:
                page = pdf_doc.new_page()
                page.insert_text((50, 50), "This PDF has highlight.")
                # 형광펜 주석 추가 (예시, 실제 annotation 속성은 PDF 규격에 따라 다를 수 있음)
                annot = page.add_annot(rect=fitz.Rect(40, 40, 150, 60), type=8) # Highlight annotation 추가
                annot.update()
                pdf_doc.save(highlight_pdf_path)

        processor_highlight = PDFProcessor(highlight_pdf_path)
        processor_highlight.open_pdf_document()
        self.assertTrue(processor_highlight.page_has_highlight(0)) # 하이라이트 있는 페이지 테스트
        processor_highlight.doc.close()
        os.remove(highlight_pdf_path) # 하이라이트 테스트 PDF 삭제


if __name__ == '__main__':
    main() # CLI 실행 (UI 개발 시 streamlit_app()으로 변경)
    # streamlit_app() # Streamlit UI 실행 (UI 개발 환경에서 활성화)
    unittest.main() # 테스트 실행 (main 함수와 분리하여 실행하거나, 명령행에서 'python -m unittest test_pdf_processor.py' 로 실행)