import sys
import os
import unittest
import fitz
import pandas as pd

# pdf_processor.py 파일이 있는 현재 디렉토리를 sys.path에 추가합니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_processor import PDFProcessor

class TestPDFProcessor(unittest.TestCase):
    """PDFProcessor 클래스에 대한 단위 테스트 (추가 테스트 포함)"""

    def setUp(self):
        """테스트 셋업: dummy_pdf 파일 생성 및 PDFProcessor 인스턴스 초기화"""
        self.dummy_pdf_path = "dummy_pdf.pdf"
        if not os.path.exists(self.dummy_pdf_path):
            with fitz.open() as pdf_doc:
                page = pdf_doc.new_page()
                # 두 줄 이상의 텍스트를 입력해 Footer 테스트에 활용
                page.insert_text((50, 50), "This is a dummy PDF file for testing.\nFooterText")
                pdf_doc.save(self.dummy_pdf_path)
        self.processor = PDFProcessor(self.dummy_pdf_path)
        self.processor.open_pdf_document()

    def tearDown(self):
        """테스트 종료 후 문서 객체 닫기 및 파일 삭제"""
        if hasattr(self.processor, 'doc') and self.processor.doc:
            self.processor.doc.close()
        if os.path.exists(self.dummy_pdf_path):
            os.remove(self.dummy_pdf_path)

    def test_normalize_text(self):
        """텍스트 정규화 기능 테스트"""
        self.assertEqual(self.processor.normalize_text("  hello   world  "), "helloworld")
        self.assertEqual(self.processor.normalize_text("No space"), "Nospace")
        self.assertEqual(self.processor.normalize_text("  \t\n  "), "")

    def test_is_header_row(self):
        """표 헤더 행 판별 기능 테스트"""
        header_row_long = ["", "보장명", "지급사유", "지급금액", "기타"]
        header_row_short = ["보장명", "지급사유", "지급금액"]
        not_header_row = ["", "내용1", "내용2", "내용3"]

        self.assertTrue(self.processor.is_header_row(header_row_long))
        self.assertTrue(self.processor.is_header_row(header_row_short))
        self.assertFalse(self.processor.is_header_row(not_header_row))
        self.assertTrue(self.processor.is_header_row(["", "  보장명  ", " 지급사유 ", "지급금액 "]))

    def test_page_has_highlight(self):
        """페이지 하이라이트 감지 기능 테스트"""
        # 기본 dummy PDF에는 형광펜 주석이 없음
        self.assertFalse(self.processor.page_has_highlight(0))

        # 형광펜 주석이 포함된 PDF 생성 테스트
        highlight_pdf_path = "highlight_pdf.pdf"
        if not os.path.exists(highlight_pdf_path):
            with fitz.open() as pdf_doc:
                page = pdf_doc.new_page()
                page.insert_text((50, 50), "This PDF has highlight.")
                annot = page.add_annot(rect=fitz.Rect(40, 40, 150, 60), type=8)  # Highlight annotation 추가
                annot.update()
                pdf_doc.save(highlight_pdf_path)

        processor_highlight = PDFProcessor(highlight_pdf_path)
        processor_highlight.open_pdf_document()
        self.assertTrue(processor_highlight.page_has_highlight(0))
        if hasattr(processor_highlight, 'doc') and processor_highlight.doc:
            processor_highlight.doc.close()
        os.remove(highlight_pdf_path)

    def test_get_page_footer(self):
        """페이지 하단 Footer 텍스트 추출 테스트"""
        footer = self.processor.get_page_footer(0)
        self.assertIn("FooterText", footer)

    def test_split_text_into_chunks(self):
        """텍스트 청크 분할 기능 테스트"""
        text = "Sentence one. Sentence two. Sentence three. Sentence four."
        # 최대 청크 크기를 작게 설정하여 여러 청크가 생성되는지 테스트
        chunks = self.processor.split_text_into_chunks(text, max_chunk_size=25)
        self.assertIsInstance(chunks, list)
        self.assertGreater(len(chunks), 0)
        for chunk in chunks:
            self.assertTrue(len(chunk) > 0)

    def test_drop_redundant_header(self):
        """중복 헤더 행 제거 기능 테스트"""
        # 중복 헤더가 포함된 DataFrame 생성
        data = [
            ["보장명", "지급사유", "지급금액"],
            ["보장명", "지급사유", "지급금액"],
            ["내용1", "내용2", "내용3"],
            ["보장명", "지급사유", "지급금액"],
            ["내용4", "내용5", "내용6"]
        ]
        df = pd.DataFrame(data)
        filtered_df = self.processor.drop_redundant_header(df)
        # 필터링 후 헤더 행이 단 한 번만 있어야 함
        header_count = sum(1 for _, row in filtered_df.iterrows() if self.processor.is_header_row(list(row)))
        self.assertEqual(header_count, 1)

if __name__ == '__main__':
    unittest.main()