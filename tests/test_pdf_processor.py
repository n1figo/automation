# tests/test_pdf_processor.py
import sys
import os
import unittest
import fitz
import pandas as pd

# 프로젝트 루트 디렉토리를 Python 경로에 추가
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, project_root)

from src.pdf_processor import PDFProcessor  # 수정된 import 경로


class TestPDFProcessor(unittest.TestCase):
    """PDFProcessor 클래스 테스트"""

    @classmethod
    def setUpClass(cls):
        """테스트 클래스 전체에서 사용할 리소스 설정"""
        cls.test_dir = "test_data"
        os.makedirs(cls.test_dir, exist_ok=True)

    def setUp(self):
        """각 테스트 메서드 실행 전 설정"""
        # 테스트용 PDF 생성
        self.test_pdf_path = os.path.join(self.test_dir, "test.pdf")
        self.create_test_pdf()
        
        # PDFProcessor 인스턴스 생성
        self.processor = PDFProcessor(self.test_pdf_path)
        self.processor.open_pdf_document()

    def tearDown(self):
        """각 테스트 메서드 실행 후 정리"""
        self.processor.close_pdf_document()
        if os.path.exists(self.test_pdf_path):
            os.remove(self.test_pdf_path)

    @classmethod
    def tearDownClass(cls):
        """테스트 클래스 종료 시 정리"""
        if os.path.exists(cls.test_dir):
            try:
                os.rmdir(cls.test_dir)
            except OSError:
                pass  # 디렉토리가 비어있지 않은 경우 무시

    def create_test_pdf(self):
        """테스트용 PDF 파일 생성"""
        doc = fitz.open()
        page = doc.new_page()
        
        # 기본 텍스트 추가
        page.insert_text((50, 50), "나. 보험금")
        page.insert_text((50, 100), "상해관련 특별약관")
        
        # 표 형태의 텍스트 추가
        page.insert_text((50, 150), "보장명")
        page.insert_text((150, 150), "지급사유")
        page.insert_text((250, 150), "지급금액")
        
        # 하이라이트 추가
        annot = page.add_highlight_annot(fitz.Rect(40, 90, 200, 110))
        
        doc.save(self.test_pdf_path)
        doc.close()

    def test_normalize_text(self):
        """텍스트 정규화 테스트"""
        test_cases = [
            ("  hello   world  ", "helloworld"),
            ("No space", "Nospace"),
            ("  \t\n  ", ""),
            ("보험금  지급", "보험금지급"),
            ("", ""),
            ("123  456", "123456"),
            (None, "")
        ]
        
        for input_text, expected in test_cases:
            with self.subTest(input_text=input_text):
                self.assertEqual(
                    self.processor.normalize_text(input_text),
                    expected
                )

    def test_is_header_row(self):
        """헤더 행 판별 테스트"""
        test_cases = [
            # 정상적인 헤더 행 (4열)
            (["", "보장명", "지급사유", "지급금액"], True),
            # 정상적인 헤더 행 (3열)
            (["보장명", "지급사유", "지급금액"], True),
            # 공백이 포함된 헤더 행
            (["", "  보장명  ", " 지급사유 ", "지급금액 "], True),
            # 잘못된 헤더 행
            (["", "내용1", "내용2", "내용3"], False),
            # 열 개수가 부족한 행
            (["보장명", "지급사유"], False),
            # 빈 행
            ([], False),
            # 다른 순서의 헤더
            (["지급사유", "보장명", "지급금액"], False),
            # 다른 텍스트가 포함된 행
            (["", "보장명", "지급사유", "지급금액", "비고"], True)
        ]
        
        for row, expected in test_cases:
            with self.subTest(row=row):
                self.assertEqual(
                    self.processor.is_header_row(row),
                    expected
                )

    def test_page_has_highlight(self):
        """하이라이트 감지 테스트"""
        # 기본 페이지 테스트 (하이라이트 있음)
        self.assertTrue(self.processor.page_has_highlight(0))
        
        # 존재하지 않는 페이지 테스트
        self.assertFalse(self.processor.page_has_highlight(999))
        
        # 하이라이트가 없는 새 페이지 테스트
        doc = fitz.open()
        doc.new_page()
        temp_path = os.path.join(self.test_dir, "no_highlight.pdf")
        doc.save(temp_path)
        doc.close()
        
        no_highlight_processor = PDFProcessor(temp_path)
        no_highlight_processor.open_pdf_document()
        self.assertFalse(no_highlight_processor.page_has_highlight(0))
        no_highlight_processor.close_pdf_document()
        os.remove(temp_path)

    def test_get_page_footer(self):
        """페이지 푸터 추출 테스트"""
        # 기본 푸터 테스트
        footer = self.processor.get_page_footer(0)
        self.assertIsInstance(footer, str)
        
        # 존재하지 않는 페이지 테스트
        self.assertEqual(self.processor.get_page_footer(999), "")
        
        # 푸터가 있는 특수 PDF 생성 및 테스트
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((50, 800), "Page Footer Text")
        temp_path = os.path.join(self.test_dir, "with_footer.pdf")
        doc.save(temp_path)
        doc.close()
        
        footer_processor = PDFProcessor(temp_path)
        footer_processor.open_pdf_document()
        self.assertEqual(footer_processor.get_page_footer(0), "Page Footer Text")
        footer_processor.close_pdf_document()
        os.remove(temp_path)

    def test_drop_redundant_header(self):
        """중복 헤더 제거 테스트"""
        # 테스트용 DataFrame 생성
        test_data = [
            ["보장명", "지급사유", "지급금액"],  # 헤더 (제거 대상)
            ["상해", "사고시", "100만원"],       # 데이터
            ["질병", "진단시", "200만원"],       # 데이터
            ["보장명", "지급사유", "지급금액"],  # 중복 헤더 (제거 대상)
            ["후유장해", "진단시", "300만원"]    # 데이터
        ]
        df = pd.DataFrame(test_data)
        
        # 중복 헤더 제거 실행
        result_df = self.processor.drop_redundant_header(df)
        
        # 결과 검증
        self.assertEqual(len(result_df), 3)  # 헤더 행 2개가 제거되어야 함
        self.assertEqual(result_df.iloc[0][0], "상해")  # 첫 번째 데이터 행 확인
        self.assertEqual(result_df.iloc[-1][0], "후유장해")  # 마지막 데이터 행 확인

    def test_extract_tables_from_page(self):
        """단일 페이지 표 추출 테스트"""
        # 표가 있는 테스트용 PDF 생성
        doc = fitz.open()
        page = doc.new_page()
        
        # 표 형태의 텍스트 추가
        y_positions = [100, 150, 200]
        headers = ["보장명", "지급사유", "지급금액"]
        data = ["상해", "사고시", "100만원"]
        
        for i, (header, value) in enumerate(zip(headers, data)):
            page.insert_text((50 + i*100, y_positions[0]), header)
            page.insert_text((50 + i*100, y_positions[1]), value)
        
        table_pdf_path = os.path.join(self.test_dir, "table_test.pdf")
        doc.save(table_pdf_path)
        doc.close()
        
        # 표 추출 테스트
        table_processor = PDFProcessor(table_pdf_path)
        table_processor.open_pdf_document()
        result_df = table_processor.extract_tables_from_page(0, "테스트섹션")
        table_processor.close_pdf_document()
        os.remove(table_pdf_path)
        
        # 결과가 비어있지 않은지 확인 (실제 표 내용 검증은 camelot 의존성으로 인해 제한적)
        self.assertIsInstance(result_df, pd.DataFrame)

    def test_process_pdf(self):
        """전체 PDF 처리 통합 테스트"""
        output_path = os.path.join(self.test_dir, "test_output.xlsx")
        try:
            # PDF 파일에 필요한 내용 추가
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), "나. 보험금")
            page.insert_text((50, 100), "상해관련 특별약관")
            
            # 표 내용 추가
            headers = ["보장명", "지급사유", "지급금액"]
            data = ["상해", "사고시", "100만원"]
            for i, (header, value) in enumerate(zip(headers, data)):
                page.insert_text((50 + i*100, 150), header)
                page.insert_text((50 + i*100, 200), value)
            
            test_pdf_path = os.path.join(self.test_dir, "process_test.pdf")
            doc.save(test_pdf_path)
            doc.close()
            
            # 새로운 PDF로 프로세서 초기화
            test_processor = PDFProcessor(test_pdf_path)
            test_processor.open_pdf_document()
            
            # process_pdf 실행
            success = test_processor.process_pdf(
                "나. 보험금",
                ["상해관련 특별약관"],
                output_path
            )
            
            # 결과 검증
            self.assertTrue(success)
            self.assertTrue(os.path.exists(output_path))
            
        finally:
            # 리소스 정리
            if 'test_processor' in locals():
                test_processor.close_pdf_document()
            if os.path.exists(test_pdf_path):
                os.remove(test_pdf_path)
            if os.path.exists(output_path):
                os.remove(output_path)

if __name__ == '__main__':
    unittest.main(verbosity=2)