# src/pdf_processor.py
import os
import re
import PyPDF2
import camelot
import pandas as pd
import fitz
import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
from llama_cpp import Llama

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_llm_model(model_path="/workspaces/automation/models/koalpaca-polyglot-5.8b-q4_k_m.gguf"):
    """KoAlpaca GGUF 모델을 로드합니다."""
    try:
        if not os.path.exists(model_path):
            os.makedirs(os.path.dirname(model_path), exist_ok=True)
            logging.warning("모델 파일을 수동으로 다운로드하여 지정된 경로에 배치해주세요.")
            return None

        llm = Llama(
            model_path=model_path,
            n_ctx=2048,
            n_threads=4,
            n_gpu_layers=0,
            verbose=False
        )
        logging.info("KoAlpaca 모델 로드 성공")
        return llm

    except Exception as e:
        logging.error(f"모델 로드 실패: {e}")
        return None

class PDFProcessor:
    """PDF 문서 처리 및 분석을 위한 클래스"""

    def __init__(self, file_path: str):
        """PDFProcessor 초기화"""
        self.file_path = file_path
        self.doc = None
        self.pdf_reader = None
        self.MAX_CHUNK_SIZE = 2000

    def open_pdf_document(self) -> bool:
        """PDF 문서를 열고 fitz 및 PyPDF2 객체를 초기화합니다."""
        try:
            self.doc = fitz.open(self.file_path)
            with open(self.file_path, "rb") as pdf_file:
                self.pdf_reader = PyPDF2.PdfReader(pdf_file)
            logging.info(f"PDF 문서 '{self.file_path}' 로드 완료.")
            return True
        except Exception as e:
            logging.error(f"PDF 문서 열기 실패: {e}")
            self.doc = None
            self.pdf_reader = None
            return False

    def close_pdf_document(self):
        """PDF 문서 객체를 안전하게 닫습니다."""
        try:
            if self.doc:
                self.doc.close()
                self.doc = None
            self.pdf_reader = None
        except Exception as e:
            logging.error(f"PDF 문서 닫기 실패: {e}")

    @staticmethod
    def normalize_text(text: str) -> str:
        """텍스트 정규화: 공백 문자 제거"""
        if text is None:
            return ""
        if not isinstance(text, str):
            text = str(text)
        return re.sub(r'\s+', '', text)

    def is_header_row(self, row, header=["보장명", "지급사유", "지급금액"]) -> bool:
        """표 헤더 행 여부 판별"""
        try:
            if len(row) >= 4:
                cells = [self.normalize_text(row[i]) for i in range(1, 4)]
            elif len(row) == 3:
                cells = [self.normalize_text(row[i]) for i in range(0, 3)]
            else:
                return False
            norm_header = [self.normalize_text(h) for h in header]
            logging.debug(f"[DEBUG] Row: {cells}, Expected: {norm_header}")
            return cells == norm_header
        except Exception as e:
            logging.error(f"Header row check error: {e}")
            return False

    def drop_redundant_header(self, df: pd.DataFrame, header=["보장명", "지급사유", "지급금액"]) -> pd.DataFrame:
        """표 데이터프레임에서 불필요한 헤더 행 제거"""
        keep_rows = []
        for idx, row in df.iterrows():
            if self.is_header_row(row, header):
                logging.info(f"Dropping redundant header row {idx}")
            else:
                keep_rows.append(idx)
        return df.loc[keep_rows]

    def page_has_highlight(self, page_no: int) -> bool:
        """페이지에 형광펜 주석이 있는지 확인"""
        try:
            page = self.doc.load_page(page_no)
            annots = page.annots()
            if annots:
                return any(annot.type[0] == 8 for annot in annots)
            return False
        except Exception as e:
            logging.error(f"Highlight check error: {e}")
            return False

    def get_page_footer(self, page_no: int) -> str:
        """페이지 하단 텍스트 추출"""
        try:
            page = self.doc.load_page(page_no)
            text = page.get_text("text")
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            return lines[-1] if lines else ""
        except Exception as e:
            logging.error(f"Footer extraction error: {e}")
            return ""

    def split_text_into_chunks(self, text: str) -> list:
        """텍스트를 청크로 분할"""
        sentences = text.split('.')
        chunks = []
        current_chunk = []
        current_size = 0

        for sentence in sentences:
            sentence = sentence.strip() + '.'
            sentence_size = len(sentence)

            if current_size + sentence_size > self.MAX_CHUNK_SIZE:
                chunks.append(' '.join(current_chunk))
                current_chunk = [sentence]
                current_size = sentence_size
            else:
                current_chunk.append(sentence)
                current_size += sentence_size

        if current_chunk:
            chunks.append(' '.join(current_chunk))
        return chunks

    def extract_tables_from_page(self, page_num: int, term: str) -> pd.DataFrame:
        """특정 페이지에서 표 추출"""
        page_str = str(page_num + 1)
        try:
            tables = camelot.read_pdf(self.file_path, pages=page_str, flavor="lattice")
            combined_df = pd.DataFrame()

            for idx, table in enumerate(tables):
                suffix = f" - P{page_str}T{idx+1}"
                table_df = table.df.copy()

                footer_text = self.get_page_footer(page_num)
                table_df.insert(0, "PDF페이지", page_str)
                table_df.insert(0, "출처 페이지", footer_text)
                table_df.insert(0, "Source", term + suffix)

                table_df = self.drop_redundant_header(table_df)
                combined_df = pd.concat([combined_df, table_df], ignore_index=True)
            return combined_df
        except Exception as e:
            logging.error(f"Table extraction error on page {page_num+1}: {e}")
            return pd.DataFrame()

    def extract_tables_parallel(self, term: str, pages: list) -> pd.DataFrame:
        """병렬 처리로 여러 페이지에서 표 추출"""
        page_dfs = []
        with ProcessPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(self.extract_tables_from_page, page_num, term) 
                      for page_num in pages]
            for future in as_completed(futures):
                df = future.result()
                if not df.empty:
                    page_dfs.append(df)

        return pd.concat(page_dfs, ignore_index=True) if page_dfs else pd.DataFrame()

    def process_pdf(self, initial_term: str, search_terms: list, output_file: str) -> bool:
        """PDF 처리 메인 함수"""
        try:
            if not self.doc or not self.pdf_reader:
                logging.error("PDF document not loaded")
                return False

            # Reopen PDF file for PyPDF2
            with open(self.file_path, "rb") as file:
                self.pdf_reader = PyPDF2.PdfReader(file)

                # Find starting page - 이 부분을 원본 코드 방식으로 수정
                start_page = None
                for i, page in enumerate(self.pdf_reader.pages):
                    text = page.extract_text()
                    if text:
                        normalized_search = self.normalize_text(initial_term)
                        normalized_text = self.normalize_text(text)
                        if normalized_search in normalized_text:
                            start_page = i
                            logging.info(f"Found initial term '{initial_term}' on page {i+1}")
                            break

                if start_page is None:
                    logging.error(f"Term '{initial_term}' not found")
                    return False

                try:
                    # Extract tables and save to Excel
                    excel_writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
                    
                    for term in search_terms:
                        pages = range(start_page, len(self.pdf_reader.pages))
                        combined_df = self.extract_tables_parallel(term, pages)
                        
                        if not combined_df.empty:
                            sheet_name = term.replace(" ", "")[:31]
                            combined_df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
                            logging.info(f"Tables extracted for {term}")
                        else:
                            logging.warning(f"No tables found for {term}")

                    excel_writer.close()
                    logging.info(f"Excel file saved: {output_file}")
                    return True

                except Exception as e:
                    logging.error(f"Excel processing error: {e}")
                    return False

        except Exception as e:
            logging.error(f"PDF processing error: {e}")
            return False

def main():
    """메인 함수"""
    INPUT_PDF_PATH = "/workspaces/automation/data/input/0211/KB Yes!365 건강보험(세만기)(무배당)(25.01)_0214_요약서_v1.1.pdf"
    OUTPUT_EXCEL_PATH = "/workspaces/automation/tests/test_data/output/extracted_tables.xlsx"
    INITIAL_TERM = "나. 보험금"
    SEARCH_TERMS = [
        "상해관련 특별약관",
        "질병관련 특별약관",
        "상해및질병관련특별약관"
    ]

    logging.info("Starting PDF analysis")
    
    processor = PDFProcessor(INPUT_PDF_PATH)
    if processor.open_pdf_document():
        success = processor.process_pdf(INITIAL_TERM, SEARCH_TERMS, OUTPUT_EXCEL_PATH)
        if success:
            logging.info("PDF analysis completed successfully")
        else:
            logging.error("PDF analysis failed")
    else:
        logging.error("Failed to open PDF document")
    
    processor.close_pdf_document()

if __name__ == "__main__":
    main()