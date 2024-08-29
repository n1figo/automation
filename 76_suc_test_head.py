import fitz
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import logging
import os
from typing import List, Tuple, Dict, Any
from collections import defaultdict
import re

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_text_for_excel(text):
    if isinstance(text, str):
        # 제어 문자 제거
        text = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text)
        # 줄바꿈을 공백으로 변경
        text = text.replace('\n', ' ').replace('\r', '')
        # 연속된 공백을 하나의 공백으로 변경
        text = re.sub(r'\s+', ' ', text)
        # 특수 문자 제거 또는 변경
        text = text.replace('•', '-').replace('Ⅱ', 'II')
        # 추가적인 특수 문자 처리가 필요한 경우 여기에 추가
    return text

class PDFTableExtractor:
    def __init__(self, pdf_path: str, tessdata_dir: str = None):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        if tessdata_dir:
            os.environ['TESSDATA_PREFIX'] = tessdata_dir
        
    def extract_tables_with_titles(self) -> List[Tuple[str, pd.DataFrame]]:
        all_tables = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            tables = self.extract_tables_from_page(page)
            titled_tables = self._assign_titles_to_tables(page, tables)
            all_tables.extend(titled_tables)
        return self._merge_tables_with_same_title(all_tables)
    
    def extract_tables_from_page(self, page: fitz.Page) -> List[Any]:
        tables = page.find_tables()
        return tables
    
    def _assign_titles_to_tables(self, page: fitz.Page, tables: List[Any]) -> List[Tuple[str, pd.DataFrame]]:
        titled_tables = []
        for table in tables:
            title = self._find_table_title(page, table)
            df = self._table_to_dataframe(table)
            titled_tables.append((title, df))
        return titled_tables
    
    def _find_table_title(self, page: fitz.Page, table: Any) -> str:
        blocks = page.get_text("dict")["blocks"]
        table_top = table.bbox[1]  # y0 좌표
        potential_titles = []
        for b in blocks:
            if 'lines' in b:
                for l in b['lines']:
                    for s in l['spans']:
                        if s['bbox'][3] < table_top and s['bbox'][3] > table_top - 50:
                            potential_titles.append(s['text'])
        
        if potential_titles:
            return " ".join(potential_titles).strip()
        return "Untitled Table"
    
    def _table_to_dataframe(self, table: Any) -> pd.DataFrame:
        df = pd.DataFrame(table.extract())
        df = df.applymap(clean_text_for_excel)
        return df
    
    def _merge_tables_with_same_title(self, tables: List[Tuple[str, pd.DataFrame]]) -> List[Tuple[str, pd.DataFrame]]:
        merged_tables = defaultdict(list)
        for title, df in tables:
            merged_tables[title].append(df)
        
        return [(title, pd.concat(dfs, ignore_index=True)) for title, dfs in merged_tables.items()]
    
    def extract_text_with_ocr(self, page_number: int, language: str = 'eng', dpi: int = 300) -> str:
        page = self.doc[page_number - 1]
        try:
            text = page.get_text()
            if not text.strip():  # 텍스트가 비어있는 경우에만 OCR 실행
                text = page.get_textpage_ocr(flags=3, language=language, dpi=dpi)
            return text
        except RuntimeError as e:
            logger.warning(f"OCR failed: {str(e)}. Falling back to normal text extraction.")
            return page.get_text()

class ExcelWriter:
    def __init__(self, output_path: str):
        self.output_path = output_path
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "Extracted Tables"
        
    def write_tables(self, tables: List[Tuple[str, pd.DataFrame]]):
        row = 1
        for title, df in tables:
            self.sheet.cell(row=row, column=1, value=clean_text_for_excel(title))
            row += 1
            
            for r in dataframe_to_rows(df, index=False, header=True):
                self.sheet.append([clean_text_for_excel(cell) for cell in r])
            row += len(df) + 2
        
        self.workbook.save(self.output_path)
        logger.info(f"Tables saved to {self.output_path}")

def main(pdf_path: str, output_excel_path: str, tessdata_dir: str = None):
    try:
        extractor = PDFTableExtractor(pdf_path, tessdata_dir)
        
        # 테이블 추출
        tables = extractor.extract_tables_with_titles()
        writer = ExcelWriter(output_excel_path)
        writer.write_tables(tables)
        
        # OCR 수행 (예: 첫 번째 페이지)
        ocr_text = extractor.extract_text_with_ocr(1, language='kor+eng')
        print("추출된 텍스트:")
        print(ocr_text)
        
        logger.info("Table extraction, writing, and text extraction completed successfully.")
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}", exc_info=True)

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    tessdata_dir = "/usr/share/tesseract-ocr/4.00/tessdata"  # Tesseract OCR 언어 데이터 파일이 있는 디렉토리
    main(pdf_path, output_excel_path, tessdata_dir)