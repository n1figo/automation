import fitz
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import logging
from typing import List, Tuple, Dict, Any
from collections import defaultdict

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFTableExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        
    def extract_tables_with_titles(self) -> List[Tuple[str, pd.DataFrame]]:
        all_tables = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            tables = self._extract_tables_from_page(page)
            titled_tables = self._assign_titles_to_tables(page, tables)
            all_tables.extend(titled_tables)
        return self._merge_tables_with_same_title(all_tables)
    
    def _extract_tables_from_page(self, page: fitz.Page) -> List[fitz.TableFinder]:
        tables = page.find_tables()
        return tables
    
    def _assign_titles_to_tables(self, page: fitz.Page, tables: List[fitz.TableFinder]) -> List[Tuple[str, pd.DataFrame]]:
        titled_tables = []
        for table in tables:
            title = self._find_table_title(page, table)
            df = self._table_to_dataframe(table)
            titled_tables.append((title, df))
        return titled_tables
    
    def _find_table_title(self, page: fitz.Page, table: fitz.TableFinder) -> str:
        # 표 위의 텍스트 블록을 찾아 제목으로 사용
        text_blocks = page.get_text("blocks")
        table_top = table.bbox.y0
        potential_titles = [block[4] for block in text_blocks if block[3] < table_top and block[3] > table_top - 50]
        
        if potential_titles:
            return potential_titles[-1].strip()
        return "Untitled Table"
    
    def _table_to_dataframe(self, table: fitz.TableFinder) -> pd.DataFrame:
        return table.to_pandas()
    
    def _merge_tables_with_same_title(self, tables: List[Tuple[str, pd.DataFrame]]) -> List[Tuple[str, pd.DataFrame]]:
        merged_tables = defaultdict(list)
        for title, df in tables:
            merged_tables[title].append(df)
        
        return [(title, pd.concat(dfs, ignore_index=True)) for title, dfs in merged_tables.items()]

class ExcelWriter:
    def __init__(self, output_path: str):
        self.output_path = output_path
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self.sheet.title = "Extracted Tables"
        
    def write_tables(self, tables: List[Tuple[str, pd.DataFrame]]):
        row = 1
        for title, df in tables:
            self.sheet.cell(row=row, column=1, value=title)
            row += 1
            
            for r in dataframe_to_rows(df, index=False, header=True):
                self.sheet.append(r)
            row += len(df) + 2
        
        self.workbook.save(self.output_path)
        logger.info(f"Tables saved to {self.output_path}")

def main(pdf_path: str, output_excel_path: str):
    try:
        extractor = PDFTableExtractor(pdf_path)
        tables = extractor.extract_tables_with_titles()
        
        writer = ExcelWriter(output_excel_path)
        writer.write_tables(tables)
        
        logger.info("Table extraction and writing completed successfully.")
    except Exception as e:
        logger.error(f"An error occurred: {str(e)}", exc_info=True)


if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/path/to/your/output/excel/file.xlsx"
    main(pdf_path, output_excel_path)