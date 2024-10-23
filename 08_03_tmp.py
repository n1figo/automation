import PyPDF2
import re
import logging
import fitz
import numpy as np
from typing import Dict, List, Tuple, Optional
import os
import pandas as pd
import camelot
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFStructureAnalyzer:
    def __init__(self):
        self.font_size_threshold = 10
        self.title_max_length = 50

    def analyze_page_structure(self, page) -> Tuple[List[Dict], List[Dict]]:
        """페이지의 표와 제목 구조 분석"""
        blocks = page.get_text("rawdict")["blocks"]
        tables = []
        titles = []

        for block in blocks:
            if block['type'] == 0:  # 텍스트 블록
                text = ' '.join(span['text'] for line in block['lines'] 
                              for span in line['spans'])
                font_info = block['lines'][0]['spans'][0]
                
                if self.is_table(text, block):
                    tables.append({
                        'text': text,
                        'bbox': block['bbox'],
                        'font_info': font_info
                    })
                elif self.is_title(text, font_info):
                    titles.append({
                        'text': text,
                        'bbox': block['bbox'],
                        'font_info': font_info
                    })

        return tables, titles

    def is_table(self, text: str, block: Dict) -> bool:
        """표 여부 확인"""
        return ("특약" in text and 
                len(block['lines']) > 1 and
                any('보장' in line['spans'][0]['text'] for line in block['lines']))

    def is_title(self, text: str, font_info: Dict) -> bool:
        """제목 여부 확인"""
        return (font_info['size'] > self.font_size_threshold and
                len(text) < self.title_max_length and
                not any(char in text for char in '※►▶') and
                ('bold' in font_info['font'].lower() or
                 font_info['size'] >= 12))

class TableExtractor:
    def __init__(self):
        self.structure_analyzer = PDFStructureAnalyzer()

    def extract_tables_from_page(self, pdf_path: str, page_num: int) -> List[Tuple[str, pd.DataFrame]]:
        """페이지에서 표와 관련 제목 추출"""
        try:
            # PDF 페이지 열기
            doc = fitz.open(pdf_path)
            page = doc[page_num]

            # 페이지 구조 분석
            tables, titles = self.structure_analyzer.analyze_page_structure(page)

            # Camelot으로 표 추출
            camelot_tables = self.extract_with_camelot(pdf_path, page_num + 1)
            results = []

            for idx, camelot_table in enumerate(camelot_tables):
                if idx < len(tables):  # 구조 분석에서 찾은 표와 매칭
                    table_info = tables[idx]
                    title = self.find_closest_title(table_info, titles)
                    
                    df = camelot_table.df
                    df = self.clean_table(df)
                    
                    if not df.empty:
                        results.append((
                            title['text'] if title else "Untitled Table",
                            df
                        ))

            doc.close()
            return results

        except Exception as e:
            logger.error(f"Error extracting tables: {str(e)}")
            return []

    def extract_with_camelot(self, pdf_path: str, page_num: int) -> List:
        """Camelot을 사용한 표 추출"""
        try:
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_num),
                flavor='lattice'
            )
            if not tables:
                tables = camelot.read_pdf(
                    pdf_path,
                    pages=str(page_num),
                    flavor='stream'
                )
            return tables
        except Exception as e:
            logger.error(f"Camelot extraction failed: {str(e)}")
            return []

    def find_closest_title(self, table: Dict, titles: List[Dict]) -> Optional[Dict]:
        """표와 가장 가까운 제목 찾기"""
        table_top = table['bbox'][1]
        above_titles = [t for t in titles if t['bbox'][3] < table_top]
        
        if not above_titles:
            return None
            
        return min(above_titles, key=lambda t: table_top - t['bbox'][3])

    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """표 데이터 정제"""
        df = df.dropna(how='all')
        df = df[~df.iloc[:, 0].str.contains("※|주)", regex=False, na=False)]
        return df

class ExcelWriter:
    @staticmethod
    def save_to_excel(data: List[Tuple[str, pd.DataFrame]], output_path: str):
        """결과를 Excel 파일로 저장"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                current_row = 0
                sheet_name = "추출된 표"
                
                for title, df in data:
                    # 제목 쓰기
                    df_title = pd.DataFrame([[title]], columns=[''])
                    df_title.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startrow=current_row,
                        index=False,
                        header=False
                    )
                    
                    # 표 데이터 쓰기
                    df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startrow=current_row + 2,  # 제목 다음에 빈 행 하나 추가
                        index=False
                    )
                    
                    current_row += len(df) + 4  # 다음 표를 위한 공간

                # 스타일 적용
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # 제목 셀 스타일링
                title_fill = PatternFill(start_color='E6E6E6',
                                       end_color='E6E6E6',
                                       fill_type='solid')
                                       
                for row in range(1, worksheet.max_row + 1, len(df) + 4):
                    cell = worksheet.cell(row=row, column=1)
                    cell.fill = title_fill

            logger.info(f"Successfully saved tables to {output_path}")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")

def main():
    try:
        # 파일 경로 설정
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        output_path = "특약표_extracted.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return

        # 표 추출
        table_extractor = TableExtractor()
        all_tables = []
        
        # PDF 페이지 수 확인
        with fitz.open(pdf_path) as doc:
            num_pages = len(doc)
            
            # 각 페이지에서 표 추출
            for page_num in range(num_pages):
                logger.info(f"Processing page {page_num + 1}/{num_pages}")
                tables = table_extractor.extract_tables_from_page(pdf_path, page_num)
                all_tables.extend(tables)

        # 결과 저장
        if all_tables:
            ExcelWriter.save_to_excel(all_tables, output_path)
        else:
            logger.error("No tables extracted")

    except Exception as e:
        logger.error(f"Processing error: {str(e)}")

if __name__ == "__main__":
    main()