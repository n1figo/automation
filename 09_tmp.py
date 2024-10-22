import PyPDF2
import re
import logging
import fitz
import numpy as np
from typing import Dict, List, Tuple, Optional
import os
import pandas as pd
import camelot

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TableExtractor:
    def __init__(self):
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }

    def extract_tables_from_range(self, pdf_path: str, start_page: int, end_page: int) -> pd.DataFrame:
        """
        지정된 페이지 범위에서 표 추출 및 병합
        """
        try:
            logger.info(f"Extracting tables from pages {start_page} to {end_page}")
            print(f"Extracting tables from pages {start_page} to {end_page}")
            
            # 페이지 범위 유효성 검사
            if start_page > end_page or start_page < 1:
                logger.error("Invalid page range specified.")
                print("Invalid page range specified.")
                return pd.DataFrame()

            pages = f"{start_page}-{end_page}"
            print(f"Pages: {pages}")
            
            # Camelot flavor 변경 (lattice -> stream) 시도
            try:
                print(f"Trying Camelot 'lattice' flavor...")
                tables = camelot.read_pdf(
                    pdf_path,
                    pages=pages,
                    flavor='lattice'
                )
                print(f"Lattice mode table count: {len(tables)}")
            except Exception as lattice_error:
                print(f"Lattice mode failed: {str(lattice_error)}, trying 'stream' flavor...")
                tables = camelot.read_pdf(
                    pdf_path,
                    pages=pages,
                    flavor='stream'
                )
                print(f"Stream mode table count: {len(tables)}")
            
            if len(tables) == 0:
                logger.warning(f"No tables found in pages {start_page}-{end_page}")
                print(f"No tables found in pages {start_page}-{end_page}")
                return pd.DataFrame()

            dfs = []
            for idx, table in enumerate(tables):
                df = table.df
                print(f"Table {idx} extracted with shape: {df.shape}")
                # 빈 행이나 불필요한 데이터 제거 (괄호 처리)
                print(f"Cleaning table {idx}...")
                df = df.dropna(how='all')
                # 정규 표현식이 아닌 단순 문자열로 검색하도록 수정 (괄호 문제 해결)
                df = df[~df.iloc[:, 0].str.contains("※|주)", regex=False, na=False)]
                print(f"Table {idx} after cleaning has shape: {df.shape}")
                dfs.append(df)

            merged_df = pd.concat(dfs, ignore_index=True)
            print(f"Merged DataFrame shape: {merged_df.shape}")
            return merged_df

        except Exception as e:
            logger.error(f"Error extracting tables: {str(e)}")
            print(f"Error extracting tables: {str(e)}")
            return pd.DataFrame()

class SectionDetector:
    def __init__(self):
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }

    def find_section_ranges(self, pdf_path: str) -> Dict[str, Tuple[int, int]]:
        """
        각 섹션의 시작과 끝 페이지 찾기
        """
        section_ranges = {}
        injury_start = None
        disease_start = None
        disease_end = None
        
        try:
            print(f"Opening PDF: {pdf_path}")
            with fitz.open(pdf_path) as doc:
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text()
                    print(f"Processing page {page_num + 1}...")

                    # 상해관련 특약 시작 찾기
                    if not injury_start and re.search(self.key_patterns['상해관련'], text, re.IGNORECASE):
                        injury_start = page_num + 1
                        logger.info(f"Found injury section start at page {injury_start}")
                        print(f"Found injury section start at page {injury_start}")
                    
                    # 질병관련 특약 시작 찾기
                    if not disease_start and re.search(self.key_patterns['질병관련'], text, re.IGNORECASE):
                        disease_start = page_num + 1
                        # 상해관련 특약 끝은 질병관련 특약 시작 직전
                        if injury_start:
                            section_ranges['상해관련'] = (injury_start, disease_start - 1)
                        logger.info(f"Found disease section start at page {disease_start}")
                        print(f"Found disease section start at page {disease_start}")
                    
                    # 다음 섹션 시작 찾기 (예: 74페이지)
                    if disease_start and page_num + 1 == 74:
                        disease_end = page_num + 1
                        section_ranges['질병관련'] = (disease_start, disease_end)
                        logger.info(f"Found disease section end at page {disease_end}")
                        print(f"Found disease section end at page {disease_end}")
                        break
            
            print(f"Section ranges: {section_ranges}")
            return section_ranges
            
        except Exception as e:
            logger.error(f"Error finding sections: {str(e)}")
            print(f"Error finding sections: {str(e)}")
            return {}

def process_pdf_and_save_tables(pdf_path: str, output_path: str):
    """
    PDF에서 표를 추출하고 Excel 파일로 저장하는 메인 함수
    """
    try:
        print(f"Processing PDF: {pdf_path}")
        # 섹션 탐지
        section_detector = SectionDetector()
        section_ranges = section_detector.find_section_ranges(pdf_path)
        
        if not section_ranges:
            logger.error("No sections found in PDF")
            print("No sections found in PDF")
            return
            
        # 표 추출
        table_extractor = TableExtractor()
        
        injury_df = pd.DataFrame()
        disease_df = pd.DataFrame()
        
        # 상해관련 특약 표 추출
        injury_range = section_ranges.get('상해관련')
        if injury_range:
            logger.info(f"Extracting injury section tables from pages {injury_range[0]} to {injury_range[1]}")
            print(f"Extracting injury section tables from pages {injury_range[0]} to {injury_range[1]}")
            injury_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                injury_range[0], 
                injury_range[1]
            )
            
        # 질병관련 특약 표 추출
        disease_range = section_ranges.get('질병관련')
        if disease_range:
            logger.info(f"Extracting disease section tables from pages {disease_range[0]} to {disease_range[1]}")
            print(f"Extracting disease section tables from pages {disease_range[0]} to {disease_range[1]}")
            disease_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                disease_range[0], 
                disease_range[1]
            )
        
        # Excel 파일로 저장
        if injury_df.empty and disease_df.empty:
            logger.error("Both DataFrames are empty, no data to save.")
            print("Both DataFrames are empty, no data to save.")
            return

        print(f"Saving data to Excel: {output_path}")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not injury_df.empty:
                print(f"Saving injury section to Excel with shape: {injury_df.shape}")
                injury_df.to_excel(
                    writer, 
                    sheet_name='특약표', 
                    index=False, 
                    startrow=0,
                    startcol=0
                )
                logger.info("Saved injury section tables")
                
            if not disease_df.empty:
                start_row = len(injury_df) + 3 if not injury_df.empty else 0
                print(f"Saving disease section to Excel with shape: {disease_df.shape}, starting at row: {start_row}")
                disease_df.to_excel(
                    writer, 
                    sheet_name='특약표',
                    index=False,
                    startrow=start_row,
                    startcol=0
                )
                logger.info("Saved disease section tables")
        
        logger.info(f"Successfully saved tables to {output_path}")
        print(f"Successfully saved tables to {output_path}")
        
    except Exception as e:
        logger.error(f"Error processing PDF and saving tables: {str(e)}")
        print(f"Error processing PDF and saving tables: {str(e)}")

def main():
    try:
        # 파일 경로
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        output_path = "특약표_combined.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            print("PDF file not found")
            return
        
        # PDF 처리 및 표 저장
        process_pdf_and_save_tables(pdf_path, output_path
