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
            # camelot으로 표 추출
            tables = camelot.read_pdf(
                pdf_path,
                pages=f"{start_page}-{end_page}",
                flavor='lattice'
            )
            
            if not tables:
                logger.warning(f"No tables found in pages {start_page}-{end_page}")
                return pd.DataFrame()

            # 모든 표 병합
            dfs = []
            for table in tables:
                df = table.df
                # 빈 행이나 불필요한 데이터 제거
                df = df.dropna(how='all')
                df = df[~df.iloc[:,0].str.contains("※|주)", na=False)]
                dfs.append(df)

            # 모든 DataFrame 병합
            merged_df = pd.concat(dfs, ignore_index=True)
            return merged_df

        except Exception as e:
            logger.error(f"Error extracting tables: {str(e)}")
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
            with fitz.open(pdf_path) as doc:
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = page.get_text()
                    
                    # 상해관련 특약 시작 찾기
                    if not injury_start and re.search(self.key_patterns['상해관련'], text, re.IGNORECASE):
                        injury_start = page_num + 1
                        logger.info(f"Found injury section start at page {injury_start}")
                    
                    # 질병관련 특약 시작 찾기
                    if not disease_start and re.search(self.key_patterns['질병관련'], text, re.IGNORECASE):
                        disease_start = page_num + 1
                        # 상해관련 특약 끝은 질병관련 특약 시작 직전
                        if injury_start:
                            section_ranges['상해관련'] = (injury_start, disease_start - 1)
                        logger.info(f"Found disease section start at page {disease_start}")
                    
                    # 다음 섹션 시작 찾기 (74페이지)
                    if disease_start and page_num + 1 == 74:
                        disease_end = page_num + 1
                        section_ranges['질병관련'] = (disease_start, disease_end)
                        logger.info(f"Found disease section end at page {disease_end}")
                        break
            
            return section_ranges
            
        except Exception as e:
            logger.error(f"Error finding sections: {str(e)}")
            return {}

def process_pdf_and_save_tables(pdf_path: str, output_path: str):
    """
    PDF에서 표를 추출하고 Excel 파일로 저장하는 메인 함수
    """
    try:
        # 섹션 탐지
        section_detector = SectionDetector()
        section_ranges = section_detector.find_section_ranges(pdf_path)
        
        if not section_ranges:
            logger.error("No sections found in PDF")
            return
            
        # 표 추출
        table_extractor = TableExtractor()
        
        # 상해관련 특약 표 추출
        injury_range = section_ranges.get('상해관련')
        if injury_range:
            logger.info(f"Extracting injury section tables from pages {injury_range[0]} to {injury_range[1]}")
            injury_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                injury_range[0], 
                injury_range[1]
            )
        else:
            injury_df = pd.DataFrame()
            
        # 질병관련 특약 표 추출
        disease_range = section_ranges.get('질병관련')
        if disease_range:
            logger.info(f"Extracting disease section tables from pages {disease_range[0]} to {disease_range[1]}")
            disease_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                disease_range[0], 
                disease_range[1]
            )
        else:
            disease_df = pd.DataFrame()
            
        # Excel 파일로 저장
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 상해관련 특약 표 저장
            if not injury_df.empty:
                injury_df.to_excel(
                    writer, 
                    sheet_name='특약표', 
                    index=False, 
                    startrow=0,
                    startcol=0
                )
                logger.info("Saved injury section tables")
                
                # 질병관련 특약 표 저장 (상해관련 특약 표 아래에)
                if not disease_df.empty:
                    # 빈 줄 추가를 위한 시작 행 계산
                    start_row = len(injury_df) + 3
                    disease_df.to_excel(
                        writer, 
                        sheet_name='특약표',
                        index=False,
                        startrow=start_row,
                        startcol=0
                    )
                    logger.info("Saved disease section tables")
            
            # 워크시트 가져오기
            worksheet = writer.sheets['특약표']
            
            # 구분선 추가 (상해관련과 질병관련 사이)
            if not injury_df.empty and not disease_df.empty:
                separator_row = len(injury_df) + 2
                for col in range(1, max(len(injury_df.columns), len(disease_df.columns)) + 1):
                    cell = worksheet.cell(row=separator_row, column=col)
                    cell.value = "=" * 50  # 구분선으로 사용할 문자
            
            # 열 너비 자동 조정
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        logger.info(f"Successfully saved tables to {output_path}")
        
    except Exception as e:
        logger.error(f"Error processing PDF and saving tables: {str(e)}")

def main():
    try:
        # 파일 경로
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        output_path = "특약표_combined.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return
        
        # PDF 처리 및 표 저장
        process_pdf_and_save_tables(pdf_path, output_path)
            
    except Exception as e:
        logger.error(f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
