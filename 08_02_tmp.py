import PyPDF2
import re
import logging
import fitz
import numpy as np
from typing import Dict, List, Tuple, Optional
import os
import pandas as pd
import camelot
import cv2
from PIL import Image
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFImageProcessor:
    @staticmethod
    def pdf_to_image(page):
        """PDF 페이지를 이미지로 변환"""
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return np.array(img)

    @staticmethod
    def detect_highlights(image):
        """이미지에서 하이라이트된 영역 감지"""
        hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
        s = hsv[:,:,1]
        v = hsv[:,:,2]

        saturation_threshold = 30
        saturation_mask = s > saturation_threshold

        _, binary = cv2.threshold(v, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        combined_mask = cv2.bitwise_and(binary, binary, mask=saturation_mask.astype(np.uint8) * 255)

        kernel = np.ones((5,5), np.uint8)
        cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
        cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)

        contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        return contours

    @staticmethod
    def get_highlight_regions(contours, image_height):
        """하이라이트된 영역의 좌표 추출"""
        regions = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            top = image_height - (y + h)
            bottom = image_height - y
            regions.append((top, bottom))
        return regions

class TableExtractor:
    def __init__(self):
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }
        self.image_processor = PDFImageProcessor()

    def extract_tables_from_range(self, pdf_path: str, start_page: int, end_page: int) -> List[Tuple[str, pd.DataFrame, List]]:
        """지정된 페이지 범위에서 표 추출 및 하이라이트 정보 포함"""
        try:
            logger.info(f"Extracting tables from pages {start_page} to {end_page}")
            results = []
            
            # PDF 문서 열기
            doc = fitz.open(pdf_path)
            
            for page_num in range(start_page - 1, end_page):
                # 이미지 처리로 하이라이트 영역 감지
                page = doc[page_num]
                image = self.image_processor.pdf_to_image(page)
                contours = self.image_processor.detect_highlights(image)
                highlight_regions = self.image_processor.get_highlight_regions(contours, image.shape[0])
                
                # 표 추출
                try:
                    tables = camelot.read_pdf(
                        pdf_path,
                        pages=str(page_num + 1),
                        flavor='lattice'
                    )
                except Exception:
                    try:
                        tables = camelot.read_pdf(
                            pdf_path,
                            pages=str(page_num + 1),
                            flavor='stream'
                        )
                    except Exception as e:
                        logger.error(f"Failed to extract tables from page {page_num + 1}: {str(e)}")
                        continue

                # 각 표 처리
                for table in tables:
                    df = table.df
                    df = self.clean_table(df)
                    
                    if not df.empty:
                        title = self.extract_table_title(page)
                        results.append((title, df, highlight_regions))

            doc.close()
            return results

        except Exception as e:
            logger.error(f"Error extracting tables: {str(e)}")
            return []

    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """표 데이터 정제"""
        df = df.dropna(how='all')
        df = df[~df.iloc[:, 0].str.contains("※|주)", regex=False, na=False)]
        return df

    def extract_table_title(self, page) -> str:
        """페이지에서 표 제목 추출"""
        text = page.get_text("text")
        lines = text.split('\n')
        for line in lines:
            if line.strip():
                return line.strip()
        return "Untitled Table"

class SectionDetector:
    def __init__(self):
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }

    def find_section_ranges(self, pdf_path: str) -> Dict[str, Tuple[int, int]]:
        """문서에서 섹션 범위 찾기"""
        section_ranges = {}
        
        try:
            with fitz.open(pdf_path) as doc:
                injury_start = None
                disease_start = None
                
                for page_num in range(len(doc)):
                    text = doc[page_num].get_text()
                    
                    # 상해관련 특약 시작점 찾기
                    if not injury_start and re.search(self.key_patterns['상해관련'], text, re.IGNORECASE):
                        injury_start = page_num + 1
                    
                    # 질병관련 특약 시작점 찾기
                    if not disease_start and re.search(self.key_patterns['질병관련'], text, re.IGNORECASE):
                        disease_start = page_num + 1
                        if injury_start:
                            section_ranges['상해관련'] = (injury_start, disease_start - 1)
                
                # 마지막 섹션 끝점 설정
                if disease_start:
                    section_ranges['질병관련'] = (disease_start, len(doc))
                elif injury_start:
                    section_ranges['상해관련'] = (injury_start, len(doc))
                    
            return section_ranges
            
        except Exception as e:
            logger.error(f"Error finding sections: {str(e)}")
            return {}

class ExcelWriter:
    @staticmethod
    def save_to_excel(data: Dict[str, List[Tuple[str, pd.DataFrame, List]]], output_path: str):
        """결과를 Excel 파일로 저장"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for section, tables in data.items():
                    current_row = 0
                    
                    for title, df, highlights in tables:
                        # 표 제목 쓰기
                        df.to_excel(writer, sheet_name=section, startrow=current_row + 1, index=False)
                        worksheet = writer.sheets[section]
                        worksheet.cell(row=current_row + 1, column=1, value=title)
                        
                        # 하이라이트된 행 표시
                        for row_idx in range(len(df)):
                            for col_idx in range(len(df.columns)):
                                cell = worksheet.cell(row=current_row + row_idx + 2, column=col_idx + 1)
                                if any(top <= row_idx <= bottom for top, bottom in highlights):
                                    cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        
                        current_row += len(df) + 3

            logger.info(f"Successfully saved tables to {output_path}")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")

def main():
    try:
        # 파일 경로 설정
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        output_path = "특약표_enhanced.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return
        
        # 섹션 감지
        section_detector = SectionDetector()
        section_ranges = section_detector.find_section_ranges(pdf_path)
        
        if not section_ranges:
            logger.error("No sections found in PDF")
            return
        
        # 표 추출
        table_extractor = TableExtractor()
        extracted_data = {}
        
        for section, (start, end) in section_ranges.items():
            tables = table_extractor.extract_tables_from_range(pdf_path, start, end)
            if tables:
                extracted_data[section] = tables
        
        # Excel 파일로 저장
        if extracted_data:
            ExcelWriter.save_to_excel(extracted_data, output_path)
        else:
            logger.error("No tables extracted")

    except Exception as e:
        logger.error(f"Processing error: {str(e)}")

if __name__ == "__main__":
    main()