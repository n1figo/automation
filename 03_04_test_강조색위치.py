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
from openpyxl.styles import PatternFill, Alignment, Font
import cv2
from PIL import Image

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class HighlightDetector:
    def __init__(self):
        self.saturation_threshold = 30
        self.kernel_size = (5, 5)
        
    def pdf_to_image(self, page: fitz.Page) -> np.ndarray:
        """PDF 페이지를 이미지로 변환"""
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return np.array(img)

    def detect_highlights(self, image: np.ndarray) -> List[Tuple[float, float, float, float]]:
        """이미지에서 형광펜 영역 감지 및 정확한 좌표 반환"""
        hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
        
        # 채도와 명도를 이용한 마스크 생성
        saturation_mask = hsv[:,:,1] > self.saturation_threshold
        _, value_mask = cv2.threshold(hsv[:,:,2], 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # 마스크 결합 및 노이즈 제거
        combined_mask = cv2.bitwise_and(value_mask, value_mask, 
                                      mask=saturation_mask.astype(np.uint8) * 255)
        kernel = np.ones(self.kernel_size, np.uint8)
        cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
        cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)
        
        # 컨투어 찾기 및 좌표 반환
        contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        highlight_regions = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            # 비율 기반 좌표 계산 (0~1 사이 값)
            x_ratio = x / image.shape[1]
            y_ratio = y / image.shape[0]
            w_ratio = w / image.shape[1]
            h_ratio = h / image.shape[0]
            highlight_regions.append((x_ratio, y_ratio, w_ratio, h_ratio))
            
        return highlight_regions

class TableProcessor:
    def __init__(self):
        self.highlight_detector = HighlightDetector()
        
    def get_table_cells_coordinates(self, table) -> List[List[Tuple[float, float, float, float]]]:
        """테이블의 각 셀의 상대 좌표 계산"""
        cells_coordinates = []
        bbox = table._bbox
        table_width = bbox[2] - bbox[0]
        table_height = bbox[3] - bbox[1]
        
        for row_idx, row in enumerate(table.cells):
            row_coords = []
            for cell_idx, cell in enumerate(row):
                # 셀의 상대 좌표 계산
                x1_ratio = (cell.x1 - bbox[0]) / table_width
                y1_ratio = (cell.y1 - bbox[1]) / table_height
                x2_ratio = (cell.x2 - bbox[0]) / table_width
                y2_ratio = (cell.y2 - bbox[1]) / table_height
                row_coords.append((x1_ratio, y1_ratio, x2_ratio, y2_ratio))
            cells_coordinates.append(row_coords)
            
        return cells_coordinates
        
    def check_cell_highlight(self, cell_coords: Tuple[float, float, float, float], 
                           highlight_regions: List[Tuple[float, float, float, float]]) -> bool:
        """셀과 하이라이트 영역의 겹침 여부 확인"""
        cell_x1, cell_y1, cell_x2, cell_y2 = cell_coords
        
        for h_x, h_y, h_w, h_h in highlight_regions:
            h_x2 = h_x + h_w
            h_y2 = h_y + h_h
            
            # 영역 겹침 확인
            if (cell_x1 < h_x2 and h_x < cell_x2 and
                cell_y1 < h_y2 and h_y < cell_y2):
                return True
                    
        return False

    def process_table(self, table, page: fitz.Page) -> pd.DataFrame:
        """테이블 처리 및 하이라이트 정보 추가"""
        # 페이지 이미지 변환 및 하이라이트 감지
        image = self.highlight_detector.pdf_to_image(page)
        highlight_regions = self.highlight_detector.detect_highlights(image)
        
        # 테이블 셀 좌표 획득
        cells_coordinates = self.get_table_cells_coordinates(table)
        
        # DataFrame 생성 및 하이라이트 정보 추가
        df = table.df.copy()
        df['변경사항'] = ''
        
        # 각 셀별 하이라이트 확인
        for row_idx, row_coords in enumerate(cells_coordinates):
            for cell_idx, cell_coords in enumerate(row_coords):
                if self.check_cell_highlight(cell_coords, highlight_regions):
                    df.iloc[row_idx, -1] = '추가'
                    break
        
        return df

class PDFTableExtractor:
    def __init__(self):
        self.table_processor = TableProcessor()
        
    def extract_tables(self, pdf_path: str, page_num: int) -> List[pd.DataFrame]:
        """페이지에서 표 추출 및 처리"""
        try:
            # Camelot으로 표 추출
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
            
            if not tables:
                return []
                
            # PDF 페이지 객체 얻기
            doc = fitz.open(pdf_path)
            page = doc[page_num - 1]
            
            # 각 표 처리
            processed_tables = []
            for table in tables:
                df = self.table_processor.process_table(table, page)
                if not df.empty:
                    processed_tables.append(df)
            
            doc.close()
            return processed_tables
            
        except Exception as e:
            logger.error(f"Table extraction error on page {page_num}: {str(e)}")
            return []

def main():
    # 설정
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    output_path = "path/to/output.xlsx"
    
    # PDF 테이블 추출기 초기화
    extractor = PDFTableExtractor()
    
    try:
        # PDF 파일 열기 및 페이지 수 확인
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        doc.close()
        
        # 각 페이지에서 표 추출
        all_tables = []
        for page_num in range(1, total_pages + 1):
            logger.info(f"Processing page {page_num}/{total_pages}")
            tables = extractor.extract_tables(pdf_path, page_num)
            all_tables.extend((table, page_num) for table in tables)
        
        # 결과를 Excel 파일로 저장
        if all_tables:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, (df, page_num) in enumerate(all_tables):
                    sheet_name = f'Table_{i+1}'
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 하이라이트 스타일 적용
                    worksheet = writer.sheets[sheet_name]
                    yellow_fill = PatternFill(start_color='FFFF00',
                                            end_color='FFFF00',
                                            fill_type='solid')
                    
                    for row_idx, row in enumerate(df.itertuples(), start=2):
                        if getattr(row, '변경사항') == '추가':
                            for col in range(1, len(df.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col)
                                cell.fill = yellow_fill
            
            logger.info(f"Successfully saved tables to {output_path}")
        else:
            logger.warning("No tables found in the PDF")
            
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        raise

if __name__ == "__main__":
    main()