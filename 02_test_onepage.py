import os
import logging
import fitz
import cv2
import numpy as np
from PIL import Image
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
import time
import gc

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    datefmt='%Y/%m/%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class PDFHighlightAnalyzer:
    def __init__(self):
        """초기화"""
        logger.info("PDFHighlightAnalyzer 초기화 시작")
        try:
            # 색상 범위 정의 (HSV)
            self.color_ranges = {
                'yellow': [(20, 60, 60), (45, 255, 255)],
                'green': [(35, 60, 60), (85, 255, 255)],
                'blue': [(95, 60, 60), (145, 255, 255)]
            }
            logger.info("색상 범위 정의 완료")
            
        except Exception as e:
            logger.error(f"초기화 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def detect_tables(self, image):
        """OpenCV를 사용한 표 감지"""
        try:
            # 그레이스케일 변환 및 이진화
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

            # 수평/수직 선 감지
            horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
            vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,40))

            horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)
            vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel)

            # 표 영역 찾기
            table_mask = cv2.add(horizontal_lines, vertical_lines)
            contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            tables = []
            for cnt in contours:
                x, y, w, h = cv2.boundingRect(cnt)
                if w > 100 and h > 100:  # 최소 크기 필터
                    tables.append({
                        'bbox': (x, y, w, h),
                        'image': image[y:y+h, x:x+w].copy()
                    })

            # 메모리 정리
            del gray, thresh, horizontal_lines, vertical_lines, table_mask
            gc.collect()

            return tables

        except Exception as e:
            logger.error(f"표 감지 중 오류: {str(e)}", exc_info=True)
            raise

    def detect_cells(self, table_img):
        """표 내의 셀 감지"""
        try:
            # 그레이스케일 변환 및 이진화
            gray = cv2.cvtColor(table_img, cv2.COLOR_BGR2GRAY)
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

            # 셀 경계 찾기
            contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            cells = []

            for cnt in contours:
                x, y, w, h = cv2.boundingRect(cnt)
                if w > 20 and h > 20:  # 최소 크기 필터
                    cells.append({
                        'bbox': (x, y, w, h),
                        'image': table_img[y:y+h, x:x+w].copy()
                    })

            # 메모리 정리
            del gray, thresh
            gc.collect()

            return cells

        except Exception as e:
            logger.error(f"셀 감지 중 오류: {str(e)}", exc_info=True)
            raise

    def analyze_color(self, cell_img):
        """셀 이미지의 하이라이트 색상 분석"""
        try:
            # 이미지 크기 제한
            max_size = 256
            height, width = cell_img.shape[:2]
            if width > max_size or height > max_size:
                scale = min(max_size/width, max_size/height)
                cell_img = cv2.resize(cell_img, (int(width*scale), int(height*scale)))

            hsv = cv2.cvtColor(cell_img, cv2.COLOR_BGR2HSV)
            detected_colors = []

            for color_name, (lower, upper) in self.color_ranges.items():
                mask = cv2.inRange(hsv, np.array(lower), np.array(upper))
                pixel_count = np.sum(mask > 0)
                coverage = pixel_count / (mask.shape[0] * mask.shape[1])

                if coverage > 0.25:  # 25% 이상 커버리지
                    detected_colors.append({
                        'color': color_name,
                        'coverage': coverage
                    })

                del mask

            del hsv
            gc.collect()

            if detected_colors:
                strongest_color = max(detected_colors, key=lambda x: x['coverage'])
                return [strongest_color['color']]
            return []

        except Exception as e:
            logger.error(f"색상 분석 중 오류: {str(e)}", exc_info=True)
            raise

    def extract_page_image(self, pdf_path, page_num):
        """PDF에서 특정 페이지 이미지 추출"""
        try:
            doc = fitz.open(pdf_path)
            page = doc[page_num - 1]

            # 이미지 추출
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x 해상도
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()

            # PIL Image를 NumPy 배열로 변환
            img_array = np.array(img)

            # 메모리 정리
            del img, pix
            gc.collect()

            return img_array

        except Exception as e:
            logger.error(f"PDF 페이지 추출 중 오류: {str(e)}", exc_info=True)
            raise

    def process_page(self, pdf_path, page_num):
        """특정 페이지 처리"""
        try:
            start_time = time.time()
            logger.info(f"페이지 {page_num} 처리 시작")

            # 메모리 정리
            gc.collect()

            # 이미지 추출
            image = self.extract_page_image(pdf_path, page_num)
            logger.info("페이지 이미지 추출 완료")

            # 표 감지
            tables = self.detect_tables(image)
            logger.info(f"감지된 표 수: {len(tables)}")

            results = []
            for table_idx, table in enumerate(tables):
                table_img = table['image']
                cells = self.detect_cells(table_img)
                
                table_data = {
                    'table_index': table_idx,
                    'bbox': table['bbox'],
                    'highlighted_cells': []
                }

                for cell_idx, cell in enumerate(cells):
                    colors = self.analyze_color(cell['image'])
                    if colors:
                        cell_info = {
                            'row': cell_idx // 10,  # 임시 행 번호
                            'col': cell_idx % 10,   # 임시 열 번호
                            'bbox': cell['bbox'],
                            'colors': colors
                        }
                        table_data['highlighted_cells'].append(cell_info)

                if table_data['highlighted_cells']:
                    results.append(table_data)

                # 메모리 정리
                del table_img, cells
                gc.collect()

            process_time = time.time() - start_time
            logger.info(f"페이지 {page_num} 처리 완료 (소요시간: {process_time:.2f}초)")
            return results

        except Exception as e:
            logger.error(f"페이지 처리 중 오류: {str(e)}", exc_info=True)
            raise

def save_to_excel(tables_data, output_path):
    """결과를 Excel 파일로 저장"""
    try:
        logger.info(f"Excel 저장 시작: {output_path}")
        
        wb = Workbook()
        wb.remove(wb.active)
        
        for table_data in tables_data:
            if table_data['highlighted_cells']:
                sheet_name = f"Table_{table_data['table_index']}"
                ws = wb.create_sheet(sheet_name)
                
                # 제목 추가
                ws.cell(row=1, column=1, value="하이라이트 분석 결과").font = Font(bold=True, size=12)
                
                # 표 정보 추가
                x, y, w, h = table_data['bbox']
                ws.cell(row=3, column=1, value=f"표 위치: ({x}, {y})")
                ws.cell(row=3, column=2, value=f"크기: {w}x{h}")
                
                # 데이터 헤더
                headers = ['행', '열', '위치', '하이라이트 색상']
                row = 5
                for col, header in enumerate(headers, 1):
                    ws.cell(row=row, column=col, value=header).font = Font(bold=True)
                
                # 데이터 추가
                yellow_fill = PatternFill(start_color='FFFF00', 
                                        end_color='FFFF00', 
                                        fill_type='solid')
                
                for i, cell in enumerate(table_data['highlighted_cells'], 1):
                    row_num = row + i
                    x, y, w, h = cell['bbox']
                    ws.cell(row=row_num, column=1, value=cell['row'])
                    ws.cell(row=row_num, column=2, value=cell['col'])
                    ws.cell(row=row_num, column=3, value=f"({x}, {y})")
                    ws.cell(row=row_num, column=4, value=', '.join(cell['colors']))
                    
                    # 하이라이트된 행 배경색 지정
                    for col in range(1, 5):
                        ws.cell(row=row_num, column=col).fill = yellow_fill
                
                # 열 너비 자동 조정
                for column_cells in ws.columns:
                    length = max(len(str(cell.value) or "") for cell in column_cells)
                    ws.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)
        
        wb.save(output_path)
        logger.info(f"Excel 파일 저장 완료: {output_path}")
        
    except Exception as e:
        logger.error(f"Excel 저장 중 오류: {str(e)}", exc_info=True)
        raise

def main():
    # 설정
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    output_path = "page_59_analysis.xlsx"
    
    try:
        total_start_time = time.time()
        logger.info("프로그램 시작")
        
        # 메모리 정리
        gc.collect()
        
        # 분석기 초기화
        analyzer = PDFHighlightAnalyzer()
        
        # 59페이지 분석
        results = analyzer.process_page(pdf_path, page_num=60)
        
        # 결과 저장
        if results:
            save_to_excel(results, output_path)
            logger.info(f"분석 완료. 결과가 {output_path}에 저장되었습니다.")
        else:
            logger.warning("감지된 하이라이트가 없습니다.")
        
        total_time = time.time() - total_start_time
        logger.info(f"프로그램 종료 (총 소요시간: {total_time:.2f}초)")
        
    except Exception as e:
        logger.error(f"실행 중 오류 발생: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()