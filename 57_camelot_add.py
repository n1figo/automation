import camelot
import pandas as pd
import numpy as np
import cv2
import os
import fitz
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 디버깅 모드 설정
DEBUG_MODE = True

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def detect_highlights(image, page_num):
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # 테이블 헤더(회색) 및 일반 텍스트(검정) 마스크 생성
    header_mask = cv2.inRange(image, (200, 200, 200), (230, 230, 230))
    text_mask = cv2.inRange(image, (0, 0, 0), (50, 50, 50))
    
    # 강조 영역 마스크 (헤더와 일반 텍스트 제외)
    highlight_mask = cv2.bitwise_and(binary, cv2.bitwise_not(cv2.bitwise_or(header_mask, text_mask)))
    
    kernel = np.ones((5,5), np.uint8)
    highlight_mask = cv2.morphologyEx(highlight_mask, cv2.MORPH_CLOSE, kernel)
    highlight_mask = cv2.morphologyEx(highlight_mask, cv2.MORPH_OPEN, kernel)

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), highlight_mask)

    contours, _ = cv2.findContours(highlight_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))

    return contours

def get_highlight_regions(contours, image_height):
    regions = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        regions.append((y, y + h))
    return regions

def extract_tables_with_camelot(pdf_path, page_number):
    print(f"Extracting tables from page {page_number} using Camelot...")
    tables = camelot.read_pdf(pdf_path, pages=str(page_number), flavor='lattice')
    print(f"Found {len(tables)} tables on page {page_number}")
    return tables

def process_tables(tables, highlight_regions, page_height):
    processed_data = []
    for i, table in enumerate(tables):
        df = table.df
        y1, x1, y2, x2 = table._bbox
        
        for row_index in range(len(df)):
            row_data = df.iloc[row_index].copy()
            
            # 행의 상단과 하단 y 좌표 계산 (PDF 좌표계에서 이미지 좌표계로 변환)
            row_top = page_height - (y1 + row_index * (y2 - y1) / len(df))
            row_bottom = page_height - (y1 + (row_index + 1) * (y2 - y1) / len(df))
            
            row_highlighted = check_highlight((row_top, row_bottom), highlight_regions)
            row_data["변경사항"] = "추가" if row_highlighted else ""
            row_data["Table_Number"] = i + 1
            processed_data.append(row_data)

    return pd.DataFrame(processed_data)

def check_highlight(row_range, highlight_regions):
    row_top, row_bottom = row_range
    for region_top, region_bottom in highlight_regions:
        if (region_top <= row_top <= region_bottom) or (region_top <= row_bottom <= region_bottom) or \
           (row_top <= region_top <= row_bottom) or (row_top <= region_bottom <= row_bottom):
            return True
    return False

def save_to_excel_with_highlight(df, output_path):
    df.to_excel(output_path, index=False)
    
    wb = load_workbook(output_path)
    ws = wb.active

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    change_col_index = df.columns.get_loc('변경사항') + 1

    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=change_col_index).value == '추가':
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = yellow_fill

    wb.save(output_path)
    print(f"Data saved to '{output_path}' with highlighted rows")

def main(pdf_path, output_excel_path):
    print("Extracting tables and detecting highlights from PDF...")

    page_number = 50  # 51페이지 (0-based index)

    # PyMuPDF로 PDF 열기
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    image = pdf_to_image(page)

    # 원본 이미지 저장
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    # 강조된 영역 탐지
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_highlight_regions(contours, image.shape[0])

    # 강조 영역이 표시된 이미지 저장
    highlighted_image = image.copy()
    for region in highlight_regions:
        cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))

    print(f"감지된 강조 영역 수: {len(highlight_regions)}")
    print(f"강조 영역: {highlight_regions}")

    # Camelot을 사용하여 표 추출
    tables = extract_tables_with_camelot(pdf_path, page_number + 1)

    # 추출된 표 처리
    processed_df = process_tables(tables, highlight_regions, image.shape[0])

    # 처리된 데이터 출력
    print(processed_df)

    # 엑셀로 저장 (하이라이트 포함)
    save_to_excel_with_highlight(processed_df, output_excel_path)

    print(f"Processed data saved to {output_excel_path}")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables_camelot_highlighted.xlsx"
    main(pdf_path, output_excel_path)