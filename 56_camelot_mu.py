import camelot
import pandas as pd
import numpy as np
import cv2
import os
import fitz
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from llama_parse import LlamaParse

DEBUG_MODE = True
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def extract_structure_with_llama_parse(pdf_path):
    parser = LlamaParse()
    result = parser.parse(pdf_path)
    return result

def extract_tables_with_camelot(pdf_path, page_number):
    print(f"Extracting tables from page {page_number} using Camelot...")
    tables = camelot.read_pdf(pdf_path, pages=str(page_number), flavor='lattice')
    print(f"Found {len(tables)} tables on page {page_number}")
    return tables

def get_highlighted_areas(llama_result, page_number):
    # LLaMA Parse 결과에서 강조 영역 추출 (실제 구현은 LLaMA Parse의 API에 따라 달라질 수 있음)
    highlighted_areas = llama_result.get_highlighted_areas(page_number)
    return highlighted_areas

def is_overlapping(bbox1, bbox2):
    x1, y1, x2, y2 = bbox1
    x3, y3, x4, y4 = bbox2
    return not (x2 < x3 or x1 > x4 or y2 < y3 or y1 > y4)

def process_tables(tables, highlighted_areas):
    processed_data = []
    for i, table in enumerate(tables):
        df = table.df
        for row_index, row in df.iterrows():
            row_data = row.copy()
            row_bbox = table.cells[row_index][0].bbox  # 행의 bounding box
            row_highlighted = any(is_overlapping(row_bbox, area) for area in highlighted_areas)
            row_data["변경사항"] = "추가" if row_highlighted else ""
            row_data["Table_Number"] = i + 1
            processed_data.append(row_data)
    return pd.DataFrame(processed_data)

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

def visualize_results(image, highlighted_areas, tables, output_path):
    viz_image = image.copy()
    # 강조 영역 시각화
    for area in highlighted_areas:
        x1, y1, x2, y2 = area
        cv2.rectangle(viz_image, (int(x1), int(y1)), (int(x2), int(y2)), (0, 255, 0), 2)
    
    # 테이블 영역 시각화
    for table in tables:
        x1, y1, x2, y2 = table._bbox
        cv2.rectangle(viz_image, (int(x1), int(y1)), (int(x2), int(y2)), (255, 0, 0), 2)
    
    cv2.imwrite(output_path, cv2.cvtColor(viz_image, cv2.COLOR_RGB2BGR))
    print(f"Visualization saved to {output_path}")

def main(pdf_path, output_excel_path):
    print("Extracting structure and tables from PDF...")

    # LLaMA Parse를 사용하여 PDF 구조 분석
    llama_result = extract_structure_with_llama_parse(pdf_path)

    page_number = 50  # 51페이지 (0-based index)

    # PyMuPDF를 사용하여 이미지 추출 (시각화용)
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    image = pdf_to_image(page)

    # 강조 영역 추출
    highlighted_areas = get_highlighted_areas(llama_result, page_number + 1)

    # Camelot을 사용하여 표 추출
    tables = extract_tables_with_camelot(pdf_path, page_number + 1)

    # 추출된 표 처리
    processed_df = process_tables(tables, highlighted_areas)

    print(processed_df)

    # 결과 시각화
    viz_output_path = os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_visualization.png')
    visualize_results(image, highlighted_areas, tables, viz_output_path)

    # 엑셀로 저장 (하이라이트 포함)
    save_to_excel_with_highlight(processed_df, output_excel_path)

    print(f"처리된 데이터가 {output_excel_path}에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables_llama_camelot.xlsx"
    main(pdf_path, output_excel_path)