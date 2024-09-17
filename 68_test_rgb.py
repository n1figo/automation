import fitz  # PyMuPDF
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
import re
import os
from collections import Counter

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 특정 색상 범위 (흰색, 검정색, 회색을 제외하는 범위)
WHITE_COLOR = (255, 255, 255)
BLACK_COLOR = (0, 0, 0)
GRAY_COLOR = (200, 200, 200)  # 회색 범위를 넓힘

# 강조 색상 정의 (살색과 주황색)
HIGHLIGHT_COLORS = {
    "살색": (255, 218, 185),
    "주황색": (255, 165, 0)
}

# 허용되지 않는 문자를 제거하는 함수
def remove_illegal_characters(text):
    ILLEGAL_CHARACTERS_RE = re.compile(
        '['
        '\x00-\x08'
        '\x0B-\x0C'
        '\x0E-\x1F'
        ']'
    )
    return ILLEGAL_CHARACTERS_RE.sub('', text)

# 텍스트를 엑셀에 맞게 정리 (줄바꿈 유지)
def clean_text_for_excel(text: str) -> str:
    if isinstance(text, str):
        text = remove_illegal_characters(text)
        return text  # 줄바꿈을 제거하지 않음
    return text

# 색상 유사도 비교 함수
def is_similar_color(c1, c2, tolerance=30):
    return all(abs(c1[i] - c2[i]) <= tolerance for i in range(3))

# 강조 색상 감지 함수
def detect_highlight_color(color):
    for name, highlight_color in HIGHLIGHT_COLORS.items():
        if is_similar_color(color, highlight_color, tolerance=50):
            return True, name
    return False, None

# 이미지에서 가장 빈번한 강조 색상 감지 함수
def detect_most_common_highlight_color(image):
    img_array = np.array(image)
    height, width, _ = img_array.shape
    
    highlight_colors = []
    for y in range(height):
        for x in range(width):
            pixel_color = tuple(img_array[y, x])
            is_highlight, color_name = detect_highlight_color(pixel_color)
            if is_highlight:
                highlight_colors.append(color_name)
    
    if not highlight_colors:
        return False, None
    
    color_counts = Counter(highlight_colors)
    most_common_color = color_counts.most_common(1)[0][0]
    
    return True, most_common_color

# 페이지에서 타겟 표를 추출하는 함수
def extract_target_tables_from_page(page, page_image, page_number):
    print(f"페이지 {page_number + 1} 처리 중...")
    tables = page.find_tables()
    print(f"페이지 {page_number + 1}에서 찾은 테이블 수: {len(tables.tables)}")
    
    table_data = []
    for table_index, table in enumerate(tables.tables):
        print(f"테이블 {table_index + 1} 처리 중...")
        table_content = table.extract()
        if not table_content:
            continue
        
        header_row = table_content[0]
        header_texts = [clean_text_for_excel(cell.strip()) if cell else '' for cell in header_row]
        header_texts_normalized = [text.replace(" ", "").replace("\n", "") for text in header_texts]
        
        if all(any(target_header == header_cell for header_cell in header_texts_normalized) for target_header in TARGET_HEADERS):
            cell_dict = {}
            for cell in table.cells:
                if len(cell) == 6:
                    cell_row = cell[0]
                    cell_col = cell[1]
                    cell_bbox = cell[2:6]
                elif len(cell) >= 3 and isinstance(cell[2], (tuple, list)):
                    cell_row = cell[0]
                    cell_col = cell[1]
                    cell_bbox = cell[2]
                else:
                    continue
                if len(cell_bbox) == 4:
                    cell_rect = fitz.Rect(*cell_bbox)
                else:
                    continue
                cell_dict[(cell_row, cell_col)] = cell_rect
            
            num_rows = len(table_content)
            num_cols = len(header_texts)
            for row_index in range(1, num_rows):
                row = table_content[row_index]
                row_data = {}
                change_detected = False
                cell_bg_color = None
                
                for col_index in range(num_cols):
                    header = header_texts[col_index].replace(" ", "").replace("\n", "")
                    cell_text = clean_text_for_excel(row[col_index].strip()) if row[col_index] else ''
                    if header in TARGET_HEADERS:
                        cell_texts = cell_text.split('\n')
                        if header == '보장명':
                            if len(cell_texts) > 1:
                                row_data['보장명1'] = cell_texts[0]
                                row_data['보장명2'] = cell_texts[1]
                            else:
                                row_data['보장명1'] = cell_text
                                row_data['보장명2'] = ''
                        elif header == '지급사유':
                            if len(cell_texts) > 1:
                                row_data['지급사유1'] = cell_texts[0]
                                row_data['지급사유2'] = cell_texts[1]
                            else:
                                row_data['지급사유1'] = cell_text
                                row_data['지급사유2'] = ''
                        else:
                            row_data[header] = cell_text
                        
                        cell_rect = cell_dict.get((row_index, col_index))
                        if cell_rect:
                            x0, y0, x1, y1 = cell_rect
                            x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
                            cell_image = page_image.crop((x0, y0, x1, y1))
                            
                            color_detected, color_name = detect_most_common_highlight_color(cell_image)
                            if color_detected:
                                change_detected = True
                                cell_bg_color = color_name
                                print(f"페이지 {page_number + 1}, 셀 ({row_index}, {col_index})에서 강조색 감지: {color_name}")
                
                if row_data:
                    row_data["페이지"] = page_number + 1
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    row_data["배경색"] = cell_bg_color if cell_bg_color else ''
                    table_data.append(row_data)
    return table_data

# 메인 함수
def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    doc = fitz.open(pdf_path)
    
    # 51페이지만 처리
    page_number = 50  # 0-based index, so 50 is actually page 51
    page = doc[page_number]
    images = convert_from_path(pdf_path, first_page=page_number+1, last_page=page_number+1, dpi=200, fmt='png')
    page_image = images[0]
    
    # 페이지에서 표 추출 및 셀 배경색 분석
    table_data = extract_target_tables_from_page(page, page_image, page_number)
    
    if table_data:
        df = pd.DataFrame(table_data)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "배경색"]:
            if col not in df.columns:
                df[col] = ''
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항", "배경색"]]
        save_revisions_to_excel(df, output_excel_path)
        print("작업이 완료되었습니다.")
    else:
        print("지정된 헤더를 가진 표를 찾을 수 없습니다.")

# 엑셀 파일로 저장하는 함수
def save_revisions_to_excel(df, output_excel_path):
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)