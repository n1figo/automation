import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import re
from PIL import Image

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def remove_illegal_characters(text):
    ILLEGAL_CHARACTERS_RE = re.compile(
        '['
        '\x00-\x08'
        '\x0B-\x0C'
        '\x0E-\x1F'
        ']'
    )
    return ILLEGAL_CHARACTERS_RE.sub('', text)

def clean_text_for_excel(text: str) -> str:
    if isinstance(text, str):
        text = remove_illegal_characters(text)
        return text  # 줄바꿈을 제거하지 않음
    return text

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def detect_highlights(image, page_num):
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
    
    # 디버깅: 마스크 이미지 저장
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)
    
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 디버깅: 윤곽선이 그려진 이미지 저장
    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))
    
    return contours

def get_capture_regions(contours, image_height, image_width):
    if not contours:
        return []

    capture_height = image_height // 3
    sorted_contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])
    
    regions = []
    current_region = None
    
    for contour in sorted_contours:
        x, y, w, h = cv2.boundingRect(contour)
        
        if current_region is None:
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
        elif y - current_region[1] < capture_height//2:
            current_region[1] = min(image_height, y + h + capture_height//2)
        else:
            regions.append(current_region)
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
    
    if current_region:
        regions.append(current_region)
    
    return regions

def extract_target_tables_from_page(page, image, page_number):
    print(f"페이지 {page_number + 1} 처리 중...")
    tables = page.find_tables()
    print(f"페이지 {page_number + 1}에서 찾은 테이블 수: {len(tables.tables)}")
    
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])
    
    # 디버깅: 강조 영역이 표시된 이미지 저장
    debug_image = image.copy()
    for start_y, end_y in highlight_regions:
        cv2.rectangle(debug_image, (0, start_y), (image.shape[1], end_y), (255, 0, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlights.png'), cv2.cvtColor(debug_image, cv2.COLOR_RGB2BGR))
    
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
            num_rows = len(table_content)
            num_cols = len(header_texts)
            for row_index in range(1, num_rows):
                row = table_content[row_index]
                row_data = {}
                change_detected = False
                
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
                
                # 강조 영역 확인
                if len(table.cells) > row_index and len(table.cells[row_index]) > 0:
                    cell = table.cells[row_index][0]
                    if isinstance(cell, (tuple, list)) and len(cell) > 2:
                        cell_rect = cell[2]
                        if isinstance(cell_rect, fitz.Rect):
                            cell_y = cell_rect.y1
                            for start_y, end_y in highlight_regions:
                                if start_y <= cell_y <= end_y:
                                    change_detected = True
                                    break
                
                if row_data:
                    row_data["페이지"] = page_number + 1
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    table_data.append(row_data)
                    
                    # 디버깅: 각 행의 데이터와 변경 사항 출력
                    print(f"행 {row_index}: 변경사항 = {row_data['변경사항']}, 데이터 = {row_data}")
    
    return table_data

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    doc = fitz.open(pdf_path)
    
    # 51페이지만 처리 (0-based index이므로 50)
    page_number = 50
    page = doc[page_number]
    image = pdf_to_image(page)
    
    # 원본 이미지 저장
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
    
    # 페이지에서 표 추출 및 강조 영역 분석
    table_data = extract_target_tables_from_page(page, image, page_number)
    
    if table_data:
        df = pd.DataFrame(table_data)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액"]:
            if col not in df.columns:
                df[col] = ''
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항"]]
        save_revisions_to_excel(df, output_excel_path)
        print("작업이 완료되었습니다.")
        
        # 디버깅: DataFrame 내용 출력
        print("\nDataFrame 내용:")
        print(df)
    else:
        print("지정된 헤더를 가진 표를 찾을 수 없습니다.")

def save_revisions_to_excel(df, output_excel_path):
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)