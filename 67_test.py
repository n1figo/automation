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
    
    # 강조 색상에 대한 조건을 설정 (예: 연한 복숭아색)
    lower_hsv = np.array([0, 20, 200])
    upper_hsv = np.array([25, 150, 255])
    
    mask = cv2.inRange(hsv, lower_hsv, upper_hsv)
    
    kernel = np.ones((5,5), np.uint8)
    cleaned_mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
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

def extract_tables_from_page(page):
    tables = []
    blocks = page.get_text("blocks")
    current_table = []
    
    for block in blocks:
        text = block[4]
        if any(header in text for header in TARGET_HEADERS):
            if current_table:
                tables.append(current_table)
            current_table = [text]
        elif current_table:
            current_table.append(text)
    
    if current_table:
        tables.append(current_table)
    
    return tables

def parse_table(table):
    rows = []
    header = table[0].split('\n')
    rows.append(header)
    
    for cell in table[1:]:
        cell_data = cell.split('\n')
        rows.append(cell_data)
    
    # 모든 행의 길이를 6으로 맞춥니다.
    max_cols = 6
    padded_rows = [row + [''] * (max_cols - len(row)) for row in rows]
    
    # DataFrame을 생성할 때 열 이름을 명시적으로 지정합니다.
    df = pd.DataFrame(padded_rows[1:], columns=['col1', 'col2', 'col3', 'col4', 'col5', 'col6'])
    
    # 원래의 헤더 정보를 사용하여 열 이름을 재지정합니다.
    for i, header in enumerate(padded_rows[0]):
        if header in TARGET_HEADERS:
            df = df.rename(columns={f'col{i+1}': f'{header}1', f'col{i+2}': f'{header}2'})
            i += 1  # 다음 열을 건너뜁니다.
        else:
            df = df.rename(columns={f'col{i+1}': header})
    
    return df

def extract_target_tables_from_page(page, image, page_number):
    print(f"페이지 {page_number + 1} 처리 중...")
    
    tables = extract_tables_from_page(page)
    print(f"페이지 {page_number + 1}에서 찾은 테이블 수: {len(tables)}")
    
    if not tables:
        print(f"페이지 {page_number + 1}에서 테이블을 찾을 수 없습니다.")
        return []
    
    # 페이지에서 강조 색상 감지
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])
    
    # 디버깅: 강조 영역이 표시된 이미지 저장
    debug_image = image.copy()
    for start_y, end_y in highlight_regions:
        cv2.rectangle(debug_image, (0, start_y), (image.shape[1], end_y), (255, 0, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlights.png'), cv2.cvtColor(debug_image, cv2.COLOR_RGB2BGR))
    
    table_data = []
    for table_index, table in enumerate(tables):
        print(f"테이블 {table_index + 1} 처리 중...")
        df = parse_table(table)
        
        if df.empty:
            continue
        
        # 지정된 헤더가 포함된 테이블만 처리
        if all(any(target_header in col for col in df.columns) for target_header in TARGET_HEADERS):
            for row_index, row in df.iterrows():
                row_data = {}
                change_detected = False
                
                # 강조 영역이 존재하면 해당 행을 "추가"로 간주
                if highlight_regions:
                    change_detected = True
                
                for header in TARGET_HEADERS:
                    if f'{header}1' in df.columns:
                        row_data[f'{header}1'] = clean_text_for_excel(str(row[f'{header}1']).strip())
                        row_data[f'{header}2'] = clean_text_for_excel(str(row[f'{header}2']).strip())
                    elif header in df.columns:
                        row_data[f'{header}1'] = clean_text_for_excel(str(row[header]).strip())
                        row_data[f'{header}2'] = ''
                
                if row_data:
                    row_data["페이지"] = page_number + 1
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    table_data.append(row_data)
                    
                    # 디버깅: 각 행의 데이터와 변경 사항 출력
                    print(f"행 {row_index + 1}: 변경사항 = {row_data['변경사항']}, 데이터 = {row_data}")
    
    return table_data

def save_revisions_to_excel(df, output_excel_path):
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    doc = fitz.open(pdf_path)
    
    # 51페이지만 처리 (0-based index이므로 50)
    page_number = 50
    if page_number >= len(doc):
        print(f"PDF에 51 페이지가 존재하지 않습니다.")
        return
    
    page = doc[page_number]
    image = pdf_to_image(page)
    
    # 원본 이미지 저장 (디버깅용)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
    
    # 페이지에서 표 추출 및 강조 영역 분석
    table_data = extract_target_tables_from_page(page, image, page_number)
    
    if table_data:
        df = pd.DataFrame(table_data)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액1", "지급금액2"]:
            if col not in df.columns:
                df[col] = ''
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액1", "지급금액2", "변경사항"]]
        save_revisions_to_excel(df, output_excel_path)
        
        # 디버깅: DataFrame 내용 출력
        if DEBUG_MODE:
            print("\nDataFrame 내용:")
            print(df)
    else:
        print("지정된 헤더를 가진 표를 찾을 수 없습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)