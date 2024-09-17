import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import re
from PIL import Image
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

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
        return text
    return text

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def detect_highlights(image, page_num):
    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    
    # 노란색 범위 설정 (이 값은 조정이 필요할 수 있습니다)
    lower_yellow = np.array([20, 100, 100])
    upper_yellow = np.array([30, 255, 255])
    
    mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
    
    kernel = np.ones((5,5), np.uint8)
    cleaned_mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)
    
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)
    
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))
    
    return contours

def get_capture_regions(contours, image_height, image_width):
    if not contours:
        return []

    regions = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        regions.append((x, y, x+w, y+h))
    
    return regions

def extract_target_tables_from_page(page, image, page_number):
    print(f"페이지 {page_number + 1} 처리 중...")
    tables = page.find_tables()
    print(f"페이지 {page_number + 1}에서 찾은 테이블 수: {len(tables.tables)}")
    
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])
    
    debug_image = image.copy()
    for x1, y1, x2, y2 in highlight_regions:
        cv2.rectangle(debug_image, (x1, y1), (x2, y2), (255, 0, 0), 2)
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
            for row_index, row in enumerate(table_content[1:], start=1):
                row_data = {}
                change_detected = False
                
                for col_index, header in enumerate(header_texts):
                    cell_text = clean_text_for_excel(row[col_index].strip()) if col_index < len(row) else ''
                    header_normalized = header.replace(" ", "").replace("\n", "")
                    if header_normalized in TARGET_HEADERS:
                        cell_texts = cell_text.split('\n')
                        if header_normalized == '보장명':
                            row_data['보장명1'] = cell_texts[0] if cell_texts else ''
                            row_data['보장명2'] = cell_texts[1] if len(cell_texts) > 1 else ''
                        elif header_normalized == '지급사유':
                            row_data['지급사유1'] = cell_texts[0] if cell_texts else ''
                            row_data['지급사유2'] = cell_texts[1] if len(cell_texts) > 1 else ''
                        elif header_normalized == '지급금액':
                            row_data['지급금액'] = cell_text
                
                # 강조 영역 확인
                cell = table.cells[row_index][0]
                if isinstance(cell, (tuple, list)) and len(cell) > 2:
                    cell_rect = cell[2]
                    if isinstance(cell_rect, fitz.Rect):
                        for x1, y1, x2, y2 in highlight_regions:
                            if (x1 <= cell_rect.x0 <= x2 and y1 <= cell_rect.y0 <= y2) or \
                               (x1 <= cell_rect.x1 <= x2 and y1 <= cell_rect.y1 <= y2):
                                change_detected = True
                                break
                
                if row_data:
                    row_data["페이지"] = page_number + 1
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    table_data.append(row_data)
    
    return table_data

def compare_and_update_excel(excel_path, table_data):
    df = pd.read_excel(excel_path)
    
    vectorizer = TfidfVectorizer()
    df_text = df['보장명1'] + ' ' + df['보장명2'] + ' ' + df['지급사유1'] + ' ' + df['지급사유2'] + ' ' + df['지급금액'].astype(str)
    df_vectors = vectorizer.fit_transform(df_text)
    
    for row in table_data:
        row_text = row['보장명1'] + ' ' + row['보장명2'] + ' ' + row['지급사유1'] + ' ' + row['지급사유2'] + ' ' + str(row['지급금액'])
        row_vector = vectorizer.transform([row_text])
        
        similarities = cosine_similarity(row_vector, df_vectors)[0]
        max_similarity_index = similarities.argmax()
        
        if similarities[max_similarity_index] > 0.8:  # 유사도 임계값
            df.at[max_similarity_index, '변경사항'] = row['변경사항']
    
    df.to_excel(excel_path, index=False)
    print(f"엑셀 파일이 업데이트되었습니다: {excel_path}")

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    doc = fitz.open(pdf_path)
    
    page_number = 50
    page = doc[page_number]
    image = pdf_to_image(page)
    
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
    
    table_data = extract_target_tables_from_page(page, image, page_number)
    
    if table_data:
        df = pd.DataFrame(table_data)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액"]:
            if col not in df.columns:
                df[col] = ''
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항"]]
        df.to_excel(output_excel_path, index=False)
        print(f"초기 엑셀 파일이 저장되었습니다: {output_excel_path}")
        
        compare_and_update_excel(output_excel_path, table_data)
        
        print("작업이 완료되었습니다.")
    else:
        print("지정된 헤더를 가진 표를 찾을 수 없습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)