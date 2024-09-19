import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
import logging

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

# 하이라이트 영역 탐지
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

    logging.info(f"하이라이트 영역 탐지 완료: {len(contours)}개의 윤곽선 발견")
    return contours

# 강조된 영역 반환
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

    logging.info(f"강조 영역 추출 완료: {len(regions)}개의 영역 발견")
    for i, region in enumerate(regions):
        logging.debug(f"강조 영역 {i+1}: y={region[0]} to y={region[1]}")
    return regions

def merge_split_rows(df):
    merged_rows = []
    current_row = {}
    main_columns = ["특약명칭(번호)", "보장명"]  # 주요 열 목록

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        
        # 새로운 행의 시작인지 확인
        if any(row_dict.get(col) for col in main_columns):
            if current_row:
                merged_rows.append(current_row)
            current_row = row_dict
        else:
            # 현재 행의 비어있지 않은 값으로 current_row 업데이트
            for key, value in row_dict.items():
                if pd.notna(value) and (key not in current_row or pd.isna(current_row[key])):
                    current_row[key] = value

    if current_row:
        merged_rows.append(current_row)

    return pd.DataFrame(merged_rows)

# PDF에서 테이블을 추출하고 강조된 행을 찾는 함수
def extract_and_process_tables(doc, page_number, highlight_regions):
    page = doc[page_number]
    table_finder = page.find_tables()
    tables = table_finder.tables
    
    logging.info(f"페이지 {page_number+1}에서 {len(tables)}개의 테이블 발견")
    
    processed_data = []

    if not tables:
        logging.warning(f"페이지 {page_number+1}에서 테이블을 찾을 수 없습니다.")
        return pd.DataFrame()

    for table_index, table in enumerate(tables):
        cells = table.extract()
        if not cells:
            logging.warning(f"Table {table_index + 1}에서 셀을 추출할 수 없습니다.")
            continue

        # 열 이름 중복 처리
        columns = cells[0]
        unique_columns = []
        for i, col in enumerate(columns):
            if col in unique_columns:
                unique_columns.append(f"{col}_{i}")
            else:
                unique_columns.append(col)

        df = pd.DataFrame(cells[1:], columns=unique_columns)
        
        # 행 병합 로직 적용
        df = merge_split_rows(df)

        table_bbox = table.bbox
        x0, y0, x1, y1 = table_bbox

        logging.info(f"Table {table_index + 1} 위치: (x0={x0}, y0={y0}, x1={x1}, y1={y1})")
        logging.info(f"Table {table_index + 1} 크기: {len(df)}행 x {len(df.columns)}열")
        logging.debug(f"Table {table_index + 1} 열: {df.columns.tolist()}")

        for row_index, row in df.iterrows():
            row_data = row.to_dict()
            row_y = y0 + (row_index + 1) * (y1 - y0) / (len(df) + 1)
            row_highlighted = check_highlight(row_y, highlight_regions)
            row_data["변경사항"] = "추가" if row_highlighted else ""
            processed_data.append(row_data)

    if not processed_data:
        logging.warning("처리된 데이터가 없습니다.")
        return pd.DataFrame()

    result_df = pd.DataFrame(processed_data)
    logging.info(f"총 {len(result_df)}개의 행이 추출되었습니다.")
    logging.debug(f"추출된 데이터 열: {result_df.columns.tolist()}")
    return result_df

# 강조 영역 확인
def check_highlight(row_y, highlight_regions):
    for region in highlight_regions:
        if region[0] <= row_y <= region[1]:
            return True
    return False

# 엑셀로 저장
def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    logging.info(f"파일이 '{output_path}'에 저장되었습니다.")

# 메인 함수
def main(pdf_path, output_excel_path):
    logging.info("PDF에서 개정된 부분을 추출합니다...")

    # PyMuPDF로 PDF 열기
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    # 원본 이미지 저장
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    # 강조된 영역 탐지
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    # 강조 영역이 표시된 이미지 저장
    highlighted_image = image.copy()
    for region in highlight_regions:
        cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))

    logging.info(f"감지된 강조 영역 수: {len(highlight_regions)}")
    logging.info(f"강조 영역: {highlight_regions}")

    # PyMuPDF를 사용하여 테이블 추출 및 처리
    processed_df = extract_and_process_tables(doc, page_number, highlight_regions)

    # 처리된 데이터 출력
    logging.info("처리된 데이터:")
    logging.info(processed_df)

    # 엑셀로 저장
    save_to_excel(processed_df, output_excel_path)

    logging.info(f"처리된 데이터가 {output_excel_path}에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)