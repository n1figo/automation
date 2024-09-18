import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
from paddleocr import PaddleOCR

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

# PaddleOCR 초기화
ocr = PaddleOCR(use_angle_cls=True, lang='korean')

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

    if DEBUG_MODE:
        cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)

    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if DEBUG_MODE:
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

def extract_text_from_image(image):
    result = ocr.ocr(image, cls=True)
    return result

def process_ocr_result(ocr_result, highlight_regions):
    processed_data = []
    current_row = {}
    
    for line in ocr_result:
        for word_info in line:
            text = word_info[1][0]
            confidence = word_info[1][1]
            bbox = word_info[0]
            
            # 텍스트의 y 좌표 (bbox의 상단 y 좌표)
            text_y = bbox[0][1]
            
            # 강조 영역 확인
            is_highlighted = any(region[0] <= text_y <= region[1] for region in highlight_regions)
            
            # 헤더 확인 및 새 행 시작
            if text in TARGET_HEADERS:
                if current_row:
                    processed_data.append(current_row)
                current_row = {header: "" for header in TARGET_HEADERS}
                current_row["변경사항"] = "추가" if is_highlighted else ""
            
            # 현재 행에 텍스트 추가
            for header in TARGET_HEADERS:
                if text.startswith(header):
                    current_row[header] = text[len(header):].strip()
                    break
    
    # 마지막 행 추가
    if current_row:
        processed_data.append(current_row)
    
    return pd.DataFrame(processed_data)

def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"파일이 '{output_path}'에 저장되었습니다.")

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")

    doc = fitz.open(pdf_path)
    page_number = 50  # 51페이지 (인덱스는 0부터 시작)

    print(f"Processing page {page_number + 1}...")
    page = doc[page_number]
    image = pdf_to_image(page)

    if DEBUG_MODE:
        cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    if DEBUG_MODE:
        highlighted_image = image.copy()
        for region in highlight_regions:
            cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
        cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))

    ocr_result = extract_text_from_image(image)
    processed_df = process_ocr_result(ocr_result, highlight_regions)

    if DEBUG_MODE:
        print(f"Page {page_number + 1} processed data:")
        print(processed_df)

    save_to_excel(processed_df, output_excel_path)

    print(f"51페이지의 처리된 데이터가 {output_excel_path}에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)