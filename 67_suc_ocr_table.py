import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image

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

    return regions

# PDF에서 테이블을 추출하고 강조된 행을 찾는 함수
def extract_and_process_tables(doc, page_number, highlight_regions):
    page = doc[page_number]
    tables = page.find_tables()
    
    processed_data = []

    for table in tables:
        df = pd.DataFrame(table.extract())
        for row_index in range(len(df)):
            row_data = df.iloc[row_index]
            row_highlighted = check_highlight(row_index, highlight_regions)  # 강조 영역 검사
            row_data["변경사항"] = "추가" if row_highlighted else "유지"
            processed_data.append(row_data)

    return pd.DataFrame(processed_data)

# 강조 영역 확인
def check_highlight(row_index, highlight_regions):
    for region in highlight_regions:
        if region[0] <= row_index <= region[1]:
            return True
    return False

# 엑셀로 저장
def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"파일이 '{output_path}'에 저장되었습니다.")

# 메인 함수
def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")

    # PyMuPDF로 PDF 열기
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    # 강조된 영역 탐지
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    # PyMuPDF를 사용하여 테이블 추출 및 처리
    processed_df = extract_and_process_tables(doc, page_number, highlight_regions)

    # 엑셀로 저장
    save_to_excel(processed_df, output_excel_path)

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)