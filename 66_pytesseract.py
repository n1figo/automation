import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import re
from PIL import Image
import pytesseract  # Tesseract로 OCR 수행

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 및 텍스트 파일 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
TEXT_OUTPUT_DIR = "/workspaces/automation/output"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)
os.makedirs(TEXT_OUTPUT_DIR, exist_ok=True)

BEFORE_HIGHLIGHT_PATH = os.path.join(TEXT_OUTPUT_DIR, "before_highlight.txt")

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
        return text.replace(" ", "").replace("\n", "").replace("\t", "")  # 공백, 개행, 탭 제거
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
    mask_path = os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png')
    cv2.imwrite(mask_path, cleaned_mask)
    if DEBUG_MODE:
        print(f"마스크 이미지 저장: {mask_path}")
    
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 디버깅: 윤곽선이 그려진 이미지 저장
    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    contours_path = os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png')
    cv2.imwrite(contours_path, cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))
    if DEBUG_MODE:
        print(f"윤곽선 이미지 저장: {contours_path}")
    
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

def extract_text_from_region(image, region):
    start_y, end_y = region
    roi = image[start_y:end_y, :]
    
    # 이미지 전처리
    gray = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                   cv2.THRESH_BINARY, 31, 2)
    denoised = cv2.medianBlur(thresh, 3)
    
    # 이미지 크기 확대 (예: 150%)
    scale_percent = 150
    width = int(denoised.shape[1] * scale_percent / 100)
    height = int(denoised.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized = cv2.resize(denoised, dim, interpolation=cv2.INTER_LINEAR)
    
    # Tesseract로 OCR 수행
    extracted_text = pytesseract.image_to_string(resized, lang='kor')
    return extracted_text.strip()

def extract_target_tables_from_page(page, image, page_number, pdf_path, before_highlight_file):
    print(f"페이지 {page_number + 1} 처리 중...")
    
    # 테이블 추출: PyMuPDF의 extract_tables() 사용
    try:
        tables = page.extract_tables()
        if DEBUG_MODE:
            print(f"페이지 {page_number + 1}에서 추출한 테이블 수: {len(tables)}")
    except Exception as e:
        print(f"페이지 {page_number + 1}에서 테이블 추출 중 오류 발생: {e}")
        tables = []
    
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])
    
    # 디버깅: 강조 영역이 표시된 이미지 저장
    debug_image = image.copy()
    for start_y, end_y in highlight_regions:
        cv2.rectangle(debug_image, (0, start_y), (image.shape[1], end_y), (255, 0, 0), 2)
    highlights_path = os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlights.png')
    cv2.imwrite(highlights_path, cv2.cvtColor(debug_image, cv2.COLOR_RGB2BGR))
    if DEBUG_MODE:
        print(f"강조 영역 이미지 저장: {highlights_path}")
    
    table_data = []
    ocr_texts = []
    
    # 하이라이트 영역에서 OCR로 텍스트 추출
    for idx, region in enumerate(highlight_regions, start=1):
        extracted_text = extract_text_from_region(image, region)
        if DEBUG_MODE:
            print(f"하이라이트 영역 {idx} ({region})에서 추출한 텍스트:\n{extracted_text}\n")
        ocr_texts.append(extracted_text)
        
        # before_highlight.txt 파일에 추출된 텍스트 기록
        with open(before_highlight_file, 'a', encoding='utf-8') as f:
            f.write(f"페이지 {page_number + 1} - 영역 {idx}:\n{extracted_text}\n")
            f.write("-" * 50 + "\n")
        if DEBUG_MODE:
            print(f"추출된 텍스트가 '{before_highlight_file}'에 기록되었습니다.")
    
    return table_data

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")

    doc = fitz.open(pdf_path)
    
    # 51페이지만 처리 (0-based index이므로 50)
    page_number = 50
    if page_number >= len(doc):
        print(f"PDF에 페이지 {page_number + 1}이 존재하지 않습니다.")
        return
    
    page = doc.load_page(page_number)
    image = pdf_to_image(page)
    
    # 원본 이미지 저장
    original_path = os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png')
    cv2.imwrite(original_path, cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
    if DEBUG_MODE:
        print(f"원본 이미지 저장: {original_path}")
    
    # before_highlight.txt 초기화
    with open(BEFORE_HIGHLIGHT_PATH, 'w', encoding='utf-8') as f:
        f.write(f"페이지 {page_number + 1}의 하이라이트된 텍스트\n")
        f.write("=" * 50 + "\n")
    
    # 페이지에서 표 추출 및 강조 영역 분석
    table_data = extract_target_tables_from_page(page, image, page_number, pdf_path, BEFORE_HIGHLIGHT_PATH)
    
    if DEBUG_MODE:
        print("추출된 테이블 데이터:", table_data)  # 디버그: table_data 출력
    
    print("작업이 완료되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)
