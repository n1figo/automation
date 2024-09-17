import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import re
from PIL import Image
import pytesseract

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

# Tesseract OCR 경로 설정 (필요 시)
# 예: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# Ubuntu/macOS 예시:
pytesseract.pytesseract.tesseract_cmd = r'/usr/bin/tesseract'

# TESSDATA_PREFIX 환경 변수 설정 (필요 시)
os.environ['TESSDATA_PREFIX'] = '/usr/share/tesseract-ocr/4.00/tessdata/'  # Ubuntu 예시
# Windows 예시:
# os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata\'

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
    # Convert to PIL Image for pytesseract
    pil_img = Image.fromarray(roi)
    text = pytesseract.image_to_string(pil_img, lang='kor')  # 언어 설정 필요 시 조정
    return text

def verify_tesseract():
    """
    Tesseract OCR가 정상적으로 작동하는지 확인하기 위한 함수.
    간단한 텍스트가 포함된 이미지를 생성하여 OCR을 수행합니다.
    """
    sample_text = "테스트"
    # Create a simple image with the sample text using OpenCV
    img = np.ones((200, 400, 3), dtype=np.uint8) * 255  # White background
    # Use OpenCV's built-in fonts
    cv2.putText(img, sample_text, (50, 150), cv2.FONT_HERSHEY_SIMPLEX, 2, (0,0,0), 3, cv2.LINE_AA)
    pil_img = Image.fromarray(img)
    extracted_text = pytesseract.image_to_string(pil_img, lang='kor')
    print(f"OCR Extracted Text: '{extracted_text.strip()}'")  # 추가된 로그
    if sample_text in extracted_text:
        print("Tesseract OCR 검증 성공: '테스트'가 인식되었습니다.")
        return True
    else:
        print("Tesseract OCR 검증 실패: '테스트'가 인식되지 않았습니다.")
        return False

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
    
    # 테이블 데이터와 OCR 텍스트 매칭
    for table_index, table in enumerate(tables):
        print(f"테이블 {table_index + 1} 처리 중...")
        if not table:
            print(f"테이블 {table_index + 1}이 비어 있습니다.")
            continue
        
        # 테이블 내용 추출
        table_content = table
        if not table_content:
            print(f"테이블 {table_index + 1}에서 내용 추출 실패.")
            continue
        
        header_row = table_content[0]
        header_texts = [clean_text_for_excel(cell.strip()) if cell else '' for cell in header_row]
        header_texts_normalized = [text.replace(" ", "").replace("\n", "").replace("\t", "") for text in header_texts]
        
        if DEBUG_MODE:
            print(f"테이블 {table_index + 1} 헤더: {header_texts}")
        
        # 헤더 매칭: 부분 일치 허용
        if all(any(target_header in header_cell for header_cell in header_texts_normalized) for target_header in TARGET_HEADERS):
            if DEBUG_MODE:
                print(f"테이블 {table_index + 1}이 타겟 헤더를 포함합니다.")
            num_rows = len(table_content)
            num_cols = len(header_texts)
            for row_index in range(1, num_rows):
                row = table_content[row_index]
                row_data = {}
                change_detected = False
                
                # 디버깅: 셀 데이터 출력
                if DEBUG_MODE:
                    print(f"행 {row_index + 1}의 셀 데이터: {row}")
                
                for col_index in range(num_cols):
                    if col_index >= len(row):
                        continue  # 셀이 부족한 경우 건너뜀
                    header = header_texts[col_index].replace(" ", "").replace("\n", "").replace("\t", "")
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
                
                # 행의 텍스트 생성 (매칭 용이하게)
                row_text = ' '.join([str(row_data.get(col, '')) for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액"]])
                if DEBUG_MODE:
                    print(f"행 {row_index + 1}의 전체 텍스트: {row_text}")
                
                # OCR로 추출한 텍스트와 매칭
                for ocr_text in ocr_texts:
                    if ocr_text and row_text in ocr_text:
                        change_detected = True
                        if DEBUG_MODE:
                            print(f"행 {row_index + 1}이 하이라이트된 텍스트와 매칭되었습니다.")
                        break  # 한 번 매칭되면 더 이상 확인하지 않음
                
                if row_data:
                    row_data["페이지"] = page_number + 1
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    table_data.append(row_data)
    
    return table_data

def output_highlighted_rows(table_data, output_path):
    if not table_data:
        print("출력할 강조된 행이 없습니다.")
    with open(output_path, 'w', encoding='utf-8') as f:
        for row in table_data:
            if row.get("변경사항") == "추가":
                f.write(f"페이지: {row.get('페이지', '')}\n")
                f.write(f"보장명1: {row.get('보장명1', '')}\n")
                f.write(f"보장명2: {row.get('보장명2', '')}\n")
                f.write(f"지급사유1: {row.get('지급사유1', '')}\n")
                f.write(f"지급사유2: {row.get('지급사유2', '')}\n")
                f.write(f"지급금액: {row.get('지급금액', '')}\n")
                f.write("-" * 50 + "\n")
    print(f"강조된 행이 '{output_path}'에 저장되었습니다.")

def save_revisions_to_excel(df, output_excel_path):
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    
    # Tesseract OCR 검증
    print("Tesseract OCR 검증 중...")
    if not verify_tesseract():
        print("Tesseract OCR 검증에 실패했습니다. 스크립트를 종료합니다.")
        return
    print("Tesseract OCR 검증에 성공했습니다.")
    
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
    
    # 강조된 행 텍스트 파일로 출력
    output_highlighted_rows(table_data, os.path.join(TEXT_OUTPUT_DIR, "highlighted_rows.txt"))
    
    # 강조된 행이 있을 경우 Excel로 저장
    highlighted_rows = [row for row in table_data if row.get("변경사항") == "추가"]
    if highlighted_rows:
        df = pd.DataFrame(highlighted_rows)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항"]:
            if col not in df.columns:
                df[col] = ''
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항"]]
        save_revisions_to_excel(df, output_excel_path)
        print("작업이 완료되었습니다.")
    else:
        print("지정된 헤더를 가진 표를 찾거나, 변경된 행이 없습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)
