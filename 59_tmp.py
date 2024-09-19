import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
import logging
import json
import requests
from dotenv import load_dotenv


# Hugging Face API 토큰 설정
load_dotenv()
HUGGINGFACE_API_TOKEN = os.getenv("HUGGINGFACE_API_TOKEN")

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

# Hugging Face API 토큰 설정 (실제 사용 시 환경 변수 등을 통해 안전하게 관리해야 합니다)
HUGGINGFACE_API_TOKEN = "your_huggingface_api_token_here"

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

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)

    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))

    logging.info(f"하이라이트 영역 탐지 완료: {len(contours)}개의 윤곽선 발견")
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

    logging.info(f"강조 영역 추출 완료: {len(regions)}개의 영역 발견")
    for i, region in enumerate(regions):
        logging.debug(f"강조 영역 {i+1}: y={region[0]} to y={region[1]}")
    return regions

def integrate_abnormal_rows(df):
    # 주요 열 정의 (이 열들의 값이 있으면 새로운 행으로 간주)
    key_columns = ['보장명', '특약명칭(번호)']
    # 병합 대상 열 (이 열들의 값은 병합 시 결합됨)
    merge_columns = ['지급사유', '지급금액']

    def is_new_entry(row):
        return any(pd.notna(row[col]) for col in key_columns)

    def merge_rows(current_row, next_row):
        for col in df.columns:
            if col in merge_columns and pd.notna(next_row[col]):
                if pd.isna(current_row[col]):
                    current_row[col] = next_row[col]
                else:
                    current_row[col] += f" {next_row[col]}"
            elif pd.isna(current_row[col]) and pd.notna(next_row[col]):
                current_row[col] = next_row[col]
        return current_row

    merged_rows = []
    current_row = None

    for _, row in df.iterrows():
        if current_row is None or is_new_entry(row):
            if current_row is not None:
                merged_rows.append(current_row)
            current_row = row.to_dict()
        else:
            current_row = merge_rows(current_row, row)

    if current_row is not None:
        merged_rows.append(current_row)

    result_df = pd.DataFrame(merged_rows)
    
    # 빈 문자열을 NaN으로 변환
    result_df = result_df.replace('', np.nan)
    
    return result_df

def verify_with_llm(original_df, processed_df, highlight_regions):
    # 강조 영역에 해당하는 행만 추출
    original_highlighted = original_df[original_df.apply(lambda row: any(region[0] <= row.name <= region[1] for region in highlight_regions), axis=1)]
    processed_highlighted = processed_df[processed_df.apply(lambda row: any(region[0] <= row.name <= region[1] for region in highlight_regions), axis=1)]
    
    original_sample = original_highlighted.to_string()
    processed_sample = processed_highlighted.to_string()
    
    prompt = f"""
    다음은 원본 PDF에서 추출한 강조된 영역의 데이터입니다:
    {original_sample}

    그리고 이는 처리 및 병합된 강조 영역의 데이터입니다:
    {processed_sample}

    두 데이터를 비교하고 다음 질문에 답해주세요:
    1. 처리된 데이터가 원본 데이터의 정보를 모두 포함하고 있나요?
    2. 데이터 병합 과정에서 정보의 손실이나 오류가 있나요?
    3. 추가적인 수정이 필요한 부분이 있다면 어떤 것인가요?

    JSON 형식으로 답변해주세요.
    """

    API_URL = "https://api-inference.huggingface.co/models/google/flan-t5-xxl"
    headers = {"Authorization": f"Bearer {HUGGINGFACE_API_TOKEN}"}

    try:
        response = requests.post(API_URL, headers=headers, json={"inputs": prompt})
        response.raise_for_status()
        
        analysis = response.json()[0]['generated_text']
        
        try:
            analysis_dict = json.loads(analysis)
        except json.JSONDecodeError:
            analysis_dict = {
                "data_completeness": "확인 필요",
                "data_errors": "분석 실패",
                "suggestions": "수동 검토 권장"
            }
    except Exception as e:
        logging.error(f"Hugging Face API 호출 중 오류 발생: {str(e)}")
        analysis_dict = {
            "data_completeness": "확인 불가",
            "data_errors": "API 오류",
            "suggestions": "수동 검토 필요"
        }
    
    return analysis_dict

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

        columns = cells[0]
        unique_columns = []
        for i, col in enumerate(columns):
            if col in unique_columns:
                unique_columns.append(f"{col}_{i}")
            else:
                unique_columns.append(col)

        original_df = pd.DataFrame(cells[1:], columns=unique_columns)
        
        # 데이터 통합
        processed_df = integrate_abnormal_rows(original_df)

        # LLM을 사용한 검증
        verification_result = verify_with_llm(original_df, processed_df, highlight_regions)
        
        logging.info("LLM 검증 결과:")
        logging.info(json.dumps(verification_result, indent=2, ensure_ascii=False))

        if verification_result['data_completeness'] != "완전함" or verification_result['data_errors'] != "없음":
            logging.warning("데이터 병합 결과에 문제가 있을 수 있습니다. 수동 검토가 필요합니다.")

        # 처리된 데이터 추가
        for row_index, row in processed_df.iterrows():
            row_data = row.to_dict()
            row_y = table.bbox[1] + (row_index + 1) * (table.bbox[3] - table.bbox[1]) / (len(processed_df) + 1)
            row_highlighted = any(region[0] <= row_y <= region[1] for region in highlight_regions)
            row_data["변경사항"] = "추가" if row_highlighted else ""
            processed_data.append(row_data)

    result_df = pd.DataFrame(processed_data)
    logging.info(f"총 {len(result_df)}개의 행이 추출되었습니다.")
    logging.debug(f"추출된 데이터 열: {result_df.columns.tolist()}")
    return result_df

def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    logging.info(f"파일이 '{output_path}'에 저장되었습니다.")

def main(pdf_path, output_excel_path):
    logging.info("PDF에서 개정된 부분을 추출합니다...")

    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    highlighted_image = image.copy()
    for region in highlight_regions:
        cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))

    logging.info(f"감지된 강조 영역 수: {len(highlight_regions)}")
    logging.info(f"강조 영역: {highlight_regions}")

    processed_df = extract_and_process_tables(doc, page_number, highlight_regions)

    logging.info("처리된 데이터:")
    logging.info(processed_df)

    save_to_excel(processed_df, output_excel_path)

    logging.info(f"처리된 데이터가 {output_excel_path}에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)