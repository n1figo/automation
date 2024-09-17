import pdfplumber
import pandas as pd
import numpy as np
import cv2
import os
import re
from PIL import Image
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

# 문장 임베딩 모델 로드
model = SentenceTransformer('sentence-transformers/xlm-r-100langs-bert-base-nli-stsb-mean-tokens')

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

# PDF 페이지를 이미지로 변환하는 함수
def pdf_to_image(page):
    # 페이지를 PIL 이미지로 변환
    pil_image = page.to_image(resolution=300).original.convert('RGB')
    return np.array(pil_image)

# 색상 유사도 비교 함수
def is_similar_color(c1, c2, tolerance=30):
    return all(abs(c1[i] - c2[i]) <= tolerance for i in range(3))

# 강조 색상을 텍스트로 변환하는 함수
def convert_color_to_text(color):
    return f"RGB({color[0]}, {color[1]}, {color[2]})"

# 배경색 추출 함수 (페이지 전체에서 강조 색상 감지)
def get_page_background_colors(image):
    img_array = np.array(image)
    height, width, _ = img_array.shape

    detected_colors = set()

    # 성능 향상을 위해 간격을 두고 색상 검사
    step = 10  # 픽셀 간격
    for y in range(0, height, step):
        for x in range(0, width, step):
            pixel_color = tuple(np.uint8(img_array[y, x]))
            if not (is_similar_color(pixel_color, (255, 255, 255)) or
                    is_similar_color(pixel_color, (0, 0, 0)) or
                    is_similar_color(pixel_color, (127, 127, 127))):
                detected_colors.add(pixel_color)

    if detected_colors:
        detected_colors_text = [convert_color_to_text(color) for color in detected_colors]
        print(f"페이지에서 감지된 강조 색상: {detected_colors_text}")
        return detected_colors_text
    return []

# 페이지에서 강조 색상을 감지하는 함수
def detect_emphasized_color_on_page(page_image):
    detected_colors = get_page_background_colors(page_image)
    if detected_colors:
        return True, detected_colors
    return False, []

# 텍스트 유사도 함수
def text_similarity(text1, text2):
    if not text1 or not text2:
        return 0.0
    embeddings = model.encode([text1, text2])
    cos_similarity = cosine_similarity([embeddings[0]], [embeddings[1]])[0][0]
    return cos_similarity

# 셀의 배경색을 추출하는 함수
def get_cell_background_color(image, cell, page_width, page_height):
    # 셀의 위치 정보 (x0, top, x1, bottom)를 기반으로 셀 영역 추출
    x0, top, x1, bottom = cell["x0"], cell["top"], cell["x1"], cell["bottom"]
    # 이미지의 크기
    img_height, img_width, _ = image.shape

    # PDF 페이지의 크기와 이미지의 크기 비율 계산
    y_ratio = img_height / page_height
    x_ratio = img_width / page_width

    # 픽셀 좌표로 변환
    x0_pixel = int(x0 * x_ratio)
    x1_pixel = int(x1 * x_ratio)
    top_pixel = int(top * y_ratio)
    bottom_pixel = int(bottom * y_ratio)

    # 셀 이미지 크롭
    cell_image = image[top_pixel:bottom_pixel, x0_pixel:x1_pixel]

    if cell_image.size == 0:
        return None

    # 주요 색상 추출 (가장 많이 나타나는 색상)
    pixels = cell_image.reshape(-1, 3)
    counts = {}
    for pixel in pixels:
        key = tuple(pixel)
        counts[key] = counts.get(key, 0) + 1
    dominant_color = max(counts, key=counts.get)

    return convert_color_to_text(dominant_color)

# 테이블 추출 및 변경사항 분석 함수
def extract_target_tables_from_page(page, image, page_number, emphasized_colors):
    print(f"페이지 {page_number + 1} 처리 중...")
    tables = page.find_tables()
    print(f"페이지 {page_number + 1}에서 찾은 테이블 수: {len(tables)}")

    if not tables:
        print(f"페이지 {page_number + 1}에서 테이블을 찾을 수 없습니다.")
        return []

    table_data = []
    page_width = page.width
    page_height = page.height

    for table_index, table in enumerate(tables):
        print(f"테이블 {table_index + 1} 처리 중...")
        table_content = table.extract()
        if not table_content:
            continue

        header_row = table_content[0]
        header_texts = [clean_text_for_excel(cell.strip()) if cell else '' for cell in header_row]
        header_texts_normalized = [text.replace(" ", "").replace("\n", "") for text in header_texts]

        # 지정된 헤더가 포함된 테이블만 처리
        if all(any(target_header in header_cell for header_cell in header_texts_normalized) for target_header in TARGET_HEADERS):
            num_rows = len(table_content)
            num_cols = len(header_texts)
            for row_index in range(1, num_rows):
                row = table_content[row_index]
                row_data = {}
                change_detected = False
                cell_bg_colors = []

                for col_index in range(num_cols):
                    header = header_texts[col_index].replace(" ", "").replace("\n", "")
                    if header in TARGET_HEADERS:
                        cell_text = clean_text_for_excel(row[col_index].strip()) if row[col_index] else ''

                        # 해당 셀을 찾기 위해 table.cells를 순회
                        # table.cells는 리스트이므로, row_index와 col_index에 맞는 셀을 찾기
                        # 주의: table.cells는 순차적으로 나열되므로, 특정 셀을 찾기 위해서는 row와 col을 비교
                        cell = None
                        for c in table.cells:
                            if c["row"] == row_index and c["col"] == col_index:
                                cell = c
                                break

                        if cell and isinstance(cell, dict):
                            cell_bg_color = get_cell_background_color(image, cell, page_width, page_height)
                            if cell_bg_color in emphasized_colors:
                                change_detected = True
                        else:
                            cell_bg_color = ''

                        # 데이터 분리
                        cell_texts = cell_text.split('\n')
                        if header == '보장명':
                            if len(cell_texts) > 1:
                                row_data['보장명1'] = cell_texts[0]
                                row_data['보장명2'] = cell_texts[1]
                            else:
                                row_data['보장명1'] = cell_texts[0] if cell_texts else ''
                                row_data['보장명2'] = ''
                        elif header == '지급사유':
                            if len(cell_texts) > 1:
                                row_data['지급사유1'] = cell_texts[0]
                                row_data['지급사유2'] = cell_texts[1]
                            else:
                                row_data['지급사유1'] = cell_texts[0] if cell_texts else ''
                                row_data['지급사유2'] = ''
                        else:
                            row_data[header] = cell_text

                        # 배경색 기록
                        if cell_bg_color:
                            cell_bg_colors.append(cell_bg_color)

                        # 디버깅: 셀의 배경색 출력
                        if cell_bg_color:
                            print(f"페이지 {page_number + 1}, 테이블 {table_index + 1}, 셀 ({row_index}, {col_index}) 배경색: {cell_bg_color}")
                        else:
                            print(f"페이지 {page_number + 1}, 테이블 {table_index + 1}, 셀 ({row_index}, {col_index}) 배경색: 없음")

                # 변경사항 및 배경색 설정
                if change_detected:
                    row_data["변경사항"] = "추가"
                    row_data["배경색"] = ', '.join(set(cell_bg_colors))
                else:
                    row_data["변경사항"] = "유지"
                    row_data["배경색"] = ''

                if row_data:
                    row_data["페이지"] = page_number + 1
                    table_data.append(row_data)

                    # 디버깅: 각 행의 데이터와 변경 사항 출력
                    print(f"행 {row_index}: 변경사항 = {row_data['변경사항']}, 데이터 = {row_data}")

    return table_data

# 엑셀에 결과를 저장하는 함수
def save_results_to_excel(df, output_excel_path):
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

# 메인 함수 (51페이지에 대해서만 처리)
def main(pdf_path, output_excel_path):
    print("PDF에서 51페이지 개정된 부분을 추출합니다...")
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        if 50 >= total_pages:
            print(f"PDF에 51 페이지가 존재하지 않습니다.")
            return

        page_number = 50  # 51페이지는 인덱스로 50
        page = pdf.pages[page_number]
        image = pdf_to_image(page)

        # 원본 이미지 저장 (디버깅용)
        cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

        # 페이지에서 강조 색상 감지
        color_detected, emphasized_colors = detect_emphasized_color_on_page(image)
        if color_detected:
            print(f"페이지 {page_number + 1}에서 강조 색상 발견: {emphasized_colors}")
        else:
            print(f"페이지 {page_number + 1}에서 강조 색상을 발견하지 못했습니다.")

        # 테이블 추출 및 변경 사항 분석
        table_data = extract_target_tables_from_page(page, image, page_number, emphasized_colors)

        if table_data:
            df = pd.DataFrame(table_data)
            for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액"]:
                if col not in df.columns:
                    df[col] = ''
            df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "배경색", "변경사항"]]
            save_results_to_excel(df, output_excel_path)

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
