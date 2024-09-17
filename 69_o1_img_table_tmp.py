import fitz  # PyMuPDF
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
import re
import os

# 디버깅 모드 설정 (True로 설정하면 색상 정보를 출력합니다)
DEBUG_MODE = False

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

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
        # 제어 문자 제거 (줄바꿈은 유지)
        text = remove_illegal_characters(text)
        return text  # 줄바꿈을 제거하지 않음
    return text

# 셀의 변경사항 여부를 판단하는 함수 (이미지 기반 하이라이트 감지)
def is_cell_highlighted(cell_image):
    img_array = np.array(cell_image)
    # 이미지가 흑백이면 RGB로 변환
    if len(img_array.shape) == 2:
        img_array = np.stack((img_array,)*3, axis=-1)
    elif img_array.shape[2] == 4:
        # 투명도 채널 제거
        img_array = img_array[:, :, :3]
    # 특정 색상의 픽셀 수 계산
    pixels = img_array.reshape(-1, 3)
    highlighted = False
    for pixel in pixels:
        r, g, b = pixel
        # 살색 음영 감지 (예시로 RGB 값 범위 설정)
        if r > 200 and g > 150 and b > 150:
            highlighted = True
            break
    return highlighted

# 셀의 배경색을 추출하는 함수
def get_cell_background_color(cell_image):
    img_array = np.array(cell_image)
    # 이미지가 흑백이면 RGB로 변환
    if len(img_array.shape) == 2:
        img_array = np.stack((img_array,)*3, axis=-1)
    elif img_array.shape[2] == 4:
        # 투명도 채널 제거
        img_array = img_array[:, :, :3]
    # 이미지에서 주요 색상 추출
    pixels = img_array.reshape(-1, 3)
    # 배경색 추출 (최빈값)
    counts = {}
    for pixel in pixels:
        key = tuple(pixel)
        counts[key] = counts.get(key, 0) + 1
    dominant_color = max(counts, key=counts.get)
    return dominant_color

# 페이지에서 타겟 표를 추출하는 함수
def extract_target_tables_from_page(page, page_image, page_number):
    print(f"페이지 {page_number + 1} 처리 중...")
    table_data = []
    tables = page.find_tables()
    for table in tables:
        # 테이블 데이터 추출
        table_content = table.extract()
        if not table_content:
            continue
        # 헤더 행 추출 및 정리
        header_row = table_content[0]
        header_texts = [clean_text_for_excel(cell.strip()) if cell else '' for cell in header_row]
        header_texts_normalized = [text.replace(" ", "").replace("\n", "") for text in header_texts]
        # 타겟 헤더가 모두 포함되어 있는지 확인
        if all(any(target_header == header_cell for header_cell in header_texts_normalized) for target_header in TARGET_HEADERS):
            # 셀 딕셔너리 생성 (셀 위치를 기준으로 매핑)
            cell_dict = {}
            for cell in table.cells:
                # cell의 구조를 확인하여 적절히 접근
                if len(cell) == 6:
                    cell_row = cell[0]
                    cell_col = cell[1]
                    cell_bbox = cell[2:6]  # x0, y0, x1, y1
                elif len(cell) >= 3 and isinstance(cell[2], (tuple, list)):
                    cell_row = cell[0]
                    cell_col = cell[1]
                    cell_bbox = cell[2]  # (x0, y0, x1, y1)
                else:
                    continue  # 알 수 없는 셀 구조인 경우 건너뜁니다.

                # cell_bbox가 4개의 값(x0, y0, x1, y1)인지 확인
                if len(cell_bbox) == 4:
                    cell_rect = fitz.Rect(*cell_bbox)
                else:
                    continue  # bbox가 올바르지 않은 경우 건너뜁니다.

                cell_dict[(cell_row, cell_col)] = cell_rect
            # 테이블 데이터 처리
            num_rows = len(table_content)
            num_cols = len(header_texts)
            for row_index in range(1, num_rows):  # 헤더 행 이후부터 처리
                row = table_content[row_index]
                row_data = {}
                change_detected = False
                cell_bg_color = None  # 초기화
                for col_index in range(num_cols):
                    header = header_texts[col_index].replace(" ", "").replace("\n", "")
                    # 셀 값 처리 시 None 체크 추가
                    cell_text = clean_text_for_excel(row[col_index].strip()) if row[col_index] else ''
                    if header in TARGET_HEADERS:
                        # 셀 내용 분할 (줄바꿈 기준)
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
                        # 셀 객체 가져오기
                        cell_rect = cell_dict.get((row_index, col_index))
                        if cell_rect:
                            # 좌표 변환 (PyMuPDF 좌표계에서 이미지 좌표계로)
                            x0, y0, x1, y1 = cell_rect
                            x0 = int(x0)
                            y0 = int(page_image.height - y1)
                            x1 = int(x1)
                            y1 = int(page_image.height - y0)
                            # 셀 이미지 크롭
                            cell_image = page_image.crop((x0, y0, x1, y1))
                            # 배경색 추출
                            bg_color = get_cell_background_color(cell_image)
                            cell_bg_color = bg_color
                            # 셀 하이라이트 여부 판단
                            if is_cell_highlighted(cell_image):
                                change_detected = True
                            # 디버깅 모드일 때 색상 정보 출력
                            if DEBUG_MODE:
                                print(f"페이지 {page_number + 1}, 셀 ({row_index}, {col_index}) - 배경색: {bg_color}, 변경사항 있음: {change_detected}")
                if row_data:
                    # 페이지 번호 추가
                    row_data["페이지"] = page_number + 1
                    # 변경사항 설정
                    row_data["변경사항"] = "추가" if change_detected else "유지"
                    # 배경색 추가
                    row_data["배경색"] = str(cell_bg_color) if cell_bg_color else ''
                    table_data.append(row_data)
    return table_data

# 메인 함수
def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    # PDF를 이미지로 변환
    images = convert_from_path(pdf_path, dpi=200, fmt='png')
    all_table_data = []
    for page_number in range(total_pages):
        page = doc[page_number]
        page_image = images[page_number]
        table_data = extract_target_tables_from_page(page, page_image, page_number)
        all_table_data.extend(table_data)
    if all_table_data:
        # 데이터프레임 생성
        df = pd.DataFrame(all_table_data)
        # 결측치 처리 (컬럼이 없을 경우 대비)
        for col in ["보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "배경색"]:
            if col not in df.columns:
                df[col] = ''
        # 컬럼 순서 지정
        df = df[["페이지", "보장명1", "보장명2", "지급사유1", "지급사유2", "지급금액", "변경사항", "배경색"]]
        # 엑셀로 저장
        print("개정된 부분을 엑셀 파일로 저장합니다...")
        save_revisions_to_excel(df, output_excel_path)
        print("작업이 완료되었습니다.")
    else:
        print("지정된 헤더를 가진 표를 찾을 수 없습니다.")

# 엑셀 파일로 저장하는 함수
def save_revisions_to_excel(df, output_excel_path):
    # 엑셀 파일 생성
    df.to_excel(output_excel_path, index=False)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)
