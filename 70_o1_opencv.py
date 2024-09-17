import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
import pytesseract
from pytesseract import Output
import numpy as np
import concurrent.futures
import os
from typing import List, Tuple
from PIL import Image
import io

# 셀의 이미지를 추출하고 Tesseract로 분석하여 음영 여부를 판단하는 함수
def is_cell_shaded(image):
    # 이미지에서 텍스트 없이 빈 공간의 픽셀 수를 계산
    # Tesseract의 신뢰도(confidence)를 이용하여 빈 영역을 찾음
    d = pytesseract.image_to_data(image, output_type=Output.DICT)
    conf = d['conf']
    n_boxes = len(d['text'])

    # 신뢰도가 -1인 영역은 빈 공간 또는 배경으로 판단
    empty_areas = sum(1 for c in conf if int(c) == -1)

    # 전체 영역 대비 빈 공간의 비율 계산
    empty_ratio = empty_areas / n_boxes if n_boxes > 0 else 0

    # 빈 공간의 비율이 높으면 음영 처리된 것으로 판단
    if empty_ratio > 0.5:  # 임계값은 조정 가능
        return True
    else:
        return False

# 한 페이지에서 표와 셀의 음영 여부를 추출하는 함수
def extract_tables_from_page(page):
    tables = page.get_text("blocks")
    page_tables = []

    for block in tables:
        if block['type'] == 0:  # 텍스트 블록인 경우
            # 블록 내의 줄과 스팬을 통해 텍스트와 위치를 얻습니다.
            lines = block['lines']
            data = []
            for line in lines:
                row_data = []
                for span in line['spans']:
                    text = span['text'].strip()
                    bbox = fitz.Rect(span['bbox'])
                    # 셀의 이미지를 추출
                    pix = page.get_pixmap(clip=bbox, colorspace=fitz.csRGB)
                    img_data = pix.tobytes(output="png")
                    image = Image.open(io.BytesIO(img_data))
                    # 셀이 음영 처리되었는지 판단
                    if is_cell_shaded(image):
                        status = "추가"
                    else:
                        status = "유지"
                    row_data.append((text, status))
                if row_data:
                    data.append(row_data)
            if data:
                page_tables.append(data)
    return page_tables

# 병렬 처리를 위한 함수
def process_page(args):
    page_number, pdf_path = args
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number)
    tables = extract_tables_from_page(page)
    return tables

# 엑셀 파일로 저장하는 함수
def save_tables_to_excel(all_tables, output_excel_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Extracted Tables"
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    row_num = 1
    for tables in all_tables:
        for table in tables:
            for row in table:
                col_num = 1
                for text, status in row:
                    cell = sheet.cell(row=row_num, column=col_num, value=text)
                    cell.alignment = Alignment(wrap_text=True)
                    if status == "추가":
                        cell.fill = yellow_fill
                    col_num += 1
                row_num += 1
            row_num += 2  # 테이블 간 간격
    workbook.save(output_excel_path)
    print(f"Data saved to {output_excel_path}")

# 메인 함수
def main(pdf_path, output_excel_path, max_workers=4):
    doc = fitz.open(pdf_path)
    num_pages = doc.page_count
    all_tables = []

    # 병렬 처리 설정
    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        # 각 페이지 번호와 필요한 인자를 튜플로 전달
        args = [(page_num, pdf_path) for page_num in range(num_pages)]
        # 결과를 받아옴
        results = executor.map(process_page, args)
        for result in results:
            all_tables.append(result)

    # 결과를 엑셀로 저장
    save_tables_to_excel(all_tables, output_excel_path)

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    max_workers = 4  # 병렬 처리 시 사용할 프로세스 수

    main(pdf_path, output_excel_path, max_workers)
