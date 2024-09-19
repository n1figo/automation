import camelot
import pandas as pd
import os
import cv2
import numpy as np
from PIL import Image

# 디버깅 모드 설정
DEBUG_MODE = True

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def extract_tables_with_camelot(pdf_path, page_number):
    print(f"Extracting tables from page {page_number} using Camelot...")
    tables = camelot.read_pdf(pdf_path, pages=str(page_number), flavor='lattice')
    print(f"Found {len(tables)} tables on page {page_number}")
    return tables

def process_tables(tables):
    processed_data = []
    for i, table in enumerate(tables):
        df = table.df
        df['Table_Number'] = i + 1
        df['변경사항'] = ""  # 여기서는 변경사항을 별도로 표시하지 않습니다
        processed_data.append(df)
    return pd.concat(processed_data, ignore_index=True)

def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"Data saved to '{output_path}'")

def main(pdf_path, output_excel_path):
    print("Extracting tables from PDF...")

    page_number = 50  # 51페이지 (0-based index)

    # Camelot을 사용하여 표 추출
    tables = extract_tables_with_camelot(pdf_path, page_number + 1)

    # 추출된 표 처리
    processed_df = process_tables(tables)

    # 처리된 데이터 출력
    print(processed_df)

    # 엑셀로 저장
    save_to_excel(processed_df, output_excel_path)

    print(f"Processed data saved to {output_excel_path}")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables_camelot.xlsx"
    main(pdf_path, output_excel_path)