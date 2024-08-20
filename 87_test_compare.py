import pytesseract
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
import cv2
import pdfplumber
import fitz  # PyMuPDF
import re
import os

def extract_table_from_image(image_path):
    print("이미지에서 표 추출 시작...")
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,40))
    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    table_mask = horizontal_lines + vertical_lines
    table_mask = cv2.dilate(table_mask, cv2.getStructuringElement(cv2.MORPH_RECT, (3,3)), iterations=1)
    
    table_area = cv2.bitwise_and(img, img, mask=table_mask)

    data = pytesseract.image_to_data(table_area, lang='kor+eng', output_type=pytesseract.Output.DATAFRAME)
    data = data[data.conf != -1]
    
    data['line_num'] = data['top'] // 10
    lines = data.groupby('line_num')['text'].apply(lambda x: ' '.join(x)).tolist()
    
    df = pd.DataFrame([line.split() for line in lines if line.strip()])
    print("이미지에서 표 추출 완료")
    return df

def is_color_highlighted(color):
    if isinstance(color, (tuple, list)) and len(color) == 3:
        return color not in [(1, 1, 1), (0.9, 0.9, 0.9)] and any(c < 0.9 for c in color)
    elif isinstance(color, int):
        return 0 < color < 230
    else:
        return False

def extract_highlighted_text_with_context(pdf_path):
    print("PDF에서 음영 처리된 텍스트 추출 시작...")
    doc = fitz.open(pdf_path)
    highlighted_texts_with_context = []
    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        lines = page.get_text("text").split('\n')
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if "color" in span and is_color_highlighted(span["color"]):
                            highlighted_text = span["text"]
                            line_index = lines.index(highlighted_text) if highlighted_text in lines else -1
                            if line_index != -1:
                                context = '\n'.join(lines[max(0, line_index-10):line_index])
                                highlighted_texts_with_context.append((context, highlighted_text))
    print("PDF에서 음영 처리된 텍스트 추출 완료")
    return highlighted_texts_with_context

def compare_dataframes(df_before, highlighted_texts_with_context):
    print("데이터프레임 비교 시작...")
    matching_rows = []

    for context, highlighted_text in highlighted_texts_with_context:
        context_lines = context.split('\n')
        for i in range(len(df_before)):
            match = True
            for j, line in enumerate(context_lines[-10:]):
                if i+j >= len(df_before) or not any(str(cell).strip() in line for cell in df_before.iloc[i+j]):
                    match = False
                    break
            if match:
                matching_rows.extend(df_before.iloc[i:i+10].index.tolist())
                break

    matching_rows = sorted(set(matching_rows))
    df_matching = df_before.loc[matching_rows].copy()
    df_matching['하단 표 삽입요망'] = '하단 표 삽입요망'
    
    print(f"데이터프레임 비교 완료. {len(matching_rows)}개의 일치하는 행 발견")
    return df_matching

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_excel_path = os.path.join(output_dir, "matching_rows.xlsx")

    df_before = extract_table_from_image(image_path)
    highlighted_texts_with_context = extract_highlighted_text_with_context(pdf_path)

    if df_before is not None and highlighted_texts_with_context:
        df_matching = compare_dataframes(df_before, highlighted_texts_with_context)
        df_matching.to_excel(output_excel_path, index=False)
        print(f"일치하는 행이 포함된 엑셀 파일이 저장되었습니다: {output_excel_path}")
    else:
        print("표 추출 또는 음영 처리된 텍스트 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()