import pytesseract
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
import cv2
import pdfplumber
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
    return df, data

def extract_table_from_pdf(pdf_path):
    print("PDF에서 표 추출 시작...")
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        tables = first_page.extract_tables()
        if tables:
            df = pd.DataFrame(tables[0][1:], columns=tables[0][0])
            print("PDF에서 표 추출 완료")
            return df
        else:
            print("PDF에서 표를 찾을 수 없습니다.")
            return None

def preprocess_dataframe(df):
    df = df.dropna(axis=1, how='all')
    df = df[(df.notna().any(axis=1)) & (df.astype(str).ne('').any(axis=1))]
    df = df.fillna('')
    return df

def compare_dataframes(df1, df2):
    print("데이터프레임 비교 시작...")
    changes = []

    df1 = preprocess_dataframe(df1)
    df2 = preprocess_dataframe(df2)

    if len(df2) > len(df1):
        for i in range(len(df1), len(df2)):
            changes.append(("새로운 행 추가", i, df2.iloc[i].to_dict()))

    for i in range(min(len(df1), len(df2))):
        for col in df1.columns:
            if col in df2.columns:
                val1 = str(df1.iloc[i][col]).strip()
                val2 = str(df2.iloc[i][col]).strip()
                if val1 != val2:
                    changes.append(("셀 변경", i, col, val1, val2))

    print(f"데이터프레임 비교 완료. {len(changes)}개의 변경 사항 발견")
    return changes

def highlight_changes_on_image(image_path, changes, ocr_data, output_path):
    print("변경 사항을 이미지에 표시하는 중...")
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    for change in changes:
        if change[0] == "새로운 행 추가":
            row_index = change[1]
            if row_index < len(ocr_data):
                y_position = ocr_data.iloc[row_index]['top']
                height = ocr_data.iloc[row_index]['height']
                draw.rectangle([10, y_position, image.width - 10, y_position + height], 
                               outline="red", fill=(255, 255, 0, 64), width=2)
                draw.text((15, y_position), "새로운 행", fill="black", font=font)

    image.save(output_path)
    print(f"변경 사항이 표시된 이미지가 저장되었습니다: {output_path}")

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_image_path = os.path.join(output_dir, "highlighted_changes.png")

    df_image, ocr_data = extract_table_from_image(image_path)
    df_pdf = extract_table_from_pdf(pdf_path)

    if df_image is not None and df_pdf is not None:
        changes = compare_dataframes(df_image, df_pdf)

        print("\n감지된 변경 사항:")
        for change in changes:
            print(change)

        highlight_changes_on_image(image_path, changes, ocr_data, output_image_path)
    else:
        print("표 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()