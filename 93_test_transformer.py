import pytesseract
from PIL import Image
import pandas as pd
import tabula
import numpy as np
import cv2

def extract_table_from_image(image_path):
    print("이미지에서 표 추출 시작...")
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # Detect horizontal lines
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
    detect_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    cnts = cv2.findContours(detect_horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    for c in cnts:
        cv2.drawContours(img, [c], -1, (255,0,0), 2)

    # Detect vertical lines
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,40))
    detect_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    cnts = cv2.findContours(detect_vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    for c in cnts:
        cv2.drawContours(img, [c], -1, (255,0,0), 2)

    # OCR
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
    data = data[data.conf != -1]
    lines = data.groupby('block_num')['text'].apply(list).tolist()
    
    df = pd.DataFrame(lines)
    print("이미지에서 표 추출 완료")
    return df

def extract_table_from_pdf(pdf_path):
    print("PDF에서 표 추출 시작...")
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    if tables:
        df = tables[0]  # Assuming we're interested in the first table
        print("PDF에서 표 추출 완료")
        return df
    else:
        print("PDF에서 표를 찾을 수 없습니다.")
        return None

def compare_dataframes(df1, df2):
    print("데이터프레임 비교 시작...")
    changes = []

    # Check for added rows
    if len(df2) > len(df1):
        for i in range(len(df1), len(df2)):
            changes.append(f"새로운 행 추가: {df2.iloc[i].to_dict()}")

    # Check for changed cells
    for i in range(min(len(df1), len(df2))):
        for col in df1.columns:
            if df1.iloc[i][col] != df2.iloc[i][col]:
                changes.append(f"셀 변경: 행 {i+1}, 열 '{col}' - "
                               f"변경 전: {df1.iloc[i][col]}, 변경 후: {df2.iloc[i][col]}")

    print(f"데이터프레임 비교 완료. {len(changes)}개의 변경 사항 발견")
    return changes

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"

    df_image = extract_table_from_image(image_path)
    df_pdf = extract_table_from_pdf(pdf_path)

    if df_image is not None and df_pdf is not None:
        changes = compare_dataframes(df_image, df_pdf)

        print("\n감지된 변경 사항:")
        for i, change in enumerate(changes, 1):
            print(f"변경 사항 {i}: {change}")
    else:
        print("표 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()