import pandas as pd
import numpy as np
import cv2
import pytesseract
import pdfplumber
import os
import csv
import codecs

def extract_table_from_image(image_path):
    print("이미지에서 표 추출 시작...")
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
    opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

    custom_config = r'--oem 3 --psm 6'
    data = pytesseract.image_to_data(opening, lang='kor+eng', config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data[data.conf != -1]
    
    data['line_num'] = data['top'] // 5
    lines = data.groupby('line_num').agg({'text': ' '.join, 'top': 'min', 'height': 'max'}).reset_index()
    
    df = pd.DataFrame([line['text'].split() for _, line in lines.iterrows()])
    print("이미지에서 표 추출 완료")
    return df

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

def save_dataframe_to_csv(df, output_path):
    try:
        with codecs.open(output_path, 'w', encoding='utf-8-sig') as f:
            df.to_csv(f, index=False, quoting=csv.QUOTE_ALL)
        print(f"CSV 파일이 성공적으로 저장되었습니다: {output_path}")
    except Exception as e:
        print(f"CSV 파일 저장 중 오류 발생: {e}")

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)

    # 이미지에서 추출한 표 저장
    df_image = extract_table_from_image(image_path)
    if df_image is not None:
        image_csv_path = os.path.join(output_dir, "table_from_image.csv")
        save_dataframe_to_csv(df_image, image_csv_path)
    else:
        print("이미지에서 표 추출에 실패했습니다.")

    # PDF에서 추출한 표 저장
    df_pdf = extract_table_from_pdf(pdf_path)
    if df_pdf is not None:
        pdf_csv_path = os.path.join(output_dir, "table_from_pdf.csv")
        save_dataframe_to_csv(df_pdf, pdf_csv_path)
    else:
        print("PDF에서 표 추출에 실패했습니다.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()