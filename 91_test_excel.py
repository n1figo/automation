import pandas as pd
import numpy as np
import cv2
import pytesseract
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import os

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

def compare_dataframes(df1, df2):
    print("데이터프레임 비교 시작...")
    changes = []

    df1 = df1.fillna('')
    df2 = df2.fillna('')

    for i in range(min(len(df1), len(df2))):
        if not df1.iloc[i].equals(df2.iloc[i]):
            changes.append(("행 변경", i, df1.iloc[i].to_dict(), df2.iloc[i].to_dict()))

    if len(df2) > len(df1):
        for i in range(len(df1), len(df2)):
            changes.append(("새로운 행 추가", i, df2.iloc[i].to_dict()))

    print(f"데이터프레임 비교 완료. {len(changes)}개의 변경 사항 발견")
    return changes

def highlight_changes_in_excel(df, changes, output_path):
    print("변경 사항을 Excel 파일에 표시하는 중...")
    wb = Workbook()
    ws = wb.active

    # 데이터프레임을 Excel 워크시트로 변환
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # 변경 사항 강조 표시
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    red_font = Font(color='FF0000', bold=True)

    for change in changes:
        if change[0] == "행 변경":
            row_index = change[1] + 2  # Excel은 1부터 시작하고 헤더가 있으므로 +2
            ws.insert_rows(row_index)
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_index, column=col)
                cell.fill = yellow_fill
                cell.value = "변경된 행"
            
            # 변경된 셀 강조
            for col, (old_val, new_val) in enumerate(zip(change[2].values(), change[3].values()), 1):
                if old_val != new_val:
                    cell = ws.cell(row=row_index+1, column=col)
                    cell.font = red_font

        elif change[0] == "새로운 행 추가":
            row_index = change[1] + 2
            ws.insert_rows(row_index)
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_index, column=col)
                cell.fill = yellow_fill
                cell.value = "새로운 행 추가"

    wb.save(output_path)
    print(f"변경 사항이 표시된 Excel 파일이 저장되었습니다: {output_path}")

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_excel_path = os.path.join(output_dir, "highlighted_changes.xlsx")

    df_image = extract_table_from_image(image_path)
    df_pdf = extract_table_from_pdf(pdf_path)

    if df_image is not None and df_pdf is not None:
        changes = compare_dataframes(df_image, df_pdf)
        highlight_changes_in_excel(df_image, changes, output_excel_path)
    else:
        print("표 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()