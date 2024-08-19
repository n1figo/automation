import pandas as pd
import numpy as np
import cv2
import pytesseract
import pdfplumber
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

    # 열 이름 통일
    common_columns = list(set(df1.columns) & set(df2.columns))
    df1 = df1[common_columns]
    df2 = df2[common_columns]

    for i in range(min(len(df1), len(df2))):
        row_changes = []
        for col in common_columns:
            if df1.iloc[i][col] != df2.iloc[i][col]:
                row_changes.append((col, df1.iloc[i][col], df2.iloc[i][col]))
        if row_changes:
            changes.append(("행 변경", i, row_changes))

    if len(df2) > len(df1):
        for i in range(len(df1), len(df2)):
            changes.append(("새로운 행 추가", i, df2.iloc[i].to_dict()))

    print(f"데이터프레임 비교 완료. {len(changes)}개의 변경 사항 발견")
    return changes

import pandas as pd
import numpy as np
import cv2
import pytesseract
import pdfplumber
import os
import csv
import codecs

# ... (이전의 extract_table_from_image, extract_table_from_pdf, compare_dataframes 함수들은 그대로 유지) ...

def mark_changes_in_csv(df, changes, output_path):
    print("변경 사항을 CSV 파일에 표시하는 중...")
    
    # 'before_mark' 컬럼 추가
    df.insert(0, 'before_mark', '')

    for change in changes:
        if change[0] == "행 변경":
            row_index = change[1]
            df.at[row_index, 'before_mark'] = 'BEFORE'
            for col_name, old_val, new_val in change[2]:
                df.at[row_index, col_name] = f"{old_val} -> {new_val}"
        elif change[0] == "새로운 행 추가":
            row_index = change[1]
            if row_index > 0:
                df.at[row_index-1, 'before_mark'] = 'BEFORE'
            new_row = pd.DataFrame({'before_mark': ['NEW'], **change[2]}, index=[row_index])
            df = pd.concat([df.iloc[:row_index], new_row, df.iloc[row_index:]]).reset_index(drop=True)

    try:
        # UTF-8 인코딩으로 CSV 파일 저장 시도
        with codecs.open(output_path, 'w', encoding='utf-8-sig') as f:
            df.to_csv(f, index=False, quoting=csv.QUOTE_ALL)
        print(f"변경 사항이 표시된 CSV 파일이 UTF-8로 저장되었습니다: {output_path}")
    except Exception as e:
        print(f"UTF-8 인코딩 저장 중 오류 발생: {e}")
        try:
            # ASCII로 저장 시도 (인코딩 불가능한 문자는 대체됩니다)
            with codecs.open(output_path, 'w', encoding='ascii', errors='replace') as f:
                df.to_csv(f, index=False, quoting=csv.QUOTE_ALL)
            print(f"변경 사항이 표시된 CSV 파일이 ASCII로 저장되었습니다: {output_path}")
        except Exception as e:
            print(f"ASCII 인코딩 저장 중 오류 발생: {e}")
            print("CSV 파일 저장에 실패했습니다.")

    return df

def main():
    print("프로그램 시작")
    image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_csv_path = os.path.join(output_dir, "marked_changes.csv")

    df_image = extract_table_from_image(image_path)
    df_pdf = extract_table_from_pdf(pdf_path)

    if df_image is not None and df_pdf is not None:
        changes = compare_dataframes(df_image, df_pdf)
        final_df = mark_changes_in_csv(df_image, changes, output_csv_path)
        print("최종 데이터프레임:")
        print(final_df)
    else:
        print("표 추출에 실패했습니다. 이미지와 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()