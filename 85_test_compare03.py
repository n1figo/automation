import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from openpyxl import Workbook
import fitz  # PyMuPDF
import os

def extract_tables_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 선택특약 또는 선택약관 섹션을 찾습니다
    special_section = soup.find('p', string=re.compile(r'선택특약|선택약관'))
    
    if special_section:
        # 선택특약/선택약관 섹션 이후의 모든 표를 찾습니다
        tables = special_section.find_all_next('table', class_='tb_default03')
    else:
        # 선택특약/선택약관 섹션을 찾지 못한 경우, 모든 표를 추출합니다
        tables = soup.find_all('table', class_='tb_default03')
    
    return tables

def process_table(table, table_index):
    headers = [th.text.strip() for th in table.find_all('th')]
    if not headers:
        headers = [td.text.strip() for td in table.find('tr').find_all('td')]
    
    # 고유한 접두사를 추가하여 컬럼 이름을 고유하게 만듭니다
    headers = [f"Table_{table_index}_{header}" for header in headers]
    headers.append(f"Table_{table_index}_변경_여부")
    
    rows = []
    for tr in table.find_all('tr')[1:]:  # Skip the header row
        row = [td.text.strip() for td in tr.find_all('td')]
        if row:
            while len(row) < len(headers) - 1:  # -1 because we added '변경 여부'
                row.append('')
            row.append('')  # '변경 여부' 컬럼에 기본값 추가
            rows.append(row)
    
    print(f"Headers: {headers}")
    print(f"First row: {rows[0] if rows else 'No rows'}")
    print(f"Number of headers: {len(headers)}, Number of columns in first row: {len(rows[0]) if rows else 0}")
    
    return headers, rows

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
                                context = '\n'.join(lines[max(0, line_index-5):line_index+5])
                                highlighted_texts_with_context.append((context, highlighted_text))
    print("PDF에서 음영 처리된 텍스트 추출 완료")
    return highlighted_texts_with_context

def compare_dataframes(df_before, highlighted_texts_with_context):
    print("데이터프레임 비교 시작...")
    matching_rows = []

    for context, highlighted_text in highlighted_texts_with_context:
        # 모든 열에 대해 검색
        for col in df_before.columns:
            matches = df_before[df_before[col].astype(str).str.contains(highlighted_text, na=False)].index.tolist()
            matching_rows.extend(matches)

    matching_rows = sorted(set(matching_rows))
    df_matching = df_before.loc[matching_rows].copy()
    df_matching['하단 표 삽입요망'] = '하단 표 삽입요망'
    
    print(f"데이터프레임 비교 완료. {len(matching_rows)}개의 일치하는 행 발견")
    return df_matching

def main():
    print("프로그램 시작")
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_excel_path = os.path.join(output_dir, "matching_rows.xlsx")

    # HTML에서 표 추출
    response = requests.get(url)
    html_content = response.text
    tables = extract_tables_from_html(html_content)
    
    # 추출한 표를 DataFrame으로 변환
    df_list = []
    for i, table in enumerate(tables):
        headers, data = process_table(table, i)
        if data:  # 데이터가 있는 경우에만 DataFrame 생성
            df_table = pd.DataFrame(data, columns=headers)
            df_list.append(df_table)

    if not df_list:
        print("추출된 데이터가 없습니다. URL을 확인해주세요.")
        return

    df_before = pd.concat(df_list, axis=1)
    print("Combined DataFrame:")
    print(df_before.head())
    print(f"Shape of combined DataFrame: {df_before.shape}")

    highlighted_texts_with_context = extract_highlighted_text_with_context(pdf_path)

    if not df_before.empty and highlighted_texts_with_context:
        df_matching = compare_dataframes(df_before, highlighted_texts_with_context)
        df_matching.to_excel(output_excel_path, index=False)
        print(f"일치하는 행이 포함된 엑셀 파일이 저장되었습니다: {output_excel_path}")
    else:
        print("표 추출 또는 음영 처리된 텍스트 추출에 실패했습니다. URL과 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()