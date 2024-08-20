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

def process_table(table):
    headers = [th.text.strip() for th in table.find_all('th')]
    if not headers:
        headers = [td.text.strip() for td in table.find('tr').find_all('td')]
    headers.append('변경 여부')
    
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

# ... (나머지 함수들은 그대로 유지)

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
    for table in tables:
        headers, data = process_table(table)
        if data:  # 데이터가 있는 경우에만 DataFrame 생성
            df_table = pd.DataFrame(data, columns=headers)
            df_list.append(df_table)

    if not df_list:
        print("추출된 데이터가 없습니다. URL을 확인해주세요.")
        return

    df_before = pd.concat(df_list, ignore_index=True)

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