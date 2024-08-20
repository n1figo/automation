import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from openpyxl import Workbook

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
    headers.append('변경 여부')
    
    rows = []
    for tr in table.find_all('tr'):
        row = [td.text.strip() for td in tr.find_all(['td', 'th'])]
        row.append('')  # '변경 여부' 컬럼에 기본값 추가
        if row:
            rows.append(row)
    
    return headers, rows

def create_excel_from_tables(tables, output_path):
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    for i, table in enumerate(tables, 1):
        headers, data = process_table(table)
        sheet = wb.create_sheet(title=f"Table {i}")
        
        # 헤더 추가
        sheet.append(headers)
        
        # 데이터 추가
        for row in data:
            sheet.append(row)

    wb.save(output_path)

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    response = requests.get(url)
    html_content = response.text
    
    tables = extract_tables_from_html(html_content)
    
    output_path = "/workspaces/automation/output/변경전_표_데이터.xlsx"
    create_excel_from_tables(tables, output_path)
    
    print(f"엑셀 파일이 생성되었습니다: {output_path}")

if __name__ == "__main__":
    main()