import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

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
    
    all_data = []
    
    for table in tables:
        headers = [th.text.strip() for th in table.find_all('th')]
        
        # '변경 여부' 컬럼 추가
        headers.append('변경 여부')
        
        rows = []
        for tr in table.find_all('tr'):
            row = [td.text.strip() for td in tr.find_all(['td', 'th'])]
            
            # '변경 여부' 컬럼에 기본값 추가
            row.append('')
            
            if row:
                rows.append(row)
        
        all_data.extend(rows)
    
    return headers, all_data

def create_excel_from_tables(headers, data, output_path):
    df = pd.DataFrame(data, columns=headers)
    df.to_excel(output_path, index=False, engine='openpyxl')

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    response = requests.get(url)
    html_content = response.text
    
    headers, data = extract_tables_from_html(html_content)
    
    output_path = "/workspaces/automation/output/변경전_표_데이터.xlsx"
    create_excel_from_tables(headers, data, output_path)
    
    print(f"엑셀 파일이 생성되었습니다: {output_path}")

if __name__ == "__main__":
    main()