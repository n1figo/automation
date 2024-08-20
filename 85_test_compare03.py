import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from openpyxl import Workbook
import fitz  # PyMuPDF
import os
from PIL import Image
import numpy as np

def extract_tables_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    special_section = soup.find('p', string=re.compile(r'선택특약|선택약관'))
    
    if special_section:
        tables = special_section.find_all_next('table', class_='tb_default03')
    else:
        tables = soup.find_all('table', class_='tb_default03')
    
    return tables

def process_table(table, table_index):
    headers = [th.text.strip() for th in table.find_all('th')]
    if not headers:
        headers = [td.text.strip() for td in table.find('tr').find_all('td')]
    
    headers = [f"Table_{table_index}_{header}" for header in headers]
    headers.append(f"Table_{table_index}_변경_여부")
    
    rows = []
    for tr in table.find_all('tr')[1:]:
        row = [td.text.strip() for td in tr.find_all('td')]
        if row:
            while len(row) < len(headers) - 1:
                row.append('')
            row.append('')
            rows.append(row)
    
    print(f"Headers: {headers}")
    print(f"First row: {rows[0] if rows else 'No rows'}")
    print(f"Number of headers: {len(headers)}, Number of columns in first row: {len(rows[0]) if rows else 0}")
    
    return headers, rows

def is_color_highlighted(color):
    r, g, b = color
    if r == g == b:
        return False
    return max(r, g, b) > 200 and (max(r, g, b) - min(r, g, b)) > 30

def detect_highlights(image):
    width, height = image.size
    img_array = np.array(image)
    
    highlighted_rows = set()
    for y in range(height):
        for x in range(width):
            if is_color_highlighted(img_array[y, x]):
                highlighted_rows.add(y)
    
    if highlighted_rows:
        start_row = max(0, min(highlighted_rows) - 10 * height // 100)
        end_row = min(height, max(highlighted_rows) + 10 * height // 100)
        return [(0, start_row, width, end_row)]
    
    return []

def extract_highlighted_text_with_context(pdf_path, max_pages=20):
    print("PDF에서 음영 처리된 텍스트 추출 시작...")
    doc = fitz.open(pdf_path)
    total_pages = min(len(doc), max_pages)
    highlighted_texts_with_context = []
    
    output_image_dir = os.path.join("output", "images")
    os.makedirs(output_image_dir, exist_ok=True)
    
    for page_num in range(total_pages):
        print(f"처리 중: {page_num + 1}/{total_pages} 페이지")
        page = doc[page_num]
        
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        highlighted_sections = detect_highlights(img)
        
        if highlighted_sections:
            section = highlighted_sections[0]
            x0, y0, x1, y1 = section
            
            highlight_img = img.crop(section)
            
            image_filename = f"page_{page_num + 1}_highlight.png"
            image_path = os.path.join(output_image_dir, image_filename)
            highlight_img.save(image_path)
            
            text = page.get_text("text", clip=section)
            if text.strip():
                context = page.get_text("text", clip=section)
                highlighted_texts_with_context.append((context, text, page_num + 1, image_path))

    print(f"PDF에서 음영 처리된 텍스트 추출 완료 (총 {total_pages} 페이지)")
    return highlighted_texts_with_context

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def process_tables(tables):
    all_data = []
    for i, table in enumerate(tables):
        headers, data = process_table(table, i)
        df = pd.DataFrame(data, columns=headers)
        
        # 구체적인 항목 추가 (예시)
        df['구체적인_항목'] = f'Table_{i+1}_상세정보'
        
        all_data.append(df)
    
    return pd.concat(all_data, axis=0, ignore_index=True)

def compare_dataframes(df_before, highlighted_texts_with_context):
    print("데이터프레임 비교 시작...")
    matching_rows = []

    for context, highlighted_text, page_num, image_path in highlighted_texts_with_context:
        context_lines = context.split('\n')
        for i in range(len(df_before)):
            match = True
            for j, line in enumerate(context_lines):
                if i+j >= len(df_before) or not any(str(cell).strip() in line for cell in df_before.iloc[i+j]):
                    match = False
                    break
            if match:
                matching_rows.extend(range(i, i+len(context_lines)))
                break

    matching_rows = sorted(set(matching_rows))
    df_matching = df_before.loc[matching_rows].copy()
    
    # 새 열 추가 (기본값으로 초기화)
    df_matching['일치'] = '일치'
    df_matching['하단 표 삽입요망'] = '하단 표 삽입요망'
    df_matching['PDF_페이지'] = ''
    df_matching['이미지_경로'] = ''
    
    # 사용 가능한 데이터로 열 업데이트
    for i, (_, _, page, path) in enumerate(highlighted_texts_with_context):
        if i < len(df_matching):
            df_matching.loc[df_matching.index[i], 'PDF_페이지'] = page
            df_matching.loc[df_matching.index[i], '이미지_경로'] = path
    
    print(f"데이터프레임 비교 완료. {len(matching_rows)}개의 일치하는 행 발견")
    return df_matching

def save_to_excel(df, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "비교 결과"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    wb.save(output_excel_path)
    print(f"결과가 {output_excel_path}에 저장되었습니다.")

def main():
    print("프로그램 시작")
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)
    output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")

    # HTML에서 표 추출
    response = requests.get(url)
    html_content = response.text
    tables = extract_tables_from_html(html_content)
    
    # 모든 표를 하나의 DataFrame으로 처리
    df_before = process_tables(tables)
    print("Combined DataFrame:")
    print(df_before.head())
    print(f"Shape of combined DataFrame: {df_before.shape}")

    highlighted_texts_with_context = extract_highlighted_text_with_context(pdf_path, max_pages=20)

    if not df_before.empty and highlighted_texts_with_context:
        df_matching = compare_dataframes(df_before, highlighted_texts_with_context)
        save_to_excel(df_matching, output_excel_path)
    else:
        print("표 추출 또는 음영 처리된 텍스트 추출에 실패했습니다. URL과 PDF를 확인해주세요.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()