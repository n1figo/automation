import requests
import pandas as pd
import fitz  # PyMuPDF
import os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import StringIO

def extract_tables_from_html(url):
    response = requests.get(url)
    html_io = StringIO(response.text)
    dfs = pd.read_html(html_io)
    
    print(f"추출된 테이블 수: {len(dfs)}")
    return dfs

def process_tables(dfs):
    all_data = []
    for i, df in enumerate(dfs):
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(col).strip() for col in df.columns.values]
        df['table_info'] = f'Table_{i+1}_상세정보'
        all_data.append(df)
    
    if not all_data:
        print("추출된 데이터가 없습니다.")
        return pd.DataFrame()
    
    result = pd.concat(all_data, axis=0, ignore_index=True)
    
    print("Final DataFrame:")
    print(f"Columns: {result.columns.tolist()}")
    print(f"Shape: {result.shape}")
    
    return result

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
    
    df_matching['일치'] = '일치'
    df_matching['하단 표 삽입요망'] = '하단 표 삽입요망'
    df_matching['PDF_페이지'] = ''
    df_matching['이미지_경로'] = ''
    
    for i, (_, _, page, path) in enumerate(highlighted_texts_with_context):
        if i < len(df_matching):
            df_matching.iloc[i, df_matching.columns.get_loc('PDF_페이지')] = page
            df_matching.iloc[i, df_matching.columns.get_loc('이미지_경로')] = path
    
    print(f"데이터프레임 비교 완료. {len(matching_rows)}개의 일치하는 행 발견")
    return df_matching

def save_to_excel(df, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "비교 결과"

    row = 1
    for table_info in df['table_info'].unique():
        table_df = df[df['table_info'] == table_info]
        
        ws.cell(row=row, column=1, value=table_info)
        row += 1
        
        for col, header in enumerate(table_df.columns, start=1):
            ws.cell(row=row, column=col, value=header)
        row += 1
        
        for _, data_row in table_df.iterrows():
            for col, value in enumerate(data_row, start=1):
                ws.cell(row=row, column=col, value=value)
            row += 1
        
        row += 2

    wb.save(output_excel_path)
    print(f"결과가 {output_excel_path}에 저장되었습니다.")

def save_original_tables_to_excel(dfs, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "원본 테이블"

    row = 1
    for i, df in enumerate(dfs, start=1):
        # 테이블 이름 쓰기
        ws.cell(row=row, column=1, value=f'Table_{i}')
        row += 1

        # MultiIndex 처리
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(col).strip() for col in df.columns.values]

        # 헤더 쓰기
        for col, header in enumerate(df.columns, start=1):
            ws.cell(row=row, column=col, value=header)
        row += 1

        # 데이터 쓰기
        for _, data_row in df.iterrows():
            for col, value in enumerate(data_row, start=1):
                ws.cell(row=row, column=col, value=value)
            row += 1

        # 테이블 간 간격
        row += 2

    wb.save(output_excel_path)
    print(f"원본 테이블이 {output_excel_path}에 저장되었습니다.")

def main():
    print("프로그램 시작")
    try:
        url = "https://www.kbinsure.co.kr/CG302120001.ec"
        pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
        output_dir = "/workspaces/automation/output"
        os.makedirs(output_dir, exist_ok=True)
        output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")
        original_excel_path = os.path.join(output_dir, "변경전.xlsx")

        dfs = extract_tables_from_html(url)
        if not dfs:
            print("표 추출에 실패했습니다. URL을 확인해주세요.")
            return

        save_original_tables_to_excel(dfs, original_excel_path)

        df_before = process_tables(dfs)
        print("Combined DataFrame:")
        print(df_before.head())
        print(f"Shape of combined DataFrame: {df_before.shape}")

        highlighted_texts_with_context = extract_highlighted_text_with_context(pdf_path, max_pages=20)

        if not df_before.empty and highlighted_texts_with_context:
            df_matching = compare_dataframes(df_before, highlighted_texts_with_context)
            save_to_excel(df_matching, output_excel_path)
        else:
            print("표 추출 또는 음영 처리된 텍스트 추출에 실패했습니다. URL과 PDF를 확인해주세요.")

    except Exception as e:
        print(f"오류 발생: {str(e)}")
    
    print("프로그램 종료")

if __name__ == "__main__":
    main()