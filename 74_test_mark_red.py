import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup
from io import StringIO
import pdfplumber


async def get_full_html_and_tables(url, output_dir):
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page()
            await page.goto(url, wait_until="networkidle")
            
            # '보장내용' 탭 클릭
            await page.click('text="보장내용"')
            await page.wait_for_load_state('networkidle')
            
            # JavaScript 실행 후 페이지 내용 가져오기
            html_content = await page.content()
            
            # 테이블 데이터 추출
            tables = await page.evaluate('''
                () => {
                    const tables = document.querySelectorAll('.tab_cont[data-list="보장내용"] table');
                    return Array.from(tables).map(table => {
                        const rows = table.querySelectorAll('tr');
                        return Array.from(rows).map(row => {
                            const cells = row.querySelectorAll('th, td');
                            return Array.from(cells).map(cell => cell.innerText);
                        });
                    });
                }
            ''')
            
            # 전체 페이지 스크린샷 캡처
            await page.screenshot(path=os.path.join(output_dir, "full_page.png"), full_page=True)
            
            # 테이블 영역 스크린샷 캡처
            table_elements = await page.query_selector_all('.tab_cont[data-list="보장내용"] table')
            for i, table_element in enumerate(table_elements):
                try:
                    await table_element.scroll_into_view_if_needed()
                    await page.wait_for_timeout(1000)
                    await table_element.screenshot(path=os.path.join(output_dir, f"table_{i+1}.png"))
                except Exception as e:
                    print(f"Failed to capture screenshot for table {i+1}: {str(e)}")
            
            await browser.close()
        
        html_file_path = os.path.join(output_dir, "source.txt")
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Full HTML source and table screenshots have been saved to {output_dir}")
        return html_file_path, tables
    except Exception as e:
        print(f"Error in get_full_html_and_tables: {str(e)}")
        return None, []

def process_tables(tables):
    all_data = []
    for i, table in enumerate(tables):
        df = pd.DataFrame(table[1:], columns=table[0])
        df['table_info'] = f'Table_{i+1}'
        all_data.append(df)
    
    if not all_data:
        print("No data extracted.")
        return pd.DataFrame()
    
    result = pd.concat(all_data, axis=0, ignore_index=True)
    
    print("Final DataFrame:")
    print(f"Columns: {result.columns.tolist()}")
    print(f"Shape: {result.shape}")
    
    return result

def save_original_tables_to_excel(dfs, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Original Tables"

    row = 1
    for i, df in enumerate(dfs, start=1):
        ws.cell(row=row, column=1, value=f'Table_{i}')
        row += 1

        for col, header in enumerate(df.columns, start=1):
            ws.cell(row=row, column=col, value=header)
        row += 1

        for _, data_row in df.iterrows():
            for col, value in enumerate(data_row, start=1):
                ws.cell(row=row, column=col, value=value)
            row += 1

        row += 2

    wb.save(output_excel_path)
    print(f"Original tables have been saved to {output_excel_path}")

def is_color_text(color):
    r, g, b = color
    is_black = r == g == b == 0
    is_white = r == g == b == 255
    is_grayscale = r == g == b
    return not (is_black or is_white or is_grayscale)

def detect_colored_text(image):
    width, height = image.size
    img_array = np.array(image)
    
    colored_rows = set()
    for y in range(height):
        for x in range(width):
            color = img_array[y, x]
            if is_color_text(color):
                colored_rows.add(y)
    
    return list(colored_rows)

def extract_highlighted_text_and_tables(pdf_path, output_dir):
    print("Starting extraction of colored text and tables from PDF...")
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        texts_with_context = []
        colored_pages = set()
        
        output_image_dir = os.path.join(output_dir, "images")
        os.makedirs(output_image_dir, exist_ok=True)
        
        for page_num, page in enumerate(pdf.pages):
            print(f"Processing page {page_num + 1}/{total_pages}")
            
            img = page.to_image()
            pil_image = img.original
            
            colored_rows = detect_colored_text(pil_image)
            
            if colored_rows:
                colored_pages.add(page_num + 1)
                
                image_filename = f"page_{page_num + 1}_full.png"
                image_path = os.path.join(output_image_dir, image_filename)
                pil_image.save(image_path)
                
                text = page.extract_text()
                texts_with_context.append((text, page_num + 1, image_path, '색상 텍스트 감지', colored_rows))
            else:
                text = page.extract_text()
                texts_with_context.append((text, page_num + 1, None, '유지', []))
        
        # 색상 텍스트가 있는 페이지의 표 추출
        table_data = {}
        for page_num in colored_pages:
            page = pdf.pages[page_num - 1]
            tables = page.extract_tables()
            if tables:
                for i, table in enumerate(tables):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    sheet_name = f"Page_{page_num}_Table_{i+1}"
                    table_data[sheet_name] = df
        
        if table_data:
            tables_excel_path = os.path.join(output_dir, "colored_text_tables.xlsx")
            with pd.ExcelWriter(tables_excel_path, engine='openpyxl') as writer:
                for sheet_name, df in table_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Saved all tables from pages with colored text to {tables_excel_path}")

    print(f"Finished extracting text and tables from PDF (total {total_pages} pages)")
    return texts_with_context

def compare_dataframes(df_before, texts_with_context, output_dir):
    print("Starting comparison of dataframes...")
    df_result = df_before.copy() if not df_before.empty else pd.DataFrame(columns=['table_info'])
    df_result['PDF_페이지'] = ''
    df_result['이미지_경로'] = ''
    df_result['상태'] = '유지'

    for text, page_num, image_path, status, colored_rows in texts_with_context:
        text_lines = text.split('\n')
        for i in range(len(df_result)):
            if any(str(cell).strip() in text for cell in df_result.iloc[i]):
                table_start = i
                while table_start > 0 and df_result.loc[table_start-1, 'table_info'] == df_result.loc[i, 'table_info']:
                    table_start -= 1
                table_end = i
                while table_end < len(df_result)-1 and df_result.loc[table_end+1, 'table_info'] == df_result.loc[i, 'table_info']:
                    table_end += 1
                
                df_result.loc[table_start:table_end, 'PDF_페이지'] = page_num
                df_result.loc[table_start:table_end, '이미지_경로'] = image_path
                
                if colored_rows:
                    df_result.loc[table_start:table_end, '상태'] = '색상 텍스트 감지'
                
                break

    cols = df_result.columns.tolist()
    cols.append(cols.pop(cols.index('상태')))
    df_result = df_result[cols]

    print(f"Finished comparison. Updated {len(df_result[df_result['상태'] != '유지'])} rows")
    return df_result

def save_to_excel(df, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"

    row = 1
    if 'table_info' in df.columns:
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
    else:
        # 'table_info' 열이 없는 경우 전체 DataFrame을 그대로 저장
        for col, header in enumerate(df.columns, start=1):
            ws.cell(row=row, column=col, value=header)
        row += 1
        
        for _, data_row in df.iterrows():
            for col, value in enumerate(data_row, start=1):
                ws.cell(row=row, column=col, value=value)
            row += 1

    wb.save(output_excel_path)
    print(f"Results have been saved to {output_excel_path}")


async def main():
    print("Program start")
    try:
        url = "https://www.kbinsure.co.kr/CG302290001.ec#"
        pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
        output_dir = "/workspaces/automation/output"
        os.makedirs(output_dir, exist_ok=True)
        output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")
        original_excel_path = os.path.join(output_dir, "변경전.xlsx")

        html_file_path, tables = await get_full_html_and_tables(url, output_dir)
        
        df_before = pd.DataFrame()
        if tables:
            df_before = process_tables(tables)
            save_original_tables_to_excel([df_before], original_excel_path)
            print("Combined DataFrame:")
            print(df_before.head())
            print(f"Shape of combined DataFrame: {df_before.shape}")
        else:
            print("No tables extracted from the website. Proceeding with an empty DataFrame.")

        texts_with_context = extract_highlighted_text_and_tables(pdf_path, output_dir)

        if texts_with_context:
            df_matching = compare_dataframes(df_before, texts_with_context, output_dir)
            print("Columns in df_matching:", df_matching.columns)
            save_to_excel(df_matching, output_excel_path)
        else:
            print("Failed to extract colored text from PDF. Please check the PDF file.")

    except Exception as e:
        print(f"Error occurred: {str(e)}")
        print("Detailed error information:")
        import traceback
        print(traceback.format_exc())
    
    print("Program end")

if __name__ == "__main__":
    asyncio.run(main())