import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import fitz  # PyMuPDF
import os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup
from io import StringIO

async def get_full_html(url, output_dir):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(url, wait_until="networkidle")
        
        await page.evaluate("""
            () => {
                const expandElements = (elements) => {
                    for (let elem of elements) {
                        if (window.getComputedStyle(elem).display === 'none') {
                            elem.style.display = 'block';
                        }
                        expandElements(elem.children);
                    }
                };
                expandElements(document.body.children);
            }
        """)
        
        html_content = await page.content()
        await browser.close()
    
    html_file_path = os.path.join(output_dir, "source.txt")
    with open(html_file_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"Full HTML source has been saved to {html_file_path}")
    return html_file_path

def extract_tables_after_section(html_file_path, section_text):
    with open(html_file_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, 'html.parser')
    section = soup.find('strong', text=lambda text: section_text in text if text else False)

    if not section:
        print(f"Section '{section_text}' not found.")
        return []

    tables = []
    current_element = section.find_next()
    while current_element:
        if current_element.name == 'table':
            tables.append(current_element)
        elif current_element.name == 'strong' and '특별약관' in current_element.text:
            break
        current_element = current_element.find_next()

    dfs = [pd.read_html(StringIO(str(table)))[0] for table in tables]
    print(f"Extracted {len(dfs)} tables after '{section_text}'.")
    return dfs

def process_tables_with_status(dfs):
    all_data = []
    for i, df in enumerate(dfs):
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [' '.join(col).strip() for col in df.columns.values]
        df['table_info'] = f'Table_{i+1}'
        df['status'] = '유지'  # Default status
        all_data.append(df)
    
    if not all_data:
        print("No data extracted.")
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
    return max(r, g, b) > 200 or (max(r, g, b) - min(r, g, b)) > 30

def detect_highlights_and_colors(image):
    width, height = image.size
    img_array = np.array(image)
    
    highlighted_rows = set()
    colored_rows = set()
    for y in range(height):
        for x in range(width):
            color = img_array[y, x]
            if is_color_highlighted(color):
                highlighted_rows.add(y)
            elif not all(c == color[0] for c in color):  # Check if not black or gray
                colored_rows.add(y)
    
    highlighted_sections = []
    if highlighted_rows or colored_rows:
        all_special_rows = sorted(highlighted_rows.union(colored_rows))
        start_row = max(0, min(all_special_rows) - 10 * height // 100)
        end_row = min(height, max(all_special_rows) + 10 * height // 100)
        highlighted_sections.append((0, start_row, width, end_row))
    
    return highlighted_sections, highlighted_rows, colored_rows

def extract_highlighted_text_with_context(pdf_path, max_pages=20):
    print("Starting extraction of highlighted and colored text from PDF...")
    doc = fitz.open(pdf_path)
    total_pages = min(len(doc), max_pages)
    texts_with_context = []
    
    output_image_dir = os.path.join("output", "images")
    os.makedirs(output_image_dir, exist_ok=True)
    
    for page_num in range(total_pages):
        print(f"Processing page {page_num + 1}/{total_pages}")
        page = doc[page_num]
        
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        highlighted_sections, highlighted_rows, colored_rows = detect_highlights_and_colors(img)
        
        for section in highlighted_sections:
            x0, y0, x1, y1 = section
            
            section_img = img.crop(section)
            
            image_filename = f"page_{page_num + 1}_highlight.png"
            image_path = os.path.join(output_image_dir, image_filename)
            section_img.save(image_path)
            
            text = page.get_text("text", clip=section)
            if text.strip():
                context = page.get_text("text", clip=section)
                status = '수정' if highlighted_rows.intersection(range(y0, y1)) and colored_rows.intersection(range(y0, y1)) else \
                         '삭제' if highlighted_rows.intersection(range(y0, y1)) else \
                         '추가' if colored_rows.intersection(range(y0, y1)) else '유지'
                texts_with_context.append((context, text, page_num + 1, image_path, status))

    print(f"Finished extracting highlighted and colored text from PDF (total {total_pages} pages)")
    return texts_with_context

def compare_dataframes(df_before, texts_with_context):
    print("Starting comparison of dataframes...")
    matching_rows = []

    for context, text, page_num, image_path, status in texts_with_context:
        context_lines = context.split('\n')
        for i in range(len(df_before)):
            match = True
            for j, line in enumerate(context_lines):
                if i+j >= len(df_before) or not any(str(cell).strip() in line for cell in df_before.iloc[i+j]):
                    match = False
                    break
            if match:
                matching_rows.extend(range(i, i+len(context_lines)))
                df_before.loc[i:i+len(context_lines)-1, 'status'] = status
                break

    matching_rows = sorted(set(matching_rows))
    df_matching = df_before.loc[matching_rows].copy()
    
    df_matching['PDF_페이지'] = ''
    df_matching['이미지_경로'] = ''
    
    for i, (_, _, page, path, _) in enumerate(texts_with_context):
        if i < len(df_matching):
            df_matching.iloc[i, df_matching.columns.get_loc('PDF_페이지')] = page
            df_matching.iloc[i, df_matching.columns.get_loc('이미지_경로')] = path
    
    print(f"Finished comparison. Found {len(matching_rows)} matching rows")
    return df_matching

def save_to_excel(df, output_excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"

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
    print(f"Results have been saved to {output_excel_path}")

async def main():
    print("Program start")
    try:
        url = "https://www.kbinsure.co.kr/CG302120001.ec"
        pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
        output_dir = "/workspaces/automation/output"
        os.makedirs(output_dir, exist_ok=True)
        output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")

        html_file_path = await get_full_html(url, output_dir)
        dfs = extract_tables_after_section(html_file_path, "상해 관련 특별약관")
        if not dfs:
            print("Failed to extract tables. Please check the HTML file.")
            return

        df_before = process_tables_with_status(dfs)
        if df_before.empty:
            print("No data processed.")
            return

        print("Combined DataFrame:")
        print(df_before.head())
        print(f"Shape of combined DataFrame: {df_before.shape}")

        texts_with_context = extract_highlighted_text_with_context(pdf_path, max_pages=20)

        if not df_before.empty and texts_with_context:
            df_matching = compare_dataframes(df_before, texts_with_context)
            save_to_excel(df_matching, output_excel_path)
        else:
            print("Failed to extract tables or highlighted text. Please check the URL and PDF.")

    except Exception as e:
        print(f"Error occurred: {str(e)}")
        print("Detailed error information:")
        import traceback
        print(traceback.format_exc())
    
    print("Program end")

if __name__ == "__main__":
    asyncio.run(main())