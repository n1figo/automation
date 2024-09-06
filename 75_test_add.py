import fitz
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import logging
import os
from typing import List, Tuple, Dict, Any
from collections import defaultdict
import re
import asyncio
from playwright.async_api import async_playwright
from PIL import Image
import numpy as np
import pdfplumber

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_text_for_excel(text):
    if isinstance(text, str):
        text = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text)
        text = text.replace('\n', ' ').replace('\r', '')
        text = re.sub(r'\s+', ' ', text)
        text = text.replace('•', '-').replace('Ⅱ', 'II')
    return text

class PDFTableExtractor:
    def __init__(self, pdf_path: str, tessdata_dir: str = None):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        if tessdata_dir:
            os.environ['TESSDATA_PREFIX'] = tessdata_dir
        
    def extract_tables_with_titles(self) -> List[Tuple[str, pd.DataFrame]]:
        all_tables = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            tables = self.extract_tables_from_page(page)
            titled_tables = self._assign_titles_to_tables(page, tables)
            all_tables.extend(titled_tables)
        return self._merge_tables_with_same_title(all_tables)
    
    def extract_tables_from_page(self, page: fitz.Page) -> List[Any]:
        tables = page.find_tables()
        return tables
    
    def _assign_titles_to_tables(self, page: fitz.Page, tables: List[Any]) -> List[Tuple[str, pd.DataFrame]]:
        titled_tables = []
        for table in tables:
            title = self._find_table_title(page, table)
            df = self._table_to_dataframe(table)
            titled_tables.append((title, df))
        return titled_tables
    
    def _find_table_title(self, page: fitz.Page, table: Any) -> str:
        blocks = page.get_text("dict")["blocks"]
        table_top = table.bbox[1]
        potential_titles = []
        for b in blocks:
            if 'lines' in b:
                for l in b['lines']:
                    for s in l['spans']:
                        if s['bbox'][3] < table_top and s['bbox'][3] > table_top - 50:
                            potential_titles.append(s['text'])
        
        if potential_titles:
            return " ".join(potential_titles).strip()
        return "Untitled Table"
    
    def _table_to_dataframe(self, table: Any) -> pd.DataFrame:
        df = pd.DataFrame(table.extract())
        df = df.applymap(clean_text_for_excel)
        return df
    
    def _merge_tables_with_same_title(self, tables: List[Tuple[str, pd.DataFrame]]) -> List[Tuple[str, pd.DataFrame]]:
        merged_tables = defaultdict(list)
        for title, df in tables:
            merged_tables[title].append(df)
        
        return [(title, pd.concat(dfs, ignore_index=True)) for title, dfs in merged_tables.items()]

async def get_full_html_and_tables(url, output_dir):
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page()
            await page.goto(url, wait_until="networkidle")
            
            await page.click('text="보장내용"')
            await page.wait_for_load_state('networkidle')
            
            html_content = await page.content()
            
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
            
            await browser.close()
        
        html_file_path = os.path.join(output_dir, "source.txt")
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        logger.info(f"Full HTML source has been saved to {output_dir}")
        return html_file_path, tables
    except Exception as e:
        logger.error(f"Error in get_full_html_and_tables: {str(e)}")
        return None, []

def process_tables(tables):
    all_data = []
    for i, table in enumerate(tables):
        df = pd.DataFrame(table[1:], columns=table[0])
        df['table_info'] = f'Table_{i+1}'
        all_data.append(df)
    
    if not all_data:
        logger.warning("No data extracted.")
        return pd.DataFrame()
    
    result = pd.concat(all_data, axis=0, ignore_index=True)
    
    logger.info(f"Final DataFrame: Columns: {result.columns.tolist()}, Shape: {result.shape}")
    
    return result

def is_color_highlighted(color):
    r, g, b = color
    if r == g == b:
        return False
    return max(r, g, b) > 200 and (max(r, g, b) - min(r, g, b)) > 30

def detect_highlights_and_colors(image):
    width, height = image.size
    img_array = np.array(image)
    
    highlighted_rows = set()
    colored_rows = set()
    for y in range(height):
        for x in range(width):
            color = img_array[y, x]
            if color[0] > 200 and color[1] > 200 and color[2] < 100:  # Yellow highlight
                highlighted_rows.add(y)
            elif is_color_highlighted(color):
                colored_rows.add(y)
    
    all_special_rows = sorted(highlighted_rows.union(colored_rows))
    return all_special_rows, highlighted_rows, colored_rows

def extract_highlighted_text_and_tables(pdf_path, output_dir):
    logger.info("Starting extraction of highlighted text and tables from PDF...")
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        texts_with_context = []
        
        for page_num, page in enumerate(pdf.pages):
            logger.info(f"Processing page {page_num + 1}/{total_pages}")
            
            img = page.to_image()
            pil_image = img.original
            
            all_special_rows, highlighted_rows, colored_rows = detect_highlights_and_colors(pil_image)
            
            text = page.extract_text()
            text_lines = text.split('\n')
            
            for i, line in enumerate(text_lines):
                status = '추가' if i in highlighted_rows or i in colored_rows else '유지'
                texts_with_context.append((line, page_num + 1, status))

    logger.info(f"Finished extracting text from PDF (total {total_pages} pages)")
    return texts_with_context

def compare_and_update_excel(df_before, texts_with_context, output_excel_path):
    logger.info("Starting comparison and Excel update...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"

    # Write headers
    headers = list(df_before.columns) + ['PDF_페이지', '상태']
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    row = 2
    for _, web_row in df_before.iterrows():
        web_text = ' '.join(str(cell) for cell in web_row if pd.notna(cell))
        matched = False
        for pdf_text, page_num, status in texts_with_context:
            if pdf_text.strip() in web_text:
                matched = True
                for col, value in enumerate(web_row, start=1):
                    ws.cell(row=row, column=col, value=value)
                ws.cell(row=row, column=len(web_row)+1, value=page_num)
                ws.cell(row=row, column=len(web_row)+2, value=status)
                break
        if not matched:
            for col, value in enumerate(web_row, start=1):
                ws.cell(row=row, column=col, value=value)
            ws.cell(row=row, column=len(web_row)+1, value='')
            ws.cell(row=row, column=len(web_row)+2, value='유지')
        row += 1

    wb.save(output_excel_path)
    logger.info(f"Comparison results saved to {output_excel_path}")

async def main():
    logger.info("Program start")
    try:
        url = "https://www.kbinsure.co.kr/CG302290001.ec#"
        pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
        output_dir = "/workspaces/automation/output"
        os.makedirs(output_dir, exist_ok=True)
        output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")

        html_file_path, tables = await get_full_html_and_tables(url, output_dir)
        
        df_before = pd.DataFrame()
        if tables:
            df_before = process_tables(tables)
            logger.info("Web tables extracted and processed")
        else:
            logger.warning("No tables extracted from the website. Proceeding with an empty DataFrame.")

        texts_with_context = extract_highlighted_text_and_tables(pdf_path, output_dir)

        if texts_with_context:
            compare_and_update_excel(df_before, texts_with_context, output_excel_path)
        else:
            logger.error("Failed to extract highlighted text from PDF. Please check the PDF file.")

    except Exception as e:
        logger.error(f"Error occurred: {str(e)}", exc_info=True)
    
    logger.info("Program end")

if __name__ == "__main__":
    asyncio.run(main())