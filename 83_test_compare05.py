import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import fitz  # PyMuPDF
import os
from PIL import Image
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

async def get_full_html(url, output_dir):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.goto(url, wait_until="networkidle")
        
        # Execute JavaScript to expand all hidden elements
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

def extract_tables_from_html(html_file_path):
    try:
        with open(html_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        tables = soup.find_all('table')
        
        if not tables:
            print("No tables found in the HTML content.")
            print(f"HTML content preview: {html_content[:1000]}...")  # Print first 1000 characters for debugging
            return []
        
        dfs = [pd.read_html(str(table))[0] for table in tables]
        print(f"Extracted {len(dfs)} tables.")
        return dfs
    except Exception as e:
        print(f"Error extracting tables: {str(e)}")
        print("Detailed error information:")
        import traceback
        print(traceback.format_exc())
        return []

# The rest of the functions remain the same

async def main():
    print("Program start")
    try:
        url = "https://www.kbinsure.co.kr/CG302120001.ec"
        pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
        output_dir = "/workspaces/automation/output"
        os.makedirs(output_dir, exist_ok=True)
        output_excel_path = os.path.join(output_dir, "comparison_results.xlsx")
        original_excel_path = os.path.join(output_dir, "변경전.xlsx")

        html_file_path = await get_full_html(url, output_dir)
        dfs = extract_tables_from_html(html_file_path)
        if not dfs:
            print("Failed to extract tables. Please check the HTML file.")
            return

        save_original_tables_to_excel(dfs, original_excel_path)

        df_before = process_tables(dfs)
        if df_before.empty:
            print("No data processed.")
            return

        print("Combined DataFrame:")
        print(df_before.head())
        print(f"Shape of combined DataFrame: {df_before.shape}")

        highlighted_texts_with_context = extract_highlighted_text_with_context(pdf_path, max_pages=20)

        if not df_before.empty and highlighted_texts_with_context:
            df_matching = compare_dataframes(df_before, highlighted_texts_with_context)
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