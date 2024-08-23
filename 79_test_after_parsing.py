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

# 나머지 함수들은 이전과 동일하게 유지...

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
            save_to_excel(df_matching, output_excel_path)
        else:
            print("Failed to extract highlighted text from PDF. Please check the PDF file.")

    except Exception as e:
        print(f"Error occurred: {str(e)}")
        print("Detailed error information:")
        import traceback
        print(traceback.format_exc())
    
    print("Program end")

if __name__ == "__main__":
    asyncio.run(main())