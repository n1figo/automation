import re
import PyPDF2
import camelot  # Camelot 라이브러리
import pandas as pd

file_path = "/workspaces/automation/data/input/0211/KB Yes!365 건강보험(세만기)(무배당)(25.01)_0214_요약서_v1.1.pdf"

search_term_initial = "나. 보험금"
search_terms = [
    "상해관련 특별약관",
    "질병관련 특별약관",
    "상해및질병관련특별약관"
]

def normalize(text):
    return re.sub(r'\s+', '', text)

with open(file_path, "rb") as pdf_file:
    reader = PyPDF2.PdfReader(pdf_file)
    start_page = None

    # 1단계: "나. 보험금"을 포함하는 페이지 찾기
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if text and normalize(search_term_initial) in normalize(text):
            start_page = i
            print(f"Found initial term '{search_term_initial}' on page {i + 1}")
            break

    if start_page is None:
        print(f"Initial term '{search_term_initial}' not found.")
    else:
        # 2단계: 초기 페이지 이후 범위에서 다른 세 검색어 페이지 번호 찾기
        results = {term: [] for term in search_terms}
        for i in range(start_page + 1, len(reader.pages)):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                normalized_text = normalize(text)
                for term in search_terms:
                    if normalize(term) in normalized_text:
                        results[term].append(i + 1)  # 1-indexed page number

        # 추출된 표를 각 특별약관 별로 구분하여 Excel 파일에 저장
        output_file = "/workspaces/automation/tests/test_data/output/extracted_tables.xlsx"
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        
        # 결과 시트: 각 검색어에 대해서 하나 이상의 시트에 데이터를 저장
        for term, pages in results.items():
            if pages:
                print(f"Term '{term}' found on page(s): {pages}")
                combined_df = pd.DataFrame()
                
                for page in pages:
                    print(f"Extracting tables for term '{term}' on page {page}:")
                    tables = camelot.read_pdf(file_path, pages=str(page), flavor="lattice")
                    for idx, table in enumerate(tables):
                        suffix = f" - P{page}T{idx+1}"
                        table_df = table.df.copy()
                        table_df.insert(0, "Source", term + suffix)
                        # Append table data to combined dataframe for 이 특별약관
                        combined_df = pd.concat([combined_df, table_df], ignore_index=True)
                        print(f"Table {idx + 1} from page {page} extracted.")
                
                # 시트 이름은 Excel의 제한(31자 이내) 고려
                sheet_name = term.replace(" ", "")[:31]
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data for '{term}' written to sheet '{sheet_name}'.")
            else:
                print(f"Term '{term}' not found in subsequent pages.")
        
        writer.close()
        print(f"Extracted tables have been saved to {output_file}")