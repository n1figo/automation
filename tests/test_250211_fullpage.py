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
    if not isinstance(text, str):
        text = str(text)
    return re.sub(r'\s+', '', text)

def is_header_row(row, header=["보장명", "지급사유", "지급금액"]):
    try:
        # 행의 열 개수에 따라 비교 대상 결정: 4개 이상이면 인덱스 1~3, 3개이면 인덱스 0~2 사용
        if len(row) >= 4:
            cells = [normalize(row[i]) for i in range(1, 4)]
        elif len(row) == 3:
            cells = [normalize(row[i]) for i in range(0, 3)]
        else:
            return False
        norm_header = [normalize(h) for h in header]
        print(f"[DEBUG] Checking row: {cells} against expected: {norm_header}")
        return cells == norm_header
    except Exception as e:
        print(f"[ERROR] Exception in is_header_row: {e}")
        return False

def drop_redundant_header(df, header=["보장명", "지급사유", "지급금액"]):
    keep_rows = []
    for idx, row in df.iterrows():
        if is_header_row(row, header):
            print(f"[INFO] Dropping row {idx} as header row")
        else:
            keep_rows.append(idx)
    return df.loc[keep_rows]

def main():
    with open(file_path, "rb") as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        start_page = None

        # 1단계: "나. 보험금"을 포함하는 첫 페이지 찾기
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text and normalize(search_term_initial) in normalize(text):
                start_page = i
                print(f"Found initial term '{search_term_initial}' on page {i+1}")
                break

        if start_page is None:
            print(f"Initial term '{search_term_initial}' not found.")
            return

        # 2단계: 초기 페이지 이후부터 각 검색어에 해당하는 페이지 번호 찾기
        results = {term: [] for term in search_terms}
        for i in range(start_page + 1, len(reader.pages)):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                normalized_text = normalize(text)
                for term in search_terms:
                    if normalize(term) in normalized_text:
                        results[term].append(i + 1)  # 페이지 번호: 1-indexed

        # 각 특별약관별로 추출 테이블을 별도 시트에 저장 (전체 페이지를 한 번에 파싱)
        output_file = "/workspaces/automation/tests/test_data/output/extracted_tables.xlsx"
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        for term, pages in results.items():
            if pages:
                pages_str = ",".join(map(str, pages))
                print(f"Term '{term}' found on page(s): {pages_str}")
                print(f"Extracting tables for term '{term}' on pages {pages_str}:")
                tables = camelot.read_pdf(file_path, pages=pages_str, flavor="lattice")
                combined_df = pd.DataFrame()

                for idx, table in enumerate(tables):
                    suffix = f" - P{pages_str}T{idx+1}"
                    table_df = table.df.copy()
                    print(f"[DEBUG] Original table df from pages {pages_str} table {idx+1}:")
                    print(table_df)
                    table_df = drop_redundant_header(table_df)
                    table_df.insert(0, "Source", term + suffix)
                    combined_df = pd.concat([combined_df, table_df], ignore_index=True)
                    print(f"Table {idx+1} from pages {pages_str} extracted.")

                sheet_name = term.replace(" ", "")[:31]
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data for '{term}' written to sheet '{sheet_name}'.")
            else:
                print(f"Term '{term}' not found in subsequent pages.")

        writer.close()
        print(f"Extracted tables have been saved to {output_file}")

if __name__ == "__main__":
    main()