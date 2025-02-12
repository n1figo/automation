import re
import PyPDF2
import camelot  # Camelot 라이브러리
import pandas as pd
import fitz     # PyMuPDF

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

def page_has_highlight(doc, page_no):
    """
    PyMuPDF를 사용하여 페이지에 하이라이트 annotation이 있는지 체크합니다.
    annotation이 존재하고, 그 유형이 하이라이트(일반적으로 type id 8)인 경우 True를 반환합니다.
    """
    page = doc.load_page(page_no)
    annots = page.annots()
    if annots:
        for annot in annots:
            try:
                # annot.type는 tuple 형태로 (type_value, type_name)으로 제공됩니다.
                if annot.type[0] == 8:
                    return True
            except Exception as ex:
                print(f"[ERROR] Annotation error: {ex}")
    return False

def main():
    # 1. PyPDF2로 PDF 전체 읽기
    with open(file_path, "rb") as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(reader.pages)
        start_page = None

        # "나. 보험금"을 포함하는 첫 페이지 찾기
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text and normalize(search_term_initial) in normalize(text):
                start_page = i
                print(f"Found initial term '{search_term_initial}' on page {i+1}")
                break

        if start_page is None:
            print(f"Initial term '{search_term_initial}' not found.")
            return

        # 2. 초기 페이지 이후부터 각 섹션에 해당하는 페이지 번호 찾기
        results = {term: [] for term in search_terms}
        for i in range(start_page + 1, total_pages):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                normalized_text = normalize(text)
                for term in search_terms:
                    if normalize(term) in normalized_text:
                        results[term].append(i + 1)   # 페이지 번호: 1-indexed

    # 3. 하이라이트 인식 범위 결정: "나. 보험금" 페이지부터 "상해및질병관련특별약관" 마지막 페이지까지
    highlight_end_page = total_pages  # 기본적으로 전체 페이지
    if results["상해및질병관련특별약관"]:
        highlight_end_page = max(results["상해및질병관련특별약관"])
    print(f"Highlight detection range: from page {start_page+1} to page {highlight_end_page}")

    # 4. PyMuPDF로 지정 범위 내에서 하이라이트(또는 색깔 글씨) annotation이 있는 페이지와 그 전후 페이지 찾기
    doc = fitz.open(file_path)
    highlight_pages = set()
    for i in range(start_page, highlight_end_page):
        if page_has_highlight(doc, i):
            highlight_pages.add(i)  # 현재 페이지
            if i - 1 >= start_page:
                highlight_pages.add(i - 1)
            if i + 1 < highlight_end_page:
                highlight_pages.add(i + 1)
    if highlight_pages:
        highlight_pages_sorted = sorted(list(highlight_pages))
        hp_str = ",".join(str(p+1) for p in highlight_pages_sorted)
        print(f"Pages with highlights (and adjacent pages): {hp_str}")
    else:
        print("No highlight annotations found within the specified range.")

    # 5. 각 특별약관별로 표 추출 후 Excel 시트에 저장 (기존 방식)
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