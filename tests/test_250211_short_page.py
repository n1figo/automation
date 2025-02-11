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

# 두번째, 세번째, 네번째 컬럼에 헤더가 있는 행을 판별하는 함수
# 추출된 테이블의 열 개수가 3개인 경우와 4개 이상인 경우를 모두 처리합니다.
def is_header_row(row, header=["보장명", "지급사유", "지급금액"]):
    try:
        # row의 길이에 따라 비교할 column 인덱스 결정
        if len(row) >= 4:
            # 4개 이상이면 두번째~네번째 컬럼 (인덱스 1,2,3) 비교
            cells = [normalize(row[i]) for i in range(1, 4)]
        elif len(row) == 3:
            # 3개이면 첫번째~세번째 컬럼 (인덱스 0,1,2) 비교
            cells = [normalize(row[i]) for i in range(0, 3)]
        else:
            return False
        norm_header = [normalize(h) for h in header]
        print(f"[DEBUG] Checking row: {cells} against expected: {norm_header}")
        return cells == norm_header
    except Exception as e:
        print(f"[ERROR] Exception in is_header_row: {e}")
        return False

# DataFrame에서 헤더 행(중간에 삽입된 행)을 제거하는 함수 + 로그 추가
def drop_redundant_header(df, header=["보장명", "지급사유", "지급금액"]):
    keep_rows = []
    for idx, row in df.iterrows():
        if is_header_row(row, header):
            print(f"[INFO] Dropping row {idx} as header row")
        else:
            keep_rows.append(idx)
    return df.loc[keep_rows]

with open(file_path, "rb") as pdf_file:
    reader = PyPDF2.PdfReader(pdf_file)
    start_page = None

    # 1단계: "나. 보험금"이 포함된 페이지 찾기
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if text and normalize(search_term_initial) in normalize(text):
            start_page = i
            print(f"Found initial term '{search_term_initial}' on page {i + 1}")
            break

    if start_page is None:
        print(f"Initial term '{search_term_initial}' not found.")
    else:
        # 2단계: 초기 페이지 이후 범위에서 각 검색어가 등장하는 페이지 번호 찾기
        results = {term: [] for term in search_terms}
        for i in range(start_page + 1, len(reader.pages)):
            page = reader.pages[i]
            text = page.extract_text()
            if text:
                normalized_text = normalize(text)
                for term in search_terms:
                    if normalize(term) in normalized_text:
                        results[term].append(i + 1)  # 1-indexed page number

        # 각 특별약관별 추출 테이블을 별도 시트에 저장 (전처리 개선 후 로그 추가)
        output_file = "/workspaces/automation/tests/test_data/output/extracted_tables.xlsx"
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

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
                        print(f"[DEBUG] Original table df from page {page} table {idx+1}:")
                        print(table_df)
                        # 헤더행 제거 전 normalize 결과와 함께 로그 기록
                        table_df = drop_redundant_header(table_df)
                        table_df.insert(0, "Source", term + suffix)
                        combined_df = pd.concat([combined_df, table_df], ignore_index=True)
                        print(f"Table {idx+1} from page {page} extracted.")
                sheet_name = term.replace(" ", "")[:31]
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data for '{term}' written to sheet '{sheet_name}'.")
            else:
                print(f"Term '{term}' not found in subsequent pages.")

        writer.close()
        print(f"Extracted tables have been saved to {output_file}")