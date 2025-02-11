import re
import PyPDF2
import camelot  # 추가: Camelot 라이브러리

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

        for term, pages in results.items():
            if pages:
                print(f"Term '{term}' found on page(s): {pages}")
                # 각 페이지에서 lattice 모드를 사용하여 표 추출
                for page in pages:
                    print(f"Extracting tables for term '{term}' on page {page}:")
                    tables = camelot.read_pdf(file_path, pages=str(page), flavor="lattice")
                    for idx, table in enumerate(tables):
                        print(f"Table {idx + 1} from page {page}:")
                        print(table.df)
                        print("-" * 80)
            else:
                print(f"Term '{term}' not found in subsequent pages.")