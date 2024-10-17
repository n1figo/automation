import PyPDF2
from sentence_transformers import SentenceTransformer
import faiss
import numpy as np
import pandas as pd

# PDF에서 텍스트 추출 및 페이지 번호 추적
def extract_text_with_page_numbers(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        page_numbers = []
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            text += page_text + "\n"
            page_numbers.extend([i+1] * len(page_text.split()))
    return text, page_numbers

# 텍스트를 청크로 나누기
def split_text_into_chunks(text, chunk_size=200, overlap=50):
    words = text.split()
    chunks = []
    for i in range(0, len(words), chunk_size - overlap):
        chunk = " ".join(words[i:i + chunk_size])
        chunks.append(chunk)
    return chunks

# 임베딩 생성 및 인덱스 구축
def create_index(chunks):
    model = SentenceTransformer('distiluse-base-multilingual-cased')
    embeddings = model.encode(chunks)
    
    dimension = embeddings.shape[1]
    index = faiss.IndexFlatL2(dimension)
    index.add(embeddings.astype('float32'))
    
    return index, model

# 검색 함수 - 페이지 번호 포함
def search(query, index, model, chunks, page_numbers, k=5):
    query_vector = model.encode([query])
    distances, indices = index.search(query_vector.astype('float32'), k)
    
    results = []
    for idx in indices[0]:
        results.append({
            'content': chunks[idx],
            'page': page_numbers[idx]
        })
    
    return results

# 결과를 DataFrame으로 변환하는 함수
def results_to_dataframe(results, query_type):
    df = pd.DataFrame(results)
    df['type'] = query_type
    return df

# 특정 키워드가 포함된 페이지 찾기
def find_pages_with_keyword(text, keyword):
    pages = []
    for i, page in enumerate(text.split('\n')):
        if keyword in page:
            pages.append(i + 1)  # 페이지 번호는 1부터 시작
    return pages

# 메인 실행 코드
def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    text, page_numbers = extract_text_with_page_numbers(pdf_path)
    chunks = split_text_into_chunks(text)
    chunk_page_numbers = [page_numbers[i] for i in range(0, len(page_numbers), 200)]  # 청크 사이즈에 맞춰 조정
    index, model = create_index(chunks)

    # 선택특약 검색
    select_query = "선택특약"
    select_results = search(select_query, index, model, chunks, chunk_page_numbers)
    select_df = results_to_dataframe(select_results, "선택특약")

    # 상해관련특약 검색
    injury_query = "상해관련특약"
    injury_results = search(injury_query, index, model, chunks, chunk_page_numbers)
    injury_df = results_to_dataframe(injury_results, "상해관련특약")

    # 결과 합치기
    final_df = pd.concat([select_df, injury_df], ignore_index=True)

    # 엑셀 파일로 저장
    output_path = "insurance_special_clauses_search_results.xlsx"
    final_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"검색 결과가 {output_path}에 저장되었습니다.")

    # 키워드가 포함된 페이지 찾기 및 출력
    select_pages = find_pages_with_keyword(text, "선택특약")
    injury_pages = find_pages_with_keyword(text, "상해관련특약")
    injury_special_pages = find_pages_with_keyword(text, "상해관련 특별약관")

    print("선택특약이 포함된 페이지:", select_pages)
    print("상해관련특약이 포함된 페이지:", injury_pages)
    print("상해관련 특별약관이 포함된 페이지:", injury_special_pages)

if __name__ == "__main__":
    main()