import PyPDF2
from sentence_transformers import SentenceTransformer
import faiss
import numpy as np
import pandas as pd
import os
from fuzzywuzzy import fuzz

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        page_numbers = []
        for i, page in enumerate(reader.pages):
            page_text = page.extract_text()
            text += page_text + "\n"
            page_numbers.extend([i+1] * len(page_text.split()))
    return text, page_numbers

def extract_text_from_html(html_path):
    with open(html_path, 'r', encoding='utf-8') as file:
        text = file.read()
    return text

def split_text_into_chunks(text, chunk_size=200, overlap=50):
    words = text.split()
    chunks = []
    for i in range(0, len(words), chunk_size - overlap):
        chunk = " ".join(words[i:i + chunk_size])
        chunks.append(chunk)
    return chunks

def create_index(chunks):
    model = SentenceTransformer('distiluse-base-multilingual-cased')
    embeddings = model.encode(chunks)
    
    dimension = embeddings.shape[1]
    index = faiss.IndexFlatL2(dimension)
    index.add(embeddings.astype('float32'))
    
    return index, model

def search(query, index, model, chunks, page_numbers=None, k=5):
    query_vector = model.encode([query])
    distances, indices = index.search(query_vector.astype('float32'), k)
    
    results = []
    for idx in indices[0]:
        result = {'content': chunks[idx]}
        if page_numbers:
            result['page'] = page_numbers[idx]
        results.append(result)
    
    return results

def fuzzy_search(query, text, threshold=80):
    words = text.split()
    results = []
    for i, word in enumerate(words):
        if fuzz.partial_ratio(query, word) > threshold:
            start = max(0, i - 5)
            end = min(len(words), i + 6)
            context = " ".join(words[start:end])
            results.append(context)
    return results

def find_insurance_types(text):
    types = []
    for type in ["1종", "2종", "3종"]:
        if type in text:
            types.append(type)
    return types

def results_to_dataframe(results, query_type, text):
    df = pd.DataFrame(results)
    df['type'] = query_type
    
    # 보험 종류 찾기
    df['insurance_types'] = df['content'].apply(lambda x: ", ".join(find_insurance_types(x)))
    
    return df

def main():
    file_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    
    _, file_extension = os.path.splitext(file_path)
    
    if file_extension.lower() == '.pdf':
        text, page_numbers = extract_text_from_pdf(file_path)
    elif file_extension.lower() in ['.html', '.htm']:
        text = extract_text_from_html(file_path)
        page_numbers = None
    else:
        print("지원되지 않는 파일 형식입니다.")
        return

    chunks = split_text_into_chunks(text)
    index, model = create_index(chunks)

    # 선택특약 검색
    select_query = "선택특약"
    select_results = search(select_query, index, model, chunks, page_numbers)
    select_df = results_to_dataframe(select_results, "선택특약", text)

    # 상해관련 특약 검색 (퍼지 매칭 사용)
    injury_query = "상해관련특약"
    injury_results = fuzzy_search(injury_query, text)
    injury_df = pd.DataFrame({'content': injury_results, 'type': '상해관련특약'})
    injury_df['insurance_types'] = injury_df['content'].apply(lambda x: ", ".join(find_insurance_types(x)))

    # 결과 합치기
    final_df = pd.concat([select_df, injury_df], ignore_index=True)

    # 엑셀 파일로 저장
    output_path = "insurance_special_clauses_search_results.xlsx"
    final_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"검색 결과가 {output_path}에 저장되었습니다.")

    if page_numbers:
        select_pages = [page for page, chunk in zip(page_numbers, chunks) if "선택특약" in chunk]
        injury_pages = [page for page, chunk in zip(page_numbers, chunks) if any(result in chunk for result in injury_results)]

        print("선택특약이 포함된 페이지:", list(set(select_pages)))
        print("상해관련특약이 포함된 페이지:", list(set(injury_pages)))

if __name__ == "__main__":
    main()