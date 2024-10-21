import os
import PyPDF2
import camelot
import pandas as pd
from sentence_transformers import SentenceTransformer
import numpy as np
import faiss

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        return {i+1: page.extract_text() for i, page in enumerate(reader.pages)}

def create_embeddings(texts):
    model = SentenceTransformer('distiluse-base-multilingual-cased-v1')
    return model.encode(texts)

def find_relevant_pages(texts_by_page, query, k=5):
    texts = list(texts_by_page.values())
    page_numbers = list(texts_by_page.keys())
    
    embeddings = create_embeddings(texts)
    query_embedding = create_embeddings([query])[0]
    
    index = faiss.IndexFlatIP(embeddings.shape[1])
    index.add(embeddings.astype('float32'))
    
    D, I = index.search(query_embedding.reshape(1, -1).astype('float32'), k=k)
    
    return [page_numbers[i] for i in I[0]]

def extract_tables_from_pages(pdf_path, pages):
    all_tables = []
    for page in pages:
        try:
            tables = camelot.read_pdf(pdf_path, pages=str(page), flavor='lattice')
            for table in tables:
                all_tables.append({
                    'dataframe': table.df,
                    'page': page
                })
        except Exception as e:
            print(f"Error extracting tables from page {page}: {e}")
    return all_tables

def save_tables_to_excel(tables, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, table in enumerate(tables):
            df = table['dataframe']
            sheet_name = f"Page_{table['page']}_Table_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Tables have been saved to {output_path}")

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"  # 실제 PDF 파일 경로로 변경하세요
    output_excel_path = "output/extracted_tables.xlsx"

    # PDF에서 텍스트 추출
    texts_by_page = extract_text_from_pdf(pdf_path)

    # RAG를 사용하여 "1종" 관련 페이지 찾기
    relevant_pages = find_relevant_pages(texts_by_page, "1종")
    print(f"'1종' 관련 페이지: {relevant_pages}")

    # 관련 페이지와 그 다음 페이지에서 표 추출
    pages_to_extract = set(relevant_pages + [p+1 for p in relevant_pages])
    tables = extract_tables_from_pages(pdf_path, pages_to_extract)

    # 추출된 표를 엑셀 파일로 저장
    if tables:
        save_tables_to_excel(tables, output_excel_path)
    else:
        print("추출된 표가 없습니다.")

if __name__ == "__main__":
    main()