import os
import re
import camelot
import fitz  # PyMuPDF
import pandas as pd
from sentence_transformers import SentenceTransformer
import faiss
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ----------------------------
# 1. 텍스트 추출 및 페이지 매핑
# ----------------------------
def extract_text_by_page(pdf_path):
    """
    PDF 파일에서 각 페이지의 텍스트를 추출하여 딕셔너리로 반환합니다.
    """
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        texts_by_page = {}
        for i, page in enumerate(reader.pages):
            page_num = i + 1
            page_text = page.extract_text()
            texts_by_page[page_num] = page_text if page_text else ""
    return texts_by_page

# ----------------------------
# 2. FAISS 인덱스 구축
# ----------------------------
def create_faiss_index(texts, model_name='distiluse-base-multilingual-cased'):
    """
    텍스트 리스트를 임베딩하고 FAISS 인덱스를 구축하여 반환합니다.
    """
    model = SentenceTransformer(model_name)
    embeddings = model.encode(list(texts.values()), convert_to_numpy=True)
    embeddings = embeddings / np.linalg.norm(embeddings, axis=1, keepdims=True)  # 정규화

    dimension = embeddings.shape[1]
    index = faiss.IndexFlatIP(dimension)  # 내적 기반 인덱스
    index.add(embeddings.astype('float32'))

    return index, model

# ----------------------------
# 3. RAG을 사용한 검색
# ----------------------------
def search_faiss(query, index, model, texts, top_k=5, threshold=0.5):
    """
    주어진 쿼리에 대해 FAISS 인덱스를 검색하여 유사한 페이지를 반환합니다.
    """
    query_embedding = model.encode([query], convert_to_numpy=True)
    query_embedding = query_embedding / np.linalg.norm(query_embedding, axis=1, keepdims=True)  # 정규화

    distances, indices = index.search(query_embedding.astype('float32'), top_k)
    results = []
    for idx, score in zip(indices[0], distances[0]):
        if score >= threshold:
            page_num = list(texts.keys())[idx]
            results.append({'page': page_num, 'score': score, 'content': texts[page_num]})
    return results

# ----------------------------
# 4. 표 위의 텍스트 추출
# ----------------------------
def extract_text_above_bbox(page, bbox):
    """
    주어진 표의 바운딩 박스 위에 있는 텍스트를 추출합니다.
    """
    x0, y0, x1, y1 = bbox
    text_blocks = page.get_text("blocks")
    texts_above = []
    for block in text_blocks:
        bx0, by0, bx1, by1, text = block[:5]
        if by1 <= y0:
            texts_above.append((by1, text.strip()))
    if texts_above:
        # y 좌표가 가장 큰 (표 바로 위에 있는) 텍스트 선택
        texts_above.sort(reverse=True)
        return texts_above[0][1]
    else:
        return "제목 없음"

# ----------------------------
# 5. Camelot과 PyMuPDF를 사용한 표 및 제목 추출
# ----------------------------
def extract_tables_with_titles(pdf_path, page_num):
    """
    주어진 페이지에서 Camelot을 사용하여 표를 추출하고, PyMuPDF를 사용하여 표 위의 제목을 추출합니다.
    """
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='lattice')
    if not tables:
        return []

    doc = fitz.open(pdf_path)
    page_obj = doc.load_page(page_num - 1)

    extracted_tables = []
    for table in tables:
        df = table.df
        bbox = table._bbox  # (x1, y1, x2, y2)
        title = extract_text_above_bbox(page_obj, bbox)
        extracted_tables.append({'title': title, 'dataframe': df, 'page': page_num})
    return extracted_tables

# ----------------------------
# 6. 엑셀 파일에 저장
# ----------------------------
def save_to_excel(tables_dict, output_path, document_title=None):
    """
    추출된 표들을 각 종류별 시트에 저장합니다.
    """
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    for sheet_name, tables in tables_dict.items():
        ws = wb.create_sheet(title=sheet_name)

        if document_title:
            ws.cell(row=1, column=1, value=f"{document_title} - {sheet_name}")
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
            ws.cell(row=1, column=1).font = Font(size=16, bold=True)
            ws.row_dimensions[1].height = 30
            current_row = 2
        else:
            current_row = 1

        for table in tables:
            title = table['title']
            df = table['dataframe']
            page = table['page']

            # 표 제목 추가
            ws.cell(row=current_row, column=1, value=f"{title} (페이지 {page})")
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=df.shape[1])
            ws.cell(row=current_row, column=1).font = Font(bold=True)
            current_row += 1

            # 표 데이터 작성
            for r in dataframe_to_rows(df, index=False, header=True):
                for c_idx, value in enumerate(r, start=1):
                    ws.cell(row=current_row, column=c_idx, value=value)
                    ws.cell(row=current_row, column=c_idx).alignment = Alignment(wrap_text=True)
                current_row += 1
            current_row += 1  # 표 사이에 빈 줄 추가

        # 열 너비 자동 조정
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
            adjusted_width = min(length + 2, 50)
            column_letter = get_column_letter(column_cells[0].column)
            ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_path)
    print(f"Tables have been saved to {output_path}")

# ----------------------------
# 7. 메인 함수
# ----------------------------
def main():
    import PyPDF2
    import numpy as np

    # 경로 설정
    uploads_folder = "uploads"
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)

    # PDF 파일 찾기
    pdf_files = [f for f in os.listdir(uploads_folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print("No PDF files found in the uploads folder.")
        return

    pdf_file = pdf_files[0]
    pdf_path = os.path.join(uploads_folder, pdf_file)
    output_excel_path = os.path.join(output_folder, f"{os.path.splitext(pdf_file)[0]}_analysis.xlsx")

    print(f"Processing PDF file: {pdf_file}")

    # 1. 페이지별 텍스트 추출
    texts_by_page = extract_text_by_page(pdf_path)

    # 2. FAISS 인덱스 생성
    index, model = create_faiss_index(texts_by_page)

    # 3. '선택특약' 검색
    select_query = "선택특약"
    select_results = search_faiss(select_query, index, model, texts_by_page, top_k=10, threshold=0.5)
    select_pages = [result['page'] for result in select_results]

    print("선택특약이 포함된 페이지:", select_pages)

    # 4. 각 선택특약 페이지 처리
    doc = fitz.open(pdf_path)
    tables_dict = {'1종': [], '2종': [], '3종': []}

    for page_num in select_pages:
        page_text = texts_by_page[page_num]
        print(f"\nProcessing page {page_num} for 선택특약...")

        # 2. 해당 페이지의 전체 텍스트 추출 및 [1종], [2종], [3종] 확인
        types_found = []
        for insurance_type in ['1종', '2종', '3종']:
            if re.search(rf'\[{insurance_type}\]', page_text):
                types_found.append(insurance_type)
                print(f"{insurance_type}이(가) 페이지 {page_num}에 존재합니다.")

        if not types_found:
            print(f"페이지 {page_num}에 [1종], [2종], [3종]이 없습니다.")
            continue

        # 3. '질병관련 특별약관' 확인하여 표 추출 중단
        if re.search(r'질병관련\s*특별약관', page_text):
            print(f"페이지 {page_num}에 '질병관련 특별약관'이 발견되어 표 추출을 중단합니다.")
            break

        # 4. Camelot과 PyMuPDF를 사용하여 표 및 제목 추출
        extracted_tables = extract_tables_with_titles(pdf_path, page_num)
        if not extracted_tables:
            print(f"페이지 {page_num}에서 표를 찾을 수 없습니다.")
            continue

        # 5. 추출된 표를 해당 종류에 할당
        for table in extracted_tables:
            for insurance_type in types_found:
                tables_dict[insurance_type].append(table)
                print(f"페이지 {page_num}에서 추출된 표을 [ {insurance_type} ] 시트에 추가했습니다.")

    # 6. 엑셀 파일로 저장
    if any(tables_dict.values()):
        # 문서 제목 추출 (첫 페이지 첫 번째 줄)
        first_page = doc.load_page(0)
        first_page_text = extract_text_above_bbox(first_page, first_page.bound())
        document_title = first_page_text if first_page_text else "Document Title"

        save_to_excel(tables_dict, output_excel_path, document_title=document_title)
    else:
        print("저장할 표가 없습니다.")

# ----------------------------
# 8. 스크립트 실행
# ----------------------------
if __name__ == "__main__":
    main()
