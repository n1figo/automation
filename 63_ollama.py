import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import json
from PIL import Image
import io
import base64
import warnings
from difflib import SequenceMatcher

from langchain_community.chat_models import ChatOllama
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain.schema import HumanMessage, Document

# 경고 메시지 무시 설정
warnings.filterwarnings("ignore", category=FutureWarning)

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 및 로그 저장 경로 설정
OUTPUT_DIR = "/workspaces/automation/output"
IMAGE_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "images")
LOG_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "logs")
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)
os.makedirs(LOG_OUTPUT_DIR, exist_ok=True)

# Ollama 모델 로드
model = ChatOllama(model="llava", temperature=0)

def log_to_file(message, filename="debug_log.txt"):
    with open(os.path.join(LOG_OUTPUT_DIR, filename), "a") as f:
        f.write(message + "\n")

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def detect_highlights(image):
    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    
    # 노란색 범위 정의 (예시, 필요에 따라 조정)
    lower_yellow = np.array([20, 100, 100])
    upper_yellow = np.array([30, 255, 255])
    
    mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
    
    kernel = np.ones((5,5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    highlighted_regions = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        highlighted_regions.append((x, y, x+w, y+h))

    return highlighted_regions

def extract_tables_to_excel(pdf_path, output_excel_path):
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    tables = page.find_tables()
    
    all_data = []

    for table in tables:
        df = pd.DataFrame(table.extract())
        all_data.append(df)

    combined_df = pd.concat(all_data, ignore_index=True)
    combined_df['변경사항'] = ''
    combined_df.to_excel(output_excel_path, index=False)
    log_to_file(f"테이블이 추출되어 '{output_excel_path}'에 저장되었습니다.")

def extract_highlighted_text_ollama(image, regions):
    extracted_texts = []

    for region in regions:
        x0, y0, x1, y1 = region
        cropped_image = Image.fromarray(image[y0:y1, x0:x1])
        
        buffered = io.BytesIO()
        cropped_image.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        prompt = f"이 이미지에서 강조된 텍스트를 추출해주세요. 강조된 영역의 좌표는 ({x0}, {y0}, {x1}, {y1})입니다."
        
        message = HumanMessage(
            content=[
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": f"data:image/png;base64,{img_str}"}
            ]
        )
        response = model.invoke([message])
        
        extracted_texts.append(response.content)

    return extracted_texts

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def create_vector_store(documents):
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=100,
    )
    docs = text_splitter.split_documents(documents)
    
    embeddings = HuggingFaceEmbeddings(
        model_name='jhgan/ko-sroberta-nli',
        model_kwargs={'device':'cpu'},
        encode_kwargs={'normalize_embeddings':True},
    )
    vectorstore = Chroma.from_documents(docs, embeddings)
    return vectorstore

def find_highlighted_text_in_excel(excel_path, highlighted_chunks):
    df = pd.read_excel(excel_path)
    
    documents = [Document(page_content=row.to_string()) for _, row in df.iterrows()]
    
    vectorstore = create_vector_store(documents)
    
    matches = []
    for chunk in highlighted_chunks:
        similar_docs = vectorstore.similarity_search(chunk, k=5)
        chunk_matches = []
        for doc in similar_docs:
            similarity = similar(chunk, doc.page_content)
            chunk_matches.append({
                "chunk": chunk,
                "excel_row": doc.page_content,
                "similarity": similarity
            })
        matches.append(chunk_matches)
    
    with open(os.path.join(LOG_OUTPUT_DIR, "detailed_matches.json"), "w") as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)
    
    return matches

def update_excel_file(excel_path, matches):
    df = pd.read_excel(excel_path)
    
    df['변경사항'] = df['변경사항'].astype(str)
    
    update_count = 0
    update_details = []
    for chunk_matches in matches:
        best_match = max(chunk_matches, key=lambda x: x['similarity'])
        if best_match['similarity'] > 0.6:
            row_index = df[df.apply(lambda row: similar(' '.join(row.astype(str)), best_match['excel_row']) > 0.8, axis=1)].index
            if not row_index.empty:
                df.loc[row_index, '변경사항'] = '추가'
                update_count += 1
                update_details.append({
                    "chunk": best_match['chunk'],
                    "excel_row": best_match['excel_row'],
                    "similarity": best_match['similarity'],
                    "excel_index": int(row_index[0])
                })
    
    df.to_excel(excel_path, index=False)
    log_to_file(f"엑셀 파일이 업데이트되었습니다: {excel_path}")
    log_to_file(f"총 {update_count}개의 행이 '추가'로 표시되었습니다.")
    
    class NumpyEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, np.integer):
                return int(obj)
            elif isinstance(obj, np.floating):
                return float(obj)
            elif isinstance(obj, np.ndarray):
                return obj.tolist()
            return super(NumpyEncoder, self).default(obj)

    with open(os.path.join(LOG_OUTPUT_DIR, "update_details.json"), "w") as f:
        json.dump(update_details, f, ensure_ascii=False, indent=2, cls=NumpyEncoder)

def main(pdf_path, excel_path):
    log_to_file("PDF에서 테이블을 추출하여 엑셀 파일로 저장합니다...")
    extract_tables_to_excel(pdf_path, excel_path)

    log_to_file("PDF에서 강조된 부분을 추출하고 엑셀 파일을 업데이트합니다...")

    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    highlighted_regions = detect_highlights(image)
    highlighted_chunks = extract_highlighted_text_ollama(image, highlighted_regions)
    matches = find_highlighted_text_in_excel(excel_path, highlighted_chunks)
    update_excel_file(excel_path, matches)

    log_to_file(f"처리된 PDF 페이지: {page_number + 1}")
    log_to_file(f"추출된 강조 영역 수: {len(highlighted_regions)}")
    log_to_file(f"추출된 강조 텍스트 청크 수: {len(highlighted_chunks)}")
    log_to_file(f"매칭된 결과 수: {len(matches)}")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, excel_path)