from langchain.embeddings import HuggingFaceEmbeddings
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
import json
from PIL import Image
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
import warnings
from difflib import SequenceMatcher

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

def log_to_file(message, filename="debug_log.txt"):
    with open(os.path.join(LOG_OUTPUT_DIR, filename), "a") as f:
        f.write(message + "\n")

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def is_highlight_color(color):
    # RGB 값을 HSV로 변환
    hsv = cv2.cvtColor(np.uint8([[color]]), cv2.COLOR_RGB2HSV)[0][0]
    
    # 흰색, 검정색, 회색 제외
    if (color[0] > 200 and color[1] > 200 and color[2] > 200) or \
       (color[0] < 50 and color[1] < 50 and color[2] < 50) or \
       (abs(color[0] - color[1]) < 10 and abs(color[1] - color[2]) < 10):
        return False
    
    # 채도가 낮은 경우 제외 (회색 계열)
    if hsv[1] < 50:
        return False
    
    return True

def detect_highlights(image, page_num):
    height, width = image.shape[:2]
    mask = np.zeros((height, width), dtype=np.uint8)
    
    for y in range(height):
        for x in range(width):
            if is_highlight_color(image[y, x]):
                mask[y, x] = 255

    kernel = np.ones((5,5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), mask)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    highlighted_regions = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        highlighted_regions.append((x, y, x+w, y+h))

    return highlighted_regions

# PDF에서 테이블을 추출하고 엑셀로 저장하는 함수
def extract_tables_to_excel(pdf_path, output_excel_path):
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    tables = page.find_tables()
    
    all_data = []

    for table in tables:
        df = pd.DataFrame(table.extract())
        all_data.append(df)

    # 모든 테이블을 하나의 DataFrame으로 결합
    combined_df = pd.concat(all_data, ignore_index=True)

    # '변경사항' 열 추가
    combined_df['변경사항'] = ''

    # 엑셀로 저장
    combined_df.to_excel(output_excel_path, index=False)
    log_to_file(f"테이블이 추출되어 '{output_excel_path}'에 저장되었습니다.")

# PDF에서 텍스트 추출 및 처리
def extract_and_process_text(doc, page_number, highlighted_regions):
    page = doc[page_number]
    highlighted_chunks = []

    for region in highlighted_regions:
        x0, y0, x1, y1 = region
        text = page.get_text("text", clip=fitz.Rect(x0, y0, x1, y1))
        if text.strip():
            highlighted_chunks.append(text.strip())

    # 디버깅: 강조된 텍스트 저장
    with open(os.path.join(LOG_OUTPUT_DIR, "highlighted_chunks.txt"), "w") as f:
        for chunk in highlighted_chunks:
            f.write(chunk + "\n\n")
    
    return highlighted_chunks

# 텍스트 유사도 계산 함수
def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

# 엑셀 파일에서 강조된 텍스트 찾기 (개선된 버전)
def find_highlighted_text_in_excel(excel_path, highlighted_chunks):
    df = pd.read_excel(excel_path)
    
    # HuggingFace Embeddings 초기화
    embeddings_model = HuggingFaceEmbeddings(
        model_name="snunlp/KR-SBERT-V40K-klueNLI-augSTS",
    )
    
    # 엑셀 데이터를 벡터화
    excel_texts = df.apply(lambda row: ' '.join(row.astype(str)), axis=1).tolist()
    excel_vectors = embeddings_model.embed_documents(excel_texts)
    
    # 강조된 텍스트를 벡터화
    highlighted_vectors = embeddings_model.embed_documents(highlighted_chunks)
    
    # FAISS 인덱스 생성
    index = FAISS.from_embeddings(zip(excel_texts, excel_vectors), embeddings_model)
    
    # 강조된 텍스트와 가장 유사한 엑셀 행 찾기
    matches = []
    for chunk, vector in zip(highlighted_chunks, highlighted_vectors):
        similar_rows = index.similarity_search_by_vector(vector, k=5)  # Top 5 결과 검색
        chunk_matches = []
        for row in similar_rows:
            similarity = similar(chunk, row.page_content)
            chunk_matches.append({
                "chunk": chunk,
                "excel_row": row.page_content,
                "similarity": similarity
            })
        matches.append(chunk_matches)
    
    # 디버깅: 매칭 결과 저장
    with open(os.path.join(LOG_OUTPUT_DIR, "detailed_matches.json"), "w") as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)
    
    return matches

# 엑셀 파일 업데이트 (개선된 버전)
def update_excel_file(excel_path, matches):
    df = pd.read_excel(excel_path)
    
    # '변경사항' 열을 문자열 타입으로 변경
    df['변경사항'] = df['변경사항'].astype(str)
    
    update_count = 0
    update_details = []
    for chunk_matches in matches:
        best_match = max(chunk_matches, key=lambda x: x['similarity'])
        if best_match['similarity'] > 0.6:  # 임계값을 0.6으로 낮춤
            # 엑셀에서 일치하는 행 찾기
            row_index = df[df.apply(lambda row: similar(' '.join(row.astype(str)), best_match['excel_row']) > 0.8, axis=1)].index
            if not row_index.empty:
                # '변경사항' 열에 '추가' 입력
                df.loc[row_index, '변경사항'] = '추가'
                update_count += 1
                update_details.append({
                    "chunk": best_match['chunk'],
                    "excel_row": best_match['excel_row'],
                    "similarity": best_match['similarity'],
                    "excel_index": int(row_index[0])  # numpy.int64를 int로 변환
                })
    
    # 업데이트된 DataFrame을 엑셀 파일로 저장
    df.to_excel(excel_path, index=False)
    log_to_file(f"엑셀 파일이 업데이트되었습니다: {excel_path}")
    log_to_file(f"총 {update_count}개의 행이 '추가'로 표시되었습니다.")
    
    # 디버깅: 업데이트 상세 정보 저장
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

# 메인 함수
def main(pdf_path, excel_path):
    log_to_file("PDF에서 테이블을 추출하여 엑셀 파일로 저장합니다...")
    extract_tables_to_excel(pdf_path, excel_path)

    log_to_file("PDF에서 강조된 부분을 추출하고 엑셀 파일을 업데이트합니다...")

    # PyMuPDF로 PDF 열기
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    # 강조된 영역 탐지
    highlighted_regions = detect_highlights(image, page_number + 1)

    # PDF에서 강조된 텍스트 추출
    highlighted_chunks = extract_and_process_text(doc, page_number, highlighted_regions)

    # 엑셀 파일에서 강조된 텍스트 찾기
    matches = find_highlighted_text_in_excel(excel_path, highlighted_chunks)

    # 엑셀 파일 업데이트
    update_excel_file(excel_path, matches)

    # 디버깅: 전체 프로세스 요약
    log_to_file(f"처리된 PDF 페이지: {page_number + 1}")
    log_to_file(f"추출된 강조 영역 수: {len(highlighted_regions)}")
    log_to_file(f"추출된 강조 텍스트 청크 수: {len(highlighted_chunks)}")
    log_to_file(f"매칭된 결과 수: {len(matches)}")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, excel_path)