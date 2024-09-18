import pypdf  # 이 줄을 추가합니다
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
from langchain_community.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from transformers import AutoTokenizer

# 디버깅 모드 설정
DEBUG_MODE = True

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

# 하이라이트 영역 탐지
def detect_highlights(image, page_num):
    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    s = hsv[:,:,1]
    v = hsv[:,:,2]

    saturation_threshold = 30
    saturation_mask = s > saturation_threshold

    _, binary = cv2.threshold(v, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    combined_mask = cv2.bitwise_and(binary, binary, mask=saturation_mask.astype(np.uint8) * 255)

    kernel = np.ones((5,5), np.uint8)
    cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
    cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)

    # 디버깅: 마스크 이미지 저장
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)

    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 디버깅: 윤곽선이 그려진 이미지 저장
    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))

    return contours

# 강조된 영역 반환
def get_capture_regions(contours, image_height, image_width):
    if not contours:
        return []

    capture_height = image_height // 3
    sorted_contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

    regions = []
    current_region = None

    for contour in sorted_contours:
        x, y, w, h = cv2.boundingRect(contour)

        if current_region is None:
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
        elif y - current_region[1] < capture_height//2:
            current_region[1] = min(image_height, y + h + capture_height//2)
        else:
            regions.append(current_region)
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]

    if current_region:
        regions.append(current_region)

    return regions

# PDF에서 텍스트 추출 및 처리
def extract_and_process_text(pdf_path, page_number, highlight_regions):
    loader = PyPDFLoader(pdf_path)
    pages = loader.load()
    
    # 특정 페이지만 처리
    page = pages[page_number]
    
    # 텍스트 분할
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=100,
        chunk_overlap=20,
        length_function=len
    )
    chunks = text_splitter.split_text(page.page_content)
    
    # 강조된 텍스트 식별
    highlighted_chunks = []
    for i, chunk in enumerate(chunks):
        if any(region[0] <= i * 100 <= region[1] for region in highlight_regions):
            highlighted_chunks.append(chunk)
    
    return highlighted_chunks

# 엑셀 파일에서 강조된 텍스트 찾기
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
        similar_rows = index.similarity_search_by_vector(vector, k=1)
        if similar_rows:
            matches.append((chunk, similar_rows[0].page_content))
    
    return matches

# 엑셀 파일 업데이트
def update_excel_file(excel_path, matches):
    df = pd.read_excel(excel_path)
    
    for highlighted_text, excel_row in matches:
        # 엑셀에서 일치하는 행 찾기
        row_index = df[df.apply(lambda row: ' '.join(row.astype(str)) == excel_row, axis=1)].index
        if not row_index.empty:
            # '변경사항' 열에 '추가' 입력
            df.loc[row_index, '변경사항'] = '추가'
    
    # 업데이트된 DataFrame을 엑셀 파일로 저장
    df.to_excel(excel_path, index=False)
    print(f"엑셀 파일이 업데이트되었습니다: {excel_path}")

# 메인 함수
def main(pdf_path, excel_path):
    print("PDF에서 개정된 부분을 추출하고 엑셀 파일을 업데이트합니다...")

    # PyMuPDF로 PDF 열기
    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    # 강조된 영역 탐지
    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    # PDF에서 강조된 텍스트 추출
    highlighted_chunks = extract_and_process_text(pdf_path, page_number, highlight_regions)

    # 엑셀 파일에서 강조된 텍스트 찾기
    matches = find_highlighted_text_in_excel(excel_path, highlighted_chunks)

    # 엑셀 파일 업데이트
    update_excel_file(excel_path, matches)

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, excel_path)