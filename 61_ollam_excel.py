import os
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
from PIL import Image
from langchain_community.llms import HuggingFaceHub
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
from dotenv import load_dotenv


# Hugging Face API 토큰 설정
load_dotenv()
HUGGINGFACE_API_TOKEN = os.getenv("HUGGINGFACE_API_TOKEN")

# 디버깅 모드 설정
DEBUG_MODE = True

# 타겟 헤더 정의
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]

# 이미지 저장 경로 설정
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

# Hugging Face 모델 초기화
llm = HuggingFaceHub(repo_id="google/flan-t5-base", model_kwargs={"temperature": 0.5, "max_length": 512})

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

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

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), cleaned_mask)

    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    contour_image = image.copy()
    cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))

    return contours

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

def extract_text_from_region(page, region):
    x0, y0, x1, y1 = 0, region[0], page.rect.width, region[1]
    return page.get_text("text", clip=(x0, y0, x1, y1))

def compare_texts(text1, text2):
    prompt = PromptTemplate(
        input_variables=["text1", "text2"],
        template="Compare the following two texts and rate their similarity on a scale from 0 to 1:\nText 1: {text1}\nText 2: {text2}\nSimilarity (0-1):"
    )
    chain = LLMChain(llm=llm, prompt=prompt)
    response = chain.run({"text1": text1, "text2": text2})
    try:
        similarity = float(response.strip())
        return max(0, min(similarity, 1))  # 0과 1 사이의 값으로 제한
    except ValueError:
        return 0  # 숫자로 변환할 수 없는 경우 0 반환

def extract_and_process_tables(doc, page_number, highlight_regions):
    page = doc[page_number]
    tables = page.find_tables()
    
    processed_data = []

    for table_index, table in enumerate(tables):
        df = pd.DataFrame(table.extract())
        x0, y0, x1, y1 = table.bbox

        print(f"Table {table_index + 1} 위치: (x0={x0}, y0={y0}, x1={x1}, y1={y1})")

        for highlight_region in highlight_regions:
            highlighted_text = extract_text_from_region(page, highlight_region)
            
            max_similarity = 0
            max_similarity_index = -1

            for row_index, row in df.iterrows():
                row_text = " ".join(row.astype(str))
                similarity = compare_texts(highlighted_text, row_text)

                if similarity > max_similarity:
                    max_similarity = similarity
                    max_similarity_index = row_index

            if max_similarity_index != -1 and max_similarity > 0.7:  # 임계값 설정
                df.at[max_similarity_index, "변경사항"] = "추가"

        for row_index, row in df.iterrows():
            if "변경사항" not in row or pd.isna(row["변경사항"]):
                row["변경사항"] = ""
            processed_data.append(row)

            print(f"Table {table_index + 1}, Row {row_index + 1}: 변경사항: {row['변경사항']}")

    return pd.DataFrame(processed_data)

def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f"파일이 '{output_path}'에 저장되었습니다.")

def main(pdf_path, output_excel_path):
    print("PDF에서 개정된 부분을 추출합니다...")

    doc = fitz.open(pdf_path)
    page_number = 50  # 페이지 번호 설정 (여기서는 51페이지)

    page = doc[page_number]
    image = pdf_to_image(page)

    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))

    contours = detect_highlights(image, page_number + 1)
    highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])

    highlighted_image = image.copy()
    for region in highlight_regions:
        cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
    cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))

    print(f"감지된 강조 영역 수: {len(highlight_regions)}")
    print(f"강조 영역: {highlight_regions}")

    processed_df = extract_and_process_tables(doc, page_number, highlight_regions)

    print(processed_df)

    save_to_excel(processed_df, output_excel_path)

    print(f"처리된 데이터가 {output_excel_path}에 저장되었습니다.")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)