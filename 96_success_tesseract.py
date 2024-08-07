import cv2
import numpy as np
import fitz  # PyMuPDF
import os
from PIL import Image

def pdf_to_image(pdf_path, page_num):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return np.array(img)

def detect_highlights(image):
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
    
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    return contours

def get_capture_regions(contours, image_height, image_width):
    if not contours:
        return []

    # 페이지 높이의 1/3을 캡처 영역의 기준 높이로 설정
    capture_height = image_height // 3
    
    # 하이라이트 영역들을 y 좌표 기준으로 정렬
    sorted_contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])
    
    regions = []
    current_region = None
    
    for contour in sorted_contours:
        x, y, w, h = cv2.boundingRect(contour)
        
        # 새로운 영역 시작
        if current_region is None:
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
        # 기존 영역과 가까운 경우 병합
        elif y - current_region[1] < capture_height//2:
            current_region[1] = min(image_height, y + h + capture_height//2)
        # 새로운 영역 시작
        else:
            regions.append(current_region)
            current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
    
    if current_region:
        regions.append(current_region)
    
    return regions

def process_pdf(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    doc.close()

    for page_num in range(total_pages):
        print(f"Processing page {page_num + 1} of {total_pages}")
        
        image = pdf_to_image(pdf_path, page_num)
        contours = detect_highlights(image)
        
        if not contours:
            print(f"No highlights detected on page {page_num + 1}")
            continue
        
        regions = get_capture_regions(contours, image.shape[0], image.shape[1])
        
        for i, (start_y, end_y) in enumerate(regions):


            
            highlighted_region = image[start_y:end_y, 0:image.shape[1]]
            
            output_path = os.path.join(output_folder, f"page_{page_num + 1}_highlights_{i + 1}.png")
            cv2.imwrite(output_path, cv2.cvtColor(highlighted_region, cv2.COLOR_RGB2BGR))

            print(f"Captured highlighted region {i + 1} on page {page_num + 1}, saved as {output_path}")

# 실행 파라미터
pdf_path = "/workspaces/automation/uploads/5. ㅇKB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
output_folder = "/workspaces/automation/highlight_images"

# 메인 실행
if __name__ == "__main__":
    print(f"Starting to process PDF: {pdf_path}")
    process_pdf(pdf_path, output_folder)
    print("Processing completed.")