import cv2
import numpy as np
import fitz  # PyMuPDF
import os
from PIL import Image

def pdf_to_images(pdf_path, start_page, end_page):
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(start_page - 1, min(end_page, len(doc))):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append((page_num + 1, np.array(img)))
    doc.close()
    return images

def detect_highlights(image):
    # RGB에서 HSV 색공간으로 변환
    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    
    # 채도(S) 채널과 밝기(V) 채널 추출
    s = hsv[:,:,1]
    v = hsv[:,:,2]
    
    # 채도가 낮은 영역(회색 등)을 마스킹
    saturation_threshold = 30  # 이 값을 조정하여 회색 제외 정도를 조절할 수 있습니다
    saturation_mask = s > saturation_threshold
    
    # Otsu's thresholding을 사용하여 밝기 기반 이진화
    _, binary = cv2.threshold(v, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # 채도 마스크와 밝기 마스크를 결합
    combined_mask = cv2.bitwise_and(binary, binary, mask=saturation_mask.astype(np.uint8) * 255)
    
    # 노이즈 제거
    kernel = np.ones((5,5), np.uint8)
    cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
    cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)
    
    # 윤곽선 찾기
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    return contours

def process_pdf(pdf_path, start_page, end_page, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    images = pdf_to_images(pdf_path, start_page, end_page)

    for page_num, image in images:
        contours = detect_highlights(image)
        
        if not contours:
            print(f"No highlights detected on page {page_num}")
            continue
        
        # 모든 하이라이트 영역의 y 좌표 범위 찾기
        min_y = min(cv2.boundingRect(contour)[1] for contour in contours)
        max_y = max(cv2.boundingRect(contour)[1] + cv2.boundingRect(contour)[3] for contour in contours)
        
        # 전체 가로폭, 하이라이트 세로폭 캡처
        highlighted_region = image[min_y:max_y, 0:image.shape[1]]
        
        # 이미지 저장
        output_path = os.path.join(output_folder, f"page_{page_num}_highlights.png")
        cv2.imwrite(output_path, cv2.cvtColor(highlighted_region, cv2.COLOR_RGB2BGR))

        print(f"Captured highlighted regions on page {page_num}, saved as {output_path}")

# 실행 파라미터
pdf_path = "/workspaces/automation/uploads/1722922992_5._KB_5.10.10_24.05__0801_v1.0.pdf"
start_page = 50
end_page = 52
output_folder = "/workspaces/automation/highlight_images"

# 메인 실행
if __name__ == "__main__":
    process_pdf(pdf_path, start_page, end_page, output_folder)
    print("Processing completed.")