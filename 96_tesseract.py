import cv2
import numpy as np
import pytesseract
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
        images.append((page_num + 1, np.array(img)))  # 페이지 번호와 함께 이미지 저장
    doc.close()
    return images

def detect_underlines(image):
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    kernel = np.ones((1, 40), np.uint8)
    detected_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    contours, _ = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return contours

def process_pdf(pdf_path, start_page, end_page, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    images = pdf_to_images(pdf_path, start_page, end_page)

    for page_num, image in images:
        contours = detect_underlines(image)

        for j, contour in enumerate(contours):
            x, y, w, h = cv2.boundingRect(contour)
            
            # 밑줄 위의 텍스트를 포함하도록 영역 확장
            y_extended = max(0, y - 30)  # 위로 30픽셀 확장
            h_extended = min(image.shape[0] - y_extended, h + 40)  # 아래로 10픽셀 확장
            
            # 전체 가로 폭 캡처
            underlined_region = image[y_extended:y_extended+h_extended, 0:image.shape[1]]
            
            # 이미지 저장
            output_path = os.path.join(output_folder, f"page_{page_num}_underline_{j+1}.png")
            cv2.imwrite(output_path, cv2.cvtColor(underlined_region, cv2.COLOR_RGB2BGR))

            print(f"Captured underlined region on page {page_num}, saved as {output_path}")

            # OCR 수행 (선택적)
            # text = pytesseract.image_to_string(underlined_region, lang='kor+eng')
            # print(f"Page {page_num}, Underline {j+1}: {text.strip()}")

# 사용 예
pdf_path = "/workspaces/automation/uploads/1722922992_5._KB_5.10.10_24.05__0801_v1.0.pdf"
start_page = 50
end_page = 52
output_folder = "/workspaces/automation/highlight_images"

process_pdf(pdf_path, start_page, end_page, output_folder)