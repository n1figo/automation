import fitz
import cv2
import numpy as np
from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright
import os

def setup_browser():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, page

def capture_tab_content(url, tab_selector, output_filename):
    playwright, browser, page = setup_browser()
    try:
        page.goto(url)
        page.wait_for_load_state('networkidle')
        
        page.click(tab_selector)
        page.wait_for_load_state('networkidle')
        
        os.makedirs('output', exist_ok=True)
        
        page.screenshot(path=f'output/{output_filename}', full_page=True)
        print(f"Screenshot saved as output/{output_filename}")
    
    finally:
        browser.close()
        playwright.stop()

def extract_underlined_text(pdf_path):
    doc = fitz.open(pdf_path)
    underlined_texts = []
    
    for page in doc:
        words = page.get_text("words")
        for word in words:
            if word[4] == 1:  # Check if the word is underlined
                underlined_texts.append((word[0], word[1], word[2], word[3], word[4]))
    
    doc.close()
    return underlined_texts

def find_text_in_image(image_path, text):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Perform OCR (you might need to install pytesseract)
    # import pytesseract
    # data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
    
    # For demonstration, let's assume we found the text at these coordinates
    # In a real scenario, you'd use the OCR results to find the actual position
    return [(100, 100, 200, 150)]  # Example coordinates

def highlight_differences(image_path, underlined_texts, output_filename):
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image, 'RGBA')
    
    for text, x0, y0, x1, y1 in underlined_texts:
        positions = find_text_in_image(image_path, text)
        for pos in positions:
            draw.rectangle(pos, outline="red", fill=(255, 255, 0, 64), width=2)
    
    image.save(f'output/{output_filename}')
    print(f"Highlighted image saved as output/{output_filename}")

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "path_to_your_pdf_summary.pdf"  # Replace with actual path
    
    capture_tab_content(url, 'a#tabexmpl', 'signup_example.png')
    capture_tab_content(url, 'a#tabguarnt', 'coverage_details.png')
    
    underlined_texts = extract_underlined_text(pdf_path)
    
    highlight_differences('output/signup_example.png', underlined_texts, 'highlighted_signup_example.png')
    highlight_differences('output/coverage_details.png', underlined_texts, 'highlighted_coverage_details.png')

if __name__ == "__main__":
    main()