import os
from playwright.sync_api import sync_playwright
from PIL import Image, ImageDraw
import fitz  # PyMuPDF
import difflib

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
        
        screenshot_path = os.path.join('output', output_filename)
        page.screenshot(path=screenshot_path, full_page=True)
        print(f"Screenshot saved as {screenshot_path}")
        return screenshot_path
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
                underlined_texts.append(word[4])
    
    doc.close()
    return underlined_texts

def find_text_in_image(image_path, text):
    # This is a placeholder function. In a real scenario, you'd use OCR here.
    # For demonstration, we'll return some dummy coordinates.
    return [(100, 100, 200, 150), (300, 300, 400, 350)]

def highlight_differences(image_path, underlined_texts, output_filename):
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image, 'RGBA')
    
    for text in underlined_texts:
        positions = find_text_in_image(image_path, text)
        for pos in positions:
            draw.rectangle(pos, outline="red", fill=(255, 255, 0, 64), width=2)
    
    output_path = os.path.join('output', output_filename)
    image.save(output_path)
    print(f"Highlighted image saved as {output_path}")
    return output_path

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. ㅇKB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"  # Replace with actual path
    
    # Capture screenshots
    signup_image_path = capture_tab_content(url, 'a#tabexmpl', 'signup_example.png')
    coverage_image_path = capture_tab_content(url, 'a#tabguarnt', 'coverage_details.png')
    
    # Extract underlined text from PDF
    underlined_texts = extract_underlined_text(pdf_path)
    
    # Highlight differences in captured images
    highlighted_signup_path = highlight_differences(signup_image_path, underlined_texts, 'highlighted_signup_example.png')
    highlighted_coverage_path = highlight_differences(coverage_image_path, underlined_texts, 'highlighted_coverage_details.png')
    
    print("Processing complete. Check the 'output' folder for results.")

if __name__ == "__main__":
    main()