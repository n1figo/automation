from sentence_transformers import SentenceTransformer, util
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont
import io

def get_html_content_and_screenshot(url, tab_selector):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url)
        page.click(tab_selector)
        page.wait_for_load_state('networkidle')
        content = page.content()
        screenshot = page.screenshot(full_page=True)
        browser.close()
    return content, screenshot

def extract_sentences(text):
    return [sent.strip() for sent in text.split('.') if sent.strip()]

def extract_text_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return ' '.join(soup.stripped_strings)

def compare_documents(text1, text2):
    model = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    sentences1 = extract_sentences(text1)
    sentences2 = extract_sentences(text2)
    
    embeddings1 = model.encode(sentences1, convert_to_tensor=True)
    embeddings2 = model.encode(sentences2, convert_to_tensor=True)
    
    cosine_scores = util.pytorch_cos_sim(embeddings1, embeddings2)
    
    changes = []
    for i, sent1 in enumerate(sentences1):
        max_score = max(cosine_scores[i])
        if max_score < 0.8:  # Threshold for considering as changed
            changes.append((sent1, sentences2[cosine_scores[i].argmax()]))
    
    return changes

def highlight_changes_on_image(screenshot, changes):
    image = Image.open(io.BytesIO(screenshot))
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()

    y_position = 10
    for old, new in changes:
        draw.rectangle([10, y_position, image.width - 10, y_position + 40], 
                       fill=(255, 255, 0, 128))  # Semi-transparent yellow
        draw.text((15, y_position), f"Old: {old[:50]}...", fill=(255, 0, 0), font=font)
        draw.text((15, y_position + 20), f"New: {new[:50]}...", fill=(0, 255, 0), font=font)
        y_position += 45

    return image

def main():
    url_original = "https://www.kbinsure.co.kr/CG302120001.ec"  # URL for original (left) document
    url_changed = "https://www.kbinsure.co.kr/CG302120001_changed.ec"  # URL for changed (right) document, replace with actual URL

    # Get content and screenshot for original document
    original_html, original_screenshot = get_html_content_and_screenshot(url_original, 'a#tabexmpl')
    original_text = extract_text_from_html(original_html)

    # Get content for changed document
    changed_html, _ = get_html_content_and_screenshot(url_changed, 'a#tabexmpl')
    changed_text = extract_text_from_html(changed_html)

    # Compare documents
    changes = compare_documents(original_text, changed_text)

    # Highlight changes on the original document's screenshot
    highlighted_image = highlight_changes_on_image(original_screenshot, changes)

    # Save the highlighted image
    highlighted_image.save("highlighted_changes.png")

    print("Changes highlighted and saved in 'highlighted_changes.png'")

    # Print changes for reference
    print("\nDetected changes:")
    for old, new in changes:
        print(f"Old: {old}")
        print(f"New: {new}")
        print("---")

if __name__ == "__main__":
    main()