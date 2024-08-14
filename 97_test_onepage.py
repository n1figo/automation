import fitz
from bs4 import BeautifulSoup
import requests
import difflib
from PIL import Image, ImageDraw
import re

def extract_pdf_text(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_web_content(url):
    response = requests.get(url)
    return response.text

def parse_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup

def find_tab_content(soup, tab_title):
    # 탭 요소 찾기
    tab = soup.find('a', string=tab_title)
    if not tab:
        return None
    
    # 탭에 연결된 콘텐츠 ID 찾기
    content_id = tab.get('href', '').strip('#')
    if not content_id:
        return None
    
    # 해당 ID를 가진 콘텐츠 찾기
    content = soup.find(id=content_id)
    return content.get_text(strip=True) if content else None

def extract_text_with_positions(soup):
    text_positions = []
    for element in soup.find_all(text=True):
        if element.parent.name not in ['script', 'style']:
            text = element.strip()
            if text:
                parent = element.parent
                style = parent.get('style', '')
                position = re.search(r'position:\s*absolute;\s*left:\s*(\d+)px;\s*top:\s*(\d+)px', style)
                if position:
                    left, top = map(int, position.groups())
                    text_positions.append((text, left, top))
    return text_positions

def compare_texts(text1, text2):
    differ = difflib.Differ()
    diff = list(differ.compare(text1.splitlines(), text2.splitlines()))
    return diff

def find_changes(diff):
    changes = []
    for line in diff:
        if line.startswith('+ ') or line.startswith('- '):
            changes.append(line[2:])
    return changes

def highlight_changes_on_html_capture(html_capture_path, changes, text_positions):
    with Image.open(html_capture_path) as img:
        draw = ImageDraw.Draw(img)
        
        for change in changes:
            matching_positions = [pos for text, left, top in text_positions if change in text]
            for left, top in matching_positions:
                # 변경된 부분에 빨간색 테두리의 노란색 네모 표시
                draw.rectangle([left-5, top-5, left+100, top+20], outline="red", fill="yellow")
                draw.text((left, top), change[:20], fill="black")

        img.save("highlighted_html_capture.png")

def main(pdf_path, web_url, html_capture_path):
    pdf_text = extract_pdf_text(pdf_path)
    html_content = extract_web_content(web_url)
    soup = parse_html(html_content)

    # 보장내용과 가입예시 탭의 내용 추출
    guarantee_content = find_tab_content(soup, "보장내용")
    example_content = find_tab_content(soup, "가입예시")

    if guarantee_content and example_content:
        web_text = guarantee_content + "\n" + example_content
    else:
        print("보장내용 또는 가입예시 탭을 찾을 수 없습니다.")
        return

    text_positions = extract_text_with_positions(soup)

    diff = compare_texts(web_text, pdf_text)
    changes = find_changes(diff)
    highlight_changes_on_html_capture(html_capture_path, changes, text_positions)

    print("처리가 완료되었습니다. 결과 이미지를 확인하세요.")

if __name__ == "__main__":
    pdf_path = "/path/to/your/pdf/file.pdf"
    web_url = "https://example.com"
    html_capture_path = "/path/to/your/html_capture.png"
    main(pdf_path, web_url, html_capture_path)