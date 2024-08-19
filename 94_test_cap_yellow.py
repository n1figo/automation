import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
import difflib
from playwright.sync_api import sync_playwright

def get_html_content(url, tab_selector):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url)
        page.click(tab_selector)
        page.wait_for_load_state('networkidle')
        content = page.content()
        browser.close()
    return content

def extract_text_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return ' '.join(soup.stripped_strings)

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def compare_texts(text1, text2):
    differ = difflib.Differ()
    diff = list(differ.compare(text1.splitlines(), text2.splitlines()))
    return diff

def highlight_differences(diff):
    highlighted = []
    for line in diff:
        if line.startswith('+ '):
            highlighted.append(f"\033[92m{line}\033[0m")  # Green for additions
        elif line.startswith('- '):
            highlighted.append(f"\033[91m{line}\033[0m")  # Red for deletions
        elif line.startswith('? '):
            continue  # Skip the '?' lines
        else:
            highlighted.append(line)
    return '\n'.join(highlighted)

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"  # Replace with actual path

    # Get HTML content
    signup_html = get_html_content(url, 'a#tabexmpl')
    coverage_html = get_html_content(url, 'a#tabguarnt')

    # Extract text from HTML
    signup_text = extract_text_from_html(signup_html)
    coverage_text = extract_text_from_html(coverage_html)

    # Extract text from PDF
    pdf_text = extract_text_from_pdf(pdf_path)

    # Compare texts
    signup_diff = compare_texts(pdf_text, signup_text)
    coverage_diff = compare_texts(pdf_text, coverage_text)

    # Highlight differences
    print("Differences in Signup tab:")
    print(highlight_differences(signup_diff))
    print("\nDifferences in Coverage tab:")
    print(highlight_differences(coverage_diff))

if __name__ == "__main__":
    main()