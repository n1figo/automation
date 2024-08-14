from playwright.sync_api import sync_playwright
import os

def setup_browser():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True)  # Always use headless mode
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, page

def capture_tab_content(url, tab_selector, output_filename):
    playwright, browser, page = setup_browser()
    try:
        page.goto(url)
        page.wait_for_load_state('networkidle')
        
        # Click on the specified tab
        page.click(tab_selector)
        page.wait_for_load_state('networkidle')
        
        # Ensure the output directory exists
        os.makedirs('output', exist_ok=True)
        
        # Capture full page screenshot
        page.screenshot(path=f'output/{output_filename}', full_page=True)
        print(f"Screenshot saved as output/{output_filename}")
    
    finally:
        browser.close()
        playwright.stop()

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    
    # Capture 가입예시 tab
    capture_tab_content(url, 'a#tabexmpl', 'signup_example.png')
    
    # Capture 보장내용 tab
    capture_tab_content(url, 'a#tabguarnt', 'coverage_details.png')

if __name__ == "__main__":
    main()