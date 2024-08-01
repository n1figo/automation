from flask import Flask, render_template, request
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import pandas as pd
import time

app = Flask(__name__)

def setup_browser():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, page

def scrape_kb_insurance(url):
    playwright, browser, page = setup_browser()
    try:
        page.goto(url)
        page.wait_for_timeout(5000)  # 페이지 로딩을 위한 대기 시간

        # JavaScript 실행 후 페이지 소스 가져오기
        page_source = page.content()
        soup = BeautifulSoup(page_source, 'html.parser')

        # 상품명 추출
        product_name_element = page.query_selector("h2.h3AreaBtn")
        product_name = product_name_element.inner_text() if product_name_element else "상품명을 찾을 수 없습니다"

        # 보장내용 추출
        coverage_element = page.query_selector("div.leftBox")
        coverage = coverage_element.inner_text() if coverage_element else "보장내용을 찾을 수 없습니다"

        # 보험기간 추출
        period_element = page.query_selector("//th[contains(text(), '보험기간')]/following-sibling::td")
        period = period_element.inner_text() if period_element else "보험기간을 찾을 수 없습니다"

        # 테이블 데이터 추출
        table_data = []
        tables = page.query_selector_all("table.tb_ty01, table.tb_ty02, table.tb_ty03, table.tb_wrap")
        for table in tables:
            rows = table.query_selector_all("tr")
            for row in rows:
                cols = row.query_selector_all("td")
                if cols:
                    table_data.append([col.inner_text() for col in cols])

        return {
            "상품명": product_name,
            "보장내용": coverage,
            "보험기간": period,
            "테이블 데이터": table_data,
            "html_content": page_source


            
        }
    finally:
        browser.close()
        playwright.stop()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        url = request.form['url']
        result = scrape_kb_insurance(url)
        
        # 테이블 데이터를 DataFrame으로 변환
        df = pd.DataFrame(result['테이블 데이터'])
        table_html = df.to_html(classes='data', header="true") if not df.empty else "테이블 데이터를 찾을 수 없습니다."
        
        return render_template('result.html', 
                               product_name=result['상품명'],
                               coverage=result['보장내용'],
                               period=result['보험기간'],
                               table_html=table_html,
                               html_content=result['html_content'])
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)