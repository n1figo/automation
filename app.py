1. streamlit에서 사용이불가하니, flask에서 구현하게 도와줘. 
2. Playwright: Selenium의 대안으로, 더 현대적이고 빠른 웹 자동화 도구입니다. 특히 동적 웹페이지 처리에 효과적입니다.
3. 2번 적용해서 소스코드 수정해.


import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def scrape_kb_insurance(url):
    driver = setup_driver()
    try:
        driver.get(url)
        time.sleep(5)  # 페이지 로딩을 위한 대기 시간

        # JavaScript 실행 후 페이지 소스 가져오기
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        # 상품명 추출
        product_name_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.h3AreaBtn"))
        )
        product_name = product_name_element.text if product_name_element else "상품명을 찾을 수 없습니다"

        # 보장내용 추출
        coverage_element = driver.find_element(By.CSS_SELECTOR, "div.leftBox")
        coverage = coverage_element.text if coverage_element else "보장내용을 찾을 수 없습니다"

        # 보험기간 추출
        period_element = driver.find_elements(By.XPATH, "//th[contains(text(), '보험기간')]/following-sibling::td")
        period = period_element[0].text if period_element else "보험기간을 찾을 수 없습니다"

        # 테이블 데이터 추출
        table_data = []
        tables = driver.find_elements(By.CSS_SELECTOR, "table.tb_ty01, table.tb_ty02, table.tb_ty03, table.tb_wrap")
        for table in tables:
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if cols:
                    table_data.append([col.text for col in cols])

        return {
            "상품명": product_name,
            "보장내용": coverage,
            "보험기간": period,
            "테이블 데이터": table_data,
            "html_content": page_source
        }
    finally:
        driver.quit()

def main():
    st.title('KB손해보험 웹 스크래핑 테스트 (Selenium)')
    
    url = st.text_input('KB손해보험 URL을 입력하세요', 'https://www.kbinsure.co.kr/CG302130001.ec')
    
    if st.button('데이터 가져오기'):
        with st.spinner('데이터를 가져오는 중... (약 15-20초 소요될 수 있습니다)'):
            result = scrape_kb_insurance(url)
            
            st.write("### 추출된 데이터")
            st.write(f"**상품명:** {result['상품명']}")
            st.write(f"**보장내용:** {result['보장내용'][:200]}...") # 긴 내용은 일부만 표시
            st.write(f"**보험기간:** {result['보험기간']}")
            
            st.write("### 테이블 데이터")
            if result['테이블 데이터']:
                df = pd.DataFrame(result['테이블 데이터'])
                st.dataframe(df)
            else:
                st.write("테이블 데이터를 찾을 수 없습니다.")

        # 디버깅을 위한 HTML 출력
        with st.expander("디버그: 전체 HTML"):
            st.code(result['html_content'], language='html')

if __name__ == '__main__':
    main()