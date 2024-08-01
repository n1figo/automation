import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
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

        # 상품명 추출
        product_name = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.h3AreaBtn"))
        ).text

        # 보장내용 추출
        coverage = driver.find_element(By.CSS_SELECTOR, "div.leftBox").text

        # 보험기간 추출
        period_element = driver.find_elements(By.XPATH, "//th[contains(text(), '보험기간')]/following-sibling::td")
        period = period_element[0].text if period_element else "보험기간을 찾을 수 없습니다"

        # 테이블 데이터 추출
        table = driver.find_elements(By.CSS_SELECTOR, "table.tb_wrap")
        table_data = []
        if table:
            rows = table[0].find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if cols:
                    table_data.append([col.text for col in cols])

        return {
            "상품명": product_name,
            "보장내용": coverage,
            "보험기간": period,
            "테이블 데이터": table_data
        }
    finally:
        driver.quit()

def main():
    st.title('KB손해보험 웹 스크래핑 테스트 (Selenium)')
    
    url = st.text_input('KB손해보험 URL을 입력하세요', 'https://www.kbinsure.co.kr/CG302130001.ec')
    
    if st.button('데이터 가져오기'):
        with st.spinner('데이터를 가져오는 중... (약 10-15초 소요될 수 있습니다)'):
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

if __name__ == '__main__':
    main()