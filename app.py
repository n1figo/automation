import streamlit as st
from playwright.sync_api import sync_playwright
import pandas as pd

def scrape_kb_insurance(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url)
        page.wait_for_load_state('networkidle')

        try:
            # 상품명 추출
            product_name = page.query_selector('h2.h3AreaBtn').inner_text()

            # 보장내용 추출
            coverage = page.query_selector('div.leftBox').inner_text()

            # 보험기간 추출
            period_element = page.query_selector('th:has-text("보험기간") + td')
            period = period_element.inner_text() if period_element else "보험기간을 찾을 수 없습니다"

            # 테이블 데이터 추출
            table = page.query_selector('table.tb_wrap')
            table_data = []
            if table:
                rows = table.query_selector_all('tr')
                for row in rows:
                    cols = row.query_selector_all('td')
                    if cols:
                        table_data.append([col.inner_text() for col in cols])

            return {
                "상품명": product_name,
                "보장내용": coverage,
                "보험기간": period,
                "테이블 데이터": table_data
            }
        except Exception as e:
            return {"error": f"데이터 추출 중 오류 발생: {str(e)}"}
        finally:
            browser.close()

def main():
    st.title('KB손해보험 웹 스크래핑 테스트 (Playwright)')
    
    url = st.text_input('KB손해보험 URL을 입력하세요', 'https://www.kbinsure.co.kr/CG302130001.ec')
    
    if st.button('데이터 가져오기'):
        with st.spinner('데이터를 가져오는 중... (약 10-15초 소요될 수 있습니다)'):
            result = scrape_kb_insurance(url)
            
            if "error" in result:
                st.error(result["error"])
            else:
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