import streamlit as st
from requests_html import HTMLSession
import pandas as pd

def scrape_kb_insurance(url):
    session = HTMLSession()
    r = session.get(url)
    r.html.render(sleep=5, keep_page=True, scrolldown=1)  # JavaScript 렌더링

    try:
        # 상품명 추출
        product_name = r.html.find('h2.h3AreaBtn', first=True).text

        # 보장내용 추출
        coverage = r.html.find('div.leftBox', first=True).text

        # 보험기간 추출
        period_element = r.html.xpath("//th[contains(text(), '보험기간')]/following-sibling::td", first=True)
        period = period_element.text if period_element else "보험기간을 찾을 수 없습니다"

        # 테이블 데이터 추출
        table = r.html.find('table.tb_wrap', first=True)
        table_data = []
        if table:
            rows = table.find('tr')
            for row in rows:
                cols = row.find('td')
                if cols:
                    table_data.append([col.text for col in cols])

        return {
            "상품명": product_name,
            "보장내용": coverage,
            "보험기간": period,
            "테이블 데이터": table_data
        }
    except Exception as e:
        return {"error": f"데이터 추출 중 오류 발생: {str(e)}"}
    finally:
        session.close()

def main():
    st.title('KB손해보험 웹 스크래핑 테스트 (requests-html)')
    
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