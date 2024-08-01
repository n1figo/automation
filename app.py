import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup

def scrape_kb_insurance(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # 상품명 추출
        product_name = soup.find('h2', class_='h3AreaBtn')
        product_name = product_name.text.strip() if product_name else "상품명을 찾을 수 없습니다"
        
        # 보장내용 추출
        coverage = soup.find('div', class_='leftBox')
        coverage = coverage.text.strip() if coverage else "보장내용을 찾을 수 없습니다"
        
        # 보험기간 추출 (예시, 실제 위치에 따라 수정 필요)
        period = soup.find('th', string='보험기간')
        period = period.find_next('td').text.strip() if period else "보험기간을 찾을 수 없습니다"
        
        # 테이블 정보 추출
        table = soup.find('table', class_='tb_wrap')
        if table:
            rows = table.find_all('tr')
            table_data = []
            for row in rows:
                cols = row.find_all(['th', 'td'])
                if cols:
                    table_data.append([col.text.strip() for col in cols])
        else:
            table_data = [["테이블을 찾을 수 없습니다", "", ""]]
        
        return {
            "상품명": product_name,
            "보장내용": coverage,
            "보험기간": period,
            "테이블 데이터": table_data
        }
    
    except requests.RequestException as e:
        return {"error": f"웹 요청 중 오류 발생: {str(e)}"}
    except Exception as e:
        return {"error": f"데이터 처리 중 오류 발생: {str(e)}"}

def main():
    st.title('KB손해보험 웹 스크래핑 테스트')
    
    url = st.text_input('KB손해보험 URL을 입력하세요', 'https://www.kbinsure.co.kr/CG302130001.ec')
    
    if st.button('데이터 가져오기'):
        with st.spinner('데이터를 가져오는 중...'):
            result = scrape_kb_insurance(url)
            
            if "error" in result:
                st.error(result["error"])
            else:
                st.success('데이터를 성공적으로 가져왔습니다.')
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
                
                st.write("### 전체 HTML (일부)")
                st.code(requests.get(url).text[:5000], language='html') # 처음 5000자만 표시

if __name__ == '__main__':
    main()