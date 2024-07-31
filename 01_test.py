import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd

def scrape_kb_insurance(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table', class_='tbl_ty03')
        
        if not table:
            return None

        data = []
        rows = table.find_all('tr')
        for row in rows[1:]:  # Skip header row
            cols = row.find_all('td')
            if len(cols) >= 4:
                data.append([
                    cols[0].text.strip(),  # 상품명
                    cols[1].text.strip(),  # 보장내용
                    cols[2].text.strip(),  # 가입대상
                    cols[3].text.strip()   # 보험기간
                ])
        
        return pd.DataFrame(data, columns=['상품명', '보장내용', '가입대상', '보험기간'])
    except Exception as e:
        st.error(f"웹 스크래핑 중 오류 발생: {str(e)}")
        return None

def parse_hwp_content(file):
    # 실제 HWP 파싱 로직은 여기에 구현해야 합니다.
    # 현재는 예시 데이터를 반환합니다.
    return pd.DataFrame([
        ["KB국민 희망플러스자녀보험", "자녀에게 필요한 종합보장", "0~25세", "3년만기"],
        ["KB생활비받는암보험", "다양한 암 진단비 보장", "20~65세", "10년만기"]
    ], columns=['상품명', '보장내용', '가입대상', '보험기간'])

def compare_dataframes(df1, df2):
    if df1 is None or df2 is None:
        return None
    
    # 두 데이터프레임의 인덱스와 컬럼을 일치시킵니다.
    df1 = df1.set_index('상품명')
    df2 = df2.set_index('상품명')
    
    # 변경사항을 저장할 데이터프레임을 생성합니다.
    changes = pd.DataFrame(columns=['필드', '기존 내용', '변경 내용'])
    
    # 각 상품에 대해 변경사항을 확인합니다.
    for product in df1.index.union(df2.index):
        if product in df1.index and product in df2.index:
            for col in df1.columns:
                if df1.loc[product, col] != df2.loc[product, col]:
                    changes = changes.append({
                        '필드': f"{product} - {col}",
                        '기존 내용': df1.loc[product, col],
                        '변경 내용': df2.loc[product, col]
                    }, ignore_index=True)
        elif product in df1.index:
            changes = changes.append({
                '필드': product,
                '기존 내용': '존재',
                '변경 내용': '삭제됨'
            }, ignore_index=True)
        else:
            changes = changes.append({
                '필드': product,
                '기존 내용': '없음',
                '변경 내용': '새로 추가됨'
            }, ignore_index=True)
    
    return changes

st.title('KB손해보험 상품 비교 서비스')

url = st.text_input('KB손해보험 URL', 'https://www.kbinsure.co.kr/CG302130001.ec')

if st.button('웹사이트 데이터 가져오기'):
    web_data = scrape_kb_insurance(url)
    if web_data is not None:
        st.subheader('웹사이트 상품 정보')
        st.dataframe(web_data)
    else:
        st.error('웹사이트에서 데이터를 가져오는 데 실패했습니다.')

uploaded_file = st.file_uploader("HWP 파일 업로드", type="hwp")

if uploaded_file is not None:
    hwp_data = parse_hwp_content(uploaded_file)
    st.subheader('HWP 파일 상품 정보')
    st.dataframe(hwp_data)

    if 'web_data' in locals():
        changes = compare_dataframes(web_data, hwp_data)
        if changes is not None and not changes.empty:
            st.subheader('변경 사항')
            st.dataframe(changes)
        else:
            st.info('변경 사항이 없습니다.')
    else:
        st.warning('웹사이트 데이터를 먼저 가져와주세요.')
