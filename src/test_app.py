import streamlit as st
import os
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
import tempfile

# 페이지 설정
st.set_page_config(page_title="PDF 보장내용 테스트", layout="wide")
st.title("PDF 보장내용 테스트")

# 파일 업로드 영역
uploaded_files = st.file_uploader("PDF 파일을 선택하세요", type="pdf", accept_multiple_files=True)

if uploaded_files:
    st.write(f"{len(uploaded_files)}개 파일이 업로드되었습니다")
    
    if st.button("테스트 시작"):
        # 결과 저장용 데이터프레임
        all_results = []
        
        # 각 파일 처리
        for pdf_file in uploaded_files:
            st.subheader(f"파일 처리 중: {pdf_file.name}")
            
            # 임시 파일 경로 생성
            temp_dir = tempfile.gettempdir()
            file_name = pdf_file.name
            temp_path = os.path.join(temp_dir, file_name)
            
            with open(temp_path, "wb") as f:
                f.write(pdf_file.getvalue())
            
            # PDF 처리 로직
            try:
                pdf_document = fitz.open(temp_path)
                
                # "나. 보험금" 검색
                insurance_pages = []
                for page_num in range(len(pdf_document)):
                    page = pdf_document[page_num]
                    text = page.get_text()
                    if "나. 보험금" in text:
                        insurance_pages.append(page_num)
                        st.write(f"'나. 보험금' 문구 발견: 페이지 {page_num + 1}")
                
                # 섹션 탐색 및 파싱 범위 설정
                for page_num in insurance_pages:
                    # 현재 페이지부터 시작해서 섹션 범위 찾기
                    start_page = page_num
                    section_name = "보험금 섹션"
                    
                    # 다음 섹션 찾기
                    end_page = None
                    for p in range(start_page + 1, len(pdf_document)):
                        text = pdf_document[p].get_text()
                        if re.search(r'[다-힣]\.\s+\w+', text):  # 다음 섹션 패턴
                            end_page = p - 1
                            break
                    
                    if end_page is None:
                        end_page = len(pdf_document) - 1
                    
                    st.write(f"파싱 범위: {section_name}, 페이지 {start_page + 1}~{end_page + 1}")
                    
                    # 강조색/색상 텍스트 찾기
                    for p in range(start_page, end_page + 1):
                        page = pdf_document[p]
                        
                        # 텍스트 스팬 검사
                        spans = page.get_text("dict")["blocks"]
                        for block in spans:
                            if "lines" in block:
                                for line in block["lines"]:
                                    for span in line["spans"]:
                                        color = span.get("color")
                                        if color and color != 0:  # 색상 있는 텍스트
                                            # 결과 저장
                                            all_results.append({
                                                "파일명": pdf_file.name,
                                                "페이지": p + 1,
                                                "섹션": section_name,
                                                "파싱 범위": f"{start_page + 1}~{end_page + 1}",
                                                "내용": span["text"],
                                                "변경사항": "있음"
                                            })
                
                pdf_document.close()
                
                # 임시 파일 삭제
                os.remove(temp_path)
                
            except Exception as e:
                st.error(f"오류 발생: {str(e)}")
        
        # 결과 표시
        if all_results:
            df = pd.DataFrame(all_results)
            st.subheader("분석 결과")
            st.dataframe(df)
            
            # 결과 다운로드 버튼
            csv = df.to_csv(index=False)
            st.download_button(
                label="CSV로 다운로드",
                data=csv,
                file_name=f"보장내용_테스트_결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        else:
            st.info("변경사항이 발견되지 않았습니다.")

def process_pdf(pdf_document):
    # "나. 보험금" 섹션 찾기
    insurance_payment_section = None
    for page_num in range(len(pdf_document)):
        page_text = pdf_document[page_num].get_text()
        if "나. 보험금" in page_text:
            insurance_payment_section = page_num
            break
    
    # "나. 보험금"이 없을 경우 대체 로직
    if insurance_payment_section is None:
        # 상해관련 특별약관과 선택특약이 함께 있는 페이지 찾기
        for page_num in range(len(pdf_document)):
            page_text = pdf_document[page_num].get_text()
            page_text_normalized = ''.join(page_text.split())  # 띄어쓰기 제거
            
            # 다양한 형태의 "상해및질병관련특별약관" 검색
            has_special_clause = any(keyword in page_text_normalized for keyword in 
                                    ["상해및질병관련특별약관", "상해및질병관련", "상해질병관련특별약관"])
            
            # 선택특약 검색
            has_optional_clause = "선택특약" in page_text
            
            if has_special_clause and has_optional_clause:
                insurance_payment_section = page_num
                break
    
    return insurance_payment_section, other_sections