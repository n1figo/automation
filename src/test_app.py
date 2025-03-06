import streamlit as st
import os
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
import tempfile
import glob

# 페이지 설정
st.set_page_config(page_title="PDF 보장내용 테스트", layout="wide")
st.title("PDF 보장내용 테스트")

# 입력 폴더 경로 설정
input_folder = "/workspaces/automation/data/input"

# 폴더가 없으면 생성
if not os.path.exists(input_folder):
    os.makedirs(input_folder)
    st.warning(f"{input_folder} 폴더를 생성했습니다. PDF 파일을 이 폴더에 저장하세요.")

# 폴더에서 PDF 파일 목록 불러오기
pdf_files = glob.glob(os.path.join(input_folder, "*.pdf"))

if not pdf_files:
    st.warning(f"{input_folder} 폴더에 PDF 파일이 없습니다. 파일을 추가하고 테스트하세요.")
else:
    # PDF 파일 목록 표시
    st.write(f"{len(pdf_files)}개 PDF 파일이 발견되었습니다.")
    
    # 파일 목록을 체크박스로 표시
    file_options = {}
    for file_path in pdf_files:
        file_name = os.path.basename(file_path)
        file_options[file_path] = st.checkbox(f"{file_name}", value=True)
    
    # 선택된 파일 처리
    selected_files = [path for path, selected in file_options.items() if selected]
    
    if st.button("테스트 시작", type="primary"):
        if not selected_files:
            st.error("처리할 PDF 파일을 선택하세요.")
        else:
            st.write(f"선택된 {len(selected_files)}개 파일 처리를 시작합니다.")
            
            # 결과 저장용 리스트
            analysis_results = []
            
            # 진행 상황 표시
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 각 파일 처리
            for i, file_path in enumerate(selected_files):
                file_name = os.path.basename(file_path)
                status_text.text(f"파일 처리 중: {file_name} ({i+1}/{len(selected_files)})")
                
                # 결과를 저장할 딕셔너리 초기화
                file_result = {
                    "파일명": file_name,
                    "나. 보험금 페이지": [],
                    "상해관련특별약관 페이지": [],
                    "질병관련특별약관 페이지": [],
                    "상해및질병관련특별약관 페이지": [],
                    "상해및질병관련특별약관 종료 페이지": None,
                    "강조색 있는 페이지": set(),
                    "취소선 있는 페이지": set(),
                    "처리 상태": "완료",
                    "상세 로그": []
                }
                
                with st.expander(f"파일: {file_name}", expanded=False):
                    # PDF 처리 로직
                    try:
                        pdf_document = fitz.open(file_path)
                        total_pages = len(pdf_document)
                        st.write(f"총 {total_pages}페이지 로드됨")
                        
                        # 1. "나. 보험금" 검색하여 파싱 시작 페이지 찾기
                        parsing_start_page = None
                        for page_num in range(total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            if "나. 보험금" in text:
                                parsing_start_page = page_num
                                file_result["나. 보험금 페이지"].append(page_num + 1)
                                st.write(f"'나. 보험금' 문구 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'나. 보험금' 문구 발견: 페이지 {page_num + 1}")
                                # 첫 번째 발견하면 중단 (일반적으로 가장 앞에 나오는 것이 기준)
                                break
                        
                        # 파싱 시작 페이지를 찾지 못한 경우 전체 문서를 대상으로 함
                        if parsing_start_page is None:
                            parsing_start_page = 0
                            st.warning("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.")
                            file_result["상세 로그"].append("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.")
                        
                        # 2. 파싱 시작 페이지부터 검색
                        # 순서: 상해관련특별약관 -> 질병관련특별약관 -> 상해및질병관련특별약관
                        
                        # 2-1. "상해관련특별약관" 검색
                        for page_num in range(parsing_start_page, total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                            if "상해관련특별약관" in text_normalized and "상해및질병관련특별약관" not in text_normalized:
                                file_result["상해관련특별약관 페이지"].append(page_num + 1)
                                st.write(f"'상해관련특별약관' 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'상해관련특별약관' 발견: 페이지 {page_num + 1}")
                        
                        # 2-2. "질병관련특별약관" 검색
                        for page_num in range(parsing_start_page, total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                            if "질병관련특별약관" in text_normalized and "상해및질병관련특별약관" not in text_normalized:
                                file_result["질병관련특별약관 페이지"].append(page_num + 1)
                                st.write(f"'질병관련특별약관' 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'질병관련특별약관' 발견: 페이지 {page_num + 1}")
                        
                        # 2-3. "상해및질병관련특별약관" 검색 및 종료 페이지 찾기
                        combined_section_found = False
                        for page_num in range(parsing_start_page, total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                            
                            # 상해및질병관련특별약관 검색
                            if any(keyword in text_normalized for keyword in ["상해및질병관련특별약관", "상해및질병관련", "상해질병관련특별약관"]):
                                file_result["상해및질병관련특별약관 페이지"].append(page_num + 1)
                                st.write(f"'상해및질병관련특별약관' 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'상해및질병관련특별약관' 발견: 페이지 {page_num + 1}")
                                combined_section_found = True
                        
                        # 종료 페이지 찾기 - 다음 주요 섹션이나 문서 끝
                        if combined_section_found:
                            # 마지막 상해및질병관련특별약관 페이지
                            last_combined_page = max(file_result["상해및질병관련특별약관 페이지"]) - 1  # 페이지 번호를 인덱스로 변환
                            
                            # 종료 페이지 찾기 - 다음 주요 섹션 시작점
                            end_page = None
                            for page_num in range(last_combined_page + 1, total_pages):
                                page = pdf_document[page_num]
                                text = page.get_text()
                                
                                # 새로운 주요 섹션이 시작되는지 확인 (예: "다. 새로운섹션")
                                if re.search(r'[가-힣]\.\s+\w+', text) and "보험금" not in text:
                                    end_page = page_num
                                    break
                            
                            # 종료 페이지를 찾지 못했다면 문서 끝까지로 간주
                            if end_page is None:
                                end_page = total_pages
                                
                            file_result["상해및질병관련특별약관 종료 페이지"] = end_page
                            st.write(f"'상해및질병관련특별약관' 종료: 페이지 {end_page}")
                            file_result["상세 로그"].append(f"'상해및질병관련특별약관' 종료: 페이지 {end_page}")
                        
                        # 3. 강조색 및 취소선 검색 - "나. 보험금" 페이지부터 파싱 끝까지만 검색
                        # 검색 범위 설정
                        search_start_page = parsing_start_page  # "나. 보험금" 페이지 (또는 기본값 0)
                        search_end_page = total_pages - 1  # 기본값은 문서 끝까지

                        # 상해및질병관련특별약관 종료 페이지가 있으면 그것을 종료 범위로 설정
                        if file_result["상해및질병관련특별약관 종료 페이지"] is not None:
                            search_end_page = file_result["상해및질병관련특별약관 종료 페이지"] - 1  # 페이지 번호를 인덱스로 변환

                        # 설정된 범위 내에서만 강조색/취소선 검색
                        st.write(f"강조색/취소선 검색 범위: 페이지 {search_start_page + 1}~{search_end_page + 1}")
                        file_result["상세 로그"].append(f"강조색/취소선 검색 범위: 페이지 {search_start_page + 1}~{search_end_page + 1}")

                        for page_num in range(search_start_page, search_end_page + 1):
                            page = pdf_document[page_num]
                            
                            # 텍스트 스팬 검사
                            spans = page.get_text("dict")["blocks"]
                            has_highlight = False
                            has_strikethrough = False
                            
                            for block in spans:
                                if "lines" in block:
                                    for line in block["lines"]:
                                        for span in line["spans"]:
                                            # 강조색 확인
                                            color = span.get("color")
                                            if color and color != 0:  # 색상 있는 텍스트
                                                has_highlight = True
                                            
                                            # 취소선 확인 (flags 속성에 취소선 비트가 설정되어 있는지)
                                            flags = span.get("flags", 0)
                                            if flags & 2**6:  # 취소선 비트 확인 (64)
                                                has_strikethrough = True
                            
                            if has_highlight:
                                file_result["강조색 있는 페이지"].add(page_num + 1)
                            
                            if has_strikethrough:
                                file_result["취소선 있는 페이지"].add(page_num + 1)
                        
                        # 집합(set)을 정렬된 리스트로 변환
                        file_result["강조색 있는 페이지"] = sorted(list(file_result["강조색 있는 페이지"]))
                        file_result["취소선 있는 페이지"] = sorted(list(file_result["취소선 있는 페이지"]))
                        
                        if file_result["강조색 있는 페이지"]:
                            st.write(f"강조색 있는 페이지: {', '.join(map(str, file_result['강조색 있는 페이지']))}")
                            file_result["상세 로그"].append(f"강조색 있는 페이지: {', '.join(map(str, file_result['강조색 있는 페이지']))}")
                        
                        if file_result["취소선 있는 페이지"]:
                            st.write(f"취소선 있는 페이지: {', '.join(map(str, file_result['취소선 있는 페이지']))}")
                            file_result["상세 로그"].append(f"취소선 있는 페이지: {', '.join(map(str, file_result['취소선 있는 페이지']))}")
                        
                        pdf_document.close()
                        
                    except Exception as e:
                        error_message = f"오류 발생: {str(e)}"
                        st.error(error_message)
                        file_result["처리 상태"] = "오류"
                        file_result["상세 로그"].append(error_message)
                
                # 결과 저장
                analysis_results.append(file_result)
                
                # 진행 상황 업데이트
                progress = (i + 1) / len(selected_files)
                progress_bar.progress(progress)
            
            status_text.text("모든 PDF 파일 처리가 완료되었습니다.")
            
            # 결과 표시
            if analysis_results:
                st.subheader("분석 결과 요약")
                
                # 요약 테이블 생성
                summary_data = []
                for result in analysis_results:
                    summary_data.append({
                        "파일명": result["파일명"],
                        "나. 보험금 페이지": ', '.join(map(str, result["나. 보험금 페이지"])) if result["나. 보험금 페이지"] else "없음",
                        "상해및질병관련특별약관 종료": str(result["상해및질병관련특별약관 종료 페이지"]) if result["상해및질병관련특별약관 종료 페이지"] is not None else "없음",
                        "상해관련특별약관 페이지": ', '.join(map(str, result["상해관련특별약관 페이지"])) if result["상해관련특별약관 페이지"] else "없음",
                        "질병관련특별약관 페이지": ', '.join(map(str, result["질병관련특별약관 페이지"])) if result["질병관련특별약관 페이지"] else "없음",
                        "강조색 있는 페이지": ', '.join(map(str, result["강조색 있는 페이지"])) if result["강조색 있는 페이지"] else "없음",
                        "취소선 있는 페이지": ', '.join(map(str, result["취소선 있는 페이지"])) if result["취소선 있는 페이지"] else "없음",
                        "처리 상태": result["처리 상태"]
                    })
                
                # 데이터프레임 생성 및 표시
                summary_df = pd.DataFrame(summary_data)
                st.dataframe(summary_df, use_container_width=True)
                
                # 상세 로그 표시
                st.subheader("상세 로그")
                for result in analysis_results:
                    with st.expander(f"파일: {result['파일명']} 상세 로그"):
                        for log_entry in result["상세 로그"]:
                            st.write(log_entry)
                
                # 결과 다운로드 버튼
                csv = summary_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="CSV로 다운로드",
                    data=csv,
                    file_name=f"PDF_분석_결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
                
                # Excel 다운로드 버튼
                buffer = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                summary_df.to_excel(buffer.name, index=False, engine='openpyxl')
                with open(buffer.name, "rb") as f:
                    excel_data = f.read()
                st.download_button(
                    label="Excel로 다운로드",
                    data=excel_data,
                    file_name=f"PDF_분석_결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.ms-excel"
                )
                os.unlink(buffer.name)  # 임시 파일 삭제
            else:
                st.info("분석된 결과가 없습니다.")