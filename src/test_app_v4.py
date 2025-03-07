import streamlit as st
import os
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
import tempfile
import glob
import cv2
import numpy as np

# 종별 시작 페이지 식별 함수
def identify_type_start_pages(pdf_document, parsing_start_page, total_pages):
    type_starts = {}
    
    # 종별 표시 패턴
    type_pattern = r'\[(\d)종\]'
    
    # 종별 시작 페이지의 특징적 구조 키워드
    structure_keywords = ["기본담보", "선택특약", "보장내용", "지급사유", "지급금액"]
    
    # 전체 종별 표시가 있는 페이지 우선 수집
    all_type_pages = {}
    for page_num in range(parsing_start_page, total_pages):
        page = pdf_document[page_num]
        text = page.get_text()
        matches = list(re.finditer(type_pattern, text))
        
        for match in matches:
            type_num = match.group(1)
            type_key = f"[{type_num}종]"
            
            if type_key not in all_type_pages:
                all_type_pages[type_key] = []
            
            all_type_pages[type_key].append(page_num)
    
    # 각 종별 첫 발견 페이지부터 10페이지 범위 내에서 시작 페이지 식별
    for type_key, pages in all_type_pages.items():
        # 첫 발견 페이지
        first_page = min(pages)
        best_page = None
        best_confidence = 0
        
        # 해당 종이 발견된 페이지들 중에서 시작 페이지 특징 분석
        for page_num in pages:
            page = pdf_document[page_num]
            text = page.get_text()
            text_lines = text.split('\n')
            
            # 1. 표 구조 확인
            table_headers = False
            for line in text_lines:
                # 표 헤더 행 식별 (여러 가지 형태 고려)
                if (("보장내용" in line or "보험금 지급사유" in line) and "지급금액" in line) or \
                   ("담보" in line and "지급사유" in line) or \
                   ("보장명" in line and "지급금액" in line):
                    table_headers = True
                    break
            
            # 2. 섹션 구조 확인
            has_sections = False
            for keyword in ["기본담보", "선택특약"]:
                if keyword in text:
                    has_sections = True
                    break
            
            # 3. "지급사유"/"지급금액" 패턴이 여러 번 등장하는지 확인
            payment_patterns = 0
            for line in text_lines:
                if "지급사유" in line or "지급금액" in line:
                    payment_patterns += 1
            
            # 종별 시작 페이지 신뢰도 계산
            confidence = 0
            if table_headers:
                confidence += 3  # 표 헤더가 있으면 높은 가중치
            if has_sections:
                confidence += 2  # 섹션 구조가 있으면 중간 가중치
            if payment_patterns >= 2:
                confidence += 1  # 지급 패턴이 반복되면 낮은 가중치
            
            # 가장 높은 신뢰도를 가진 페이지 선택
            if confidence > best_confidence:
                best_confidence = confidence
                best_page = page_num
                
        # 신뢰도가 충분한 페이지를 시작 페이지로 판단
        if best_page is not None and best_confidence >= 2:
            type_starts[type_key] = {
                "page": best_page + 1,  # 1부터 시작하는 페이지 번호
                "confidence": best_confidence,
                "features": []
            }
            
            # 발견된 특징 기록
            page = pdf_document[best_page]
            text = page.get_text()
            text_lines = text.split('\n')
            
            if any(("보장내용" in line or "보험금 지급사유" in line) and "지급금액" in line for line in text_lines):
                type_starts[type_key]["features"].append("표 헤더 포함")
            
            if any(keyword in text for keyword in ["기본담보", "선택특약"]):
                type_starts[type_key]["features"].append("섹션 구조 포함")
            
            payment_patterns = sum(1 for line in text_lines if "지급사유" in line or "지급금액" in line)
            if payment_patterns >= 2:
                type_starts[type_key]["features"].append("지급 패턴 반복")
    
    return type_starts

# 종별 범위 계산 함수
def calculate_type_ranges(type_starts, total_pages):
    # 시작 페이지 기준으로 정렬
    sorted_types = sorted(type_starts.items(), key=lambda x: x[1]["page"])
    
    type_ranges = {}
    for i, (type_key, info) in enumerate(sorted_types):
        start_page = info["page"]
        
        # 마지막 종이 아니면 다음 종의 시작 페이지 - 1이 종료 페이지
        if i < len(sorted_types) - 1:
            end_page = sorted_types[i+1][1]["page"] - 1
        else:
            # 마지막 종은 문서 끝까지
            end_page = total_pages
        
        type_ranges[type_key] = {
            "start_page": start_page,
            "end_page": end_page,
            "confidence": info["confidence"],
            "features": info["features"]
        }
    
    return type_ranges

def detect_highlights_with_opencv(pdf_document, page_num):
    """OpenCV를 사용하여 강조된 영역(형광펜, 색상 표시) 감지"""
    # 페이지 렌더링 (고해상도로)
    zoom = 2.0
    mat = fitz.Matrix(zoom, zoom)
    pix = pdf_document[page_num].get_pixmap(matrix=mat, alpha=False)
    
    # OpenCV 형식으로 변환
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    
    # HSV 색상 모델로 변환
    hsv = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2HSV)
    
    # 다양한 형광색 범위 정의
    # 노란색 형광펜
    yellow_lower = np.array([20, 100, 200])  # 밝은 노란색
    yellow_upper = np.array([40, 255, 255])
    yellow_mask = cv2.inRange(hsv, yellow_lower, yellow_upper)
    
    # 주황색 형광펜
    orange_lower = np.array([5, 100, 200])
    orange_upper = np.array([20, 255, 255]) 
    orange_mask = cv2.inRange(hsv, orange_lower, orange_upper)
    
    # 파란색 형광펜
    blue_lower = np.array([90, 100, 200])
    blue_upper = np.array([130, 255, 255])
    blue_mask = cv2.inRange(hsv, blue_lower, blue_upper)
    
    # 녹색 형광펜
    green_lower = np.array([40, 100, 200])
    green_upper = np.array([80, 255, 255])
    green_mask = cv2.inRange(hsv, green_lower, green_upper)
    
    # 보라색 형광펜
    purple_lower = np.array([130, 50, 200])
    purple_upper = np.array([170, 255, 255])
    purple_mask = cv2.inRange(hsv, purple_lower, purple_upper)
    
    # 회색 음영 (취소선과 함께 있는 영역)
    # 회색은 채도가 낮고 명도가 중간 정도인 영역
    gray_lower = np.array([0, 0, 80])  # S가 낮음
    gray_upper = np.array([180, 40, 200])  # V가 중간
    gray_mask = cv2.inRange(hsv, gray_lower, gray_upper)

    # 모든 마스크 결합
    combined_mask = yellow_mask | orange_mask | blue_mask | green_mask | purple_mask | gray_mask
    
    # 노이즈 제거
    kernel = np.ones((3, 3), np.uint8)
    combined_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_OPEN, kernel)
    
    # 색상별 검출 영역 비율 계산
    total_pixels = hsv.shape[0] * hsv.shape[1]
    
    yellow_ratio = np.sum(yellow_mask > 0) / total_pixels * 100
    orange_ratio = np.sum(orange_mask > 0) / total_pixels * 100
    blue_ratio = np.sum(blue_mask > 0) / total_pixels * 100
    green_ratio = np.sum(green_mask > 0) / total_pixels * 100
    purple_ratio = np.sum(purple_mask > 0) / total_pixels * 100
    gray_ratio = np.sum(gray_mask > 0) / total_pixels * 100
    
    # 일정 비율 이상이면 강조색이 있다고 판단
    highlight_threshold = 0.2  # 페이지의 0.2% 이상이 강조색일 경우
    return {
        "has_highlight": np.sum(combined_mask > 0) / total_pixels * 100 > highlight_threshold,
        "colors": {
            "yellow": yellow_ratio > highlight_threshold,
            "orange": orange_ratio > highlight_threshold,
            "blue": blue_ratio > highlight_threshold,
            "green": green_ratio > highlight_threshold,
            "purple": purple_ratio > highlight_threshold,
            "gray": gray_ratio > highlight_threshold
        }
    }

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
            st.error("처리할 PDF 파일을 선택해주세요.")
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
                    "종별_모든발견": {},  # 모든 종별 발견 페이지
                    "종별_시작": {},      # 종별 시작 페이지
                    "종별_범위": {},      # 종별 페이지 범위
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
                                break
                        
                        # 파싱 시작 페이지를 찾지 못한 경우 전체 문서를 대상으로 함
                        if parsing_start_page is None:
                            parsing_start_page = 0
                            st.warning("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.")
                            file_result["상세 로그"].append("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.")
                        
                        # 종별 검색 추가 (모든 종별 발견 페이지 검색)
                        st.write("종별 패턴 검색 시작...")
                        file_result["상세 로그"].append("종별 패턴 검색 시작...")
                        
                        type_pattern = r'\[(\d)종\]'
                        for page_num in range(parsing_start_page, total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            matches = re.finditer(type_pattern, text)
                            
                            for match in matches:
                                type_num = match.group(1)
                                type_key = f"[{type_num}종]"
                                
                                if type_key not in file_result["종별_모든발견"]:
                                    file_result["종별_모든발견"][type_key] = []
                                
                                file_result["종별_모든발견"][type_key].append(page_num + 1)
                                st.write(f"'{type_key}' 패턴 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'{type_key}' 패턴 발견: 페이지 {page_num + 1}")
                        
                        if not file_result["종별_모든발견"]:
                            st.write("종별 정보를 찾을 수 없습니다.")
                            file_result["상세 로그"].append("종별 정보를 찾을 수 없습니다.")
                        else:
                            st.write(f"총 {len(file_result['종별_모든발견'])}개 종 발견")
                            file_result["상세 로그"].append(f"총 {len(file_result['종별_모든발견'])}개 종 발견")
                            
                            # 종별 시작 페이지 식별
                            st.write("종별 시작 페이지 식별 중...")
                            file_result["상세 로그"].append("종별 시작 페이지 식별 중...")
                            
                            type_starts = identify_type_start_pages(pdf_document, parsing_start_page, total_pages)
                            
                            if type_starts:
                                for type_key, info in type_starts.items():
                                    file_result["종별_시작"][type_key] = info
                                    st.write(f"'{type_key}' 시작 페이지: {info['page']} (신뢰도: {info['confidence']})")
                                    st.write(f"특징: {', '.join(info['features'])}")
                                    file_result["상세 로그"].append(f"'{type_key}' 시작 페이지: {info['page']} (신뢰도: {info['confidence']})")
                                    file_result["상세 로그"].append(f"특징: {', '.join(info['features'])}")
                                
                                # 종별 범위 계산
                                st.write("종별 페이지 범위 계산 중...")
                                file_result["상세 로그"].append("종별 페이지 범위 계산 중...")
                                
                                type_ranges = calculate_type_ranges(type_starts, total_pages)
                                file_result["종별_범위"] = type_ranges
                                
                                for type_key, range_info in type_ranges.items():
                                    st.write(f"'{type_key}' 범위: {range_info['start_page']}~{range_info['end_page']}페이지")
                                    file_result["상세 로그"].append(f"'{type_key}' 범위: {range_info['start_page']}~{range_info['end_page']}페이지")
                            else:
                                st.write("종별 시작 페이지를 식별할 수 없습니다.")
                                file_result["상세 로그"].append("종별 시작 페이지를 식별할 수 없습니다.")
                        
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

                        # 상해관련특별약관이 발견되지 않은 경우에만 추가 패턴 검색
                        if not file_result["상해관련특별약관 페이지"]:
                            st.write("'상해관련특별약관'을 찾을 수 없어 추가 패턴 검색 중...")
                            file_result["상세 로그"].append("'상해관련특별약관'을 찾을 수 없어 추가 패턴 검색 중...")
                            
                            for page_num in range(parsing_start_page, total_pages):
                                page = pdf_document[page_num]
                                text = page.get_text()
                                text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                                
                                # 상해관련특약 패턴
                                if (("상해관련특약" in text_normalized or 
                                     "상해관련 특약" in text_normalized or 
                                     "상해 관련 특약" in text_normalized or
                                     ("선택특약" in text_normalized and "상해" in text_normalized and "질병" not in text_normalized)) and
                                    "상해및질병" not in text_normalized):
                                    
                                    found_pattern = ""
                                    if "상해관련특약" in text_normalized:
                                        found_pattern = "상해관련특약"
                                    elif "상해관련 특약" in text_normalized or "상해 관련 특약" in text_normalized:
                                        found_pattern = "상해관련 특약"
                                    elif "선택특약" in text_normalized and "상해" in text_normalized:
                                        found_pattern = "선택특약(상해)"
                                        
                                    file_result["상해관련특별약관 페이지"].append(page_num + 1)
                                    st.write(f"'{found_pattern}' 발견: 페이지 {page_num + 1}")
                                    file_result["상세 로그"].append(f"'{found_pattern}' 발견: 페이지 {page_num + 1}")

                        # 2-2. "질병관련특별약관" 검색
                        for page_num in range(parsing_start_page, total_pages):
                            page = pdf_document[page_num]
                            text = page.get_text()
                            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                            if "질병관련특별약관" in text_normalized and "상해및질병관련특별약관" not in text_normalized:
                                file_result["질병관련특별약관 페이지"].append(page_num + 1)
                                st.write(f"'질병관련특별약관' 발견: 페이지 {page_num + 1}")
                                file_result["상세 로그"].append(f"'질병관련특별약관' 발견: 페이지 {page_num + 1}")

                        # 질병관련특별약관이 발견되지 않은 경우에만 추가 패턴 검색
                        if not file_result["질병관련특별약관 페이지"]:
                            st.write("'질병관련특별약관'을 찾을 수 없어 추가 패턴 검색 중...")
                            file_result["상세 로그"].append("'질병관련특별약관'을 찾을 수 없어 추가 패턴 검색 중...")
                            
                            for page_num in range(parsing_start_page, total_pages):
                                page = pdf_document[page_num]
                                text = page.get_text()
                                text_normalized = ''.join(text.split())  # 띄어쓰기 제거
                                
                                # 질병관련특약 패턴
                                if (("질병관련특약" in text_normalized or 
                                     "질병관련 특약" in text_normalized or 
                                     "질병 관련 특약" in text_normalized or
                                     ("선택특약" in text_normalized and "질병" in text_normalized and "상해" not in text_normalized)) and
                                    "상해및질병" not in text_normalized):
                                    
                                    found_pattern = ""
                                    if "질병관련특약" in text_normalized:
                                        found_pattern = "질병관련특약"
                                    elif "질병관련 특약" in text_normalized or "질병 관련 특약" in text_normalized:
                                        found_pattern = "질병관련 특약"
                                    elif "선택특약" in text_normalized and "질병" in text_normalized:
                                        found_pattern = "선택특약(질병)"
                                        
                                    file_result["질병관련특별약관 페이지"].append(page_num + 1)
                                    st.write(f"'{found_pattern}' 발견: 페이지 {page_num + 1}")
                                    file_result["상세 로그"].append(f"'{found_pattern}' 발견: 페이지 {page_num + 1}")

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
                            
                            # 취소선 검사 (기존 코드 유지)
                            has_strikethrough = False
                            spans = page.get_text("dict")["blocks"]
                            for block in spans:
                                if "lines" in block:
                                    for line in block["lines"]:
                                        for span in line["spans"]:
                                            flags = span.get("flags", 0)
                                            if flags & 2**6:  # 취소선 비트 확인 (64)
                                                has_strikethrough = True
                                                break
                                        if has_strikethrough:
                                            break
                                    if has_strikethrough:
                                        break
                            
                            # OpenCV로 강조색 검사
                            try:
                                highlight_result = detect_highlights_with_opencv(pdf_document, page_num)
                                has_highlight = highlight_result["has_highlight"]
                                
                                if has_highlight:
                                    file_result["강조색 있는 페이지"].add(page_num + 1)
                                    
                                    # 색상별 정보 로깅 (선택 사항)
                                    detected_colors = []
                                    for color_name, is_detected in highlight_result["colors"].items():
                                        if is_detected:
                                            detected_colors.append(color_name)
                                    
                                    if detected_colors:
                                        color_info = f"페이지 {page_num + 1} 감지된 색상: {', '.join(detected_colors)}"
                                        file_result["상세 로그"].append(color_info)
                                    
                            except Exception as e:
                                error_msg = f"페이지 {page_num + 1} 강조색 분석 중 오류: {str(e)}"
                                file_result["상세 로그"].append(error_msg)
                                st.warning(error_msg)
                            
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
                        
                        # 종별 범위별 강조색/취소선 페이지 분류
                        if file_result["종별_범위"]:
                            st.write("종별 범위별 강조색/취소선 페이지 분류:")
                            file_result["상세 로그"].append("종별 범위별 강조색/취소선 페이지 분류:")
                            
                            for type_key, range_info in file_result["종별_범위"].items():
                                start_page = range_info["start_page"]
                                end_page = range_info["end_page"]
                                
                                # 해당 범위 내 강조색 페이지
                                highlight_pages = [p for p in file_result["강조색 있는 페이지"] if start_page <= p <= end_page]
                                # 해당 범위 내 취소선 페이지
                                strikethrough_pages = [p for p in file_result["취소선 있는 페이지"] if start_page <= p <= end_page]
                                
                                if highlight_pages:
                                    st.write(f"{type_key} 범위 내 강조색 페이지: {', '.join(map(str, highlight_pages))}")
                                    file_result["상세 로그"].append(f"{type_key} 범위 내 강조색 페이지: {', '.join(map(str, highlight_pages))}")
                                
                                if strikethrough_pages:
                                    st.write(f"{type_key} 범위 내 취소선 페이지: {', '.join(map(str, strikethrough_pages))}")
                                    file_result["상세 로그"].append(f"{type_key} 범위 내 취소선 페이지: {', '.join(map(str, strikethrough_pages))}")
                        
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
                    # 종별 정보 문자열로 변환
                    all_types_str = ""
                    if result["종별_모든발견"]:
                        types_arr = []
                        for type_key, pages in result["종별_모든발견"].items():
                            types_arr.append(f"{type_key}: {', '.join(map(str, pages))}")
                        all_types_str = "; ".join(types_arr)
                    else:
                        all_types_str = "없음"
                    
                    # 종별 범위 정보 문자열로 변환
                    type_ranges_str = ""
                    if result["종별_범위"]:
                        ranges_arr = []
                        for type_key, range_info in result["종별_범위"].items():
                            ranges_arr.append(f"{type_key}: {range_info['start_page']}~{range_info['end_page']}")
                        type_ranges_str = "; ".join(ranges_arr)
                    else:
                        type_ranges_str = "없음"
                    
                    summary_data.append({
                        "파일명": result["파일명"],
                        "나. 보험금 페이지": ', '.join(map(str, result["나. 보험금 페이지"])) if result["나. 보험금 페이지"] else "없음",
                        "종별 범위": type_ranges_str,  # 종별 범위 정보 추가
                        "종별 모든 발견": all_types_str,  # 모든 종별 발견 정보
                        "상해관련특별약관 페이지": ', '.join(map(str, result["상해관련특별약관 페이지"])) if result["상해관련특별약관 페이지"] else "없음",
                        "질병관련특별약관 페이지": ', '.join(map(str, result["질병관련특별약관 페이지"])) if result["질병관련특별약관 페이지"] else "없음",
                        "상해및질병관련특별약관 종료": str(result["상해및질병관련특별약관 종료 페이지"]) if result["상해및질병관련특별약관 종료 페이지"] is not None else "없음",
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