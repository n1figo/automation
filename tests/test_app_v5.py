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
import camelot
import logging
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from PIL import Image
import io

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pdf_analyzer.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("pdf_parser")

# ===== 종별 분석 함수 =====

def identify_type_start_pages(pdf_document, parsing_start_page, total_pages):
    """
    PDF 문서에서 각 종별([1종], [2종] 등) 시작 페이지를 식별합니다.
    """
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
            
            # 3. 종별 강조 확인
            type_emphasis = text.count(type_key) > 0
            
            # 4. 페이지 신뢰도 계산
            confidence = 0
            if table_headers:
                confidence += 3
            if has_sections:
                confidence += 2
            if type_emphasis:
                confidence += 2
            if page_num == first_page:
                confidence += 1
            
            # 가장 유력한 시작 페이지 갱신
            if best_page is None or confidence > best_confidence:
                best_page = page_num
                best_confidence = confidence
        
        # 충분한 신뢰도를 가진 경우에만 시작 페이지로 등록
        if best_confidence >= 3:
            type_starts[type_key] = best_page
    
    return type_starts

def calculate_type_ranges(type_starts, total_pages):
    """
    각 종별 시작 페이지를 기준으로 페이지 범위를 계산합니다.
    """
    type_ranges = {}
    sorted_types = sorted(type_starts.items(), key=lambda x: x[1])
    
    for i, (type_key, start_page) in enumerate(sorted_types):
        if i < len(sorted_types) - 1:
            end_page = sorted_types[i+1][1] - 1
        else:
            end_page = total_pages - 1
        
        type_ranges[type_key] = (start_page, end_page)
    
    return type_ranges

# ===== 하이라이트 감지 함수 =====

def detect_highlights_with_opencv(pdf_document, page_num):
    """
    OpenCV를 사용하여 PDF 페이지의 하이라이트된 영역을 감지합니다.
    """
    pix = pdf_document[page_num].get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
    img_bytes = pix.tobytes("png")
    
    # OpenCV로 이미지 로드
    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    
    # BGR에서 HSV 색 공간으로 변환
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    
    # 노란색 하이라이트 감지를 위한 범위 설정
    # 노란색 범위는 환경에 따라 조정이 필요할 수 있음
    lower_yellow = np.array([20, 100, 100])
    upper_yellow = np.array([30, 255, 255])
    yellow_mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
    
    # 빨간색 하이라이트 감지 (HSV에서 빨간색은 두 범위로 나뉨)
    lower_red1 = np.array([0, 100, 100])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([160, 100, 100])
    upper_red2 = np.array([180, 255, 255])
    
    red_mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    red_mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    red_mask = cv2.bitwise_or(red_mask1, red_mask2)
    
    # 모든 하이라이트 마스크 합치기
    highlight_mask = cv2.bitwise_or(yellow_mask, red_mask)
    
    # 하이라이트된 영역 찾기
    contours, _ = cv2.findContours(highlight_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 너무 작은 영역은 제외
    min_area = 100  # 최소 영역 크기 임계값
    highlight_regions = []
    
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > min_area:
            x, y, w, h = cv2.boundingRect(contour)
            highlight_regions.append((x, y, w, h))
    
    # 이미지 크기와 함께 하이라이트 영역 반환
    return highlight_regions, img.shape[:2]

def detect_highlights_with_adaptive_threshold(pdf_document, page_num):
    """
    동적 임계값을 사용한 하이라이트 감지 - 기존 함수보다 더 정교한 감지가 가능합니다.
    """
    pix = pdf_document[page_num].get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
    img_bytes = pix.tobytes("png")
    
    # OpenCV로 이미지 로드
    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    
    # BGR에서 HSV 색 공간으로 변환
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    
    # 이미지 히스토그램 분석
    h_hist = cv2.calcHist([hsv], [0], None, [180], [0, 180])
    s_hist = cv2.calcHist([hsv], [1], None, [256], [0, 256])
    
    # 노란색 하이라이트의 동적 범위 결정
    # 기본 노란색 범위
    lower_yellow = np.array([20, 100, 100])
    upper_yellow = np.array([30, 255, 255])
    
    # 채도(S)와 명도(V) 임계값 적응형 조정
    s_thresh = max(100, np.mean(s_hist[100:200]) * 0.8)
    lower_yellow[1] = s_thresh
    
    yellow_mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
    
    # 빨간색 하이라이트 감지 - 동적 범위
    lower_red1 = np.array([0, s_thresh, 100])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([160, s_thresh, 100])
    upper_red2 = np.array([180, 255, 255])
    
    red_mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    red_mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    red_mask = cv2.bitwise_or(red_mask1, red_mask2)
    
    # 모든 하이라이트 마스크 합치기
    highlight_mask = cv2.bitwise_or(yellow_mask, red_mask)
    
    # 노이즈 제거를 위한 모폴로지 연산
    kernel = np.ones((3,3), np.uint8)
    highlight_mask = cv2.morphologyEx(highlight_mask, cv2.MORPH_OPEN, kernel)
    highlight_mask = cv2.morphologyEx(highlight_mask, cv2.MORPH_CLOSE, kernel)
    
    # 하이라이트된 영역 찾기
    contours, _ = cv2.findContours(highlight_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 최소 영역 크기 임계값 - 페이지 크기에 따라 동적으로 조정
    page_area = img.shape[0] * img.shape[1]
    min_area = page_area * 0.0005  # 페이지 면적의 0.05%
    
    highlight_regions = []
    
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > min_area:
            x, y, w, h = cv2.boundingRect(contour)
            highlight_regions.append((x, y, w, h))
    
    # 이미지 크기와 함께 하이라이트 영역 반환
    return highlight_regions, img.shape[:2]

def map_highlights_to_text(pdf_document, page_num, highlight_regions, img_size):
    """
    감지된 하이라이트 영역과 텍스트 블록을 매핑합니다.
    """
    page = pdf_document[page_num]
    img_height, img_width = img_size
    
    # 페이지의 실제 크기와 이미지 크기 사이의 비율 계산
    page_rect = page.rect
    scale_x = img_width / page_rect.width
    scale_y = img_height / page_rect.height
    
    # 페이지의 텍스트 블록 가져오기
    text_blocks = page.get_text("blocks")
    
    # 하이라이트 영역과 겹치는 텍스트 블록 찾기
    highlighted_texts = []
    
    for block in text_blocks:
        block_rect = fitz.Rect(block[:4])
        
        # 텍스트 블록의 좌표를 이미지 좌표계로 변환
        block_img_rect = fitz.Rect(
            block_rect.x0 * scale_x,
            block_rect.y0 * scale_y,
            block_rect.x1 * scale_x,
            block_rect.y1 * scale_y
        )
        
        for h_region in highlight_regions:
            h_x, h_y, h_w, h_h = h_region
            h_rect = fitz.Rect(h_x, h_y, h_x + h_w, h_y + h_h)
            
            # 텍스트 블록과 하이라이트 영역의 겹침 확인
            if block_img_rect.intersects(h_rect):
                overlap_area = block_img_rect.intersect(h_rect).get_area()
                block_area = block_img_rect.get_area()
                
                # 30% 이상 겹치면 하이라이트된 텍스트로 간주
                if overlap_area / block_area > 0.3:
                    highlighted_texts.append({
                        "text": block[4],
                        "rect": block[:4],
                        "highlight_region": h_region
                    })
                    break
    
    return highlighted_texts

# ===== Camelot을 사용한 테이블 추출 함수 =====

def extract_tables_with_camelot(pdf_path, page_num):
    """
    Camelot을 사용하여 PDF 페이지에서 테이블을 추출합니다.
    Lattice와 Stream 모드를 모두 시도하고 더 나은 결과를 선택합니다.
    """
    try:
        # Lattice 모드로 테이블 추출 (격자선이 있는 테이블)
        tables_lattice = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num + 1),  # Camelot은 1-indexed 페이지 번호 사용
            flavor='lattice'
        )
        
        # Stream 모드로 테이블 추출 (공백으로 구분된 테이블)
        tables_stream = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num + 1),
            flavor='stream'
        )
        
        # 결과 평가
        lattice_tables = []
        stream_tables = []
        
        for table in tables_lattice:
            if table.df.size > 0 and table.accuracy > 50:
                lattice_tables.append(table)
        
        for table in tables_stream:
            if table.df.size > 0 and table.accuracy > 50:
                stream_tables.append(table)
        
        # 더 많은 테이블을 추출한 방식 선택
        if len(lattice_tables) >= len(stream_tables):
            return lattice_tables
        else:
            return stream_tables
            
    except Exception as e:
        logger.error(f"테이블 추출 오류 (페이지 {page_num+1}): {str(e)}")
        return []

def classify_table_section(table_df):
    """
    테이블 내용을 분석하여 상해, 질병, 상해및질병 섹션으로 분류합니다.
    """
    # 테이블을 텍스트로 변환하여 분석
    table_text = table_df.to_string()
    
    # 섹션 분류 키워드
    injury_keywords = ["상해", "골절", "재해", "화상", "외상", "사고"]
    disease_keywords = ["질병", "암", "진단", "수술", "입원", "요양", "장애", "사망"]
    
    # 키워드 출현 빈도 계산
    injury_count = sum(table_text.count(keyword) for keyword in injury_keywords)
    disease_count = sum(table_text.count(keyword) for keyword in disease_keywords)
    
    # 관계적 맥락 분석을 통한 분류
    if "상해 및 질병" in table_text or "상해및질병" in table_text:
        return "상해및질병"
    elif injury_count > 0 and disease_count > 0:
        # 두 키워드 유형이 모두 발견되고, 상해/질병이 명시적으로 구분되지 않은 경우
        return "상해및질병"
    elif injury_count > disease_count:
        return "상해"
    elif disease_count > injury_count:
        return "질병"
    else:
        # 명확한 패턴이 없는 경우 기본값
        return "기타"

def process_single_page(pdf_document, pdf_path, page_num, detect_highlights=True):
    """
    단일 페이지 처리 함수 - 병렬 처리 시 사용됩니다.
    """
    result = {
        "page_num": page_num,
        "has_highlights": False,
        "highlighted_texts": [],
        "tables": []
    }
    
    # 하이라이트 감지 (필요한 경우)
    if detect_highlights:
        highlight_regions, img_size = detect_highlights_with_adaptive_threshold(pdf_document, page_num)
        if highlight_regions:
            result["has_highlights"] = True
            result["highlighted_texts"] = map_highlights_to_text(pdf_document, page_num, highlight_regions, img_size)
    
    # 테이블 추출
    tables = extract_tables_with_camelot(pdf_path, page_num)
    for i, table in enumerate(tables):
        table_data = {
            "df": table.df,
            "accuracy": table.accuracy,
            "section": classify_table_section(table.df)
        }
        result["tables"].append(table_data)
    
    return result

def extract_tables_for_type(pdf_document, pdf_path, type_key, page_range, highlighted_pages_only=False):
    """
    특정 종(type_key)의 페이지 범위에서 테이블을 추출합니다.
    """
    start_page, end_page = page_range
    results = []
    
    # 병렬 처리를 위한 작업 큐 생성
    tasks = []
    for page_num in range(start_page, end_page + 1):
        if not highlighted_pages_only or page_num in highlighted_pages_only:
            tasks.append((pdf_document, pdf_path, page_num))
    
    # ThreadPoolExecutor를 사용한 병렬 처리
    with ThreadPoolExecutor(max_workers=4) as executor:
        page_results = list(executor.map(
            lambda args: process_single_page(*args),
            tasks
        ))
    
    return {
        "type_key": type_key,
        "page_range": page_range,
        "pages": page_results
    }

# ===== 엑셀 출력 함수 =====

def generate_excel_output(results, output_path):
    """
    분석 결과를 엑셀로 출력합니다.
    시트는 종별로 구성되며, 각 시트는 상해/질병/상해및질병 섹션으로 구분됩니다.
    """
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    
    # 하이라이트 셀 서식
    highlight_format = workbook.add_format({'bg_color': '#FFEB9C'})
    
    # 제목 셀 서식
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9D9D9',
        'border': 1
    })
    
    # 종별 요약 시트 생성
    summary_df = pd.DataFrame(columns=['종류', '페이지 범위', '하이라이트 페이지 수', '테이블 수'])
    
    # 각 종별 데이터 처리
    for type_result in results:
        type_key = type_result["type_key"]
        page_range = type_result["page_range"]
        pages = type_result["pages"]
        
        # 하이라이트된 페이지 및 테이블 수 계산
        highlighted_pages = sum(1 for page in pages if page["has_highlights"])
        total_tables = sum(len(page["tables"]) for page in pages)
        
        # 요약 정보 추가
        summary_df = pd.concat([summary_df, pd.DataFrame([{
            '종류': type_key,
            '페이지 범위': f"{page_range[0]+1}-{page_range[1]+1}",
            '하이라이트 페이지 수': highlighted_pages,
            '테이블 수': total_tables
        }])], ignore_index=True)
        
        # 해당 종의 모든 테이블 데이터 수집
        all_tables = []
        for page in pages:
            for table_data in page["tables"]:
                # 기본 테이블 데이터
                table_info = {
                    "페이지": page["page_num"] + 1,
                    "섹션": table_data["section"],
                    "정확도": table_data["accuracy"],
                    "하이라이트": "있음" if page["has_highlights"] else "없음"
                }
                
                # 테이블 데이터프레임 처리
                df = table_data["df"].copy()
                
                # 컬럼 인덱스가 여러 레벨인 경우 처리
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [' '.join(col).strip() for col in df.columns.values]
                
                # 메타데이터 칼럼 추가
                for key, value in table_info.items():
                    df[key] = value
                
                all_tables.append(df)
        
        if not all_tables:
            continue
        
        # 테이블 병합
        combined_df = pd.concat(all_tables, ignore_index=True)
        
        # 섹션별 시트 생성
        sections = ['상해', '질병', '상해및질병', '기타']
        
        for section in sections:
            section_df = combined_df[combined_df["섹션"] == section].copy()
            
            if not section_df.empty:
                sheet_name = f"{type_key}_{section}"[:31]  # 시트 이름 길이 제한
                
                # 시트에 데이터 저장
                section_df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                
                # 하이라이트 셀에 서식 적용
                for i, row in section_df.iterrows():
                    if row["하이라이트"] == "있음":
                        for col_idx in range(len(section_df.columns)):
                            worksheet.write(i+1, col_idx, section_df.iloc[i, col_idx], highlight_format)
    
    # 요약 시트 저장
    summary_df.to_excel(writer, sheet_name='요약', index=False)
    
    writer.close()
    return output_path

# ===== 메인 파싱 파이프라인 =====

def process_pdf(pdf_path, parsing_start_page=0, highlighted_pages_only=True, output_dir="output"):
    """
    PDF 파싱 전체 파이프라인을 실행합니다.
    """
    logger.info(f"PDF 분석 시작: {pdf_path}")
    
    # 출력 디렉토리 생성
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # PDF 문서 열기
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)
    
    # 1. 종별 시작 페이지 식별
    logger.info("종별 시작 페이지 식별 중...")
    type_start_pages = identify_type_start_pages(pdf_document, parsing_start_page, total_pages)
    logger.info(f"식별된 종별 시작 페이지: {type_start_pages}")
    
    # 2. 종별 페이지 범위 계산
    type_ranges = calculate_type_ranges(type_start_pages, total_pages)
    logger.info(f"종별 페이지 범위: {type_ranges}")
    
    # 3. 하이라이트된 페이지 식별 (필요한 경우)
    highlighted_pages = set()
    if highlighted_pages_only:
        logger.info("하이라이트된 페이지 스캔 중...")
        for page_num in range(parsing_start_page, total_pages):
            highlight_regions, _ = detect_highlights_with_adaptive_threshold(pdf_document, page_num)
            if highlight_regions:
                highlighted_pages.add(page_num)
        logger.info(f"하이라이트된 페이지: {highlighted_pages}")
    
    # 4. 종별로 테이블 추출
    logger.info("테이블 추출 중...")
    results = []
    for type_key, page_range in type_ranges.items():
        type_result = extract_tables_for_type(
            pdf_document, 
            pdf_path,
            type_key, 
            page_range, 
            highlighted_pages if highlighted_pages_only else None
        )
        results.append(type_result)
    
    # 5. 엑셀 출력
    output_filename = os.path.join(output_dir, f"분석결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    output_path = generate_excel_output(results, output_filename)
    logger.info(f"분석 결과 저장 완료: {output_path}")
    
    # PDF 문서 닫기
    pdf_document.close()
    
    return output_path

# ===== Streamlit UI =====

def main():
    st.title("보험약관 분석 도구 v5")
    
    st.markdown("""
    ### 기능
    - PDF 보험약관에서 종별로 테이블 추출
    - 하이라이트된 내용 감지 및 표시
    - 상해/질병/상해및질병 섹션 자동 분류
    - 종별 엑셀 시트 생성
    """)
    
    uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type="pdf")
    
    if uploaded_file:
        st.write("파일 업로드 완료:", uploaded_file.name)
        
        # 임시 파일로 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.read())
            pdf_path = tmp_file.name
        
        # 설정 옵션
        parsing_start_page = st.number_input("분석 시작 페이지 (0부터 시작)", min_value=0, value=0)
        highlighted_only = st.checkbox("하이라이트된 페이지만 분석", value=True)
        
        if st.button("분석 시작"):
            with st.spinner("PDF 분석 중..."):
                try:
                    output_path = process_pdf(
                        pdf_path, 
                        parsing_start_page=parsing_start_page,
                        highlighted_pages_only=highlighted_only
                    )
                    
                    # 결과 다운로드 링크 제공
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="분석 결과 다운로드",
                            data=file,
                            file_name=f"분석결과_{uploaded_file.name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    st.success("분석이 완료되었습니다.")
                except Exception as e:
                    st.error(f"오류 발생: {str(e)}")
                    logger.error(f"분석 오류: {str(e)}", exc_info=True)
                finally:
                    # 임시 파일 삭제
                    if os.path.exists(pdf_path):
                        os.unlink(pdf_path)

if __name__ == "__main__":
    main()
