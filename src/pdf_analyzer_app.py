import os
import sys
import re
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
import cv2
import json
import pickle
from datetime import datetime
import tempfile
import glob
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
from PIL import Image, ImageTk
import io
import threading
import queue
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 전역 변수
analysis_results = []
current_file_index = 0
progress_var = None
status_var = None
log_text = None
result_frame = None

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

# 업무정의서 형식으로 엑셀 생성
def create_business_definition_excel(file_path, parsing_range, output_path=None):
    """테이블 데이터로 업무정의서 형식의 엑셀 생성"""
    # 파일명에서 상품명 추출
    product_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # Workbook 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "보장내용 개정사항"
    
    # 스타일 정의
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # 제목 행 추가
    ws.merge_cells('A1:F1')
    cell = ws.cell(row=1, column=1, value=f"{product_name} - 보장내용 개정사항")
    cell.font = Font(size=14, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 헤더 행 추가
    headers = ["페이지", "보장명", "지급사유", "지급금액", "비고", "강조여부"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 테이블 데이터 추가
    row_idx = 4
    
    # PDF 문서 열기
    pdf_document = fitz.open(file_path)
    
    # 파싱 범위 내의 각 페이지 처리
    for page_num in range(parsing_range[0], parsing_range[1] + 1):
        page = pdf_document[page_num]
        
        # 페이지에서 테이블 추출 시도
        # 여기서는 간단한 텍스트 기반 테이블 추출 로직 사용
        text = page.get_text()
        lines = text.split('\n')
        
        # 테이블 헤더 찾기
        header_idx = -1
        for i, line in enumerate(lines):
            if (("보장내용" in line or "보험금 지급사유" in line) and "지급금액" in line) or \
               ("담보" in line and "지급사유" in line) or \
               ("보장명" in line and "지급금액" in line):
                header_idx = i
                break
        
        # 헤더를 찾았으면 테이블 데이터 추출
        if header_idx >= 0:
            # 헤더 행 분석
            header_line = lines[header_idx]
            header_parts = []
            
            # 간단한 열 구분 (공백 기반)
            if "보장명" in header_line or "보장내용" in header_line or "담보" in header_line:
                header_parts.append("보장명")
            if "지급사유" in header_line or "보험금 지급사유" in header_line:
                header_parts.append("지급사유")
            if "지급금액" in header_line or "보험금액" in header_line:
                header_parts.append("지급금액")
            
            # 데이터 행 처리 (헤더 다음 행부터)
            for i in range(header_idx + 1, min(header_idx + 20, len(lines))):
                line = lines[i].strip()
                if not line:  # 빈 줄 건너뛰기
                    continue
                
                # 데이터 행 추가
                ws.cell(row=row_idx, column=1, value=page_num + 1).border = border  # 페이지 번호
                
                # 간단한 데이터 분할 (실제로는 더 복잡한 로직이 필요할 수 있음)
                parts = line.split('  ')  # 두 개 이상의 공백으로 분할
                parts = [p for p in parts if p.strip()]  # 빈 문자열 제거
                
                # 데이터 채우기
                for j, part in enumerate(parts[:3]):  # 최대 3개 열까지
                    if j < len(header_parts):
                        col_idx = headers.index(header_parts[j]) + 1
                        ws.cell(row=row_idx, column=col_idx, value=part.strip()).border = border
                
                # 비고와 강조여부는 빈칸으로 두기
                ws.cell(row=row_idx, column=5, value="").border = border  # 비고
                ws.cell(row=row_idx, column=6, value="").border = border  # 강조여부
                
                row_idx += 1
    
    # PDF 문서 닫기
    pdf_document.close()
    
    # 열 너비 조정
    ws.column_dimensions['A'].width = 10  # 페이지
    ws.column_dimensions['B'].width = 25  # 보장명
    ws.column_dimensions['C'].width = 40  # 지급사유
    ws.column_dimensions['D'].width = 25  # 지급금액
    ws.column_dimensions['E'].width = 15  # 비고
    ws.column_dimensions['F'].width = 10  # 강조여부
    
    # 엑셀 파일 저장
    if output_path:
        wb.save(output_path)
        return output_path
    else:
        # 기본 저장 경로
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"{product_name}_보장내용개정사항_{timestamp}.xlsx")
        wb.save(output_file)
        return output_file

# 결과 저장 함수
def save_analysis_results(analysis_results):
    """분석 결과를 파일로 저장"""
    # 출력 디렉토리 확인/생성
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "analysis_results")
    os.makedirs(output_dir, exist_ok=True)
    
    # 타임스탬프 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 결과 객체를 JSON과 피클로 저장
    for result in analysis_results:
        # 파일명에서 확장자 제거
        base_name = os.path.splitext(result["파일명"])[0]
        # 특수문자 제거
        safe_name = re.sub(r'[^\w\s-]', '', base_name)
        
        # 파일 경로 설정
        json_path = os.path.join(output_dir, f"{safe_name}_{timestamp}.json")
        pickle_path = os.path.join(output_dir, f"{safe_name}_{timestamp}.pkl")
        
        # JSON으로 저장 (set을 리스트로 변환하여 직렬화 가능하게 함)
        json_safe_result = result.copy()
        # set 객체를 리스트로 변환
        for key in json_safe_result:
            if isinstance(json_safe_result[key], set):
                json_safe_result[key] = sorted(list(json_safe_result[key]))
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_safe_result, f, ensure_ascii=False, indent=2)
        
        # 피클로 저장 (객체 그대로 유지)
        with open(pickle_path, 'wb') as f:
            pickle.dump(result, f)
    
    # 요약 파일도 저장
    summary_path = os.path.join(output_dir, f"summary_{timestamp}.json")
    
    # 요약 정보만 포함
    summary_data = []
    for result in analysis_results:
        summary_info = {
            "파일명": result["파일명"],
            "처리 상태": result["처리 상태"],
            "나. 보험금 페이지": result["나. 보험금 페이지"],
            "종별_범위": result["종별_범위"],
            "저장 시간": timestamp
        }
        summary_data.append(summary_info)
    
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(summary_data, f, ensure_ascii=False, indent=2)
    
    return output_dir, timestamp

# 로그 기록 함수
def log_results(analysis_results):
    """테스트 결과를 로그 파일에 기록하는 함수"""
    # 로그 폴더 확인/생성
    log_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_log")
    os.makedirs(log_folder, exist_ok=True)
    
    # 로그 파일명 설정
    log_filename = os.path.join(log_folder, f"test_log.txt")
    
    # 현재 시간
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 실행 파일 경로 및 파일명 가져오기
    script_path = os.path.abspath(__file__)
    script_name = os.path.basename(script_path)
    
    # 로그 기록 (기존 파일에 추가)
    with open(log_filename, 'a', encoding='utf-8') as log_file:
        log_file.write(f"\n\n===== 테스트 실행: {current_time} =====\n")
        log_file.write(f"실행 파일: {script_name}\n")
        log_file.write(f"파일 경로: {script_path}\n\n")
        
        for result in analysis_results:
            log_file.write(f"파일명: {result['파일명']}\n")
            log_file.write(f"처리 상태: {result['처리 상태']}\n")
            log_file.write(f"나. 보험금 페이지: {', '.join(map(str, result['나. 보험금 페이지'])) if result['나. 보험금 페이지'] else '없음'}\n")
            
            # 종별 범위 정보
            if result["종별_범위"]:
                log_file.write("종별 범위:\n")
                for type_key, range_info in result["종별_범위"].items():
                    log_file.write(f"  {type_key}: {range_info['start_page']}~{range_info['end_page']}페이지\n")
            else:
                log_file.write("종별 범위: 없음\n")
            
            # 특별약관 정보
            log_file.write(f"상해관련특별약관 페이지: {', '.join(map(str, result['상해관련특별약관 페이지'])) if result['상해관련특별약관 페이지'] else '없음'}\n")
            log_file.write(f"질병관련특별약관 페이지: {', '.join(map(str, result['질병관련특별약관 페이지'])) if result['질병관련특별약관 페이지'] else '없음'}\n")
            
            # 강조색 및 취소선
            log_file.write(f"강조색 있는 페이지: {', '.join(map(str, result['강조색 있는 페이지'])) if result['강조색 있는 페이지'] else '없음'}\n")
            log_file.write(f"취소선 있는 페이지: {', '.join(map(str, result['취소선 있는 페이지'])) if result['취소선 있는 페이지'] else '없음'}\n")
            
            # 구분선 추가
            log_file.write("--------------------------------------------------\n")
        
        log_file.write(f"\n테스트 완료: 총 {len(analysis_results)}개 파일 처리됨\n")
        log_file.write("==================================================\n")
    
    return log_filename

# GUI 클래스 정의
class PDFAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 보장내용 분석기 v12")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # 전역 변수 초기화
        self.analysis_results = []
        self.selected_files = []
        self.processing_thread = None
        self.message_queue = queue.Queue()
        
        # 메인 프레임 설정
        self.main_frame = ttk.Frame(root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 파일 선택 영역
        self.file_frame = ttk.LabelFrame(self.main_frame, text="PDF 파일 선택", padding=10)
        self.file_frame.pack(fill=tk.X, pady=5)
        
        self.file_btn = ttk.Button(self.file_frame, text="파일 선택", command=self.select_files)
        self.file_btn.pack(side=tk.LEFT, padx=5)
        
        self.file_label = ttk.Label(self.file_frame, text="선택된 파일 없음")
        self.file_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 진행 상황 영역
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="진행 상황", padding=10)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="대기 중...")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_var)
        self.status_label.pack(anchor=tk.W, pady=5)
        
        # 로그 영역
        self.log_frame = ttk.LabelFrame(self.main_frame, text="로그", padding=10)
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(self.log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 결과 영역
        self.result_frame = ttk.LabelFrame(self.main_frame, text="분석 결과", padding=10)
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 결과 표시 영역은 처음에는 비어 있음
        
        # 버튼 영역
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=10)
        
        self.start_btn = ttk.Button(self.button_frame, text="분석 시작", command=self.start_analysis)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(self.button_frame, text="결과 저장", command=self.save_results, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        self.excel_btn = ttk.Button(self.button_frame, text="엑셀 내보내기", command=self.export_excel, state=tk.DISABLED)
        self.excel_btn.pack(side=tk.LEFT, padx=5)
        
        # 메시지 처리 타이머 시작
        self.root.after(100, self.process_messages)
    
    def select_files(self):
        """파일 선택 대화상자 표시"""
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        )
        
        if files:
            self.selected_files = list(files)
            if len(files) == 1:
                self.file_label.config(text=os.path.basename(files[0]))
            else:
                self.file_label.config(text=f"{len(files)}개 파일 선택됨")
            
            # 로그에 선택된 파일 표시
            self.log_text.delete(1.0, tk.END)
            self.log_text.insert(tk.END, "선택된 파일:\n")
            for file in files:
                self.log_text.insert(tk.END, f"- {os.path.basename(file)}\n")
            
            # 버튼 활성화
            self.start_btn.config(state=tk.NORMAL)
        else:
            self.selected_files = []
            self.file_label.config(text="선택된 파일 없음")
    
    def start_analysis(self):
        """분석 시작"""
        if not self.selected_files:
            messagebox.showwarning("경고", "분석할 PDF 파일을 선택하세요.")
            return
        
        # 이전 결과 초기화
        self.analysis_results = []
        self.progress_var.set(0)
        self.status_var.set("분석 준비 중...")
        
        # 결과 영역 초기화
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        
        # 버튼 비활성화
        self.start_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.excel_btn.config(state=tk.DISABLED)
        
        # 로그 초기화
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "분석 시작...\n")
        
        # 분석 스레드 시작
        self.processing_thread = threading.Thread(target=self.process_files)
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def process_files(self):
        """파일 처리 스레드"""
        try:
            total_files = len(self.selected_files)
            
            for i, file_path in enumerate(self.selected_files):
                # 진행 상황 업데이트
                progress = (i / total_files) * 100
                self.message_queue.put(("progress", progress))
                
                # 파일 처리
                file_result = process_pdf_file(file_path, self.message_queue)
                
                # 결과 저장
                if file_result:
                    self.analysis_results.append(file_result)
            
            # 완료 메시지
            self.message_queue.put(("progress", 100))
            self.message_queue.put(("status", "분석 완료"))
            self.message_queue.put(("complete", None))
            
        except Exception as e:
            self.message_queue.put(("error", f"분석 중 오류 발생: {str(e)}"))
    
    def process_messages(self):
        """메시지 큐에서 메시지 처리"""
        try:
            while not self.message_queue.empty():
                msg_type, msg_data = self.message_queue.get_nowait()
                
                if msg_type == "progress":
                    self.progress_var.set(msg_data)
                
                elif msg_type == "status":
                    self.status_var.set(msg_data)
                    self.log_text.insert(tk.END, f"{msg_data}\n")
                    self.log_text.see(tk.END)
                
                elif msg_type == "log":
                    self.log_text.insert(tk.END, f"{msg_data}\n")
                    self.log_text.see(tk.END)
                
                elif msg_type == "warning":
                    self.log_text.insert(tk.END, f"경고: {msg_data}\n")
                    self.log_text.see(tk.END)
                
                elif msg_type == "error":
                    self.log_text.insert(tk.END, f"오류: {msg_data}\n")
                    self.log_text.see(tk.END)
                    messagebox.showerror("오류", msg_data)
                
                elif msg_type == "complete":
                    self.display_results()
                    self.start_btn.config(state=tk.NORMAL)
                    self.save_btn.config(state=tk.NORMAL)
                    self.excel_btn.config(state=tk.NORMAL)
        
        except Exception as e:
            print(f"메시지 처리 오류: {str(e)}")
        
        # 100ms 후 다시 호출
        self.root.after(100, self.process_messages)
    
    def display_results(self):
        """분석 결과 표시"""
        # 결과 영역 초기화
        for widget in self.result_frame.winfo_children():
            widget.destroy()
        
        if not self.analysis_results:
            ttk.Label(self.result_frame, text="분석 결과가 없습니다.").pack(pady=10)
            return
        
        # 결과 표시를 위한 노트북 생성
        notebook = ttk.Notebook(self.result_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 요약 탭
        summary_frame = ttk.Frame(notebook, padding=10)
        notebook.add(summary_frame, text="요약")
        
        # 요약 테이블 생성
        columns = ["파일명", "나. 보험금 페이지", "종별 범위", "강조색 페이지", "취소선 페이지"]
        summary_tree = ttk.Treeview(summary_frame, columns=columns, show="headings")
        
        # 열 설정
        for col in columns:
            summary_tree.heading(col, text=col)
            summary_tree.column(col, width=100)
        
        # 데이터 추가
        for result in self.analysis_results:
            # 종별 범위 문자열 생성
            type_ranges_str = ""
            if result["종별_범위"]:
                ranges_arr = []
                for type_key, range_info in result["종별_범위"].items():
                    ranges_arr.append(f"{type_key}: {range_info['start_page']}~{range_info['end_page']}")
                type_ranges_str = "; ".join(ranges_arr)
            
            # 강조색 페이지 문자열
            highlight_str = ", ".join(map(str, result["강조색 있는 페이지"])) if result["강조색 있는 페이지"] else "없음"
            
            # 취소선 페이지 문자열
            strikethrough_str = ", ".join(map(str, result["취소선 있는 페이지"])) if result["취소선 있는 페이지"] else "없음"
            
            # 트리뷰에 추가
            summary_tree.insert("", tk.END, values=(
                result["파일명"],
                ", ".join(map(str, result["나. 보험금 페이지"])) if result["나. 보험금 페이지"] else "없음",
                type_ranges_str,
                highlight_str,
                strikethrough_str
            ))
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(summary_frame, orient=tk.VERTICAL, command=summary_tree.yview)
        summary_tree.configure(yscrollcommand=scrollbar.set)
        
        # 배치
        summary_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 파일별 상세 탭 추가
        for result in self.analysis_results:
            detail_frame = ttk.Frame(notebook, padding=10)
            notebook.add(detail_frame, text=result["파일명"])
            
            # 상세 정보 표시
            detail_text = scrolledtext.ScrolledText(detail_frame, wrap=tk.WORD)
            detail_text.pack(fill=tk.BOTH, expand=True)
            
            # 상세 정보 추가
            detail_text.insert(tk.END, f"파일명: {result['파일명']}\n")
            detail_text.insert(tk.END, f"처리 상태: {result['처리 상태']}\n\n")
            
            # 나. 보험금 페이지
            detail_text.insert(tk.END, f"나. 보험금 페이지: {', '.join(map(str, result['나. 보험금 페이지'])) if result['나. 보험금 페이지'] else '없음'}\n\n")
            
            # 종별 범위
            if result["종별_범위"]:
                detail_text.insert(tk.END, "종별 범위:\n")
                for type_key, range_info in result["종별_범위"].items():
                    detail_text.insert(tk.END, f"  {type_key}: {range_info['start_page']}~{range_info['end_page']}페이지\n")
                    detail_text.insert(tk.END, f"  특징: {', '.join(range_info['features'])}\n")
                detail_text.insert(tk.END, "\n")
            
            # 특별약관 정보
            detail_text.insert(tk.END, f"상해관련특별약관 페이지: {', '.join(map(str, result['상해관련특별약관 페이지'])) if result['상해관련특별약관 페이지'] else '없음'}\n")
            detail_text.insert(tk.END, f"질병관련특별약관 페이지: {', '.join(map(str, result['질병관련특별약관 페이지'])) if result['질병관련특별약관 페이지'] else '없음'}\n")
            detail_text.insert(tk.END, f"상해및질병관련특별약관 페이지: {', '.join(map(str, result['상해및질병관련특별약관 페이지'])) if result['상해및질병관련특별약관 페이지'] else '없음'}\n")
            detail_text.insert(tk.END, f"상해및질병관련특별약관 종료 페이지: {result['상해및질병관련특별약관 종료 페이지'] if result['상해및질병관련특별약관 종료 페이지'] is not None else '없음'}\n\n")
            
            # 강조색 및 취소선
            detail_text.insert(tk.END, f"강조색 있는 페이지: {', '.join(map(str, result['강조색 있는 페이지'])) if result['강조색 있는 페이지'] else '없음'}\n")
            detail_text.insert(tk.END, f"취소선 있는 페이지: {', '.join(map(str, result['취소선 있는 페이지'])) if result['취소선 있는 페이지'] else '없음'}\n\n")
            
            # 상세 로그
            detail_text.insert(tk.END, "상세 로그:\n")
            for log_entry in result["상세 로그"]:
                detail_text.insert(tk.END, f"{log_entry}\n")
            
            # 읽기 전용으로 설정
            detail_text.config(state=tk.DISABLED)
    
    def save_results(self):
        """분석 결과 저장"""
        if not self.analysis_results:
            messagebox.showwarning("경고", "저장할 분석 결과가 없습니다.")
            return
        
        try:
            output_dir, timestamp = save_analysis_results(self.analysis_results)
            messagebox.showinfo("저장 완료", f"분석 결과가 저장되었습니다.\n저장 위치: {output_dir}")
            
            # 로그에 저장 정보 추가
            self.log_text.insert(tk.END, f"\n분석 결과 저장 완료: {output_dir} (타임스탬프: {timestamp})\n")
            self.log_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"결과 저장 중 오류 발생: {str(e)}")
    
    def export_excel(self):
        """엑셀 내보내기"""
        if not self.analysis_results or not self.selected_files:
            messagebox.showwarning("경고", "내보낼 분석 결과가 없습니다.")
            return
        
        try:
            # 파일 선택 대화상자
            file_path = filedialog.asksaveasfilename(
                title="엑셀 파일 저장",
                filetypes=[("Excel 파일", "*.xlsx")],
                defaultextension=".xlsx"
            )
            
            if not file_path:
                return
            
            # 첫 번째 파일에 대해 엑셀 생성
            result = self.analysis_results[0]
            file_path_orig = self.selected_files[0]
            
            # 파싱 범위 설정
            parsing_start = 0
            if result["나. 보험금 페이지"]:
                parsing_start = result["나. 보험금 페이지"][0] - 1
            
            parsing_end = None
            if result["상해및질병관련특별약관 종료 페이지"]:
                parsing_end = result["상해및질병관련특별약관 종료 페이지"] - 1
            else:
                with fitz.open(file_path_orig) as doc:
                    parsing_end = len(doc) - 1
            
            # 엑셀 생성
            output_path = create_business_definition_excel(
                file_path_orig,
                (parsing_start, parsing_end),
                file_path
            )
            
            messagebox.showinfo("내보내기 완료", f"엑셀 파일이 생성되었습니다.\n저장 위치: {output_path}")
            
            # 로그에 저장 정보 추가
            self.log_text.insert(tk.END, f"\n엑셀 파일 생성 완료: {output_path}\n")
            self.log_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("내보내기 오류", f"엑셀 생성 중 오류 발생: {str(e)}")

# 분석 작업 실행 함수
def process_pdf_file(file_path, queue):
    """PDF 파일 분석 작업 실행"""
    try:
        file_name = os.path.basename(file_path)
        queue.put(("status", f"파일 처리 중: {file_name}"))
        
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
        
        # PDF 처리 로직
        pdf_document = fitz.open(file_path)
        total_pages = len(pdf_document)
        queue.put(("log", f"총 {total_pages}페이지 로드됨"))
        
        # 1. "나. 보험금" 검색하여 파싱 시작 페이지 찾기
        parsing_start_page = None
        for page_num in range(total_pages):
            page = pdf_document[page_num]
            text = page.get_text()
            if "나. 보험금" in text:
                parsing_start_page = page_num
                file_result["나. 보험금 페이지"].append(page_num + 1)
                queue.put(("log", f"'나. 보험금' 문구 발견: 페이지 {page_num + 1}"))
                file_result["상세 로그"].append(f"'나. 보험금' 문구 발견: 페이지 {page_num + 1}")
                break
        
        # 파싱 시작 페이지를 찾지 못한 경우 전체 문서를 대상으로 함
        if parsing_start_page is None:
            parsing_start_page = 0
            queue.put(("warning", "'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다."))
            file_result["상세 로그"].append("'나. 보험금' 문구를 찾을 수 없어 첫 페이지부터 검색합니다.")
        
        # 종별 검색 추가 (모든 종별 발견 페이지 검색)
        queue.put(("log", "종별 패턴 검색 시작..."))
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
                queue.put(("log", f"'{type_key}' 패턴 발견: 페이지 {page_num + 1}"))
                file_result["상세 로그"].append(f"'{type_key}' 패턴 발견: 페이지 {page_num + 1}")
        
        if not file_result["종별_모든발견"]:
            queue.put(("log", "종별 정보를 찾을 수 없습니다."))
            file_result["상세 로그"].append("종별 정보를 찾을 수 없습니다.")
        else:
            queue.put(("log", f"총 {len(file_result['종별_모든발견'])}개 종 발견"))
            file_result["상세 로그"].append(f"총 {len(file_result['종별_모든발견'])}개 종 발견")
            
            # 종별 시작 페이지 식별
            queue.put(("log", "종별 시작 페이지 식별 중..."))
            file_result["상세 로그"].append("종별 시작 페이지 식별 중...")
            
            type_starts = identify_type_start_pages(pdf_document, parsing_start_page, total_pages)
            
            if type_starts:
                for type_key, info in type_starts.items():
                    file_result["종별_시작"][type_key] = info
                    queue.put(("log", f"'{type_key}' 시작 페이지: {info['page']} (신뢰도: {info['confidence']})"))
                    queue.put(("log", f"특징: {', '.join(info['features'])}"))
                    file_result["상세 로그"].append(f"'{type_key}' 시작 페이지: {info['page']} (신뢰도: {info['confidence']})")
                    file_result["상세 로그"].append(f"특징: {', '.join(info['features'])}")
                
                # 종별 범위 계산
                queue.put(("log", "종별 페이지 범위 계산 중..."))
                file_result["상세 로그"].append("종별 페이지 범위 계산 중...")
                
                type_ranges = calculate_type_ranges(type_starts, total_pages)
                file_result["종별_범위"] = type_ranges
                
                for type_key, range_info in type_ranges.items():
                    queue.put(("log", f"'{type_key}' 범위: {range_info['start_page']}~{range_info['end_page']}페이지"))
                    file_result["상세 로그"].append(f"'{type_key}' 범위: {range_info['start_page']}~{range_info['end_page']}페이지")
            else:
                queue.put(("log", "종별 시작 페이지를 식별할 수 없습니다."))
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
                queue.put(("log", f"'상해관련특별약관' 발견: 페이지 {page_num + 1}"))
                file_result["상세 로그"].append(f"'상해관련특별약관' 발견: 페이지 {page_num + 1}")

        # 상해관련특별약관이 발견되지 않은 경우에만 추가 패턴 검색
        if not file_result["상해관련특별약관 페이지"]:
            queue.put(("log", "'상해관련특별약관'을 찾을 수 없어 추가 패턴 검색 중..."))
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
                    queue.put(("log", f"'{found_pattern}' 발견: 페이지 {page_num + 1}"))
                    file_result["상세 로그"].append(f"'{found_pattern}' 발견: 페이지 {page_num + 1}")

        # 2-2. "질병관련특별약관" 검색
        for page_num in range(parsing_start_page, total_pages):
            page = pdf_document[page_num]
            text = page.get_text()
            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
            if "질병관련특별약관" in text_normalized and "상해및질병관련특별약관" not in text_normalized:
                file_result["질병관련특별약관 페이지"].append(page_num + 1)
                queue.put(("log", f"'질병관련특별약관' 발견: 페이지 {page_num + 1}"))
                file_result["상세 로그"].append(f"'질병관련특별약관' 발견: 페이지 {page_num + 1}")

        # 질병관련특별약관이 발견되지 않은 경우에만 추가 패턴 검색
        if not file_result["질병관련특별약관 페이지"]:
            queue.put(("log", "'질병관련특별약관'을 찾을 수 없어 추가 패턴 검색 중..."))
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
                    queue.put(("log", f"'{found_pattern}' 발견: 페이지 {page_num + 1}"))
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
                queue.put(("log", f"'상해및질병관련특별약관' 발견: 페이지 {page_num + 1}"))
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
            queue.put(("log", f"'상해및질병관련특별약관' 종료: 페이지 {end_page}"))
            file_result["상세 로그"].append(f"'상해및질병관련특별약관' 종료: 페이지 {end_page}")
        
        # 3. 강조색 및 취소선 검색 - "나. 보험금" 페이지부터 파싱 끝까지만 검색
        # 검색 범위 설정
        search_start_page = parsing_start_page  # "나. 보험금" 페이지 (또는 기본값 0)
        search_end_page = total_pages - 1  # 기본값은 문서 끝까지

        # 상해및질병관련특별약관 종료 페이지가 있으면 그것을 종료 범위로 설정
        if file_result["상해및질병관련특별약관 종료 페이지"] is not None:
            search_end_page = file_result["상해및질병관련특별약관 종료 페이지"] - 1  # 페이지 번호를 인덱스로 변환

        # 설정된 범위 내에서만 강조색/취소선 검색
        queue.put(("log", f"강조색/취소선 검색 범위: 페이지 {search_start_page + 1}~{search_end_page + 1}"))
        file_result["상세 로그"].append(f"강조색/취소선 검색 범위: 페이지 {search_start_page + 1}~{search_end_page + 1}")

        for page_num in range(search_start_page, search_end_page + 1):
            page = pdf_document[page_num]
            
            # 취소선 검사
            has_strikethrough = False
            spans = page.get_text("dict")["blocks"]
            for block in spans:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            flags = span.get("flags", 0)
