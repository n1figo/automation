# 범위성공
# 강조색 엑셀 성공
# 보완 : 형광펜 페이지 일부 잘리는 현상
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
import json
import pickle
import camelot
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# 종별 시작 페이지 식별 함수 (기존 함수 유지)
def identify_type_start_pages(pdf_document, parsing_start_page, total_pages):
    # (기존 코드 유지)
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

# 종별 범위 계산 함수 (기존 함수 유지)
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

# 개선된 색상 분석과 강조 영역 감지를 위한 클래스
class HighlightAnalyzer:
    def __init__(self):
        # 색상 범위 설정
        self.color_ranges = {
            "yellow": {
                "lower": np.array([20, 100, 200]),
                "upper": np.array([40, 255, 255])
            },
            "light_yellow": {
                "lower": np.array([20, 30, 200]),
                "upper": np.array([40, 100, 255])
            },
            "light_blue": {
                "lower": np.array([85, 30, 200]),
                "upper": np.array([115, 100, 255])
            },
            "orange": {
                "lower": np.array([5, 100, 200]),
                "upper": np.array([20, 255, 255])
            },
            "green": {
                "lower": np.array([40, 100, 200]),
                "upper": np.array([80, 255, 255])
            },
            "purple": {
                "lower": np.array([130, 50, 200]),
                "upper": np.array([170, 255, 255])
            },
            "gray": {
                "lower": np.array([0, 0, 80]),
                "upper": np.array([180, 40, 200])
            }
        }
    
    def analyze_document_colors(self, hsv_img):
        """문서 특성에 따른 적응형 색상 분석"""
        # 이미지 히스토그램 분석
        h_hist = cv2.calcHist([hsv_img], [0], None, [180], [0, 180])
        s_hist = cv2.calcHist([hsv_img], [1], None, [256], [0, 256])
        v_hist = cv2.calcHist([hsv_img], [2], None, [256], [0, 256])
        
        # 주요 색상 식별
        h_peaks = [i for i in range(1, 179) if h_hist[i] > h_hist[i-1] and h_hist[i] > h_hist[i+1]]
        
        # 문서 특성 분석 (밝기, 채도 분포)
        avg_v = np.average(np.arange(256), weights=v_hist.flatten())
        avg_s = np.average(np.arange(256), weights=s_hist.flatten())
        
        # 문서 특성에 맞게 임계값 조정
        thresholds = {
            "bright": 0.05 if avg_v > 200 else 0.1,  # 밝은 문서는 낮은 임계값
            "dark": 0.15 if avg_v < 150 else 0.1,    # 어두운 문서는 높은 임계값
            "colorful": 0.08 if avg_s > 100 else 0.1  # 채도가 높은 문서
        }
        
        # 적응형 임계값 반환
        return {
            "base_threshold": thresholds["bright"] if avg_v > 200 else 
                             thresholds["dark"] if avg_v < 150 else 0.1,
            "yellow_threshold": thresholds["bright"] * 0.8,  # 노란색은 더 낮은 임계값
            "gray_threshold": thresholds["dark"] * 1.2       # 회색은 더 높은 임계값
        }
    
    def detect_highlights(self, pdf_document, page_num):
        """개선된 강조색 감지 함수"""
        try:
            # 페이지 렌더링 (고해상도로)
            zoom = 2.0
            mat = fitz.Matrix(zoom, zoom)
            page = pdf_document[page_num]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # OpenCV 형식으로 변환
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            
            # HSV 색상 모델로 변환
            hsv = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2HSV)
            
            # 문서 특성에 맞는 임계값 계산
            thresholds = self.analyze_document_colors(hsv)
            
            # 색상별 마스크 생성
            masks = {}
            for color_name, ranges in self.color_ranges.items():
                mask = cv2.inRange(hsv, ranges["lower"], ranges["upper"])
                kernel = np.ones((3, 3), np.uint8)
                mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
                masks[color_name] = mask
            
            # 강조색 마스크 결합
            highlight_mask = masks["yellow"] | masks["light_yellow"] | masks["light_blue"] | masks["orange"] | masks["green"] | masks["purple"]
            
            # 색상별 검출 영역 비율 계산
            total_pixels = hsv.shape[0] * hsv.shape[1]
            
            color_ratios = {}
            for color_name, mask in masks.items():
                color_ratios[color_name] = np.sum(mask > 0) / total_pixels * 100
            
            # 일정 비율 이상이면 강조색이 있다고 판단
            highlight_threshold = thresholds["base_threshold"] * 100 # 백분율로 변환
            
            result = {
                "has_highlight": np.sum(highlight_mask > 0) / total_pixels * 100 > highlight_threshold,
                "has_gray": color_ratios["gray"] > thresholds["gray_threshold"] * 100,
                "colors": {color: ratio > highlight_threshold for color, ratio in color_ratios.items()},
                "color_ratios": color_ratios
            }
            
            # 텍스트 위치 정보 확인 (취소선)
            has_strikethrough = False
            text_blocks = page.get_text("dict")["blocks"]
            for block in text_blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            if span.get("flags", 0) & 2**6:  # 취소선 비트 확인 (64)
                                has_strikethrough = True
                                break
                        if has_strikethrough:
                            break
                    if has_strikethrough:
                        break
            
            result["has_strikethrough"] = has_strikethrough
            
            return result
            
        except Exception as e:
            print(f"강조색 감지 중 오류: {str(e)}")
            return {
                "has_highlight": False,
                "has_gray": False,
                "has_strikethrough": False,
                "colors": {},
                "color_ratios": {}
            }

# 개선된 테이블 추출 함수
def extract_tables_with_camelot(pdf_path, page_num):
    """Camelot을 사용하여 테이블 추출 - 최적화된 옵션"""
    try:
        # Camelot 테이블 추출 - lattice 모드 개선
        tables = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num + 1),
            flavor='lattice',
            line_scale=40,            
            line_tol=2,               
            strip_text='\n'           
        )
        
        if len(tables) == 0:
            # lattice 모드에서 테이블을 찾지 못한 경우 stream 모드 시도
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_num + 1),
                flavor='stream',
                edge_tol=50,          
                row_tol=10            
            )
        
        if len(tables) == 0:
            return []
            
        results = []
        for i in range(len(tables)):
            table = tables[i]
            
            # 빈 행/열 제거
            df = table.df.copy()
            
            # 문자열인 경우에만 strip 적용
            for col in df.columns:
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            
            df = df.replace('', np.nan)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            if not df.empty and df.size > 0:
                # 셀 정보 저장
                cell_data = {}
                
                # 병합된 셀 정보 저장 - 속성이 존재하는지 확인
                if hasattr(table, 'spanning_cells'):
                    for spanning_cell in table.spanning_cells:
                        try:
                            start_row, start_col, end_row, end_col = spanning_cell
                            for r in range(start_row, end_row + 1):
                                for c in range(start_col, end_col + 1):
                                    cell_data[(r, c)] = {
                                        "is_spanning": True,
                                        "spanning_coords": spanning_cell
                                    }
                        except ValueError:
                            continue
                
                # 모든 셀 좌표 정보 저장 - 안전하게 처리
                if hasattr(table, 'cells'):
                    for cell in table.cells:
                        try:
                            # 셀이 리스트 형태인 경우 특별 처리
                            if isinstance(cell, list):
                                for idx, subcell in enumerate(cell):
                                    if hasattr(subcell, 'x1') and hasattr(subcell, 'y1'):
                                        # 리스트 내 Cell 객체의 좌표 정보 추출
                                        row_idx = idx // df.shape[1] if df.shape[1] > 0 else 0
                                        col_idx = idx % df.shape[1] if df.shape[1] > 0 else 0
                                        
                                        if (row_idx, col_idx) not in cell_data:
                                            cell_data[(row_idx, col_idx)] = {}
                                        
                                        cell_data[(row_idx, col_idx)].update({
                                            "bbox": [
                                                float(subcell.x1),
                                                float(subcell.y1),
                                                float(subcell.x2),
                                                float(subcell.y2)
                                            ]
                                        })
                            else:
                                # 기존 코드 유지 (4개 값 언패킹)
                                r1, c1, r2, c2 = cell
                                
                                if (r1, c1) not in cell_data:
                                    cell_data[(r1, c1)] = {}
                                
                                cell_data[(r1, c1)].update({
                                    "bbox": [
                                        float(table.cells[r1][c1][0]),
                                        float(table.cells[r1][c1][1]),
                                        float(table.cells[r2-1][c2-1][2]),
                                        float(table.cells[r2-1][c2-1][3])
                                    ]
                                })
                        except ValueError:
                            # 4개 값 언패킹 실패 시 경고만 표시하고 계속 진행
                            st.warning(f"셀 정보 처리 중 오류: 예상한 4개 값을 받지 못했습니다 - {cell}")
                            continue
                
                results.append({
                    'page': page_num + 1,
                    'table_index': i,
                    'df': df,
                    'accuracy': table.parsing_report.get('accuracy', 0),
                    'cells': cell_data, 
                    'table': table,
                    'coords': table._bbox
                })
        
        return results
    except Exception as e:
        st.warning(f"페이지 {page_num+1}의 테이블 추출 중 오류 발생: {str(e)}")
        return []

# 테이블 셀에 강조색 매핑하는 개선된 함수
def map_highlights_to_tables(pdf_path, tables, page_num):
    """테이블 셀에 형광색/취소선 영역 매핑 - 수정된 버전"""
    try:
        # 문서 열기
        pdf_document = fitz.open(pdf_path)
        page = pdf_document[page_num]
        
        # 테이블이 없으면 빈 결과 반환
        if not tables:
            pdf_document.close()
            return []
        
        # 하이라이트 분석기 초기화
        analyzer = HighlightAnalyzer()
        
        # 강조색 감지 - 높은 해상도로 렌더링
        zoom = 3.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        hsv = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2HSV)
        
        # 적응형 임계값 계산
        thresholds = analyzer.analyze_document_colors(hsv)
        
        # 색상 마스크 생성 - 노란색 감지 범위 확장
        masks = {}
        # 노란색 감지 범위 확장
        analyzer.color_ranges["yellow"]["lower"] = np.array([15, 50, 180])
        analyzer.color_ranges["yellow"]["upper"] = np.array([45, 255, 255])
        
        for color_name, ranges in analyzer.color_ranges.items():
            mask = cv2.inRange(hsv, ranges["lower"], ranges["upper"])
            kernel = np.ones((5, 5), np.uint8)
            mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
            mask = cv2.dilate(mask, kernel, iterations=1)
            masks[color_name] = mask
        
        # 모든 강조색 마스크 통합
        highlight_mask = masks["yellow"] | masks["light_yellow"] | masks["light_blue"] | masks["orange"] | masks["green"] | masks["purple"]
        
        # 테이블의 셀 정보와 강조색 매핑
        for table_info in tables:
            df = table_info['df']
            rows, cols = df.shape
            
            # 셀별 강조색과 취소선 정보 저장
            highlight_cells = []
            gray_cells = []
            missing_text_cells = []
            
            # 각 셀의 내용과 좌표 정보 확인
            for r in range(rows):
                for c in range(cols):
                    cell_content = df.iloc[r, c] if r < rows and c < cols else ""
                    cell_has_highlight = False
                    
                    # 셀 영역 좌표 (안전하게 추출)
                    cell_coords = None
                    
                    # 방법 1: cells 딕셔너리에서 추출 (가장 안전한 방법)
                    if 'cells' in table_info:
                        cell_key = (r, c)
                        if cell_key in table_info['cells'] and 'bbox' in table_info['cells'][cell_key]:
                            cell_coords = table_info['cells'][cell_key]['bbox']
                    
                    # 방법 2: 테이블 좌표에서 계산 (cells 정보가 없는 경우)
                    if cell_coords is None and 'coords' in table_info:
                        x0, y0, x1, y1 = table_info['coords']
                        cell_width = (x1 - x0) / cols if cols > 0 else 0
                        cell_height = (y1 - y0) / rows if rows > 0 else 0
                        cell_coords = [
                            x0 + c * cell_width,
                            y0 + r * cell_height,
                            x0 + (c + 1) * cell_width,
                            y0 + (r + 1) * cell_height
                        ]
                    
                    # 방법 3: table 객체에서 직접 추출 시도 (조심해서 사용)
                    if cell_coords is None and 'table' in table_info:
                        try:
                            cells_attr = getattr(table_info['table'], 'cells', None)
                            if cells_attr and isinstance(cells_attr, list) and r < len(cells_attr):
                                if isinstance(cells_attr[r], list) and c < len(cells_attr[r]):
                                    cell_obj = cells_attr[r][c]
                                    # 객체 타입에 따라 다르게 처리
                                    if hasattr(cell_obj, 'x1') and hasattr(cell_obj, 'y1'):
                                        # Cell 객체인 경우 속성으로 접근
                                        cell_coords = [
                                            float(cell_obj.x1),
                                            float(cell_obj.y1),
                                            float(cell_obj.x2),
                                            float(cell_obj.y2)
                                        ]
                                    elif isinstance(cell_obj, (list, tuple)) and len(cell_obj) >= 4:
                                        # 좌표 리스트/튜플인 경우
                                        cell_coords = [float(v) for v in cell_obj[:4]]
                        except (IndexError, AttributeError, TypeError) as e:
                            # 오류 발생 시 무시하고 계속 진행
                            continue
                    
                    # 셀 좌표가 있으면 강조색 확인
                    if cell_coords:
                        # 픽셀 좌표로 변환
                        pixel_x0 = int(cell_coords[0] * zoom)
                        pixel_y0 = int(cell_coords[1] * zoom)
                        pixel_x1 = int(cell_coords[2] * zoom)
                        pixel_y1 = int(cell_coords[3] * zoom)
                        
                        # 범위 제한
                        pixel_x0 = max(0, min(pixel_x0, hsv.shape[1]-1))
                        pixel_y0 = max(0, min(pixel_y0, hsv.shape[0]-1))
                        pixel_x1 = max(0, min(pixel_x1, hsv.shape[1]-1))
                        pixel_y1 = max(0, min(pixel_y1, hsv.shape[0]-1))
                        
                        # 셀 영역 내 강조색 비율 계산
                        if pixel_y1 > pixel_y0 and pixel_x1 > pixel_x0:  # 유효한 영역인지 확인
                            cell_region = highlight_mask[pixel_y0:pixel_y1, pixel_x0:pixel_x1]
                            if cell_region.size > 0:
                                highlight_ratio = np.sum(cell_region > 0) / cell_region.size
                                
                                # 더 낮은 임계값 적용 (3%)
                                if highlight_ratio > 0.03:
                                    highlight_cells.append((r, c))
                                    cell_has_highlight = True
                                    
                                    # 내용이 비어있고 강조색이 있으면 텍스트 복원 시도
                                    if pd.isna(cell_content) or str(cell_content).strip() == '':
                                        # PDF 원본에서 직접 텍스트 추출
                                        clip_rect = fitz.Rect(cell_coords[0], cell_coords[1], cell_coords[2], cell_coords[3])
                                        cell_text = page.get_text("text", clip=clip_rect).strip()
                                        if cell_text:
                                            df.iloc[r, c] = cell_text
                                            missing_text_cells.append((r, c))
            
            # 테이블 정보에 강조 셀 정보와 복원된 텍스트 정보 추가
            table_info['highlight_cells'] = highlight_cells
            table_info['gray_cells'] = gray_cells
            table_info['missing_text_cells'] = missing_text_cells
            table_info['df'] = df  # 업데이트된 데이터프레임 저장
        
        pdf_document.close()
        return tables
        
    except Exception as e:
        st.warning(f"페이지 {page_num+1}의 강조색 매핑 중 오류 발생: {str(e)}")
        # 오류 발생해도, 테이블은 반환하여 계속 진행할 수 있도록 함
        return tables

# 확장된 테이블 처리 함수 (보장내용 내려받기 기능용)
def process_tables_for_export(pdf_path, page_range):
    """지정된 페이지 범위에서 테이블을 추출하고 처리"""
    all_tables = []
    
    with st.spinner(f"페이지 {page_range[0]+1}~{page_range[1]+1} 테이블 추출 중..."):
        for page_num in range(page_range[0], page_range[1] + 1):
            try:
                # Camelot으로 테이블 추출
                tables = extract_tables_with_camelot(pdf_path, page_num)
                
                if tables and len(tables) > 0:
                    # 강조색 및 취소선 정보 매핑
                    tables = map_highlights_to_tables(pdf_path, tables, page_num)
                    
                    all_tables.extend(tables)
                    st.info(f"페이지 {page_num+1}에서 {len(tables)}개 테이블 발견")
            except Exception as e:
                st.warning(f"페이지 {page_num+1} 처리 중 오류: {str(e)}")
    
    return all_tables

# 업무정의서 형식으로 엑셀 생성 (개선된 버전)
def create_business_definition_excel(all_tables, pdf_filename):
    """테이블 데이터로 업무정의서 형식의 엑셀 생성 - 강조색 처리 개선"""
    if not all_tables:
        return None
    
    # 파일명에서 상품명 추출
    product_name = os.path.splitext(os.path.basename(pdf_filename))[0]
    
    # Workbook 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "보장내용 개정사항"
    
    # 스타일 정의
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # 노란 형광색용
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")    # 회색용
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
    for table_info in all_tables:
        df = table_info['df']
        page_num = table_info['page']
        
        # 강조색 및 복원된 텍스트 정보
        highlight_cells = table_info.get('highlight_cells', [])
        missing_text_cells = table_info.get('missing_text_cells', [])
        gray_cells = table_info.get('gray_cells', [])
        
        # 데이터프레임 행, 열 수 확인
        rows, cols = df.shape
        
        # 모든 행 포함 (모든 행을 이제 포함하여 형광펜만 있는 행도 표시)
        for df_row_idx in range(rows):
            # 해당 행이 강조 처리된 행인지 확인
            row_has_highlight = any((df_row_idx, c) in highlight_cells for c in range(cols))
            
            # 셀 값 준비
            row_values = [page_num]  # 첫 번째 열은 페이지 번호
            
            # 데이터프레임의 해당 행 값 추가 (컬럼 최대 4개까지만)
            for df_col_idx in range(min(cols, 4)):
                cell_value = df.iloc[df_row_idx, df_col_idx]
                if pd.isna(cell_value):
                    cell_value = ""
                row_values.append(str(cell_value))
            
            # 부족한 열은 빈 문자열로 채움
            while len(row_values) < 5:
                row_values.append("")
            
            # 강조여부 열 추가
            highlight_status = []
            if row_has_highlight:
                highlight_status.append("강조")
            if any((df_row_idx, c) in gray_cells for c in range(cols)):
                highlight_status.append("취소선")
            if any((df_row_idx, c) in missing_text_cells for c in range(cols)):
                highlight_status.append("형광펜 텍스트 복원")
            
            row_values.append(", ".join(highlight_status) if highlight_status else "")
            
            # 엑셀에 행 추가
            for col_idx, value in enumerate(row_values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                
                # 강조색 적용 (열이 4개 이상인 경우에도 정확히 반영)
                df_col = col_idx - 2 if col_idx >= 2 and col_idx <= min(cols+1, 5) else -1
                
                # 데이터 셀에만 형식 적용
                if 2 <= col_idx <= 5 and df_col >= 0:
                    # 강조색 적용
                    if (df_row_idx, df_col) in highlight_cells:
                        cell.fill = yellow_fill
                    # 회색 적용 (취소선 있는 경우)
                    elif (df_row_idx, df_col) in gray_cells:
                        cell.fill = gray_fill
                        # 취소선 텍스트 스타일 적용
                        cell.font = Font(strike=True)
            
            row_idx += 1
    
    # 열 너비 조정
    ws.column_dimensions['A'].width = 10  # 페이지
    ws.column_dimensions['B'].width = 25  # 보장명
    ws.column_dimensions['C'].width = 40  # 지급사유
    ws.column_dimensions['D'].width = 25  # 지급금액
    ws.column_dimensions['E'].width = 15  # 비고
    ws.column_dimensions['F'].width = 15  # 강조여부
    
    # 엑셀 파일 저장 (메모리에)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

# 결과 저장 함수
def save_analysis_results(analysis_results, output_dir="/workspaces/automation/analysis_results"):
    """분석 결과를 파일로 저장"""
    # 출력 디렉토리 확인/생성
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
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

# 결과 불러오기 함수
def load_analysis_results(timestamp=None, output_dir="/workspaces/automation/analysis_results"):
    """저장된 분석 결과 불러오기"""
    if not os.path.exists(output_dir):
        return None, []
    
    # 요약 파일 목록 가져오기
    summary_files = glob.glob(os.path.join(output_dir, "summary_*.json"))
    
    if not summary_files:
        return None, []
    
    # 타임스탬프로 필터링
    if timestamp:
        summary_file = os.path.join(output_dir, f"summary_{timestamp}.json")
        if not os.path.exists(summary_file):
            return None, []
        summary_files = [summary_file]
    
    # 가장 최근 파일 선택 (타임스탬프 기준 정렬)
    summary_files.sort(reverse=True)
    latest_summary = summary_files[0]
    
    # 타임스탬프 추출
    timestamp_match = re.search(r'summary_(\d+_\d+).json', latest_summary)
    if timestamp_match:
        timestamp = timestamp_match.group(1)
    else:
        return None, []
    
    # 요약 정보 로드
    with open(latest_summary, 'r', encoding='utf-8') as f:
        summary_data = json.load(f)
    
    # 각 파일별 상세 결과 로드
    analysis_results = []
    for summary in summary_data:
        filename = summary["파일명"]
        base_name = os.path.splitext(filename)[0]
        safe_name = re.sub(r'[^\w\s-]', '', base_name)
        
        pickle_path = os.path.join(output_dir, f"{safe_name}_{timestamp}.pkl")
        
        if os.path.exists(pickle_path):
            with open(pickle_path, 'rb') as f:
                result = pickle.load(f)
                analysis_results.append(result)
    
    return timestamp, analysis_results

# 로그 기록 함수
def log_results(analysis_results):
    """테스트 결과를 로그 파일에 기록하는 함수"""
    # 로그 폴더 확인/생성
    log_folder = "/workspaces/automation/test_log"
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
        st.info(f"로그 폴더를 생성했습니다: {log_folder}")
    
    # 로그 파일명 설정 (날짜별 파일)
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
    
    st.success(f"테스트 결과가 로그 파일에 저장되었습니다: {log_filename}")

# 결과 표시 함수
def display_analysis_results(analysis_results):
    """분석 결과를 화면에 표시하는 함수"""
    if not analysis_results:
        st.info("표시할 분석 결과가 없습니다.")
        return
    
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
        
        # 강조색 있는 페이지 처리 (set 타입일 수 있음)
        if isinstance(result["강조색 있는 페이지"], set):
            highlight_pages = sorted(list(result["강조색 있는 페이지"]))
        else:
            highlight_pages = result["강조색 있는 페이지"]
            
        # 취소선 있는 페이지 처리 (set 타입일 수 있음)
        if isinstance(result["취소선 있는 페이지"], set):
            strikethrough_pages = sorted(list(result["취소선 있는 페이지"]))
        else:
            strikethrough_pages = result["취소선 있는 페이지"]
        
        summary_data.append({
            "파일명": result["파일명"],
            "나. 보험금 페이지": ', '.join(map(str, result["나. 보험금 페이지"])) if result["나. 보험금 페이지"] else "없음",
            "종별 범위": type_ranges_str,  # 종별 범위 정보 추가
            "종별 모든 발견": all_types_str,  # 모든 종별 발견 정보
            "상해관련특별약관 페이지": ', '.join(map(str, result["상해관련특별약관 페이지"])) if result["상해관련특별약관 페이지"] else "없음",
            "질병관련특별약관 페이지": ', '.join(map(str, result["질병관련특별약관 페이지"])) if result["질병관련특별약관 페이지"] else "없음",
            "상해및질병관련특별약관 종료": str(result["상해및질병관련특별약관 종료 페이지"]) if result["상해및질병관련특별약관 종료 페이지"] is not None else "없음",
            "강조색 있는 페이지": ', '.join(map(str, highlight_pages)) if highlight_pages else "없음",
            "취소선 있는 페이지": ', '.join(map(str, strikethrough_pages)) if strikethrough_pages else "없음",
            "처리 상태": result["처리 상태"]
        })
    
    # 데이터프레임 생성 및 표시
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True)
    
    # 상세 로그 표시
    st.subheader("상세 로그")
    for result in analysis_results:
        with st.expander(f"파일: {result['파일명']} 상세 로그"):
            for log_entry in result.get("상세 로그", []):
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

# 페이지 설정
st.set_page_config(page_title="PDF 보장내용 테스트", layout="wide")
st.title("PDF 보장내용 테스트")

# 메인 탭
tab1, tab2 = st.tabs(["새 분석 실행", "이전 분석 결과"])

with tab1:
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
                            
                            # 개선된 하이라이트 분석기 초기화
                            highlight_analyzer = HighlightAnalyzer()

                            for page_num in range(search_start_page, search_end_page + 1):
                                # 개선된 강조색 감지 메서드 사용
                                try:
                                    highlight_result = highlight_analyzer.detect_highlights(pdf_document, page_num)
                                    
                                    if highlight_result["has_highlight"]:
                                        file_result["강조색 있는 페이지"].add(page_num + 1)
                                        
                                        # 색상별 정보 로깅 (선택 사항)
                                        detected_colors = []
                                        for color_name, is_detected in highlight_result["colors"].items():
                                            if is_detected and color_name not in ["gray"]:
                                                detected_colors.append(color_name)
                                        
                                        if detected_colors:
                                            color_info = f"페이지 {page_num + 1} 감지된 색상: {', '.join(detected_colors)}"
                                            file_result["상세 로그"].append(color_info)
                                    
                                    # 취소선 확인 
                                    if highlight_result["has_strikethrough"] or highlight_result["has_gray"]:
                                        file_result["취소선 있는 페이지"].add(page_num + 1)
                                    
                                except Exception as e:
                                    error_msg = f"페이지 {page_num + 1} 강조색 분석 중 오류: {str(e)}"
                                    file_result["상세 로그"].append(error_msg)
                                    st.warning(error_msg)
                            
                            # 집합(set)을 정렬된 리스트로 변환
                            file_result["강조색 있는 페이지"] = sorted(list(file_result["강조색 있는 페이지"]))
                            file_result["취소선 있는 페이지"] = sorted(list(file_result["취소선 있는 페이지"]))
                            
                            if file_result["강조색 있는 페이지"]:
                                st.write(f"강조색 있는 페이지: {', '.join(map(str, file_result['강조색 있는 페이지']))}")
                                file_result["상세 로그"].append(f"강조색 있는 페이지: {', '.join(map(str, file_result['강조색 있는 페이지']))}")
                            else:
                                st.write("강조색이 발견되지 않았습니다. 전체 파싱 범위를 처리합니다.")
                                file_result["상세 로그"].append("강조색이 발견되지 않았습니다. 전체 파싱 범위를 처리합니다.")
                            
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

                # 자동 테이블 추출 및 Excel 변환
                st.subheader("테이블 자동 추출 결과")

                # 테이블 저장 폴더 생성
                tables_folder = os.path.join(input_folder, "extracted_tables")
                if not os.path.exists(tables_folder):
                    os.makedirs(tables_folder)
                    st.info(f"테이블 저장 폴더를 생성했습니다: {tables_folder}")

                # 각 파일별로 테이블 추출 및 엑셀 생성
                table_results = []
                for idx, result in enumerate(analysis_results):
                    file_name = result["파일명"]
                    file_path = selected_files[idx]
                    
                    # 진행 표시
                    status_placeholder = st.empty()
                    status_placeholder.info(f"{file_name} 테이블 추출 중...")
                    
                    # 파싱 범위 설정
                    if result["나. 보험금 페이지"]:
                        parsing_start = result["나. 보험금 페이지"][0] - 1
                        
                        # 종료 페이지 설정
                        if result["상해및질병관련특별약관 종료 페이지"]:
                            parsing_end = result["상해및질병관련특별약관 종료 페이지"] - 1
                        else:
                            with fitz.open(file_path) as doc:
                                parsing_end = len(doc) - 1
                    else:
                        # 파싱 범위가 없는 경우 전체 문서 처리
                        with fitz.open(file_path) as doc:
                            parsing_start = 0
                            parsing_end = len(doc) - 1
                    
                    # 테이블 추출
                    try:
                        all_tables = process_tables_for_export(file_path, (parsing_start, parsing_end))
                        
                        if all_tables:
                            # 안전한 파일명 생성
                            base_name = os.path.splitext(file_name)[0]
                            safe_name = re.sub(r'[^\w\s-]', '', base_name)
                            excel_filename = f"{safe_name}_테이블_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            excel_path = os.path.join(tables_folder, excel_filename)
                            
                            # 엑셀 파일 생성
                            excel_output = create_business_definition_excel(all_tables, file_name)
                            
                            if excel_output:
                                # 파일로 저장
                                with open(excel_path, "wb") as f:
                                    f.write(excel_output.getvalue())
                                
                                table_results.append({
                                    "file_name": file_name,
                                    "table_count": len(all_tables),
                                    "excel_path": excel_path,
                                    "excel_filename": excel_filename,
                                    "success": True
                                })
                                status_placeholder.success(f"{file_name}: {len(all_tables)}개 테이블 추출 완료")
                            else:
                                table_results.append({
                                    "file_name": file_name,
                                    "success": False,
                                    "error": "엑셀 생성 실패"
                                })
                                status_placeholder.error(f"{file_name}: 엑셀 생성 실패")
                        else:
                            table_results.append({
                                "file_name": file_name,
                                "success": False,
                                "error": "테이블을 찾을 수 없음"
                            })
                            status_placeholder.warning(f"{file_name}에서 테이블을 찾을 수 없습니다.")
                    except Exception as e:
                        table_results.append({
                            "file_name": file_name,
                            "success": False,
                            "error": str(e)
                        })
                        status_placeholder.error(f"{file_name} 처리 중 오류: {str(e)}")
                
                # 테이블 추출 요약
                st.subheader("테이블 추출 요약")
                if table_results:
                    # 테이블 결과 표시
                    table_summary = []
                    for res in table_results:
                        table_summary.append({
                            "파일명": res["file_name"],
                            "테이블 수": res.get("table_count", 0) if res["success"] else 0,
                            "상태": "성공" if res["success"] else "실패",
                            "오류": res.get("error", "") if not res["success"] else ""
                        })
                    
                    # 결과를 데이터프레임으로 표시
                    summary_df = pd.DataFrame(table_summary)
                    st.dataframe(summary_df)
                    
                    # 성공한 경우 엑셀 파일 다운로드 링크 제공
                    for res in table_results:
                        if res["success"] and "excel_path" in res:
                            with open(res["excel_path"], "rb") as f:
                                excel_data = f.read()
                            st.download_button(
                                label=f"{res['file_name']} 테이블 다운로드",
                                data=excel_data,
                                file_name=res["excel_filename"],
                                mime="application/vnd.ms-excel"
                            )
                else:
                    st.info("추출된 테이블이 없습니다.")