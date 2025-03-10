import fitz  # PyMuPDF
import camelot
import pandas as pd
import numpy as np
import cv2
import re
import os
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

class TableHighlightAnalyzer:
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
            "gray": {
                "lower": np.array([0, 0, 80]),
                "upper": np.array([180, 40, 200])
            }
        }
        
    def extract_tables(self, pdf_path, page_num):
        """개선된 테이블 추출 함수"""
        try:
            # 1. Camelot으로 테이블 추출
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num + 1),
                flavor='lattice',
                line_scale=40,
                process_background=True
            )
            
            if len(tables) == 0:
                # 격자 구조가 없는 경우 stream 방식 시도
                tables = camelot.read_pdf(
                    pdf_path, 
                    pages=str(page_num + 1),
                    flavor='stream',
                    edge_tol=50
                )
            
            # 2. 추출 결과 처리
            processed_tables = []
            for i, table in enumerate(tables):
                if table.df.empty:
                    continue
                
                # 셀 정보 저장
                cell_data = {}
                # 병합된 셀 정보 저장
                for spanning_cell in table.spanning_cells:
                    start_row, start_col, end_row, end_col = spanning_cell
                    for r in range(start_row, end_row + 1):
                        for c in range(start_col, end_col + 1):
                            cell_data[(r, c)] = {
                                "is_spanning": True,
                                "spanning_coords": spanning_cell
                            }
                
                # 모든 셀 좌표 정보 저장
                for cell in table.cells:
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
                
                # 데이터프레임 정제
                df = self.clean_dataframe(table.df)
                
                processed_tables.append({
                    'page': page_num + 1,
                    'table_index': i,
                    'df': df,
                    'accuracy': table.parsing_report.get('accuracy', 0),
                    'cells': cell_data,
                    'table': table,
                    'coords': table._bbox
                })
            
            return processed_tables
            
        except Exception as e:
            print(f"테이블 추출 중 오류: {str(e)}")
            return []
    
    def clean_dataframe(self, df):
        """데이터프레임 정제 함수"""
        # 빈 행/열 제거
        df = df.replace('', np.nan).dropna(how='all').dropna(axis=1, how='all')
        
        # 열 이름 정리
        df.columns = [str(col).strip() for col in df.columns]
        
        # 내용 정제
        for col in df.columns:
            df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            df[col] = df[col].replace(r'\s+', ' ', regex=True)
        
        return df
    
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
    
    def map_highlights_to_tables(self, pdf_path, tables, page_num):
        """개선된 강조색 매핑 함수"""
        try:
            # 1. PDF 문서 열기
            pdf_document = fitz.open(pdf_path)
            page = pdf_document[page_num]
            
            # 테이블이 없으면 빈 결과 반환
            if not tables:
                pdf_document.close()
                return []
            
            # 2. 페이지 이미지 렌더링 (고해상도)
            zoom = 2.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            hsv = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2HSV)
            
            # 3. 적응형 임계값 계산
            thresholds = self.analyze_document_colors(hsv)
            
            # 4. 색상 마스크 생성
            masks = {}
            for color_name, ranges in self.color_ranges.items():
                mask = cv2.inRange(hsv, ranges["lower"], ranges["upper"])
                kernel = np.ones((3, 3), np.uint8)
                mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
                masks[color_name] = mask
            
            # 모든 강조색 마스크 통합
            highlight_mask = masks["yellow"] | masks["light_yellow"] | masks["light_blue"]
            
            # 5. 텍스트 위치 정보 추출
            text_blocks = page.get_text("dict")["blocks"]
            text_positions = []
            
            for block in text_blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            # 스팬 좌표 (x0, y0, x1, y1)
                            bbox = span["bbox"]
                            
                            # 취소선 확인
                            has_strikethrough = bool(span.get("flags", 0) & 2**6)
                            
                            text_positions.append({
                                "text": span["text"],
                                "bbox": bbox,
                                "has_strikethrough": has_strikethrough
                            })
            
            # 6. 각 테이블 처리
            for table_info in tables:
                df = table_info['df']
                rows, cols = df.shape
                
                # 셀별 강조색과 취소선 정보 저장
                highlight_cells = []  # 강조색이 있는 셀
                gray_cells = []       # 취소선(회색)이 있는 셀
                
                # 셀 좌표 정보가 있는 경우만 처리
                if 'cells' in table_info:
                    # 7. 혼합 접근법: 이미지 기반 + 텍스트 위치 기반
                    for r in range(rows):
                        for c in range(cols):
                            # 셀 내용
                            cell_content = df.iloc[r, c] if r < rows and c < cols else ""
                            
                            # 셀 좌표 정보 확인
                            cell_key = (r, c)
                            if cell_key in table_info['cells'] and 'bbox' in table_info['cells'][cell_key]:
                                cell_bbox = table_info['cells'][cell_key]['bbox']
                                
                                # 이미지 기반 분석 (픽셀 좌표로 변환)
                                pixel_x0 = int(cell_bbox[0] * zoom)
                                pixel_y0 = int(cell_bbox[1] * zoom)
                                pixel_x1 = int(cell_bbox[2] * zoom)
                                pixel_y1 = int(cell_bbox[3] * zoom)
                                
                                # 범위 제한
                                pixel_x0 = max(0, min(pixel_x0, hsv.shape[1]-1))
                                pixel_y0 = max(0, min(pixel_y0, hsv.shape[0]-1))
                                pixel_x1 = max(0, min(pixel_x1, hsv.shape[1]-1))
                                pixel_y1 = max(0, min(pixel_y1, hsv.shape[0]-1))
                                
                                # 셀 영역 내 강조색 픽셀 비율
                                cell_region_highlight = highlight_mask[pixel_y0:pixel_y1, pixel_x0:pixel_x1]
                                highlight_ratio = np.sum(cell_region_highlight > 0) / cell_region_highlight.size if cell_region_highlight.size > 0 else 0
                                
                                # 셀 영역 내 회색 픽셀 비율
                                cell_region_gray = masks["gray"][pixel_y0:pixel_y1, pixel_x0:pixel_x1]
                                gray_ratio = np.sum(cell_region_gray > 0) / cell_region_gray.size if cell_region_gray.size > 0 else 0
                                
                                # 텍스트 위치 기반 분석
                                cell_has_strikethrough = False
                                cell_text_highlighted = False
                                
                                # 해당 셀 영역과 겹치는 텍스트 위치 확인
                                for pos in text_positions:
                                    text_bbox = pos["bbox"]
                                    
                                    # 셀과 텍스트 영역이 겹치는지 확인
                                    if (cell_bbox[0] <= text_bbox[2] and cell_bbox[2] >= text_bbox[0] and
                                        cell_bbox[1] <= text_bbox[3] and cell_bbox[3] >= text_bbox[1]):
                                        
                                        # 취소선 확인
                                        if pos["has_strikethrough"]:
                                            cell_has_strikethrough = True
                                        
                                        # 텍스트 영역과 강조색 겹침 확인
                                        text_pixel_x0 = int(text_bbox[0] * zoom)
                                        text_pixel_y0 = int(text_bbox[1] * zoom)
                                        text_pixel_x1 = int(text_bbox[2] * zoom)
                                        text_pixel_y1 = int(text_bbox[3] * zoom)
                                        
                                        # 범위 제한
                                        text_pixel_x0 = max(0, min(text_pixel_x0, hsv.shape[1]-1))
                                        text_pixel_y0 = max(0, min(text_pixel_y0, hsv.shape[0]-1))
                                        text_pixel_x1 = max(0, min(text_pixel_x1, hsv.shape[1]-1))
                                        text_pixel_y1 = max(0, min(text_pixel_y1, hsv.shape[0]-1))
                                        
                                        # 텍스트 영역 내 강조색 확인
                                        text_region = highlight_mask[text_pixel_y0:text_pixel_y1, text_pixel_x0:text_pixel_x1]
                                        if text_region.size > 0 and np.sum(text_region > 0) / text_region.size > thresholds["yellow_threshold"]:
                                            cell_text_highlighted = True
                                
                                # 하이브리드 분석 결과 통합
                                # 텍스트 기반과 이미지 기반 결과 중 더 신뢰할 수 있는 쪽 선택
                                is_highlighted = highlight_ratio > thresholds["base_threshold"] or cell_text_highlighted
                                is_gray = (gray_ratio > thresholds["gray_threshold"] and cell_has_strikethrough)
                                
                                if is_highlighted:
                                    highlight_cells.append((r, c))
                                if is_gray:
                                    gray_cells.append((r, c))
                
                # 테이블 정보에 강조 정보 추가
                table_info['highlight_cells'] = highlight_cells
                table_info['gray_cells'] = gray_cells
            
            pdf_document.close()
            return tables
        
        except Exception as e:
            print(f"강조색 매핑 중 오류: {str(e)}")
            return tables
    
    def create_excel_with_highlights(self, tables, output_path, product_name=""):
        """강조 표시가 정확히 반영된 엑셀 생성"""
        if not tables:
            return None
        
        # Workbook 생성
        wb = Workbook()
        ws = wb.active
        ws.title = "보장내용 개정사항"
        
        # 스타일 정의
        header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid") 
        gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 제목 행 추가
        ws.merge_cells('A1:G1')
        cell = ws.cell(row=1, column=1, value=f"{product_name} - 보장내용 개정사항")
        cell.font = Font(size=14, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 헤더 행 추가
        headers = ["페이지", "테이블", "행", "열", "내용", "강조여부", "비고"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # 테이블 데이터 추가
        row_idx = 4
        
        # 테이블별 처리
        for table_info in tables:
            df = table_info['df']
            page_num = table_info['page']
            table_idx = table_info['table_index']
            
            # 강조색 및 취소선 정보
            highlight_cells = table_info.get('highlight_cells', [])
            gray_cells = table_info.get('gray_cells', [])
            
            # 병합된 셀 정보 처리
            merged_cells = []
            if 'table' in table_info and hasattr(table_info['table'], 'spanning_cells'):
                merged_cells = table_info['table'].spanning_cells
            
            # 강조된 셀만 엑셀에 추가
            all_marked_cells = set(highlight_cells + gray_cells)
            
            if all_marked_cells:
                rows, cols = df.shape
                
                for r, c in sorted(all_marked_cells):
                    if r < rows and c < cols:
                        # 셀 값 가져오기
                        cell_value = df.iloc[r, c]
                        
                        # 강조 상태 확인
                        is_highlighted = (r, c) in highlight_cells
                        is_gray = (r, c) in gray_cells
                        
                        highlight_status = []
                        if is_highlighted:
                            highlight_status.append("강조")
                        if is_gray:
                            highlight_status.append("취소선")
                        
                        # 엑셀에 행 추가
                        ws.cell(row=row_idx, column=1, value=page_num).border = border  # 페이지
                        ws.cell(row=row_idx, column=2, value=table_idx + 1).border = border  # 테이블
                        ws.cell(row=row_idx, column=3, value=r + 1).border = border  # 행
                        ws.cell(row=row_idx, column=4, value=c + 1).border = border  # 열
                        
                        # 내용 셀
                        content_cell = ws.cell(row=row_idx, column=5, value=str(cell_value))
                        content_cell.border = border
                        
                        # 강조 여부
                        status_cell = ws.cell(row=row_idx, column=6, value=", ".join(highlight_status))
                        status_cell.border = border
                        
                        # 강조색 적용
                        if is_highlighted:
                            content_cell.fill = yellow_fill
                        elif is_gray:
                            content_cell.fill = gray_fill
                            content_cell.font = Font(strike=True)  # 취소선 적용
                        
                        row_idx += 1
        
        # 열 너비 조정
        ws.column_dimensions['A'].width = 10  # 페이지
        ws.column_dimensions['B'].width = 10  # 테이블
        ws.column_dimensions['C'].width = 8   # 행
        ws.column_dimensions['D'].width = 8   # 열
        ws.column_dimensions['E'].width = 40  # 내용
        ws.column_dimensions['F'].width = 15  # 강조여부
        ws.column_dimensions['G'].width = 15  # 비고
        
        # 엑셀 파일 저장
        wb.save(output_path)
        
        return output_path

def test_highlight_mapping():
    """테스트 실행 함수"""
    # 경로 설정
    pdf_path = "test_document.pdf"  # 테스트할 PDF 경로
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    # 분석기 초기화
    analyzer = TableHighlightAnalyzer()
    
    # 테스트할 페이지 지정
    test_page = 5  # 강조색이 있는 페이지 번호 (0부터 시작)
    
    print(f"테스트 시작: {pdf_path} 페이지 {test_page+1}")
    
    # 1. 테이블 추출
    tables = analyzer.extract_tables(pdf_path, test_page)
    if not tables:
        print("테이블을 추출할 수 없습니다.")
        return
    
    print(f"{len(tables)}개 테이블 추출 완료")
    
    # 2. 강조색 매핑
    tables_with_highlights = analyzer.map_highlights_to_tables(pdf_path, tables, test_page)
    
    # 3. 강조색 정보 출력
    for i, table_info in enumerate(tables_with_highlights):
        print(f"\n테이블 {i+1}:")
        print(f"- 강조된 셀: {len(table_info.get('highlight_cells', []))}개")
        print(f"- 취소선 셀: {len(table_info.get('gray_cells', []))}개")
    
    # 4. 엑셀 생성
    output_path = os.path.join(output_dir, f"테스트_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    result_path = analyzer.create_excel_with_highlights(tables_with_highlights, output_path, "테스트문서")
    
    if result_path:
        print(f"\n엑셀 파일 생성 완료: {result_path}")
    else:
        print("엑셀 파일 생성 실패")

# Streamlit 애플리케이션
def streamlit_app():
    st.title("PDF 강조 표시 추출 테스트")
    
    # 파일 업로드
    uploaded_file = st.file_uploader("PDF 파일 선택", type="pdf")
    
    if uploaded_file:
        # 임시 파일로 저장
        with open("temp.pdf", "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # PDF 정보 표시
        with fitz.open("temp.pdf") as doc:
            st.write(f"총 {len(doc)}페이지")
            
            # 페이지 선택
            page_num = st.number_input("분석할 페이지 선택 (1부터 시작)", 
                                     min_value=1, max_value=len(doc), value=1) - 1
            
            # 분석 시작
            if st.button("강조 표시 분석"):
                with st.spinner("분석 중..."):
                    analyzer = TableHighlightAnalyzer()
                    
                    # 테이블 추출
                    tables = analyzer.extract_tables("temp.pdf", page_num)
                    
                    if tables:
                        st.success(f"{len(tables)}개 테이블 추출 완료")
                        
                        # 강조색 매핑
                        tables_with_highlights = analyzer.map_highlights_to_tables("temp.pdf", tables, page_num)
                        
                        # 강조색 정보 출력
                        for i, table_info in enumerate(tables_with_highlights):
                            highlight_cells = table_info.get('highlight_cells', [])
                            gray_cells = table_info.get('gray_cells', [])
                            
                            st.write(f"테이블 {i+1}:")
                            st.write(f"- 강조된 셀: {len(highlight_cells)}개")
                            st.write(f"- 취소선 셀: {len(gray_cells)}개")
                            
                            # 데이터프레임 표시
                            df = table_info['df'].copy()
                            
                            # 강조 표시를 위한 스타일링
                            def highlight_cells(x):
                                df_styler = pd.DataFrame('', index=x.index, columns=x.columns)
                                
                                for r, c in highlight_cells:
                                    if r < len(df_styler) and c < len(df_styler.columns):
                                        df_styler.iloc[r, c] = 'background-color: yellow'
                                
                                for r, c in gray_cells:
                                    if r < len(df_styler) and c < len(df_styler.columns):
                                        df_styler.iloc[r, c] = 'background-color: lightgray; text-decoration: line-through'
                                
                                return df_styler
                            
                            # 스타일링된 데이터프레임 표시
                            st.dataframe(df.style.apply(highlight_cells, axis=None))
                        
                        # Excel 파일 생성
                        output = io.BytesIO()
                        result_path = analyzer.create_excel_with_highlights(
                            tables_with_highlights, output, uploaded_file.name)
                        
                        # 다운로드 버튼
                        st.download_button(
                            label="엑셀 파일 다운로드",
                            data=output.getvalue(),
                            file_name=f"강조표시_{uploaded_file.name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("테이블을 추출할 수 없습니다.")
        
        # 임시 파일 삭제
        if os.path.exists("temp.pdf"):
            os.remove("temp.pdf")

if __name__ == "__main__":
    # 명령줄에서 실행 시
    # test_highlight_mapping()
    
    # Streamlit으로 실행 시
    streamlit_app()