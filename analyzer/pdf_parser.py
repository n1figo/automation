import fitz  # PyMuPDF
import re
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Set, Optional, Any
import camelot

class PDFParser:
    """PDF 문서 파싱 및 분석 클래스"""
    
    def __init__(self, pdf_path: str):
        """
        PDF 파서 초기화
        
        Args:
            pdf_path: PDF 파일 경로
        """
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        
        # 정규표현식 패턴 정의
        self.patterns = {
            "insurance_payment": r"나\.\s*보험금",
            "insurance_type": r'\[(\d)종\]',
            "injury_special": r'상해관련특별약관|상해관련특약|상해\s*관련\s*특약',
            "disease_special": r'질병관련특별약관|질병관련특약|질병\s*관련\s*특약',
            "combined_special": r'상해및질병관련특별약관|상해및질병관련|상해질병관련특별약관',
            "section_start": r'[가-힣]\.\s+\w+',
        }
        
    def __del__(self):
        """소멸자: PDF 문서 닫기"""
        if hasattr(self, 'doc') and self.doc:
            self.doc.close()
    
    def get_total_pages(self) -> int:
        """총 페이지 수 반환"""
        return len(self.doc)
    
    def find_insurance_payment_page(self) -> Optional[int]:
        """'나. 보험금' 페이지 찾기"""
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            if re.search(self.patterns["insurance_payment"], text):
                return page_num
        return None
    
    def find_all_insurance_types(self) -> Dict[str, List[int]]:
        """모든 종별([1종], [2종] 등) 발견 페이지 찾기"""
        type_pages = {}
        
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            matches = re.finditer(self.patterns["insurance_type"], text)
            
            for match in matches:
                type_num = match.group(1)
                type_key = f"[{type_num}종]"
                
                if type_key not in type_pages:
                    type_pages[type_key] = []
                
                type_pages[type_key].append(page_num + 1)  # 1-based 페이지 번호
        
        return type_pages
    
    def identify_type_start_pages(self, parsing_start_page: int) -> Dict[str, Dict[str, Any]]:
        """종별 시작 페이지 식별"""
        type_starts = {}
        total_pages = len(self.doc)
        
        # 종별 표시 패턴
        type_pattern = self.patterns["insurance_type"]
        
        # 종별 시작 페이지의 특징적 구조 키워드
        structure_keywords = ["기본담보", "선택특약", "보장내용", "지급사유", "지급금액"]
        
        # 전체 종별 표시가 있는 페이지 우선 수집
        all_type_pages = {}
        for page_num in range(parsing_start_page, total_pages):
            page = self.doc[page_num]
            text = page.get_text()
            matches = list(re.finditer(type_pattern, text))
            
            for match in matches:
                type_num = match.group(1)
                type_key = f"[{type_num}종]"
                
                if type_key not in all_type_pages:
                    all_type_pages[type_key] = []
                
                all_type_pages[type_key].append(page_num)
        
        # 각 종별 첫 발견 페이지부터 시작 페이지 식별
        for type_key, pages in all_type_pages.items():
            # 첫 발견 페이지
            first_page = min(pages)
            best_page = None
            best_confidence = 0
            
            # 해당 종이 발견된 페이지들 중에서 시작 페이지 특징 분석
            for page_num in pages:
                page = self.doc[page_num]
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
                page = self.doc[best_page]
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
    
    def calculate_type_ranges(self, type_starts: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """종별 범위 계산"""
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
                end_page = len(self.doc)
            
            type_ranges[type_key] = {
                "start_page": start_page,
                "end_page": end_page,
                "confidence": info["confidence"],
                "features": info["features"]
            }
        
        return type_ranges
    
    def find_injury_special_sections(self) -> List[int]:
        """상해관련특별약관 찾기"""
        results = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
            
            if re.search(self.patterns["injury_special"], text_normalized) and \
               not re.search(self.patterns["combined_special"], text_normalized):
                results.append(page_num + 1)  # 1-based 페이지 번호
        
        return results
    
    def find_disease_special_sections(self) -> List[int]:
        """질병관련특별약관 찾기"""
        results = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
            
            if re.search(self.patterns["disease_special"], text_normalized) and \
               not re.search(self.patterns["combined_special"], text_normalized):
                results.append(page_num + 1)  # 1-based 페이지 번호
        
        return results
    
    def find_combined_special_sections(self) -> List[int]:
        """상해및질병관련특별약관 찾기"""
        results = []
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            text_normalized = ''.join(text.split())  # 띄어쓰기 제거
            
            if re.search(self.patterns["combined_special"], text_normalized):
                results.append(page_num + 1)  # 1-based 페이지 번호
        
        return results
    
    def find_combined_section_end_page(self) -> Optional[int]:
        """상해및질병관련특별약관 종료 페이지 찾기"""
        combined_pages = self.find_combined_special_sections()
        if not combined_pages:
            return None
        
        # 마지막 상해및질병관련특별약관 페이지
        last_combined_page = max(combined_pages) - 1  # 인덱스로 변환
        
        # 종료 페이지 찾기 - 다음 주요 섹션 시작점
        for page_num in range(last_combined_page + 1, len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            
            # 새로운 주요 섹션이 시작되는지 확인 (예: "다. 새로운섹션")
            if re.search(self.patterns["section_start"], text) and "보험금" not in text:
                return page_num + 1  # 1-based 페이지 번호
        
        # 종료 페이지를 찾지 못했다면 문서 끝까지로 간주
        return len(self.doc)
    
    def detect_all_highlights(self, start_page: int, end_page: int, 
                              highlight_analyzer) -> Tuple[List[int], List[int]]:
        """강조색 및 취소선 검출 (페이지 번호 리스트 반환)"""
        highlight_pages = []
        strikethrough_pages = []
        
        for page_num in range(start_page, end_page + 1):
            highlight_result = highlight_analyzer.detect_highlights(self.doc, page_num)
            
            if highlight_result["has_highlight"]:
                highlight_pages.append(page_num + 1)  # 1-based 페이지 번호
            
            if highlight_result["has_strikethrough"] or highlight_result["has_gray"]:
                strikethrough_pages.append(page_num + 1)  # 1-based 페이지 번호
        
        return highlight_pages, strikethrough_pages
    
    def extract_tables_from_range(self, start_page: int, end_page: int) -> List[Dict[str, Any]]:
        """지정된 페이지 범위에서 테이블 추출"""
        all_tables = []
        
        for page_num in range(start_page, end_page + 1):
            try:
                # Camelot으로 테이블 추출
                tables = self.extract_tables_with_camelot(page_num)
                
                if tables and len(tables) > 0:
                    # 테이블에 페이지 정보 추가
                    for table in tables:
                        table['page'] = page_num + 1  # 1-based 페이지 번호
                    
                    all_tables.extend(tables)
            except Exception as e:
                print(f"페이지 {page_num+1} 처리 중 오류: {str(e)}")
        
        return all_tables
    
    def extract_tables_with_camelot(self, page_num: int) -> List[Dict[str, Any]]:
        """Camelot을 사용하여 테이블 추출"""
        try:
            # Camelot 테이블 추출 - lattice 모드 시도
            tables = camelot.read_pdf(
                self.pdf_path, 
                pages=str(page_num + 1),
                flavor='lattice',
                line_scale=40,
                process_background=True,
                line_tol=2,
                strip_text='\n'
            )
            
            if len(tables) == 0:
                # lattice 모드에서 테이블을 찾지 못한 경우 stream 모드 시도
                tables = camelot.read_pdf(
                    self.pdf_path,
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
                    results.append({
                        'table_index': i,
                        'df': df,
                        'accuracy': table.parsing_report.get('accuracy', 0)
                    })
            
            return results
            
        except Exception as e:
            print(f"페이지 {page_num+1}의 테이블 추출 중 오류 발생: {str(e)}")
            return []