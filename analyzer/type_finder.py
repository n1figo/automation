import re
from typing import Dict, List, Tuple, Optional, Any, Set

class TypeFinder:
    """종별 정보 찾기 클래스"""
    
    def __init__(self, pdf_document):
        """
        TypeFinder 초기화
        
        Args:
            pdf_document: PyMuPDF 문서 객체
        """
        self.pdf_document = pdf_document
        self.type_pattern = r'\[(\d)종\]'
        self.structure_keywords = ["기본담보", "선택특약", "보장내용", "지급사유", "지급금액"]
    
    def find_all_types(self) -> Dict[str, List[int]]:
        """
        문서 전체에서 모든 종별 패턴([1종], [2종] 등)을 찾아 페이지 번호와 함께 반환
        
        Returns:
            Dict[str, List[int]]: {종별: [페이지 번호 리스트]}
        """
        type_pages = {}
        
        for page_num in range(len(self.pdf_document)):
            page = self.pdf_document[page_num]
            text = page.get_text()
            matches = re.finditer(self.type_pattern, text)
            
            for match in matches:
                type_num = match.group(1)
                type_key = f"[{type_num}종]"
                
                if type_key not in type_pages:
                    type_pages[type_key] = []
                
                type_pages[type_key].append(page_num + 1)  # 1-based 페이지 번호
        
        return type_pages
    
    def identify_start_pages(self, parsing_start_page: int = 0) -> Dict[str, Dict[str, Any]]:
        """
        각 종별 시작 페이지 식별
        
        Args:
            parsing_start_page: 파싱 시작 페이지 인덱스 (기본값: 0)
            
        Returns:
            Dict[str, Dict[str, Any]]: 종별 시작 페이지 정보
        """
        type_starts = {}
        total_pages = len(self.pdf_document)
        
        # 전체 종별 표시가 있는 페이지 우선 수집
        all_type_pages = {}
        for page_num in range(parsing_start_page, total_pages):
            page = self.pdf_document[page_num]
            text = page.get_text()
            matches = list(re.finditer(self.type_pattern, text))
            
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
                page = self.pdf_document[page_num]
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
                page = self.pdf_document[best_page]
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
    
    def calculate_ranges(self, type_starts: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """
        종별 범위 계산
        
        Args:
            type_starts: 종별 시작 페이지 정보
            
        Returns:
            Dict[str, Dict[str, Any]]: 종별 범위 정보
        """
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
                end_page = len(self.pdf_document)
            
            type_ranges[type_key] = {
                "start_page": start_page,
                "end_page": end_page,
                "confidence": info["confidence"],
                "features": info["features"]
            }
        
        return type_ranges