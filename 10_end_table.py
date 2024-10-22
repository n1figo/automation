import PyPDF2
import re
import logging
import fitz
import numpy as np
from typing import Dict, List, Tuple, Optional
import os

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PageStructureAnalyzer:
    """
    페이지 구조를 분석하는 클래스
    """
    def analyze_page_structure(self, page) -> dict:
        """
        페이지의 구조적 특징 분석
        """
        blocks = page.get_text("blocks")
        lines = page.get_text("text").split('\n')
        
        # 여백 계산
        margins = self._calculate_margins(page)
        
        return {
            'text_blocks': len(blocks),
            'line_spacing': self._calculate_line_spacing(blocks),
            'margins': margins,
            'table_presence': self._detect_table(page),
            'avg_line_length': np.mean([len(line) for line in lines if line.strip()])
        }

    def _calculate_margins(self, page) -> dict:
        """
        페이지 여백 계산
        """
        blocks = page.get_text("blocks")
        if not blocks:
            return {'left': 0, 'right': 0, 'top': 0, 'bottom': 0}
            
        page_width = page.rect.width
        page_height = page.rect.height
        
        left_margins = []
        right_margins = []
        top_margins = []
        bottom_margins = []
        
        for block in blocks:
            x0, y0, x1, y1, text = block[:5]
            left_margins.append(x0)
            right_margins.append(page_width - x1)
            top_margins.append(y0)
            bottom_margins.append(page_height - y1)
            
        return {
            'left': min(left_margins) if left_margins else 0,
            'right': min(right_margins) if right_margins else 0,
            'top': min(top_margins) if top_margins else 0,
            'bottom': min(bottom_margins) if bottom_margins else 0
        }

    def _calculate_line_spacing(self, blocks) -> float:
        """
        평균 줄 간격 계산
        """
        if len(blocks) < 2:
            return 0
            
        spacings = []
        for i in range(len(blocks)-1):
            _, y1, _, _, _ = blocks[i][:5]
            _, y2, _, _, _ = blocks[i+1][:5]
            spacings.append(y2 - y1)
            
        return np.mean(spacings) if spacings else 0

    def _detect_table(self, page) -> bool:
        """
        표 존재 여부 탐지
        """
        text = page.get_text()
        table_markers = ['│', '─', '├', '┤', '┌', '┐', '└', '┘']
        return any(marker in text for marker in table_markers)

class TitleDetector:
    def __init__(self):
        self.previous_title = None
        self.current_page = None
        self.structure_analyzer = PageStructureAnalyzer()
        
        self.title_patterns = [
            r'보험료\s*납입면제\s*관련\s*특약',
            r'간병\s*관련\s*특약',
            r'실손의료비\s*보장\s*특약',
            r'기타\s*특약',
            r'제\s*\d+\s*장',
            r'보장내용\s*요약서',
            r'주요\s*보장내용'
        ]

    def is_likely_title(self, text: str, page_structure: dict) -> bool:
        """
        텍스트가 제목일 가능성을 구조적 특징과 함께 확인
        """
        # 기본 필터링
        if not text or len(text) > 50:
            return False
            
        if text.strip().endswith(('.', '다', '요', '음')):
            return False
            
        # 구조적 특징 확인
        if page_structure['table_presence'] and len(text.strip()) < 20:
            return False  # 표 안의 짧은 텍스트는 제외
            
        # 알려진 제목 패턴 체크
        for pattern in self.title_patterns:
            if re.search(pattern, text, re.IGNORECASE):
                return True
                
        # 기타 제목 특성 체크
        if text.strip().startswith('제') and '장' in text:
            return True
        if '특약' in text and len(text) < 30:
            return True
            
        return False

    def detect_next_title(self, page: fitz.Page, page_num: int) -> Optional[Tuple[str, int]]:
        """
        다음 제목과 해당 페이지 번호 반환
        """
        text = page.get_text()
        structure = self.structure_analyzer.analyze_page_structure(page)
        
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if (self.is_likely_title(line, structure) and 
                line != self.previous_title):
                self.previous_title = line
                self.current_page = page_num
                return line, page_num
                
        return None

class SectionDetector:
    def __init__(self):
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }
        self.title_detector = TitleDetector()
        self.structure_analyzer = PageStructureAnalyzer()

    def is_significant_change(self, last_structure: dict, current_structure: dict) -> bool:
        """
        구조적 변화의 유의미성 판단
        """
        thresholds = {
            'text_blocks': 0.3,  # 블록 수 변화 임계값
            'line_spacing': 0.2,  # 줄 간격 변화 임계값
            'avg_line_length': 0.25  # 평균 라인 길이 변화 임계값
        }
        
        changes = {
            key: abs(current_structure[key] - last_structure[key]) / max(last_structure[key], 1)
            for key in thresholds.keys()
            if key in current_structure and key in last_structure
        }
        
        return any(change > thresholds[key] for key, change in changes.items())

    def find_sections(self, pdf_path: str) -> List[Tuple[str, int]]:
        """
        주요 섹션들과 그 다음에 나오는 제목들 찾기
        """
        results = []
        found_sections = set()
        found_disease = False
        last_structure = None
        
        try:
            with fitz.open(pdf_path) as doc:
                for page_num in range(1, len(doc) + 1):
                    page = doc[page_num-1]
                    text = page.get_text()
                    current_structure = self.structure_analyzer.analyze_page_structure(page)
                    
                    if not found_disease:
                        for section, pattern in self.key_patterns.items():
                            if section not in found_sections and re.search(pattern, text, re.IGNORECASE):
                                found_sections.add(section)
                                results.append((f"Found {section}", page_num))
                                logger.info(f"{section} found on page {page_num}")
                                
                                if section == '질병관련':
                                    found_disease = True
                        continue
                    
                    # 구조적 변화 확인 및 제목 탐지
                    if last_structure and self.is_significant_change(last_structure, current_structure):
                        title_info = self.title_detector.detect_next_title(page, page_num)
                        if title_info:
                            title, page = title_info
                            results.append((f"Next section: {title}", page))
                            logger.info(f"New section found: {title} on page {page}")
                    
                    last_structure = current_structure
                        
            return results
            
        except Exception as e:
            logger.error(f"Error processing PDF: {str(e)}")
            return []

def main():
    try:
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return
            
        detector = SectionDetector()
        results = detector.find_sections(pdf_path)
        
        if results:
            print("\n발견된 섹션들:")
            for text, page_num in results:
                print(f"페이지 {page_num}: {text}")
        else:
            print("\n섹션을 찾지 못했습니다.")
            
    except Exception as e:
        logger.error(f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()