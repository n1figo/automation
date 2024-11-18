import PyPDF2
import re
import logging
from typing import Dict, List, Tuple
from dataclasses import dataclass
from collections import defaultdict

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class Section:
    start_page: int
    end_page: int = None
    title: str = None

class PDFSectionFinder:
    def __init__(self):
        # 검색할 섹션 패턴 정의
        self.section_patterns = {
            '상해': [
                r'[◇◆■□▶]([\s]*)(상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
                r'상해([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'상해([\s]*)보장([\s]*)(특약|특별약관)',
                r'상해([\s]*)담보([\s]*)(특약|특별약관)',
            ],
            '질병': [
                r'[◇◆■□▶]([\s]*)(질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
                r'질병([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'질병([\s]*)보장([\s]*)(특약|특별약관)',
                r'질병([\s]*)담보([\s]*)(특약|특별약관)',
            ],
            '상해및질병': [
                r'[◇◆■□▶]([\s]*)(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'(상해\s*및\s*질병|상해와\s*질병)([\s]*)보장([\s]*)(특약|특별약관)',
            ]
        }

    def find_section_boundaries(self, pdf_path: str) -> Dict[str, List[Section]]:
        """PDF에서 각 섹션의 범위를 찾습니다"""
        # 각 카테고리별 섹션 정보를 저장할 딕셔너리
        sections = defaultdict(list)
        section_starts = []  # 모든 섹션의 시작 페이지를 저장
        
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                logger.info(f"PDF 총 페이지 수: {total_pages}")
                
                # 첫 번째 패스: 모든 섹션의 시작 위치 찾기
                for page_num in range(total_pages):
                    text = reader.pages[page_num].extract_text()
                    
                    for category, patterns in self.section_patterns.items():
                        for pattern in patterns:
                            matches = re.finditer(pattern, text, re.IGNORECASE)
                            for match in matches:
                                found_text = match.group()
                                section = Section(
                                    start_page=page_num + 1,
                                    title=found_text
                                )
                                sections[category].append(section)
                                section_starts.append((page_num + 1, category))
                                logger.info(f"{category} 섹션 발견: '{found_text}' (페이지 {page_num + 1})")
                
                # 섹션 시작 페이지 정렬
                section_starts.sort()
                
                # 두 번째 패스: 섹션의 끝 페이지 설정
                for category in sections:
                    for i, section in enumerate(sections[category]):
                        # 현재 섹션의 시작 페이지 이후에 오는 첫 번째 다른 섹션 찾기
                        next_start = total_pages
                        for start_page, start_category in section_starts:
                            if start_page > section.start_page:
                                next_start = start_page - 1
                                break
                        section.end_page = next_start
                
            return sections
            
        except Exception as e:
            logger.error(f"PDF 처리 중 오류 발생: {str(e)}")
            return defaultdict(list)

    def analyze_sections(self, sections: Dict[str, List[Section]]) -> None:
        """섹션 분석 결과를 출력합니다"""
        print("\n=== 섹션 분석 결과 ===")
        
        for category, section_list in sections.items():
            if section_list:
                print(f"\n{category} 섹션:")
                for section in section_list:
                    print(f"- {section.title}")
                    print(f"  페이지 범위: {section.start_page} ~ {section.end_page}")
                    print(f"  섹션 길이: {section.end_page - section.start_page + 1} 페이지")
                
                # 섹션 간 간격 분석
                if len(section_list) > 1:
                    print(f"\n{category} 섹션 간격 분석:")
                    for i in range(1, len(section_list)):
                        gap = section_list[i].start_page - section_list[i-1].end_page - 1
                        if gap > 0:
                            print(f"  섹션 {i} → {i+1} 사이 간격: {gap} 페이지")
        
        # 섹션 간 중첩 분석
        self._analyze_overlaps(sections)
    
    def _analyze_overlaps(self, sections: Dict[str, List[Section]]) -> None:
        """섹션 간 중첩을 분석합니다"""
        print("\n=== 섹션 중첩 분석 ===")
        
        for cat1 in sections:
            for cat2 in sections:
                if cat1 < cat2:  # 중복 비교 방지
                    for sec1 in sections[cat1]:
                        for sec2 in sections[cat2]:
                            if (sec1.start_page <= sec2.end_page and 
                                sec2.start_page <= sec1.end_page):
                                print(f"\n중첩 발견:")
                                print(f"- {cat1}: 페이지 {sec1.start_page}~{sec1.end_page}")
                                print(f"- {cat2}: 페이지 {sec2.start_page}~{sec2.end_page}")

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    
    finder = PDFSectionFinder()
    sections = finder.find_section_boundaries(pdf_path)
    finder.analyze_sections(sections)

if __name__ == "__main__":
    main()