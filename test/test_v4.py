import PyPDF2
import re
import logging
from typing import Dict, List, Tuple
from dataclasses import dataclass

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class SectionInfo:
    title: str
    start_page: int
    end_page: int
    category: str
    type_number: str = None  # [1종], [2종] 등을 저장

class PDFAnalyzer:
    def __init__(self):
        self.type_pattern = r'\[(\d)종\]'
        self.section_patterns = {
            '상해': r'[◇◆■□▶]([\s]*)(상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
            '질병': r'[◇◆■□▶]([\s]*)(질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
            '상해및질병': r'[◇◆■□▶]([\s]*)(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)'
        }

    def find_types(self, pdf_path: str) -> List[Tuple[int, str]]:
        """[1종], [2종] 등이 나오는 페이지를 찾습니다"""
        type_pages = []
        
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                text = reader.pages[page_num].extract_text()
                matches = re.finditer(self.type_pattern, text)
                for match in matches:
                    type_num = match.group(1)
                    type_pages.append((page_num + 1, f"[{type_num}종]"))
        
        return sorted(type_pages, key=lambda x: (int(re.search(r'\[(\d+)종\]', x[1]).group(1)), x[0]))

    def find_first_occurrences(self, pdf_path: str) -> Dict[str, List[int]]:
        """각 섹션이 처음 나오는 페이지들을 찾습니다"""
        first_occurrences = {
            '상해': set(),
            '질병': set(),
            '상해및질병': set()
        }
        
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                text = reader.pages[page_num].extract_text()
                
                for category, pattern in self.section_patterns.items():
                    if re.search(pattern, text):
                        first_occurrences[category].add(page_num + 1)
        
        return {k: sorted(list(v)) for k, v in first_occurrences.items()}

    def analyze_sections_by_type(self, pdf_path: str) -> None:
        """종별로 섹션을 분석합니다"""
        # 1. 먼저 종 구분이 있는지 확인
        type_pages = self.find_types(pdf_path)
        
        if not type_pages:
            logger.info("종 구분이 없습니다. 전체 섹션 분석을 진행합니다.")
            self.analyze_whole_sections(pdf_path)
            return
        
        logger.info("=== 종별 페이지 ===")
        for page, type_num in type_pages:
            logger.info(f"{type_num}: 페이지 {page}")
            
        # 종별 섹션 범위 분석
        self.analyze_sections_for_each_type(pdf_path, type_pages)

    def analyze_whole_sections(self, pdf_path: str) -> None:
        """종 구분이 없을 때의 전체 섹션 분석"""
        first_occurrences = self.find_first_occurrences(pdf_path)
        
        print("\n=== 섹션 시작 페이지 ===")
        for category, pages in first_occurrences.items():
            print(f"{category}: {pages}")

    def analyze_sections_for_each_type(self, pdf_path: str, type_pages: List[Tuple[int, str]]) -> None:
        """각 종별로 섹션을 분석합니다"""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            for i, (type_start, type_num) in enumerate(type_pages):
                # 현재 종의 끝 페이지 결정
                type_end = type_pages[i + 1][0] - 1 if i < len(type_pages) - 1 else len(reader.pages)
                
                print(f"\n=== {type_num} 섹션 분석 ({type_start} ~ {type_end} 페이지) ===")
                
                # 이 종 내에서의 섹션 찾기
                sections = self.find_sections_in_range(pdf_path, type_start, type_end)
                
                for category, section_pages in sections.items():
                    if section_pages:
                        print(f"{category} 섹션:")
                        for start, end in section_pages:
                            print(f"페이지: {start} ~ {end}")

    def find_sections_in_range(self, pdf_path: str, start_page: int, end_page: int) -> Dict[str, List[Tuple[int, int]]]:
        """지정된 페이지 범위 내에서 섹션을 찾습니다"""
        sections = {
            '상해': [],
            '질병': [],
            '상해및질병': []
        }
        
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            current_section = None
            section_start = None
            
            for page_num in range(start_page - 1, end_page):
                text = reader.pages[page_num].extract_text()
                
                for category, pattern in self.section_patterns.items():
                    if re.search(pattern, text):
                        if current_section:
                            sections[current_section].append((section_start, page_num))
                        current_section = category
                        section_start = page_num + 1
            
            # 마지막 섹션 처리
            if current_section and section_start:
                sections[current_section].append((section_start, end_page))
        
        return sections

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    analyzer = PDFAnalyzer()
    analyzer.analyze_sections_by_type(pdf_path)

if __name__ == "__main__":
    main()