import PyPDF2
import re
import logging
from typing import Dict, List, Tuple, NamedTuple
from dataclasses import dataclass
import pandas as pd

class Position(NamedTuple):
    page: int
    offset: int

@dataclass
class SectionInfo:
    title: str
    start: Position
    end: Position
    category: str
    type_number: str = None

class PDFAnalyzer:
    def __init__(self):
        self.type_pattern = r'\[(\d)종\]'
        self.section_patterns = {
            '상해': r'[◇◆■□▶]([\s]*)(상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
            '질병': r'[◇◆■□▶]([\s]*)(질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
            '상해및질병': r'[◇◆■□▶]([\s]*)(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)'
        }
        self.logger = logging.getLogger(__name__)

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
        
        # 종 번호와 페이지 번호로 정렬
        sorted_pages = sorted(type_pages, 
            key=lambda x: (int(re.search(r'\[(\d+)종\]', x[1]).group(1)), x[0]))
        
        self.logger.info(f"발견된 종 구분: {[page[1] for page in sorted_pages]}")
        return sorted_pages

    def find_section_positions(self, text: str, page_num: int) -> List[Tuple[str, Position]]:
        """페이지 내에서 각 섹션의 정확한 위치를 찾습니다"""
        positions = []
        lines = text.split('\n')
        
        for line_num, line in enumerate(lines):
            for category, pattern in self.section_patterns.items():
                if re.search(pattern, line):
                    position = Position(page_num, line_num)
                    positions.append((category, position))
        
        return sorted(positions, key=lambda x: x[1].offset)

    def find_sections_in_range(self, pdf_path: str, start_page: int, end_page: int) -> Dict[str, List[SectionInfo]]:
        """페이지 범위 내에서 섹션을 찾고 정확한 위치 정보를 포함합니다"""
        sections = {
            '상해': [],
            '질병': [],
            '상해및질병': []
        }
        
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            current_info = None
            
            for page_num in range(start_page - 1, end_page):
                text = reader.pages[page_num].extract_text()
                positions = self.find_section_positions(text, page_num + 1)
                
                # 페이지 내 여러 섹션 처리
                for idx, (category, position) in enumerate(positions):
                    # 이전 섹션이 있으면 종료
                    if current_info:
                        end_position = Position(position.page, position.offset)
                        current_info.end = end_position
                        sections[current_info.category].append(current_info)
                    
                    # 새로운 섹션 시작
                    title = self._extract_section_title(text, position.offset)
                    current_info = SectionInfo(
                        title=title,
                        start=position,
                        end=None,
                        category=category
                    )
                
                # 페이지의 마지막 섹션 처리
                if current_info and current_info.end is None and page_num == end_page - 1:
                    current_info.end = Position(page_num + 1, self._get_page_line_count(text))
                    sections[current_info.category].append(current_info)
        
        return sections

    def _extract_section_title(self, text: str, offset: int) -> str:
        """섹션 제목을 추출합니다"""
        lines = text.split('\n')
        if offset < len(lines):
            return lines[offset].strip()
        return ""

    def _get_page_line_count(self, text: str) -> int:
        """페이지의 총 줄 수를 반환합니다"""
        return len(text.split('\n'))

    def analyze_sections_by_type(self, pdf_path: str) -> None:
        """종별로 섹션을 상세 분석합니다"""
        type_pages = self.find_types(pdf_path)
        
        if not type_pages:
            self.logger.info("종 구분이 없습니다. 전체 섹션 분석을 진행합니다.")
            self.analyze_whole_sections(pdf_path)
            return
        
        self.logger.info("\n=== 종별 상세 분석 결과 ===")
        
        for i, (type_start, type_num) in enumerate(type_pages):
            type_end = type_pages[i + 1][0] - 1 if i < len(type_pages) - 1 else self._get_total_pages(pdf_path)
            
            print(f"\n{type_num} 분석 ({type_start} ~ {type_end} 페이지)")
            sections = self.find_sections_in_range(pdf_path, type_start, type_end)
            
            for category, section_infos in sections.items():
                if section_infos:
                    print(f"\n{category} 섹션:")
                    for info in section_infos:
                        print(f"- {info.title}")
                        print(f"  시작: {info.start.page}페이지 {info.start.offset}번째 줄")
                        if info.end:
                            print(f"  종료: {info.end.page}페이지 {info.end.offset}번째 줄")

    def _get_total_pages(self, pdf_path: str) -> int:
        """PDF 총 페이지 수를 반환합니다"""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            return len(reader.pages)

    def analyze_whole_sections(self, pdf_path: str) -> None:
        """종 구분이 없을 때의 전체 섹션 분석"""
        total_pages = self._get_total_pages(pdf_path)
        sections = self.find_sections_in_range(pdf_path, 1, total_pages)
        
        print("\n=== 전체 섹션 분석 결과 ===")
        for category, section_infos in sections.items():
            if section_infos:
                print(f"\n{category} 섹션:")
                for info in section_infos:
                    print(f"- {info.title}")
                    print(f"  시작: {info.start.page}페이지 {info.start.offset}번째 줄")
                    if info.end:
                        print(f"  종료: {info.end.page}페이지 {info.end.offset}번째 줄")

def main():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    analyzer = PDFAnalyzer()
    analyzer.analyze_sections_by_type(pdf_path)

if __name__ == "__main__":
    main()