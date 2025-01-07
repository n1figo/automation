import PyPDF2
import re
import logging
from typing import Dict, List, Tuple, NamedTuple
from dataclasses import dataclass

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SectionMatch(NamedTuple):
    page: int
    title: str
    position: int
    text_before: str
    text_after: str
    type_number: str = None  # [1종], [2종] 등을 저장

@dataclass
class Section:
    category: str
    start_page: int
    end_page: int
    start_position: int = 0
    end_position: int = -1
    type_number: str = None

class PDFSectionAnalyzer:
    def __init__(self):
        self.type_pattern = r'\[(\d)종\]'
        self.section_patterns = {
            '상해': r'[◇◆■□▶]([\s]*)(상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
            '질병': r'[◇◆■□▶]([\s]*)(질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
            '상해및질병': r'[◇◆■□▶]([\s]*)(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)'
        }
        self.context_size = 200

    def find_types(self, text: str) -> List[str]:
        """페이지 내 종 구분([1종], [2종] 등)을 찾습니다"""
        return [f"[{match.group(1)}종]" for match in re.finditer(self.type_pattern, text)]

    def find_section_matches(self, pdf_path: str) -> List[SectionMatch]:
        """페이지 내 위치 정보와 종 구분을 포함하여 섹션을 찾습니다"""
        matches = []
        
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                text = reader.pages[page_num].extract_text()
                types_in_page = self.find_types(text)
                current_type = types_in_page[0] if types_in_page else None
                
                for category, pattern in self.section_patterns.items():
                    for match in re.finditer(pattern, text):
                        start_pos = max(0, match.start() - self.context_size)
                        end_pos = min(len(text), match.end() + self.context_size)
                        
                        matches.append(SectionMatch(
                            page=page_num + 1,
                            title=match.group(),
                            position=match.start(),
                            text_before=text[start_pos:match.start()].strip(),
                            text_after=text[match.end():end_pos].strip(),
                            type_number=current_type
                        ))
        
        return sorted(matches, key=lambda x: (x.page, x.position))

    def determine_section_boundaries(self, matches: List[SectionMatch]) -> List[Section]:
        """섹션의 경계를 결정합니다"""
        sections = []
        current_section = None
        
        for i, match in enumerate(matches):
            category = self._determine_category(match.title)
            
            if current_section is None:
                current_section = Section(
                    category=category,
                    start_page=match.page,
                    end_page=match.page,
                    start_position=match.position,
                    type_number=match.type_number
                )
            else:
                # 종이 변경되거나 다른 섹션이 시작되는 경우
                if (match.type_number != current_section.type_number or 
                    (current_section.category == '상해' and category == '질병')):
                    
                    if match.page == current_section.start_page:
                        current_section.end_position = match.position - 1
                    else:
                        current_section.end_page = match.page - 1
                        current_section.end_position = -1
                    
                    sections.append(current_section)
                    current_section = Section(
                        category=category,
                        start_page=match.page,
                        end_page=match.page,
                        start_position=match.position,
                        type_number=match.type_number
                    )
                else:
                    # 같은 종 내에서 페이지가 변경되는 경우
                    current_section.end_page = match.page
                    current_section.end_position = match.position - 1
        
        # 마지막 섹션 처리
        if current_section:
            sections.append(current_section)
        
        return sections

    def _determine_category(self, title: str) -> str:
        """제목에서 카테고리를 결정합니다"""
        if '상해및질병' in title or '상해 및 질병' in title:
            return '상해및질병'
        elif '상해' in title:
            return '상해'
        elif '질병' in title:
            return '질병'
        return 'unknown'

    def analyze_pdf(self, pdf_path: str) -> None:
        """PDF 분석을 실행하고 결과를 출력합니다"""
        matches = self.find_section_matches(pdf_path)
        sections = self.determine_section_boundaries(matches)
        
        # 종별 그룹화
        type_sections: Dict[str, List[Section]] = {}
        for section in sections:
            type_num = section.type_number if section.type_number else "종구분없음"
            if type_num not in type_sections:
                type_sections[type_num] = []
            type_sections[type_num].append(section)

        # 결과 출력
        print("\n=== PDF 섹션 분석 결과 ===")
        for type_num, type_sections_list in type_sections.items():
            print(f"\n■ {type_num}")
            for section in type_sections_list:
                print(f"\n{section.category} 섹션:")
                if section.start_page == section.end_page:
                    print(f"페이지 {section.start_page} (위치: {section.start_position} ~ {section.end_position})")
                else:
                    print(f"시작: 페이지 {section.start_page} (위치: {section.start_position})")
                    print(f"끝: 페이지 {section.end_page}" + 
                          (f" (위치: {section.end_position})" if section.end_position >= 0 else ""))

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    analyzer = PDFSectionAnalyzer()
    analyzer.analyze_pdf(pdf_path)

if __name__ == "__main__":
    main()