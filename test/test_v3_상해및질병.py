import PyPDF2
import re
import logging
from typing import Dict, List
from dataclasses import dataclass

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class Section:
    title: str
    start_page: int
    end_page: int = None

class PDFSectionFinder:
    def __init__(self):
        self.section_patterns = {
            '상해': [
                r'[◇◆■□▶]([\s]*)(상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
                r'상해([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'상해([\s]*)보장([\s]*)(특약|특별약관)',
            ],
            '질병': [
                r'[◇◆■□▶]([\s]*)(질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
                r'질병([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'질병([\s]*)보장([\s]*)(특약|특별약관)',
            ],
            '상해및질병': [
                r'[◇◆■□▶]([\s]*)(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)',
                r'(상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)',
            ]
        }

    def find_all_section_starts(self, text: str, page_num: int) -> Dict[str, List[tuple]]:
        """페이지에서 모든 섹션 시작점을 찾습니다"""
        found_sections = {}
        
        for category, patterns in self.section_patterns.items():
            for pattern in patterns:
                matches = re.finditer(pattern, text, re.IGNORECASE)
                for match in matches:
                    if category not in found_sections:
                        found_sections[category] = []
                    found_sections[category].append((page_num, match.group()))
        
        return found_sections

    def find_section_boundaries(self, pdf_path: str) -> Dict[str, List[Section]]:
        """PDF에서 각 섹션의 범위를 찾습니다"""
        sections = {
            '상해': [],
            '질병': [],
            '상해및질병': []
        }
        
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                logger.info(f"PDF 총 페이지 수: {total_pages}")

                # 모든 섹션 시작점 찾기
                section_starts = []
                for page_num in range(total_pages):
                    text = reader.pages[page_num].extract_text()
                    found = self.find_all_section_starts(text, page_num)
                    
                    for category, matches in found.items():
                        for page, title in matches:
                            section_starts.append((page + 1, category, title))
                
                # 시작점 시간순 정렬
                section_starts.sort(key=lambda x: x[0])
                
                # 각 섹션의 범위 결정
                for i, (start_page, category, title) in enumerate(section_starts):
                    # 다음 섹션의 시작점 찾기
                    end_page = total_pages
                    if i < len(section_starts) - 1:
                        next_category = section_starts[i + 1][1]
                        if ((category == '상해' and next_category == '질병') or
                            (category == '질병' and (next_category == '상해및질병' or next_category != '질병'))):
                            end_page = section_starts[i + 1][0] - 1
                    
                    # 불필요한 중복 제거
                    if not any(s.start_page == start_page for s in sections[category]):
                        section = Section(title=title, start_page=start_page, end_page=end_page)
                        sections[category].append(section)
                        logger.info(f"{category} 섹션 찾음: {start_page}~{end_page} ({title})")
            
            return sections
            
        except Exception as e:
            logger.error(f"PDF 처리 중 오류 발생: {str(e)}")
            return sections

    def analyze_sections(self, sections: Dict[str, List[Section]]) -> None:
        """섹션 분석 결과를 출력합니다"""
        print("\n=== 섹션 분석 결과 ===")
        
        for category, section_list in sections.items():
            if section_list:
                print(f"\n{category} 섹션:")
                for section in section_list:
                    print(f"페이지: {section.start_page} ~ {section.end_page}")
                    print(f"길이: {section.end_page - section.start_page + 1} 페이지")
                
                # 섹션 간 간격 분석
                if len(section_list) > 1:
                    print(f"\n{category} 섹션 간격:")
                    for i in range(1, len(section_list)):
                        gap = section_list[i].start_page - section_list[i-1].end_page - 1
                        if gap > 0:
                            print(f"섹션 {i} → {i+1} 사이 간격: {gap} 페이지")

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    
    finder = PDFSectionFinder()
    sections = finder.find_section_boundaries(pdf_path)
    finder.analyze_sections(sections)

if __name__ == "__main__":
    main()