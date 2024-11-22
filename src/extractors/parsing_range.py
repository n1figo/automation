import fitz
import re
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
import logging

@dataclass
class SectionInfo:
    title: str
    start_page: int
    end_page: int
    category: str
    line_number: int
    end_line_number: Optional[int] = field(default=None)

class PDFSectionAnalyzer:
    def __init__(self, doc: fitz.Document):
        self.doc = doc
        self.section_patterns = {
            '상해': r'[◇◆■□▶]*\s*(상해.*(?:특약|특별약관|보장특약)|상해\s*관련\s*(?:특약|특별약관|보장특약))',
            '질병': r'[◇◆■□▶]*\s*(질병.*(?:특약|특별약관|보장특약)|질병\s*관련\s*(?:특약|특별약관|보장특약))',
            '상해및질병': r'[◇◆■□▶]*\s*((?:상해\s*및\s*질병|상해와\s*질병).*(?:특약|특별약관|보장특약))'
        }
        self.type_pattern = r'\[(\d+)종\]'
        self.additional_patterns = {
            'title_patterns': [
                r'보험금\s*지급사유\s*및\s*지급금액',
                r'보험금의\s*지급사유\s*및\s*금액',
                r'보장내용\s*요약',
                r'기본계약\s*보장공통',
                r'보험금\s*지급사유,\s*지급금액\s*및\s*지급제한사항'
            ]
        }
        self.logger = self.setup_logging()

    def setup_logging(self):
        logger = logging.getLogger("PDFSectionAnalyzer")
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger

    def find_type_ranges(self) -> List[Tuple[str, int, int]]:
        """종별 페이지 범위를 찾습니다"""
        type_pages = []
        current_type = None
        start_page = None

        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            type_match = re.search(self.type_pattern, text)

            if type_match:
                type_num = type_match.group(1)
                if current_type != type_num:
                    if current_type is not None:
                        type_pages.append((f"[{current_type}종]", start_page, page_num - 1))
                    current_type = type_num
                    start_page = page_num

        if current_type is not None:
            type_pages.append((f"[{current_type}종]", start_page, len(self.doc) - 1))

        return type_pages

    def find_section_range(self, sections: Dict[str, List[SectionInfo]]) -> Tuple[Optional[int], Optional[int]]:
        """섹션의 전체 범위를 찾습니다"""
        all_sections = []
        for category in ['상해', '질병', '상해및질병']:
            if sections[category]:
                all_sections.extend(sections[category])
        
        if not all_sections:
            return None, None
        
        start_section = min(all_sections, key=lambda s: (s.start_page, s.line_number))
        end_section = max(all_sections, key=lambda s: (s.end_page, s.end_line_number or float('inf')))
        
        return start_section.start_page, end_section.end_page

    def find_sections_in_range(self, start_page: int, end_page: int) -> Dict[str, List[SectionInfo]]:
        """페이지 범위 내에서 섹션을 찾고 시작/끝 위치를 결정합니다"""
        sections = {'상해': [], '질병': [], '상해및질병': []}
        
        for page_num in range(start_page, end_page + 1):
            text = self.doc[page_num].get_text()
            lines = text.split('\n')

            for line_num, line in enumerate(lines, 1):
                stripped_line = line.strip().replace(' ', '')
                for category, pattern in self.section_patterns.items():
                    if re.search(pattern, stripped_line, re.IGNORECASE):
                        section = SectionInfo(
                            title=line.strip(),
                            start_page=page_num,
                            end_page=end_page,
                            category=category,
                            line_number=line_num
                        )
                        sections[category].append(section)

        # 섹션 범위 조정
        all_sections = []
        for category_sections in sections.values():
            all_sections.extend(category_sections)
        all_sections.sort(key=lambda s: (s.start_page, s.line_number))

        for i, section in enumerate(all_sections):
            if i < len(all_sections) - 1:
                next_section = all_sections[i + 1]
                if next_section.start_page > section.start_page:
                    section.end_page = next_section.start_page - 1
                    section.end_line_number = None
                elif next_section.start_page == section.start_page:
                    section.end_page = section.start_page
                    section.end_line_number = next_section.line_number - 1
            else:
                section.end_page = min(section.start_page + 10, end_page)
                section.end_line_number = None

        result_sections = {'상해': [], '질병': [], '상해및질병': []}
        for section in all_sections:
            result_sections[section.category].append(section)

        return result_sections

    def analyze_and_print_results(self) -> Dict[str, Dict[str, int]]:
        """분석 실행 및 결과 출력"""
        type_ranges = self.find_type_ranges()
        print("\n=== 특약 유형별 페이지 분석 결과 ===")

        results = {}
        for type_info in type_ranges:
            type_name, type_start, type_end = type_info
            print(f"\n{type_name}")

            sections = self.find_sections_in_range(type_start, type_end)
            
            # 각 특약 유형별 페이지 출력
            print(f"  상해 특약 페이지: ", end="")
            if sections['상해']:
                pages = sorted(set(s.start_page + 1 for s in sections['상해']))
                print(f"{', '.join(map(str, pages))}")
            else:
                print("없음")

            print(f"  질병 특약 페이지: ", end="")
            if sections['질병']:
                pages = sorted(set(s.start_page + 1 for s in sections['질병']))
                print(f"{', '.join(map(str, pages))}")
            else:
                print("없음")

            print(f"  상해및질병 특약 페이지: ", end="")
            if sections['상해및질병']:
                pages = sorted(set(s.start_page + 1 for s in sections['상해및질병']))
                print(f"{', '.join(map(str, pages))}")
            else:
                print("없음")

            # 결과 저장
            results[type_name] = {
                "start_page": type_start,
                "end_page": type_end,
                "sections": {}
            }

            for category in ['상해', '질병', '상해및질병']:
                if sections[category]:
                    results[type_name]["sections"][category] = {
                        "pages": sorted(set(s.start_page + 1 for s in sections[category]))
                    }

        return results

def analyze_pdf(pdf_path: str) -> Optional[Dict[str, Dict[str, int]]]:
    """PDF 분석을 수행하는 메인 함수"""
    with fitz.open(pdf_path) as doc:
        analyzer = PDFSectionAnalyzer(doc)
        
        type_ranges = analyzer.find_type_ranges()
        if type_ranges:
            return analyzer.analyze_and_print_results()
        
        print("\n종 구분이 없는 문서입니다.")
        sections = analyzer.find_sections_in_range(0, len(doc) - 1)
        
        # 종 구분이 없는 경우의 페이지 출력
        print("\n=== 특약 유형별 페이지 분석 결과 ===")
        print("\n전체 문서")
        print(f"  상해 특약 페이지: ", end="")
        if sections['상해']:
            pages = sorted(set(s.start_page + 1 for s in sections['상해']))
            print(f"{', '.join(map(str, pages))}")
        else:
            print("없음")

        print(f"  질병 특약 페이지: ", end="")
        if sections['질병']:
            pages = sorted(set(s.start_page + 1 for s in sections['질병']))
            print(f"{', '.join(map(str, pages))}")
        else:
            print("없음")

        print(f"  상해및질병 특약 페이지: ", end="")
        if sections['상해및질병']:
            pages = sorted(set(s.start_page + 1 for s in sections['상해및질병']))
            print(f"{', '.join(map(str, pages))}")
        else:
            print("없음")

        # 결과 반환
        if any(sections.values()):
            return {
                "DefaultType": {
                    "sections": {
                        category: {
                            "pages": sorted(set(s.start_page + 1 for s in sect))
                        }
                        for category, sect in sections.items() if sect
                    }
                }
            }
        
        return None

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python parsing_range.py <pdf_path>")
    else:
        pdf_path = sys.argv[1]
        analyze_pdf(pdf_path)
