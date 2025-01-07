import PyPDF2
import re
import logging
from typing import Dict, List

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

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
            ]
        }

    def find_sections(self, pdf_path: str) -> Dict[str, List[int]]:
        """PDF에서 각 섹션의 페이지를 찾습니다"""
        section_pages = {category: [] for category in self.section_patterns.keys()}
        
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                logger.info(f"PDF 총 페이지 수: {total_pages}")
                
                for page_num in range(total_pages):
                    text = reader.pages[page_num].extract_text()
                    
                    # 각 카테고리와 패턴으로 검색
                    for category, patterns in self.section_patterns.items():
                        for pattern in patterns:
                            matches = re.finditer(pattern, text, re.IGNORECASE)
                            for match in matches:
                                if page_num + 1 not in section_pages[category]:  # 중복 제거
                                    section_pages[category].append(page_num + 1)
                                    found_text = match.group()
                                    logger.info(f"{category} 섹션 발견: '{found_text}' (페이지 {page_num + 1})")
                                    
                                    # 발견된 부분 전후 컨텍스트 출력
                                    start_idx = max(0, match.start() - 50)
                                    end_idx = min(len(text), match.end() + 150)
                                    context = text[start_idx:end_idx].replace('\n', ' ').strip()
                                    logger.info(f"컨텍스트: ...{context}...")
                
                # 각 카테고리의 페이지 번호 정렬
                for category in section_pages:
                    section_pages[category] = sorted(section_pages[category])
            
            return section_pages
            
        except Exception as e:
            logger.error(f"PDF 처리 중 오류 발생: {str(e)}")
            return {category: [] for category in self.section_patterns.keys()}

    def analyze_page_gaps(self, pages: List[int], category: str) -> None:
        """페이지 간격을 분석하고 출력합니다"""
        if len(pages) > 1:
            print(f"\n{category} 섹션 페이지 간격 분석:")
            for i in range(1, len(pages)):
                gap = pages[i] - pages[i-1]
                print(f"페이지 {pages[i-1]} → {pages[i]} (간격: {gap}페이지)")
                
                # 간격이 큰 경우 경고
                if gap > 20:
                    logger.warning(f"큰 페이지 간격 발견 ({gap}페이지), 중간에 누락된 섹션이 있을 수 있습니다.")

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    
    finder = PDFSectionFinder()
    results = finder.find_sections(pdf_path)
    
    # 결과 출력
    print("\n=== 검색 결과 ===")
    for category, pages in results.items():
        if pages:
            print(f"\n{category} 관련 섹션 위치:")
            print(f"발견된 페이지: {pages}")
            finder.analyze_page_gaps(pages, category)
        else:
            print(f"\n{category} 관련 섹션을 찾을 수 없습니다.")
    
    # 섹션 간 관계 분석
    all_pages = set()
    for pages in results.values():
        all_pages.update(pages)
    
    if len(all_pages) > 0:
        print("\n=== 전체 분석 ===")
        print(f"총 발견된 고유 페이지 수: {len(all_pages)}")
        all_pages = sorted(list(all_pages))
        print(f"전체 페이지 순서: {all_pages}")
        
        # 섹션이 같은 페이지에 있는 경우 확인
        overlapping_pages = set()
        for category1 in results:
            for category2 in results:
                if category1 < category2:  # 중복 비교 방지
                    common_pages = set(results[category1]) & set(results[category2])
                    if common_pages:
                        print(f"\n{category1}와 {category2} 섹션이 같은 페이지에 있음: {sorted(list(common_pages))}")
                        overlapping_pages.update(common_pages)

if __name__ == "__main__":
    main()