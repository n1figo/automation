import PyPDF2
import re
import logging
import fitz
from typing import Dict, List, Tuple, Optional
import os

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TitleDetector:
    def __init__(self):
        self.previous_title = None
        self.current_page = None
        
        # 제목으로 확인된 패턴들
        self.title_patterns = [
            r'보험료\s*납입면제\s*관련\s*특약',
            r'간병\s*관련\s*특약',
            r'실손의료비\s*보장\s*특약',
            r'기타\s*특약',
            r'제\s*\d+\s*장',
            r'보장내용\s*요약서',
            r'주요\s*보장내용'
        ]

    def is_likely_title(self, text: str) -> bool:
        """
        텍스트가 제목일 가능성 확인
        """
        # 기본 필터링
        if not text or len(text) > 50:
            return False
            
        # 문장 끝 패턴 체크 (제목은 보통 문장 부호로 끝나지 않음)
        if text.strip().endswith(('.', '다', '요', '음')):
            return False
            
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

    def detect_next_title(self, text: str, page_num: int) -> Optional[Tuple[str, int]]:
        """
        다음 제목과 해당 페이지 번호 반환
        """
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if (self.is_likely_title(line) and 
                line != self.previous_title):
                
                self.previous_title = line
                self.current_page = page_num
                return line, page_num
                
        return None

class SectionDetector:
    def __init__(self):
        # 상해/질병 특약 패턴
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }
        self.title_detector = TitleDetector()

    def find_sections(self, pdf_path: str) -> List[Tuple[str, int]]:
        """
        주요 섹션들과 그 다음에 나오는 제목들 찾기
        """
        results = []
        found_sections = set()
        found_disease = False
        
        try:
            # PDF 텍스트 추출
            with fitz.open(pdf_path) as doc:
                for page_num in range(1, len(doc) + 1):
                    page = doc[page_num-1]
                    text = page.get_text()
                    
                    # 아직 주요 섹션을 못 찾은 경우
                    if not found_disease:
                        for section, pattern in self.key_patterns.items():
                            if section not in found_sections and re.search(pattern, text, re.IGNORECASE):
                                found_sections.add(section)
                                results.append((f"Found {section}", page_num))
                                logger.info(f"{section} found on page {page_num}")
                                
                                if section == '질병관련':
                                    found_disease = True
                        continue
                    
                    # 질병관련 특약 이후의 제목 찾기
                    title_info = self.title_detector.detect_next_title(text, page_num)
                    if title_info:
                        title, page = title_info
                        results.append((f"Next section: {title}", page))
                        logger.info(f"New section found: {title} on page {page}")
                        
            return results
            
        except Exception as e:
            logger.error(f"Error processing PDF: {str(e)}")
            return []

def main():
    try:
        # PDF 파일 경로
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return
            
        # 섹션 검출
        detector = SectionDetector()
        results = detector.find_sections(pdf_path)
        
        # 결과 출력
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