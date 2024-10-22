
import PyPDF2
import re
import logging
from typing import Dict, List, Tuple

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SimplePageDetector:
    def __init__(self):
        # 찾고자 하는 패턴 정의
        self.target_patterns = [
            # 상해관련 특약
            r'상해관련\s*특약',
            # 질병관련 특약
            r'질병관련\s*특약',
            # 다음 섹션 패턴들
            r'보험료\s*납입면제\s*관련\s*특약',
            r'간병\s*관련\s*특약',
            r'실손의료비\s*보장\s*특약',
            r'기타\s*특약',
            r'제\s*\d+\s*장',
            r'보장내용\s*요약서'
        ]

    def find_pages(self, texts_by_page: Dict[int, str]) -> List[Tuple[int, str]]:
        """
        각 패턴이 처음 등장하는 페이지 번호를 찾음
        """
        results = []
        found_patterns = set()
        
        for page_num in sorted(texts_by_page.keys()):
            text = texts_by_page[page_num]
            
            for pattern in self.target_patterns:
                # 이미 찾은 패턴은 건너뜀
                if pattern in found_patterns:
                    continue
                    
                if re.search(pattern, text, re.IGNORECASE):
                    found_text = re.search(pattern, text, re.IGNORECASE).group()
                    results.append((page_num, found_text))
                    found_patterns.add(pattern)
                    logger.info(f"페이지 {page_num}에서 '{found_text}' 발견")

        return results

def main():
    try:
        # PDF 파일 경로
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        
        # PDF 텍스트 추출
        logger.info("PDF 처리 시작")
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            texts_by_page = {i+1: page.extract_text() for i, page in enumerate(reader.pages)}
        
        logger.info(f"총 {len(texts_by_page)} 페이지 로드됨")
        
        # 페이지 검출
        detector = SimplePageDetector()
        results = detector.find_pages(texts_by_page)
        
        # 결과 출력
        if results:
            print("\n발견된 섹션들:")
            for page_num, text in results:
                print(f"페이지 {page_num}: {text}")
        else:
            print("\n섹션을 찾지 못했습니다.")
            
    except FileNotFoundError:
        logger.error(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
    except Exception as e:
        logger.error(f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()
