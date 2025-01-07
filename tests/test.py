import PyPDF2
import re
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)

def find_section_pages(pdf_path: str) -> list:
    """
    PDF에서 '상해관련 특약' 섹션이 시작하는 페이지들을 찾습니다.
    """
    section_pages = []
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
            logger.info(f"PDF 총 페이지 수: {total_pages}")
            
            # 각 페이지 검사
            for page_num in range(total_pages):
                text = reader.pages[page_num].extract_text()
                
                # 상해관련 특약 섹션 찾기
                if "◇상해관련 특약" in text:
                    section_pages.append(page_num + 1)
                    logger.info(f"'상해관련 특약' 섹션 발견: 페이지 {page_num + 1}")
                    
                    # 섹션 내용 미리보기
                    preview = text[:200].replace('\n', ' ')
                    logger.info(f"섹션 미리보기: {preview}...")
        
        return section_pages
        
    except Exception as e:
        logger.error(f"PDF 처리 중 오류 발생: {str(e)}")
        return []

def main():
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    
    section_pages = find_section_pages(pdf_path)
    
    if section_pages:
        print("\n=== 상해관련 특약 섹션 위치 ===")
        for page in section_pages:
            print(f"페이지 {page}")
    else:
        print("\n상해관련 특약 섹션을 찾을 수 없습니다.")

if __name__ == "__main__":
    main()