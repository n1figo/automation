import os
import logging
import layoutparser as lp
import fitz
import numpy as np
from PIL import Image
import torch
from typing import Dict, List, Tuple, Optional

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class LayoutAnalyzer:
    def __init__(self):
        """
        LayoutParser 모델 초기화
        """
        # GPU 사용 가능시 cuda 사용
        self.device = 'cuda' if torch.cuda.is_available() else 'cpu'
        
        # LayoutParser 모델 로드
        self.model = lp.Detectron2LayoutModel(
            config_path='lp://PubLayNet/faster_rcnn_R_50_FPN_3x/config',
            label_map={0: "Text", 1: "Title", 2: "List", 3: "Table", 4: "Figure"},
            extra_config=["MODEL.DEVICE", self.device]
        )
        
        logger.info(f"LayoutParser model initialized on {self.device}")

    def convert_pdf_to_images(self, pdf_path: str, page_num: int) -> Image.Image:
        """
        PDF 페이지를 이미지로 변환
        """
        doc = fitz.open(pdf_path)
        page = doc[page_num-1]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 해상도 2배 증가
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def analyze_page_layout(self, pdf_path: str, page_num: int) -> Tuple[list, dict]:
        """
        페이지의 레이아웃을 분석하여 제목과 표 위치 반환
        """
        try:
            # PDF 페이지를 이미지로 변환
            image = self.convert_pdf_to_images(pdf_path, page_num)
            
            # 레이아웃 분석
            layout = self.model.detect(image)
            
            # 제목과 표 영역 분리
            titles = [block for block in layout if block.type == "Title"]
            tables = [block for block in layout if block.type == "Table"]
            
            # 위치 정보를 포함한 결과 반환
            layout_info = {
                'titles': titles,
                'tables': tables,
                'page_size': image.size
            }
            
            return layout, layout_info
            
        except Exception as e:
            logger.error(f"Error in analyzing page layout: {str(e)}")
            return None, None

class SectionDetector:
    def __init__(self):
        self.layout_analyzer = LayoutAnalyzer()
        # 찾고자 하는 섹션 패턴
        self.key_patterns = {
            '상해관련': r'상해관련\s*특약',
            '질병관련': r'질병관련\s*특약'
        }
        
    def is_target_title(self, title_block, title_text: str) -> bool:
        """
        주어진 title block이 찾고자 하는 제목인지 확인
        """
        for section, pattern in self.key_patterns.items():
            if section in title_text:
                return True
        return False

    def extract_text_from_pdf(self, pdf_path: str, page_num: int, block) -> str:
        """
        PDF에서 특정 블록의 텍스트 추출
        """
        doc = fitz.open(pdf_path)
        page = doc[page_num-1]
        
        # 블록의 좌표를 PDF 좌표계로 변환
        x0, y0, x1, y1 = block.coordinates
        rect = fitz.Rect(x0, y0, x1, y1)
        
        # 해당 영역의 텍스트 추출
        text = page.get_text("text", clip=rect)
        return text.strip()
    
    
    def find_sections(self, pdf_path: str) -> List[Tuple[str, int]]:
        """
        주요 섹션들과 그 다음에 나오는 제목들 찾기
        """
        results = []
        found_sections = set()
        found_disease = False
        last_table_bottom = 0
        
        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            
            for page_num in range(1, total_pages + 1):
                logger.info(f"Processing page {page_num}/{total_pages}")
                
                # 레이아웃 분석
                layout, layout_info = self.layout_analyzer.analyze_page_layout(pdf_path, page_num)
                
                if not layout_info:
                    continue
                
                # 페이지의 텍스트 블록들 분석
                for title in layout_info['titles']:
                    title_text = self.extract_text_from_pdf(pdf_path, page_num, title)
                    
                    # 아직 질병관련 특약을 못 찾은 경우
                    if not found_disease:
                        for section, pattern in self.key_patterns.items():
                            if section not in found_sections and section in title_text:
                                found_sections.add(section)
                                results.append((f"Found {section}", page_num))
                                logger.info(f"{section} found on page {page_num}")
                                
                                if section == '질병관련':
                                    found_disease = True
                                    # 질병관련 특약 표의 위치 저장
                                    tables = layout_info['tables']
                                    if tables:
                                        last_table_bottom = max(table.coordinates[3] for table in tables)
                        continue
                    
                    # 질병관련 특약을 찾은 후의 처리
                    if found_disease:
                        # 제목이 표보다 아래에 있는지 확인
                        if title.coordinates[1] > last_table_bottom:
                            # 새로운 섹션의 시작으로 판단되는 제목 발견
                            if len(title_text.strip()) > 5:  # 최소 길이 조건
                                results.append((f"Next section: {title_text}", page_num))
                                logger.info(f"New section found: {title_text} on page {page_num}")
                                return results  # 다음 섹션을 찾으면 종료
                
                # 현재 페이지의 표 위치 업데이트
                tables = layout_info['tables']
                if tables:
                    last_table_bottom = max(table.coordinates[3] for table in tables)
            
            return results
            
        except Exception as e:
            logger.error(f"Error in find_sections: {str(e)}")
            return []

def main():
    try:
        # PDF 파일 경로
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return
        
        # LayoutParser 사용 가능 확인
        try:
            import layoutparser
        except ImportError:
            logger.error("LayoutParser not installed. Installing required packages...")
            os.system("pip install layoutparser")
            os.system("pip install 'detectron2@git+https://github.com/facebookresearch/detectron2.git@v0.6#egg=detectron2'")
        
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

