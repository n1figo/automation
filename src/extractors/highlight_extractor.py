import numpy as np
import cv2
from PIL import Image
import fitz
import logging
from typing import List, Tuple

logger = logging.getLogger(__name__)

class HighlightExtractor:
    """PDF 문서의 하이라이트된 영역을 추출하는 클래스"""
    
    def __init__(self):
        self.saturation_threshold = 30
        self.kernel_size = (5, 5)
        self.logger = logging.getLogger(__name__)

    def pdf_to_image(self, page: fitz.Page) -> np.ndarray:
        """PDF 페이지를 이미지로 변환"""
        try:
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            return np.array(img)
        except Exception as e:
            self.logger.error(f"PDF 페이지 이미지 변환 실패: {e}")
            raise

    def detect_highlights(self, image: np.ndarray) -> List[np.ndarray]:
        """이미지에서 하이라이트된 영역 감지"""
        try:
            # RGB를 HSV로 변환
            hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
            s = hsv[:, :, 1]  # 채도
            v = hsv[:, :, 2]  # 명도

            # 채도 마스크 생성
            saturation_mask = s > self.saturation_threshold

            # Otsu's 방법으로 이진화
            _, binary = cv2.threshold(v, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            combined_mask = cv2.bitwise_and(binary, binary, 
                                          mask=saturation_mask.astype(np.uint8) * 255)

            # 모폴로지 연산으로 노이즈 제거
            kernel = np.ones(self.kernel_size, np.uint8)
            cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
            cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)

            # 윤곽선 찾기
            contours, _ = cv2.findContours(cleaned_mask, 
                                         cv2.RETR_EXTERNAL, 
                                         cv2.CHAIN_APPROX_SIMPLE)
            
            return contours

        except Exception as e:
            self.logger.error(f"하이라이트 감지 실패: {e}")
            raise

    def get_highlight_regions(self, contours: List[np.ndarray], 
                            image_height: int) -> List[Tuple[float, float]]:
        """윤곽선을 PDF 좌표계의 영역으로 변환"""
        try:
            regions = []
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                # 이미지 좌표를 PDF 좌표로 변환
                top = image_height - (y + h)
                bottom = image_height - y
                regions.append((top, bottom))
            return regions

        except Exception as e:
            self.logger.error(f"하이라이트 영역 변환 실패: {e}")
            raise

    def check_highlight(self, row_range: Tuple[float, float], 
                       highlight_regions: List[Tuple[float, float]]) -> bool:
        """행이 하이라이트된 영역과 겹치는지 확인"""
        try:
            row_top, row_bottom = row_range
            for region_top, region_bottom in highlight_regions:
                # 영역 겹침 확인
                if (region_top <= row_top <= region_bottom) or \
                   (region_top <= row_bottom <= region_bottom) or \
                   (row_top <= region_top <= row_bottom) or \
                   (row_top <= region_bottom <= row_bottom):
                    return True
            return False

        except Exception as e:
            self.logger.error(f"하이라이트 확인 실패: {e}")
            return False

    def process_page(self, page: fitz.Page) -> Tuple[List[np.ndarray], List[Tuple[float, float]]]:
        """페이지 전체 처리"""
        try:
            image = self.pdf_to_image(page)
            contours = self.detect_highlights(image)
            regions = self.get_highlight_regions(contours, image.shape[0])
            return contours, regions
            
        except Exception as e:
            self.logger.error(f"페이지 처리 실패: {e}")
            raise