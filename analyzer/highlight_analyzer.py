import numpy as np
import cv2
import fitz  # PyMuPDF
from typing import Dict, Any, Union

class HighlightAnalyzer:
    """형광펜 표시, 취소선 등의 강조 표시를 감지하는 클래스"""
    
    def __init__(self):
        """HighlightAnalyzer 초기화"""
        # 색상 범위 설정
        self.color_ranges = {
            "yellow": {
                "lower": np.array([20, 100, 200]),
                "upper": np.array([40, 255, 255])
            },
            "light_yellow": {
                "lower": np.array([20, 30, 200]),
                "upper": np.array([40, 100, 255])
            },
            "light_blue": {
                "lower": np.array([85, 30, 200]),
                "upper": np.array([115, 100, 255])
            },
            "orange": {
                "lower": np.array([5, 100, 200]),
                "upper": np.array([20, 255, 255])
            },
            "green": {
                "lower": np.array([40, 100, 200]),
                "upper": np.array([80, 255, 255])
            },
            "purple": {
                "lower": np.array([130, 50, 200]),
                "upper": np.array([170, 255, 255])
            },
            "gray": {
                "lower": np.array([0, 0, 80]),
                "upper": np.array([180, 40, 200])
            }
        }
    
    def analyze_document_colors(self, hsv_img: np.ndarray) -> Dict[str, float]:
        """
        문서 특성에 따른 적응형 색상 분석
        
        Args:
            hsv_img: HSV 색상 모델로 변환된 이미지
            
        Returns:
            Dict[str, float]: 문서 특성에 따른 임계값
        """
        # 이미지 히스토그램 분석
        h_hist = cv2.calcHist([hsv_img], [0], None, [180], [0, 180])
        s_hist = cv2.calcHist([hsv_img], [1], None, [256], [0, 256])
        v_hist = cv2.calcHist([hsv_img], [2], None, [256], [0, 256])
        
        # 주요 색상 식별
        h_peaks = [i for i in range(1, 179) if h_hist[i] > h_hist[i-1] and h_hist[i] > h_hist[i+1]]
        
        # 문서 특성 분석 (밝기, 채도 분포)
        avg_v = np.average(np.arange(256), weights=v_hist.flatten())
        avg_s = np.average(np.arange(256), weights=s_hist.flatten())
        
        # 문서 특성에 맞게 임계값 조정
        thresholds = {
            "bright": 0.05 if avg_v > 200 else 0.1,  # 밝은 문서는 낮은 임계값
            "dark": 0.15 if avg_v < 150 else 0.1,    # 어두운 문서는 높은 임계값
            "colorful": 0.08 if avg_s > 100 else 0.1  # 채도가 높은 문서
        }
        
        # 적응형 임계값 반환
        return {
            "base_threshold": thresholds["bright"] if avg_v > 200 else 
                             thresholds["dark"] if avg_v < 150 else 0.1,
            "yellow_threshold": thresholds["bright"] * 0.8,  # 노란색은 더 낮은 임계값
            "gray_threshold": thresholds["dark"] * 1.2       # 회색은 더 높은 임계값
        }
    
    def detect_highlights(self, pdf_document, page_num: int) -> Dict[str, Any]:
        """
        강조색 감지 함수
        
        Args:
            pdf_document: PyMuPDF 문서 객체
            page_num: 페이지 번호 (0-based)
            
        Returns:
            Dict[str, Any]: 강조색 감지 결과
        """
        try:
            # 페이지 렌더링 (고해상도로)
            zoom = 2.0
            mat = fitz.Matrix(zoom, zoom)
            page = pdf_document[page_num]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # OpenCV 형식으로 변환
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            img_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            
            # HSV 색상 모델로 변환
            hsv = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2HSV)
            
            # 문서 특성에 맞는 임계값 계산
            thresholds = self.analyze_document_colors(hsv)
            
            # 색상별 마스크 생성
            masks = {}
            for color_name, ranges in self.color_ranges.items():
                mask = cv2.inRange(hsv, ranges["lower"], ranges["upper"])
                kernel = np.ones((3, 3), np.uint8)
                mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
                masks[color_name] = mask
            
            # 강조색 마스크 결합
            highlight_mask = masks["yellow"] | masks["light_yellow"] | masks["light_blue"] | masks["orange"] | masks["green"] | masks["purple"]
            
            # 색상별 검출 영역 비율 계산
            total_pixels = hsv.shape[0] * hsv.shape[1]
            
            color_ratios = {}
            for color_name, mask in masks.items():
                color_ratios[color_name] = np.sum(mask > 0) / total_pixels * 100
            
            # 일정 비율 이상이면 강조색이 있다고 판단
            highlight_threshold = thresholds["base_threshold"] * 100  # 백분율로 변환
            
            # 텍스트 위치 정보 확인 (취소선)
            has_strikethrough = False
            text_blocks = page.get_text("dict")["blocks"]
            for block in text_blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            if span.get("flags", 0) & 2**6:  # 취소선 비트 확인 (64)
                                has_strikethrough = True
                                break
                        if has_strikethrough:
                            break
                    if has_strikethrough:
                        break
            
            result = {
                "has_highlight": np.sum(highlight_mask > 0) / total_pixels * 100 > highlight_threshold,
                "has_gray": color_ratios["gray"] > thresholds["gray_threshold"] * 100,
                "has_strikethrough": has_strikethrough,
                "colors": {color: ratio > highlight_threshold for color, ratio in color_ratios.items()},
                "color_ratios": color_ratios
            }
            
            return result
            
        except Exception as e:
            print(f"강조색 감지 중 오류: {str(e)}")
            return {
                "has_highlight": False,
                "has_gray": False,
                "has_strikethrough": False,
                "colors": {},
                "color_ratios": {}
            }