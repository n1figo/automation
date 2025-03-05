import fitz
import camelot
import pandas as pd
from pathlib import Path
import re
from datetime import datetime
import logging
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import cv2
import numpy as np
from PIL import Image
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
from src.utils.excel_writer import ExcelWriter
from src.config.extraction_config import TableExtractionConfig
import logging, sys, os
import xml.etree.ElementTree as ET
from xml.dom import minidom
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

import fitz
import camelot
import pandas as pd
import numpy as np
import cv2
from pathlib import Path
import re
from datetime import datetime
import logging
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
from sklearn.cluster import DBSCAN
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import sys

@dataclass
class ParsingRange:
    start_page: int
    end_page: int
    section_type: str
    insurance_type: Optional[str] = None

class ImageProcessor:
    def __init__(self, dpi=200):
        self.dpi = dpi
        # HSV 색상 범위 설정
        self.color_ranges = {
            'red': [
                ((0, 120, 70), (10, 255, 255)),
                ((160, 120, 70), (180, 255, 255))
            ],
            'yellow': [((20, 100, 100), (30, 255, 255))],
            'blue': [((100, 120, 70), (130, 255, 255))]
        }

    def pdf_page_to_image(self, page: fitz.Page) -> np.ndarray:
        """PDF 페이지를 OpenCV 이미지로 변환"""
        pix = page.get_pixmap(matrix=fitz.Matrix(self.dpi/72, self.dpi/72))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    def detect_color_regions(self, img: np.ndarray, color: str) -> np.ndarray:
        """특정 색상 영역 검출"""
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        masks = []
        
        for lower, upper in self.color_ranges[color]:
            mask = cv2.inRange(hsv, np.array(lower), np.array(upper))
            masks.append(mask)
        
        return cv2.bitwise_or(*masks) if masks else np.zeros_like(img[:,:,0])

    def find_colored_text_blocks(self, page: fitz.Page, color: str = 'red') -> List[Dict]:
        """색상 텍스트 블록 찾기"""
        img = self.pdf_page_to_image(page)
        color_mask = self.detect_color_regions(img, color)
        
        text_blocks = page.get_text("dict")["blocks"]
        colored_texts = []
        
        for block in text_blocks:
            if 'lines' not in block:
                continue
                
            x0, y0, x1, y1 = map(int, block["bbox"])
            block_mask = color_mask[y0:y1, x0:x1]
            
            # 색상 픽셀 비율 계산
            color_ratio = np.count_nonzero(block_mask) / (block_mask.size + 1e-6)
            
            if color_ratio > 0.3:
                colored_texts.append({
                    "text": "".join([span["text"] for line in block["lines"] 
                                   for span in line["spans"]]),
                    "confidence": color_ratio,
                    "bbox": block["bbox"],
                    "color": color
                })
        
        return colored_texts

logger = logging.getLogger(__name__)




class TitleMatcher:
    def __init__(self):
        try:
            # 현재 실행 파일의 디렉토리 얻기
            if getattr(sys, 'frozen', False):
                current_dir = os.path.dirname(sys.executable)
            else:
                current_dir = os.path.dirname(os.path.abspath(__file__))
                # src/analyzers에서 상위 디렉토리로 이동
                current_dir = os.path.dirname(os.path.dirname(current_dir))
            
            # models 폴더 경로 설정
            model_path = os.path.join(current_dir, "models", "distiluse-base-multilingual-cased-v1")
            
            logger.info(f"모델 경로: {model_path}")
            
            # 모델 파일 존재 확인
            model_file = os.path.join(model_path, "pytorch_model.bin")
            if not os.path.exists(model_file):
                raise FileNotFoundError(f"모델 파일을 찾을 수 없습니다: {model_file}")
            
            # 오프라인 모드로 모델 로드
            self.model = SentenceTransformer(model_path, device='cpu')
            logger.info("로컬 모델 로드 성공")
            
            # 섹션 패턴 정의
            self.section_patterns = {
                '상해': [
                    r'상해[\s]*관련[\s]*특약',
                    r'상해[\s]*특별약관',
                    r'상해[\s]*담보',
                    r'상해[\s]*보장'
                ],
                '질병': [
                    r'질병[\s]*관련[\s]*특약',
                    r'질병[\s]*특별약관',
                    r'질병[\s]*담보',
                    r'질병[\s]*보장'
                ],
                '상해및질병': [
                    r'상해[\s]*및[\s]*질병[\s]*관련[\s]*특약',
                    r'상해[\s]*및[\s]*질병[\s]*특별약관',
                    r'상해[\s]*및[\s]*질병[\s]*담보',
                    r'상해[\s]*및[\s]*질병[\s]*보장',
                    r'상해와[\s]*질병[\s]*관련[\s]*특약'
                ]
            }
            
            self.similarity_threshold = 0.7
            self.max_distance = 50
            
        except Exception as e:
            error_msg = f"TitleMatcher 초기화 오류: {str(e)}"
            logger.error(error_msg)
            raise

    def get_embedding(self, text: str) -> np.ndarray:
        """텍스트의 임베딩 벡터를 반환"""
        try:
            return self.model.encode(text, show_progress_bar=False)
        except Exception as e:
            logger.error(f"임베딩 생성 실패: {str(e)}")
            raise

    def calculate_similarity(self, text1: str, text2: str) -> float:
        """두 텍스트 간의 유사도 계산"""
        try:
            emb1 = self.get_embedding(text1)
            emb2 = self.get_embedding(text2)
            return 1 - cosine(emb1, emb2)
        except Exception as e:
            logger.error(f"유사도 계산 실패: {str(e)}")
            raise
    
    def get_section_type(self, text: str) -> str:
        """텍스트의 섹션 유형 판단"""
        text = text.strip()
        
        # 패턴 매칭으로 먼저 시도
        for section_type, patterns in self.section_patterns.items():
            if any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns):
                logger.info(f"섹션 유형 감지: {section_type} (텍스트: {text})")
                return section_type

        # 유사도 기반 매칭 시도
        try:
            reference_texts = {
                '상해': '상해보장특약',
                '질병': '질병보장특약',
                '상해및질병': '상해및질병보장특약'
            }
            
            similarities = {}
            for section_type, ref_text in reference_texts.items():
                similarity = self.calculate_similarity(text, ref_text)
                similarities[section_type] = similarity
                logger.debug(f"유사도 ({section_type}): {similarity:.3f}")
            
            best_match = max(similarities.items(), key=lambda x: x[1])
            if best_match[1] > self.similarity_threshold:
                logger.info(f"섹션 유형 감지(유사도): {best_match[0]} (점수: {best_match[1]:.3f}, 텍스트: {text})")
                return best_match[0]

        except Exception as e:
            logger.warning(f"유사도 계산 중 오류 발생: {str(e)}")
        
        logger.warning(f"섹션 유형 판단 실패 (텍스트: {text})")
        return '기타'


class PDFAnalyzer:
    def __init__(self, pdf_path: str, logger=None):
        # 결과 저장 경로 설정
        self.output_dir = Path('output')
        self.xml_dir = self.output_dir / 'xml'
        self.excel_dir = self.output_dir / 'excel'
        
        # 디렉토리 생성
        self.xml_dir.mkdir(parents=True, exist_ok=True)
        self.excel_dir.mkdir(parents=True, exist_ok=True)

        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.logger = logger or logging.getLogger(__name__)
        self.image_processor = ImageProcessor()
        self.title_matcher = TitleMatcher()

        for dir_path in [self.xml_dir, self.excel_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)
        
        self.markers = {
            'payment_section':  r'나.?\s*보험금\s*',
            'insurance_types': r'\[(\d)종\]',
            'sections': {
                '상해관련 특별약관': r'.*상해.*특약|.*상해.*특별약관',
                '질병관련 특별약관': r'.*질병.*특약|.*질병.*특별약관',
                '상해및질병관련 특별약관': r'.*상해.*질병.*특약|.*상해.*질병.*특별약관'
                # '상해관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
                # '질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
                # '상해및질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)?'
            }
        }

    def detect_colored_text_enhanced(self, page: fitz.Page) -> List[Dict]:
        """향상된 색상 텍스트 검출"""
        # PyMuPDF 기반 검출
        original_results = []
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if 'color' in span:
                        color = span["color"]
                        r = (color >> 16) & 0xff
                        g = (color >> 8) & 0xff
                        b = color & 0xff
                        
                        if r > 200 and g < 100 and b < 100:  # 빨간색 조건
                            original_results.append({
                                "text": span["text"].strip(), # 
                                "color": (r, g, b),
                                "bbox": span["bbox"]
                            })

        # OpenCV 기반 검출
        cv_results = self.image_processor.find_colored_text_blocks(page)
        
        # 결과 통합
        combined = []
        seen_texts = set()
        
        for result in original_results + cv_results:
            text = result["text"].strip()
            if text and text not in seen_texts:
                seen_texts.add(text)
                combined.append(result)

        return combined


    def process_table(self, table: pd.DataFrame, page: fitz.Page) -> Tuple[pd.DataFrame, List[List[bool]]]:
        """표 데이터 처리"""
        colored_texts = self.detect_colored_text_enhanced(page)
        highlights = []
        df = table.copy()
        df['변경사항'] = ''  # 초기값 빈 문자열
        
        for idx, row in df.iterrows():
            row_highlights = []
            has_changes = False
            
            for col_idx, cell in enumerate(row):
                cell_text = str(cell).strip()
                is_highlighted = any(
                    colored['text'] in cell_text
                    for colored in colored_texts
                )
                
                row_highlights.append(is_highlighted)
                if is_highlighted:
                    has_changes = True
            
            highlights.append(row_highlights)
            if has_changes:
                df.at[idx, '변경사항'] = '변경사항 있음'
        
        return df, highlights

    def create_xml(self, tables: List[pd.DataFrame], page_numbers: List[int]) -> ET.Element:
        """테이블을 XML로 변환"""
        root = ET.Element('document')
        root.set('created', datetime.now().isoformat())
        root.set('filename', Path(self.pdf_path).name)
        
        for idx, (df, page_num) in enumerate(zip(tables, page_numbers)):
            table_elem = ET.SubElement(root, 'table')
            table_elem.set('id', str(idx + 1))
            table_elem.set('page', str(page_num))
            
            # 헤더 추가
            header = ET.SubElement(table_elem, 'header')
            for col in df.columns:
                col_elem = ET.SubElement(header, 'column')
                col_elem.text = str(col)
            
            # 데이터 행 추가
            rows = ET.SubElement(table_elem, 'rows')
            for _, row in df.iterrows():
                row_elem = ET.SubElement(rows, 'row')
                for col, value in row.items():
                    cell = ET.SubElement(row_elem, 'cell')
                    cell.set('column', str(col))
                    cell.text = str(value) if pd.notna(value) else ''
        
        return root

    def _identify_section_type(self, table_df: pd.DataFrame, page_num: int, sections_info: dict) -> str:
        """표의 섹션 유형 식별"""
        try:
            # 페이지 번호에 따른 섹션 결정
            if page_num < sections_info['질병관련 특별약관']:  
                return '상해'
            elif page_num < sections_info['상해및질병관련 특별약관']:
                return '질병'
            else:
                return '상해및질병'
                
        except Exception as e:
            self.logger.error(f"섹션 유형 식별 중 오류: {str(e)}")
            return '기타'

    def _process_table_with_highlights(self, table, highlight_regions, page_height, 
                                   page_num, parsing_range, sections_info=None) -> pd.DataFrame:
        try:
            # sections_info 대신 self.section_pages 사용
            df = self.clean_table(table.df)
            if df.empty:
                return df

            # 페이지 번호에 따른 구분 설정
            질병_시작 = self.section_pages.get('질병관련 특별약관', float('inf'))
            상해및질병_시작 = self.section_pages.get('상해및질병관련 특별약관', float('inf'))

            if page_num < 질병_시작:
                section_type = '상해'
            elif page_num < 상해및질병_시작:
                section_type = '질병'
            else:
                section_type = '상해및질병'

            # 메타데이터 추가
            df['페이지'] = page_num + 1
            df['구분'] = section_type
            if parsing_range.insurance_type:
                df['보험종류'] = parsing_range.insurance_type

            return df

        except Exception as e:
            self.logger.error(f"표 처리 중 오류: {str(e)}")
            return pd.DataFrame()

    # ... [나머지 메서드들은 동일하게 유지] ...
    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """추출된 표 데이터 정제"""
        try:
            # 기본 정제
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # 원본 컬럼 매핑 (컬럼 수에 따라 적절히 매핑)
            if len(df.columns) >= 5:
                df.columns = ['구분', '담보명', '지급사유', '지급금액', '비고'] + list(df.columns[5:])
            elif len(df.columns) >= 4:
                df.columns = ['구분', '담보명', '지급사유', '지급금액']
            elif len(df.columns) >= 3:
                df.columns = ['담보명', '지급사유', '지급금액']
            
            # 컬럼 정제
            for col in df.columns:
                if df[col].dtype == "object":
                    # 기본 정제
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].str.replace('\n', ' ')
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                    
                    # 특수문자 정제
                    df[col] = df[col].str.replace(r'[^\w\s가-힣.\(\)]', '', regex=True)
                    
                    # 빈 문자열을 None으로 변환
                    df[col] = df[col].replace('', None)
                    df[col] = df[col].replace('nan', None)
                    df[col] = df[col].replace('None', None)
            
            # 필터링 패턴 (헤더나 불필요한 행 제거)
            filter_patterns = [
                r'보\s*장\s*명\s*지\s*급\s*사\s*유\s*지\s*급\s*금\s*액'
            ]
            
            # 필터링 적용
            for pattern in filter_patterns:
                df = df[~df.apply(lambda row: any(pattern in str(cell) for cell in row), axis=1)]
            
            # 모든 값이 None인 행 제거
            df = df.dropna(how='all')
            
            # 중복 행 제거
            df = df.drop_duplicates()
            
            # None을 빈 문자열로 변환
            df = df.fillna('')
            
            return df
            
        except Exception as e:
            self.logger.error(f"표 정제 중 오류: {str(e)}")
            return df
        
    def determine_parsing_mode(self) -> str:
        """
        파싱 모드 선택 메서드
        
        Return:
        - 'alternative': 선택특약 모드
        - 'default': 기본 모드(상해/질병/상해및질병 특별약관)
        - 'custom': 사용자 지정 페이지에서 파싱
        """
        while True:
            print("\n=== PDF 파싱 모드 선택 ===")
            print("1. 선택특약 모드로 파싱")
            print("2. 기본 모드(상해/질병/상해및질병 특별약관) 파싱")
            print("3. 사용자 지정 페이지에서 파싱")  # 추가된 옵션
            print("4. 테스트 1페이지 모드 (4페이지 파싱)")  # 새 옵션 추가
            
            try:
                choice = input("모드를 선택해주세요 (1, 2, 3 또는 4): ").strip()
                
                if choice == '1':
                    print("\n선택특약 모드로 파싱을 진행합니다.")
                    return 'alternative'
                elif choice == '2':
                    print("\n기본 모드로 파싱을 진행합니다.")
                    return 'default'
                elif choice == '3':  # 사용자 지정 페이지 모드
                    print("\n사용자 지정 페이지에서 파싱을 진행합니다.")
                    return 'custom'
                if choice == '4':  # 새 모드
                    print("\n테스트 모드로 4페이지를 파싱합니다.")
                    return 'test_page_4'
                else:
                    print("잘못된 입력입니다. 1, 2, 3 또는 4을 입력해주세요.")
            
            except Exception as e:
                print(f"입력 중 오류가 발생했습니다: {e}")

    def determine_parsing_ranges(self) -> List[ParsingRange]:
        """사용자 선택에 따른 파싱 범위 결정"""
        try:
            ranges = []
            payment_patterns = [r'나.?\s*보험금']

            # 1. 보험금 지급 섹션 시작점 검색
            payment_start = None
            for page_num in range(len(self.doc)):
                text = self.doc[page_num].get_text()
                if any(re.search(pattern, text, re.IGNORECASE) for pattern in payment_patterns):
                    payment_start = page_num
                    self.logger.info(f"보험금 지급 섹션 시작 페이지: {page_num+1}")
                    break

            # 2. 모드 선택 UI
            while True:
                print("\n" + "="*30)
                print("PDF 파싱 모드 선택")
                print("1. 기본 모드 (상해/질병/상해및질병 특별약관)")
                print("2. 선택특약 모드")
                print("3. 사용자 지정 범위 파싱")
                print("4. 테스트 1페이지 모드")
                print("="*30)
                
                choice = input("\n모드를 선택하세요 (1-4): ").strip()
                
                # 4. 테스트 1페이지 모드
                if choice == '4':
                    while True:
                        try:
                            input_page = input("\n파싱할 페이지 번호를 입력하세요 (예: 4): ").strip()
                            if not input_page.isdigit():
                                raise ValueError("숫자만 입력 가능합니다")
                                
                            page_num = int(input_page) - 1  # 1-based → 0-based 변환
                            
                            if 0 <= page_num < len(self.doc):
                                self.logger.info(f"테스트 모드 활성화: {input_page}페이지 파싱")
                                return [ParsingRange(
                                    start_page=page_num,
                                    end_page=page_num,
                                    section_type=f'테스트_{input_page}페이지'
                                )]
                            else:
                                print(f"※ 문서는 총 {len(self.doc)}페이지입니다. 유효한 페이지를 입력하세요.")
                        except ValueError as e:
                            print(f"※ 오류: {str(e)}")
                    continue
                
                # 1. 기본 모드
                elif choice == '1':
                    self.logger.info("기본 모드 선택")
                    return self._parse_default_sections()
                
                # 2. 선택특약 모드
                elif choice == '2':
                    self.logger.info("선택특약 모드 선택")
                    return self._parse_alternative_sections()
                
                # 3. 사용자 지정 모드
                elif choice == '3':
                    try:
                        start = int(input("시작 페이지 번호: ")) - 1
                        end = int(input("종료 페이지 번호: ")) - 1
                        
                        if start < 0 or end >= len(self.doc):
                            print(f"※ 유효한 페이지 범위: 1 ~ {len(self.doc)}")
                            continue
                            
                        if start > end:
                            print("※ 시작 페이지는 종료 페이지보다 작아야 합니다")
                            continue
                            
                        self.logger.info(f"사용자 지정 범위: {start+1}~{end+1}페이지")
                        return [ParsingRange(
                            start_page=start,
                            end_page=end,
                            section_type='사용자_지정'
                        )]
                    except ValueError:
                        print("※ 숫자만 입력 가능합니다")
                    continue
                
                else:
                    print("※ 1~4 사이의 숫자만 입력 가능합니다")
                    continue

        except Exception as e:
            self.logger.error(f"파싱 범위 결정 실패: {str(e)}")
            return [ParsingRange(
                start_page=0,
                end_page=len(self.doc)-1,
                section_type='전체_문서'
            )]

    def _parse_default_sections(self) -> List[ParsingRange]:
        """기본 모드 섹션 파싱"""
        try:
            ranges = []
            self.section_pages = {}

            # 보험금 지급 섹션 찾기
            payment_start = self.find_payment_section()  # 변경된 부분
            if payment_start is None:
                payment_start = 0
            
            self.logger.info(f"보험금 지급 섹션 시작: {payment_start + 1}페이지")

            # 특정 섹션(상해/질병/상해및질병 특별약관) 찾기
            section_pages = {}
            for page_num in range(payment_start, len(self.doc)):
                page_text = self.doc[page_num].get_text()
                
                # TitleMatcher 활용하여 섹션 유형 판단
                section_type = self.title_matcher.get_section_type(page_text)
                
                if section_type in ['상해', '질병', '상해및질병']:
                    section_name = f"{section_type}관련 특별약관"
                    section_pages[section_name] = page_num
                    print(f"발견: {section_name:<30} - {page_num + 1}페이지")

            # 섹션 페이지들 중 최소/최대 페이지 찾기
            if section_pages:
                start_page = payment_start  # 보험금 지급 섹션부터 시작
                try:
                    end_page = min(
                        page for section, page in section_pages.items() 
                        if '상해및질병관련 특별약관' in section
                    ) - 1
                except ValueError:
                    end_page = len(self.doc) - 1
                
                print("\n=== 섹션 파싱 범위 ===")
                print(f"시작: {start_page + 1}페이지")
                print(f"종료: {end_page + 1}페이지")
                
                ranges.append(ParsingRange(
                    start_page=start_page,
                    end_page=end_page,
                    section_type='섹션포함문서'
                ))
            
            # ranges가 비어있는 경우 전체 문서 파싱
            if not ranges:
                print("\n=== 전체 문서 파싱 ===")
                start_page = payment_start  # 보험금 지급 섹션부터 시작
                end_page = len(self.doc) - 1
                
                print(f"시작: {start_page + 1}페이지")
                print(f"종료: {end_page + 1}페이지")
                
                ranges.append(ParsingRange(
                    start_page=start_page,
                    end_page=end_page,
                    section_type='전체문서'
                ))

            # 결과 요약
            print("\n=== 최종 파싱 범위 요약 ===")
            for r in ranges:
                info = f"페이지 {r.start_page + 1} ~ {r.end_page + 1}"
                if r.insurance_type:
                    info = f"{r.insurance_type}: {info}"
                print(f"- {info}")
            print("="*50)

            return ranges

        except Exception as e:
            self.logger.error(f"기본 모드 파싱 중 오류: {str(e)}")
            # 오류 발생시 전체 문서를 하나의 범위로 설정
            return [ParsingRange(
                start_page=0,
                end_page=len(self.doc) - 1,
                section_type='전체'
            )]

    def _parse_alternative_sections(self) -> List[ParsingRange]:
        """대체 섹션(선택특약) 파싱 메서드"""
        ranges = []
        
        # 대체 섹션 찾기 (선택특약, 간병인 특별약관 등)
        alternative_start_pattern = r'선택계약|선택\s*계약'
        alternative_end_pattern = r'간병인\s*입원일당\s*관련\s*독립특별약관'
        
        alt_start_page = None
        alt_end_page = None

        # 보험금 지급 섹션 찾기
        payment_sections = self._find_payment_sections()
        min_payment_page = min(payment_sections) if payment_sections else 0

        # 보험금 지급 섹션의 최소 페이지부터 검색 시작
        for page_num in range(min_payment_page, len(self.doc)):
            text = self.doc[page_num].get_text()
            
            # 시작 페이지 찾기
            if not alt_start_page and re.search(alternative_start_pattern, text):
                alt_start_page = page_num
                print(f"대체 시작 페이지 발견: {page_num + 1}페이지")
            
            # 종료 페이지 찾기
            if re.search(alternative_end_pattern, text):
                alt_end_page = page_num
                print(f"대체 종료 페이지 발견: {page_num + 1}페이지")
            
            # 시작과 종료 페이지를 모두 찾았으면 반복 종료
            if alt_start_page is not None and alt_end_page is not None:
                break

        # 대체 섹션 범위 설정
        if alt_start_page is not None and alt_end_page is not None:
            print("\n=== 대체 파싱 범위 ===")
            start_page = alt_start_page
            end_page = alt_end_page
            
            print(f"시작: {start_page + 1}페이지 (선택특약)")
            print(f"종료: {end_page + 1}페이지 (간병인 입원일당 관련 독립특별약관)")
            
            ranges.append(ParsingRange(
                start_page=start_page,
                end_page=end_page,
                section_type='선택특약'
            ))
        else:
            # 대체 섹션을 찾지 못한 경우 전체 문서 파싱
            print("\n=== 대체 섹션을 찾지 못해 전체 문서 파싱 ===")
            start_page = 0
            end_page = len(self.doc) - 1
            
            print(f"시작: {start_page + 1}페이지")
            print(f"종료: {end_page + 1}페이지")
            
            ranges.append(ParsingRange(
                start_page=start_page,
                end_page=end_page,
                section_type='전체문서'
            ))

        # 결과 요약
        print("\n=== 최종 파싱 범위 요약 ===")
        for r in ranges:
            info = f"페이지 {r.start_page + 1} ~ {r.end_page + 1}"
            if r.insurance_type:
                info = f"{r.insurance_type}: {info}"
            print(f"- {info}")
        print("="*50)

        return ranges
    

    def find_payment_section(self) -> Optional[int]:
        """보험금 지급 섹션의 시작 페이지 찾기"""
        payment_patterns = [
            r'나.?\s*보험금\s*'
        ]
        
        self.logger.info("보험금 지급 섹션 검색 시작")
        
        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            for pattern in payment_patterns:
                if re.search(pattern, text, re.IGNORECASE):
                    self.logger.info(f"보험금 지급 섹션 발견: {page_num + 1}페이지")
                    return page_num
                    
        # 섹션을 찾지 못한 경우 첫 페이지부터 시작
        self.logger.warning("보험금 지급 섹션을 찾지 못해 첫 페이지부터 시작합니다.")
        return 0
    

    def _analyze_pdf_structure(self):
        """PDF 구조 상세 분석"""
        print("\n=== PDF 구조 분석 ===")
        self.logger.info("PDF 구조 분석 시작")
        
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            
            # 페이지의 블록 정보 추출
            try:
                blocks = page.get_text("dict")['blocks']
                
                print(f"\n페이지 {page_num + 1} - 블록 수: {len(blocks)}")
                self.logger.info(f"페이지 {page_num + 1} 블록 분석")
                
                for block_index, block in enumerate(blocks, 1):
                    if block['type'] == 0:  # 텍스트 블록
                        block_text = ''
                        for line in block['lines']:
                            line_text = ''.join([span['text'] for span in line['spans']])
                            block_text += line_text + '\n'
                        
                        # 블록 길이가 너무 길면 자르기
                        block_text = block_text[:500] + '...' if len(block_text) > 500 else block_text
                        
                        print(f"  블록 {block_index}:")
                        print(f"  내용 (일부): {block_text}")
            
            except Exception as e:
                print(f"페이지 {page_num + 1} 분석 중 오류: {e}")
                self.logger.error(f"페이지 {page_num + 1} 분석 중 오류: {e}")


    def analyze(self) -> dict:
        """PDF 분석 실행"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        result = {'clean': '', 'table_count': 0}
        
        try:
            ranges = self.determine_parsing_ranges()
            if not ranges:
                return result

            all_tables = []
            all_highlights = []
            all_changes = []
            
            for parsing_range in ranges:
                tables = self.extract_tables(parsing_range)
                for table in tables:
                    page_num = parsing_range.start_page
                    page = self.doc[page_num]
                    
                    # 표 처리 및 하이라이트 정보 추출
                    processed_df, highlights = self.process_table(table, page)
                    
                    # 변경 유형 감지
                    changes = []
                    for row_highlights in highlights:
                        row_changes = []
                        for is_highlighted in row_highlights:
                            if is_highlighted:
                                # 여기서 변경 유형 판별 로직 추가 가능
                                # 예: 색상이나 다른 특징으로 added/deleted/modified 판별
                                row_changes.append('modified')
                            else:
                                row_changes.append('')
                        changes.append(row_changes)
                    
                    # 메타데이터 추가
                    processed_df = self._add_metadata(processed_df, page_num, parsing_range)
                    
                    all_tables.append(processed_df)
                    all_highlights.append(highlights)
                    all_changes.append(changes)

            if all_tables:
                # ExcelWriter를 사용하여 결과 저장
                output_path = self.output_dir / f"보험약관분석_{timestamp}.xlsx"
                excel_writer = ExcelWriter(str(output_path))
                
                # 섹션별로 데이터 작성
                current_section = None
                for idx, (df, highlights, changes) in enumerate(zip(all_tables, all_highlights, all_changes)):
                    # 섹션 변경 확인
                    section = df['구분'].iloc[0] if '구분' in df.columns else f'Section_{idx+1}'
                    if section != current_section:
                        excel_writer.write_section_header(section)
                        current_section = section
                    
                    # 메타데이터 준비
                    metadata = {
                        '페이지': df['페이지'].iloc[0] if '페이지' in df.columns else '',
                        '보험종류': df['보험종류'].iloc[0] if '보험종류' in df.columns else ''
                    }
                    
                    # 표 작성
                    excel_writer.write_table(
                        df=df,
                        highlights=highlights,
                        change_types=changes,
                        metadata=metadata
                    )
                
                excel_writer.save()
                result['clean'] = str(output_path)
                result['table_count'] = len(all_tables)

            return result

        except Exception as e:
            self.logger.error(f"분석 실패: {str(e)}", exc_info=True)
            return result
        finally:
            if hasattr(self, 'doc'):
                self.doc.close()


    def _generate_highlight_matrix(self, df: pd.DataFrame, colored_texts: list) -> list:
        """HWP 변환 문서의 테이블 구조에 맞춘 하이라이트 매트릭스 생성"""
        highlight_matrix = []
        for _, row in df.iterrows():
            row_highlights = []
            for cell in row:
                is_highlighted = any(
                    colored['text'].strip() in str(cell) 
                    for colored in colored_texts
                )
                row_highlights.append(is_highlighted)
            highlight_matrix.append(row_highlights)
        return highlight_matrix

    def _add_metadata(self, df: pd.DataFrame, page_num: int, parsing_range: ParsingRange) -> pd.DataFrame:
        """HWP 변환 문서의 메타데이터 추가"""
        df.insert(0, '추출페이지', page_num + 1)
        df.insert(1, '문서구분', parsing_range.section_type)
        
        if parsing_range.insurance_type:
            df.insert(2, '보험종류', parsing_range.insurance_type)
            
        return df

    def find_insurance_types(self) -> List[Tuple[int, str]]:
        """종별([1종], [2종] 등) 마커가 있는 페이지 찾기"""
        type_pages = []
        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            matches = re.finditer(self.markers['insurance_types'], text)
            for match in matches:
                type_num = match.group(1)
                type_pages.append((page_num, f"[{type_num}종]"))
            
        # 종 번호와 페이지 번호로 정렬
        sorted_pages = sorted(type_pages, 
                            key=lambda x: (int(re.search(r'\[(\d+)종\]', x[1]).group(1)), x[0]))
        
        if sorted_pages:
            self.logger.info(f"발견된 보험종류: {[t[1] for t in sorted_pages]}")
        
        return sorted_pages
    
    
    # pdf_analyzer.py 내 extract_tables() 수정
    # extract_tables() 메서드 수정
    def extract_tables(self, parsing_range: ParsingRange) -> List[pd.DataFrame]:
        """특정 페이지 범위에서 표 추출"""
        tables = []
        for page_num in range(parsing_range.start_page, parsing_range.end_page + 1):
            try:
                page_tables = camelot.read_pdf(
                    self.pdf_path,
                    pages=str(page_num + 1),
                    flavor='lattice',
                    **TableExtractionConfig.get_lattice_config()
                )
                
                for table in page_tables:
                    # Camelot 테이블을 DataFrame으로 변환
                    df = pd.DataFrame(table.data)
                    if not df.empty:
                        tables.append(df)
                        
            except Exception as e:
                self.logger.error(f"페이지 {page_num+1} 표 추출 실패: {str(e)}")
                continue
                
        return tables
    
    def find_payment_sections(self):
        payment_sections = []
        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            for pattern in self.markers['payment_section']:
                if re.search(pattern, text):
                    payment_sections.append(page_num)
                    break
        return payment_sections
    

    def map_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """컬럼 매핑 및 데이터 타입 변환"""
        try:
            # 컬럼 매핑
            column_mapping = {
                '구분': ['구분', '종류', '급부종류'],
                '담보명': ['담보명', '보장명', '급부명'],
                '지급사유': ['지급사유', '보장사유', '급부사유'],
                '지급금액': ['지급금액', '보험금', '보장금액']
            }
            
            # 컬럼명 처리
            df.columns = [col.strip() for col in df.columns]
            
            # 매핑 적용
            for target_col, possible_cols in column_mapping.items():
                for col in df.columns:
                    if col in possible_cols:
                        df = df.rename(columns={col: target_col})
            
            # 숫자 데이터 변환
            if '지급금액' in df.columns:
                df['지급금액'] = df['지급금액'].apply(
                    lambda x: re.sub(r'[^\d,.]', '', str(x)) if pd.notnull(x) else '')
            
            return df
            
        except Exception as e:
            self.logger.error(f"컬럼 매핑 중 오류: {str(e)}")
            return df
    

    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """추출된 표 데이터 정제 (컬럼 구조 유지)"""
        try:
            # 기존 컬럼명 변경 로직 제거
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # 필터링 패턴 제거 (원본 데이터 보존)
            df = df.reset_index(drop=True)
            
            # 문자열 정제만 수행
            for col in df.columns:
                if df[col].dtype == "object":
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].str.replace('\n', ' ', regex=False)
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                    df[col] = df[col].replace(['', 'nan', 'None'], None)
                    
            return df
            
        except Exception as e:
            self.logger.error(f"표 정제 중 오류: {str(e)}")
            return df
        
    def save_results(self, tables: List[pd.DataFrame], page_numbers: List[int]) -> Dict[str, str]:
        """결과를 XML과 Excel로 저장"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = f"tables_{timestamp}"
        
        # XML 생성 및 저장
        root = self.create_xml(tables, page_numbers)
        xml_str = ET.tostring(root, encoding='unicode')
        pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")
        
        xml_path = self.xml_dir / f"{base_filename}.xml"
        with open(xml_path, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
            
        # Excel 생성 및 저장
        all_data = []
        for df, page_num in zip(tables, page_numbers):
            df = df.copy()
            df['페이지'] = page_num
            all_data.append(df)
            
        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            excel_path = self.excel_dir / f"{base_filename}.xlsx"
            combined_df.to_excel(excel_path, index=False, engine='openpyxl')
        
        return {
            'xml_path': str(xml_path),
            'excel_path': str(excel_path)
        }
    
    def detect_strike_through(self, page: fitz.Page) -> List[Dict]:
        """취소선이 포함된 텍스트 영역 추출"""
        strike_blocks = []
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span["flags"] & fitz.TEXT_STRIKE_THROUGH:  # 취소선 플래그 확인
                        strike_blocks.append({
                            "text": span["text"],
                            "bbox": fitz.Rect(span["bbox"])
                        })
        return strike_blocks
    
    def detect_colored_text(self, page: fitz.Page, 
                       black_threshold: int = 50) -> List[Dict]:
        """검정색이 아닌 컬러 텍스트 감지 (빨강/파랑/녹색 등)"""
        colored_blocks = []
        blocks = page.get_text("dict")["blocks"]
        
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    # RGB 값 추출 (PyMuPDF 기준 0~255 범위)
                    if 'color' not in span:
                        continue
                        
                    color = span["color"]
                    r = (color >> 16) & 0xff  # Red 채널
                    g = (color >> 8) & 0xff   # Green 채널
                    b = color & 0xff          # Blue 채널
                    
                    # 검정색 판별 조건 (모든 채널이 임계값 미만)
                    is_black = all([c < black_threshold for c in (r, g, b)])
                    
                    # 검정이 아니고, 실제 텍스트가 있는 경우만 처리
                    if not is_black and span["text"].strip():
                        colored_blocks.append({
                            "text": span["text"],
                            "color": (r, g, b),  # RGB 값 저장
                            "bbox": fitz.Rect(span["bbox"])
                        })
        
        return colored_blocks
    
    def _process_table_with_highlights(self, table, highlight_regions, page_height, 
                                       page_num, parsing_range, sections_info=None) -> pd.DataFrame:
        try:
            # 1. 기본 테이블 처리
            df = self.clean_table(table)
            if df.empty:
                return df, []  # 하이라이트 정보 반환 추가

            # 2. 색상 텍스트 및 음영 감지
            page = self.doc[page_num]
            colored_texts = self.detect_colored_text(page)
            highlight_matrix = []

            # 3. 변경사항 컬럼 초기화 및 하이라이트 매트릭스 생성
            df['변경사항'] = ''
            for row_idx, row in df.iterrows():
                row_highlights = []
                for col_idx, cell in enumerate(row):
                    # 색상 텍스트 매칭 검사
                    is_highlighted = any(
                        colored['text'].strip() in str(cell)
                        for colored in colored_texts
                    )
                    row_highlights.append(is_highlighted)
                    if is_highlighted:
                        df.iat[row_idx, col_idx] = f"{cell} [색상]"  # 셀 내용에 마킹
                highlight_matrix.append(row_highlights)

            return df, highlight_matrix  # 하이라이트 정보 반환

        except Exception as e:
            self.logger.error(f"표 처리 오류: {str(e)}")
            return pd.DataFrame(), []
        
    def detect_colored_areas(self, page: fitz.Page) -> List[Dict]:
        """음영 영역 감지 (이미지 기반)"""
        img = self.pdf_to_image(page)
        hsv = cv2.cvtColor(img, cv2.COLOR_RGB2HSV)
        
        # 노란색 범위 정의 (H: 20-30, S: 100-255, V: 100-255)
        lower_yellow = np.array([20, 100, 100])
        upper_yellow = np.array([30, 255, 255])
        
        mask = cv2.inRange(hsv, lower_yellow, upper_yellow)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        colored_areas = []
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            colored_areas.append({
                "bbox": (x, y, x+w, y+h),
                "color": "YELLOW"
            })
        return colored_areas
    
    
        
        