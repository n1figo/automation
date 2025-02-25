import fitz
import camelot
import pandas as pd
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
import re

logger = logging.getLogger(__name__)

class PDFExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.similarity_threshold = 0.7
        
        # 섹션 패턴 정의
        self.section_patterns = {
            '상해관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
            '질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
            '상해및질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)?'
        }
        
        # 섹션 초기화
        self.sections = {
            "상해관련 특별약관": {"start": None, "end": None},
            "질병관련 특별약관": {"start": None, "end": None},
            "상해및질병관련 특별약관": {"start": None, "end": None}
        }
        
        # 모델 초기화
        self._init_model()

    def _init_model(self):
        """문장 임베딩 모델 초기화"""
        try:
            model_path = "models/distiluse-base-multilingual-cased-v1"
            self.model = SentenceTransformer(model_path)
            logger.info("문장 임베딩 모델 로드 성공")
        except Exception as e:
            logger.error(f"모델 로드 실패: {str(e)}")
            raise

    def find_sections(self) -> Dict[str, Dict[str, int]]:
        """섹션 위치 찾기"""
        try:
            print("\n=== 섹션 검색 시작 ===")
            
            # 전체 페이지 검색
            for page_num in range(len(self.doc)):
                text = self.doc[page_num].get_text()
                
                # 상해및질병관련 특별약관 먼저 찾기 (다른 섹션의 범위 제한을 위해)
                both_match = re.search(self.section_patterns['상해및질병관련 특별약관'], text)
                if both_match:
                    self.sections['상해및질병관련 특별약관']['start'] = page_num
                    
                    # 이전 섹션들 종료
                    if (self.sections['상해관련 특별약관']['start'] is not None and 
                        self.sections['상해관련 특별약관']['end'] is None):
                        self.sections['상해관련 특별약관']['end'] = page_num - 1
                        print(f"[종료] 상해관련 특별약관: {page_num}페이지 (상해및질병 섹션 시작)")
                        logger.info(f"상해 섹션 종료 (상해및질병 시작): {page_num}페이지")
                    
                    if (self.sections['질병관련 특별약관']['start'] is not None and 
                        self.sections['질병관련 특별약관']['end'] is None):
                        self.sections['질병관련 특별약관']['end'] = page_num - 1
                        print(f"[종료] 질병관련 특별약관: {page_num}페이지 (상해및질병 섹션 시작)")
                        logger.info(f"질병 섹션 종료 (상해및질병 시작): {page_num}페이지")
                    continue

                # 상해관련 특별약관 찾기
                if re.search(self.section_patterns['상해관련 특별약관'], text):
                    if self.sections['상해관련 특별약관']['start'] is None:
                        self.sections['상해관련 특별약관']['start'] = page_num
                        print(f"[시작] 상해관련 특별약관: {page_num + 1}페이지")
                        logger.info(f"상해 섹션 시작: {page_num + 1}페이지")

                # 질병관련 특별약관 찾기
                if re.search(self.section_patterns['질병관련 특별약관'], text):
                    # 상해 섹션이 끝나지 않았다면 여기서 종료
                    if (self.sections['상해관련 특별약관']['start'] is not None and 
                        self.sections['상해관련 특별약관']['end'] is None):
                        self.sections['상해관련 특별약관']['end'] = page_num - 1
                        print(f"[종료] 상해관련 특별약관: {page_num}페이지 (질병 섹션 시작)")
                        logger.info(f"상해 섹션 종료 (질병 시작): {page_num}페이지")
                    
                    # 새로운 질병 섹션 시작
                    if self.sections['질병관련 특별약관']['start'] is None:
                        self.sections['질병관련 특별약관']['start'] = page_num
                        print(f"[시작] 질병관련 특별약관: {page_num + 1}페이지")
                        logger.info(f"질병 섹션 시작: {page_num + 1}페이지")

            # 상해및질병 섹션이 없는 경우의 종료 처리
            if self.sections['상해및질병관련 특별약관']['start'] is None:
                # 질병 섹션 종료
                if (self.sections['질병관련 특별약관']['start'] is not None and 
                    self.sections['질병관련 특별약관']['end'] is None):
                    self.sections['질병관련 특별약관']['end'] = len(self.doc) - 1
                    print(f"[종료] 질병관련 특별약관: {len(self.doc)}페이지")
                    logger.info(f"질병 섹션 종료 (문서 끝): {len(self.doc)}페이지")
                
                # 상해 섹션 종료
                elif (self.sections['상해관련 특별약관']['start'] is not None and 
                    self.sections['상해관련 특별약관']['end'] is None):
                    self.sections['상해관련 특별약관']['end'] = len(self.doc) - 1
                    print(f"[종료] 상해관련 특별약관: {len(self.doc)}페이지")
                    logger.info(f"상해 섹션 종료 (문서 끝): {len(self.doc)}페이지")

            return self.sections

        except Exception as e:
            logger.error(f"섹션 검색 중 오류: {str(e)}")
            return self.sections

    def extract_tables(self, start_page: int, end_page: int) -> List[pd.DataFrame]:
        """페이지 범위에서 테이블 추출"""
        tables = []
        
        try:
            for page_num in range(start_page, end_page + 1):
                logger.info(f"Processing page-{page_num + 1}")
                
                # lattice 모드로 시도
                extracted_tables = camelot.read_pdf(
                    self.pdf_path,
                    pages=str(page_num + 1),
                    flavor='lattice',
                    line_scale=40,
                    line_tol=2
                )
                
                # 실패시 stream 모드 시도
                if not extracted_tables:
                    extracted_tables = camelot.read_pdf(
                        self.pdf_path,
                        pages=str(page_num + 1),
                        flavor='stream',
                        edge_tol=150,
                        row_tol=5
                    )
                
                # 테이블 처리
                for table in extracted_tables:
                    if isinstance(table.df, pd.DataFrame) and not table.df.empty:
                        df = table.df.copy()
                        df = self._clean_table(df, page_num + 1)
                        if not df.empty and len(df.columns) >= 3:
                            tables.append(df)
                
        except Exception as e:
            logger.error(f"페이지 범위 {start_page + 1} ~ {end_page + 1} 테이블 추출 실패: {str(e)}")
        
        return tables

    def _clean_table(self, df: pd.DataFrame, page_num: int) -> pd.DataFrame:
        """테이블 데이터 정제"""
        try:
            if df.empty:
                return pd.DataFrame()
                
            # 컬럼명 정리
            df.columns = [str(col).strip() for col in df.columns]
            
            # 불필요한 행 제거
            df = df[~df.apply(lambda x: x.astype(str).str.contains('피보험자님의 가입내용').any(), axis=1)]
            df = df.dropna(how='all')
            
            # 데이터 정제
            for col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
                df[col] = df[col].replace('nan', '')
                df[col] = df[col].replace('None', '')
            
            # 페이지 정보 추가
            df['페이지'] = page_num
            
            return df
            
        except Exception as e:
            logger.error(f"테이블 정제 중 오류: {str(e)}")
            return pd.DataFrame()

    def analyze_sections(self, text: str) -> Dict[str, Tuple[int, int]]:
        """섹션 범위 분석"""
        sections = {}
        
        try:
            # 상해 섹션 시작 찾기
            injury_match = re.search(r'상해관련\s*특별약관.*?(\d+)페이지', text)
            if injury_match:
                injury_start = int(injury_match.group(1))
                logger.info(f"상해관련 특별약관 시작: {injury_start}페이지")
                
                # 상해및질병 섹션 시작 찾기
                both_match = re.search(r'상해\s*및\s*질병관련\s*특별약관.*?(\d+)페이지', text)
                if both_match:
                    both_start = int(both_match.group(1))
                    logger.info(f"상해및질병관련 특별약관 시작: {both_start}페이지")
                    
                    # 상해 섹션 범위 설정 (상해 시작 ~ 상해및질병 시작 전)
                    sections['상해'] = (injury_start, both_start - 1)
                    logger.info(f"상해 섹션 파싱 범위: {injury_start} ~ {both_start - 1}페이지")
            
            return sections
            
        except Exception as e:
            logger.error(f"섹션 범위 분석 중 오류: {str(e)}")
            return sections

    def close(self):
        """리소스 정리"""
        if self.doc:
            self.doc.close()