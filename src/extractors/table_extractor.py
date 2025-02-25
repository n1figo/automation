import camelot
import pandas as pd
import fitz
import logging
from typing import List, Tuple
from .highlight_extractor import HighlightExtractor

logger = logging.getLogger(__name__)

class TableExtractor:
    def __init__(self):
        self.highlight_extractor = HighlightExtractor()
        self.logger = logging.getLogger(__name__)

    def extract_tables_from_section(self, pdf_path: str, start_page: int, end_page: int) -> List[Tuple[str, pd.DataFrame, int]]:
        """지정된 페이지 범위에서 표 추출"""
        results = []
        
        if start_page is None or end_page is None:
            self.logger.warning("시작 페이지 또는 끝 페이지가 지정되지 않았습니다")
            return results
            
        doc = None
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(start_page, end_page + 1):
                try:
                    if page_num >= len(doc):
                        self.logger.warning(f"페이지 번호 {page_num + 1}가 문서 범위를 벗어났습니다")
                        continue

                    # PDF 페이지 처리
                    page = doc[page_num]
                    if page is None:
                        self.logger.warning(f"페이지 {page_num + 1}를 불러올 수 없습니다")
                        continue
                        
                    self.logger.info(f"Processing page-{page_num + 1}")
                    
                    # 하이라이트 감지
                    contours, regions = self.highlight_extractor.process_page(page)

                    # 페이지 텍스트 추출 (제목용)
                    page_text = page.get_text()
                    
                    # 테이블 추출 시도
                    tables = self._extract_tables(pdf_path, page_num + 1)

                    # 추출된 테이블 처리
                    if tables:
                        for table_idx, table in enumerate(tables):
                            df = self.process_table_with_highlights(
                                table, 
                                regions, 
                                page.rect.height,
                                table_idx,
                                page_num + 1
                            )
                            if not df.empty:
                                results.append((page_text, df, page_num + 1))

                except Exception as e:
                    self.logger.error(f"페이지 {page_num + 1} 처리 중 오류 발생: {str(e)}")
                    continue

        except Exception as e:
            self.logger.error(f"PDF 처리 중 오류 발생: {str(e)}")
            
        finally:
            if doc:
                doc.close()

        return results

    def _extract_tables(self, pdf_path: str, page_num: int) -> List:
        """테이블 추출 시도"""
        try:
            # 먼저 lattice 방식으로 시도
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num), 
                flavor='lattice',
                line_scale=40,
                process_background=True
            )

            # lattice 방식으로 실패하면 stream 방식 시도
            if len(tables) == 0:
                tables = camelot.read_pdf(
                    pdf_path, 
                    pages=str(page_num), 
                    flavor='stream',
                    edge_tol=500,
                    row_tol=10
                )

            return tables

        except Exception as e:
            self.logger.error(f"테이블 추출 중 오류: {str(e)}")
            return []

    def process_table_with_highlights(self, table, highlight_regions, page_height, table_idx, page_num):
        """하이라이트 적용 및 표 처리"""
        try:
            df = table.df.copy()
            if df.empty:
                return df

            # 표 위치 정보
            try:
                x1, y1, x2, y2 = table._bbox
                row_height = (y2 - y1) / len(df)
            except Exception as e:
                self.logger.warning(f"표 위치 정보 추출 실패: {str(e)}")
                row_height = page_height / len(df)
                y2 = page_height

            # 메타데이터 열 추가
            df['변경사항'] = ''
            df['Table_Number'] = table_idx + 1
            df['페이지'] = page_num

            # 하이라이트 확인
            for row_index in range(len(df)):
                try:
                    row_top = y2 - (row_index + 1) * row_height
                    row_bottom = y2 - row_index * row_height
                    
                    # 하이라이트 여부 확인
                    row_highlighted = self.highlight_extractor.check_highlight(
                        (row_top, row_bottom), 
                        highlight_regions
                    )

                    if row_highlighted:
                        df.at[row_index, '변경사항'] = '추가'
                except Exception as e:
                    self.logger.warning(f"행 {row_index} 하이라이트 처리 실패: {str(e)}")
                    continue

            # 데이터 정제
            df = self._clean_table_data(df)
            
            return df

        except Exception as e:
            self.logger.error(f"테이블 처리 중 오류: {str(e)}")
            return pd.DataFrame()

    def _clean_table_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """표 데이터 정제"""
        try:
            # 빈 행/열 제거
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # 문자열 데이터 정제
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].str.replace('\n', ' ')
                    df[col] = df[col].str.replace(r'\s+', ' ', regex=True)

            return df
            
        except Exception as e:
            self.logger.error(f"데이터 정제 중 오류: {str(e)}")
            return df