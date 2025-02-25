from bs4 import BeautifulSoup
from typing import Dict, List, Optional, Union
import pandas as pd
import logging
from pathlib import Path
import email
import os
import chardet

logger = logging.getLogger(__name__)

class HTMLFileExtractor:
    def __init__(self, file_path: str):
        """
        HTML/MHTML 파일 추출기 초기화
        
        Args:
            file_path: HTML 또는 MHTML 파일 경로
        """
        self.file_path = Path(file_path)
        self.setup_logging()
        self.soup = self._load_file()

    def setup_logging(self):
        """로깅 설정"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler()
            ]
        )

    def _detect_encoding(self, content: bytes) -> str:
        """파일 인코딩 감지"""
        result = chardet.detect(content)
        logger.info(f"감지된 인코딩: {result['encoding']}")
        return result['encoding'] or 'utf-8'

    def _try_decode(self, content: bytes, fallback_encodings=['utf-8', 'euc-kr', 'cp949']) -> str:
        """다양한 인코딩으로 디코딩 시도"""
        logger.info("디코딩 시도 중...")
        detected_encoding = self._detect_encoding(content)
        try:
            return content.decode(detected_encoding)
        except UnicodeDecodeError:
            logger.warning("감지된 인코딩으로 디코딩 실패, 폴백 인코딩 시도 중...")
            pass

        for encoding in fallback_encodings:
            try:
                return content.decode(encoding)
            except UnicodeDecodeError:
                logger.warning(f"{encoding}로 디코딩 실패")
                continue

        raise UnicodeDecodeError("모든 인코딩 시도 실패")

    def _load_file(self) -> Optional[BeautifulSoup]:
        """파일 형식에 따라 적절한 로더 선택"""
        try:
            if not self.file_path.exists():
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {self.file_path}")

            logger.info(f"파일 로드 중: {self.file_path}")
            if self.file_path.suffix.lower() == '.mhtml':
                return self._load_mhtml()
            else:
                return self._load_html()
        except Exception as e:
            logger.error(f"파일 로드 중 오류 발생: {str(e)}")
            return None

    def _load_mhtml(self) -> Optional[BeautifulSoup]:
        """MHTML 파일 로드"""
        try:
            logger.info(f"MHTML 파일 로드 시작: {self.file_path}")
            with open(self.file_path, 'rb') as f:
                msg = email.message_from_binary_file(f)
                
            html_parts = []
            for part in msg.walk():
                content_type = part.get_content_type()
                logger.info(f"Found content type: {content_type}")
                
                if content_type == 'text/html':
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset()
                        try:
                            if charset:
                                html_content = payload.decode(charset)
                            else:
                                html_content = self._try_decode(payload)
                            html_parts.append(html_content)
                        except Exception as e:
                            logger.error(f"Part 디코딩 실패: {str(e)}")
            
            if html_parts:
                combined_html = "\n".join(html_parts)
                logger.info("MHTML에서 HTML 콘텐츠를 성공적으로 추출했습니다.")
                return BeautifulSoup(combined_html, 'html.parser')
            else:
                logger.error("MHTML 파일에서 HTML 콘텐츠를 찾을 수 없습니다.")
                raise ValueError("MHTML 파일에서 HTML 콘텐츠를 찾을 수 없습니다.")
                
        except Exception as e:
            logger.error(f"MHTML 파일 처리 중 오류 발생: {str(e)}")
            return None

    def _load_html(self) -> Optional[BeautifulSoup]:
        """HTML 파일 로드"""
        try:
            logger.info(f"HTML 파일 로드 시작: {self.file_path}")
            with open(self.file_path, 'rb') as f:
                content = f.read()
                
            html_content = self._try_decode(content)
            logger.info("HTML 파일을 성공적으로 읽었습니다.")
            return BeautifulSoup(html_content, 'html.parser')
        except Exception as e:
            logger.error(f"HTML 파일 읽기 실패: {str(e)}")
            return None

    def extract_tables(self) -> List[pd.DataFrame]:
        """파일에서 시작과 끝 태그 사이의 모든 표 추출"""
        try:
            if not self.soup:
                logger.error("BeautifulSoup 객체가 초기화되지 않았습니다.")
                return []

            # 시작 지점 찾기
            start_div = self.soup.find('div', id='exmpl')
            if not start_div:
                logger.error("시작 div를 찾을 수 없습니다.")
                return []

            # 끝 지점 찾기
            end_div = self.soup.find('div', class_='guideBoxWrap')
            if not end_div:
                logger.error("끝 div를 찾을 수 없습니다.")
                return []

            # 시작과 끝 사이의 모든 테이블 찾기
            tables = []
            current = start_div
            while current and current != end_div:
                if current.name == 'table':
                    # 동적으로 생성되는 테이블의 특성 확인
                    data_list = current.get('data-list')
                    if data_list:  # 동적 테이블인 경우
                        logger.info(f"동적 테이블 발견: data-list={data_list}")
                    
                    df = self._convert_table_to_df(current)
                    if not df.empty and len(df.columns) >= 2:
                        tables.append(df)
                        logger.info(f"테이블 추가됨: {len(df)} 행, {len(df.columns)} 열")
                
                current = current.find_next()

            if not tables:
                logger.warning("추출된 테이블이 없습니다")
            else:
                logger.info(f"{len(tables)}개의 표를 추출했습니다.")

            return tables

        except Exception as e:
            logger.error(f"표 추출 중 오류 발생: {str(e)}")
            return []

    def _convert_table_to_df(self, table) -> pd.DataFrame:
        try:
            logger.info("테이블 변환 시작...")
            # 모든 행 추출 (헤더 포함)
            rows = []
            for tr in table.find_all('tr'):
                row = []
                for cell in tr.find_all(['th', 'td']):
                    colspan = int(cell.get('colspan', 1))
                    cell_text = cell.get_text(strip=True)
                    row.extend([cell_text] * colspan)
                if row:
                    rows.append(row)

            if not rows:
                return pd.DataFrame()

            # 최대 열 수 계산
            max_cols = max(len(row) for row in rows)
            
            # 모든 행의 길이를 최대 길이로 맞춤
            padded_rows = [row + [''] * (max_cols - len(row)) for row in rows]

            # 첫 번째 행을 헤더로 사용
            headers = padded_rows[0]
            data = padded_rows[1:]

            df = pd.DataFrame(data, columns=headers)
            logger.info("테이블 변환 완료")
            return self._clean_dataframe(df)

        except Exception as e:
            logger.error(f"테이블 변환 중 오류 발생: {str(e)}")
            return pd.DataFrame()

    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """DataFrame 정제"""
        try:
            logger.info("DataFrame 정제 시작...")
            # 빈 행/열 제거
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # 컬럼명 정제
            df.columns = [str(col).strip() if isinstance(col, str) else str(col) for col in df.columns]
            
            # 데이터 정제
            for col in df.columns:
                # Series를 문자열로 변환하고 정제
                df[col] = df[col].astype(str).apply(lambda x: x.strip() if not pd.isna(x) else '')
                df[col] = df[col].replace(r'\s+', ' ', regex=True)
                df[col] = df[col].replace('', '')  # 빈 문자열을 공백으로
            
            # 중복 행 제거
            df = df.drop_duplicates()
            
            # 모든 셀이 빈 값인 행 제거
            df = df.dropna(how='all')
            
            logger.info("DataFrame 정제 완료.")
            return df

        except Exception as e:
            logger.error(f"DataFrame 정제 중 오류 발생: {str(e)}")
            return df

    def save_tables_to_excel(self, tables: List[pd.DataFrame], output_path: str):
        """추출한 모든 테이블을 하나의 시트에 저장"""
        try:
            if not tables:
                logger.error("저장할 테이블이 없습니다.")
                return

            logger.info(f"테이블을 {output_path}에 저장 중...")
            
            # 모든 테이블에 대해 NA 및 빈 문자열 처리
            processed_tables = []
            for df in tables:
                # NA와 빈 문자열을 빈 칸으로 대체
                df = df.fillna('')
                df = df.applymap(lambda x: '' if pd.isna(x) or str(x).strip() == '' else str(x))
                processed_tables.append(df)
            
            # 모든 테이블을 하나로 합치기
            combined_df = pd.concat(processed_tables, axis=0, ignore_index=True)
            
            # 구분자 열 추가
            row_counts = [len(df) for df in processed_tables]
            table_labels = []
            for i, count in enumerate(row_counts, 1):
                table_labels.extend([f'Table_{i}'] * count)
            combined_df.insert(0, '구분', table_labels)
            
            # Excel 저장 전 최종 NA 체크
            combined_df = combined_df.fillna('')
            
            # Excel 파일로 저장
            combined_df.to_excel(output_path, sheet_name='통합_테이블', index=False)
            
            logger.info(f"통합 테이블을 {output_path}에 성공적으로 저장했습니다.")
            logger.info(f"총 {len(combined_df)} 행의 데이터가 저장되었습니다.")
            
        except Exception as e:
            logger.error(f"Excel 파일 저장 중 오류 발생: {str(e)}")