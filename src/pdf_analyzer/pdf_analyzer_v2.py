import fitz
import camelot
import pandas as pd 
from pathlib import Path
import re
from datetime import datetime
import logging
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font  # 수정된 임포트문

@dataclass
class ParsingRange:
    start_page: int
    end_page: int
    section_type: str
    insurance_type: Optional[str] = None

class PDFAnalyzer:
    def __init__(self, pdf_path: str, logger=None):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.logger = logger or logging.getLogger(__name__)
        
        self.markers = {
            'payment_section': "나. 보험금 지급",
            'insurance_types': r'\[(\d)종\]',
            'sections': {
                '상해관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해|상해관련|상해 관련)([\s]*)(특약|특별약관)',
                '질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>질병|질병관련|질병 관련)([\s]*)(특약|특별약관)',
                '상해및질병관련 특별약관': r'[◇◆■□▶]([\s]*)(?P<title>상해\s*및\s*질병|상해와\s*질병)([\s]*)(관련)?([\s]*)(특약|특별약관)?'
            }
        }

    def determine_parsing_ranges(self) -> List[ParsingRange]:
        """섹션과 종별 기준으로 파싱 범위 결정"""
        ranges = []
        
        # 보험금 지급 섹션 찾기
        payment_start = self.find_payment_section()
        if not payment_start:
            self.logger.error("보험금 지급 섹션을 찾을 수 없습니다")
            return ranges

        # 보험금 지급 페이지에서 [1종/2종/3종] 패턴 확인
        payment_page_text = self.doc[payment_start].get_text()
        is_single_type = bool(re.search(r'\[(1|2|3)종/(1|2|3)종/(1|2|3)종\]', payment_page_text))
        
        if is_single_type:
            # 종별 구분이 없는 경우 - 직접 섹션 검색
            self.logger.info("종별 구분 없음 (통합본)")
            section_pages = {}
            
            # payment_start 페이지부터 섹션 검색
            for page_num in range(payment_start, len(self.doc)):
                text = self.doc[page_num].get_text()
                for section, pattern in self.markers['sections'].items():
                    if re.search(pattern, text):
                        section_pages[section] = page_num
                        self.logger.info(f"{section} 섹션 발견: {page_num + 1}페이지")
            
            # 섹션별 범위 설정
            if section_pages:
                sorted_sections = sorted(section_pages.items(), key=lambda x: x[1])
                for i, (section_name, start_page) in enumerate(sorted_sections):
                    end_page = (sorted_sections[i + 1][1] - 1 
                            if i < len(sorted_sections) - 1 
                            else len(self.doc) - 1)
                    
                    ranges.append(ParsingRange(
                        start_page=start_page,
                        end_page=end_page,
                        section_type=section_name
                    ))
                    self.logger.info(f"{section_name} 파싱 범위: {start_page + 1}~{end_page + 1}페이지")
        
        else:
            # 종별 구분이 있는 경우
            type_pages = self.find_insurance_types()
            if not type_pages:
                self.logger.warning("종별 구분을 찾을 수 없습니다")
                return ranges
                
            self.logger.info(f"종별 구분 발견: {[t[1] for t in type_pages]}")
            
            # 중복 종 제거 및 정렬
            unique_type_pages = []
            seen_types = set()
            for page, type_name in type_pages:
                if type_name not in seen_types:
                    unique_type_pages.append((page, type_name))
                    seen_types.add(type_name)
            type_pages = sorted(unique_type_pages)
            
            # 각 종별로 섹션 찾기
            for i, (type_page, type_name) in enumerate(type_pages):
                next_type_page = type_pages[i + 1][0] if i + 1 < len(type_pages) else len(self.doc)
                
                # 해당 종 범위 내의 섹션 찾기
                type_sections = {}
                for page_num in range(type_page, next_type_page):
                    text = self.doc[page_num].get_text()
                    for section, pattern in self.markers['sections'].items():
                        if re.search(pattern, text):
                            type_sections[section] = page_num
                            self.logger.info(f"{type_name} - {section} 섹션 발견: {page_num + 1}페이지")
                
                # 섹션별 범위 설정
                if type_sections:
                    sorted_sections = sorted(type_sections.items(), key=lambda x: x[1])
                    for j, (section_name, start_page) in enumerate(sorted_sections):
                        end_page = (sorted_sections[j + 1][1] - 1 
                                if j < len(sorted_sections) - 1 
                                else next_type_page - 1)
                        
                        ranges.append(ParsingRange(
                            start_page=start_page,
                            end_page=end_page,
                            section_type=section_name,
                            insurance_type=type_name
                        ))
                        self.logger.info(f"{type_name} {section_name} 파싱 범위: {start_page + 1}~{end_page + 1}페이지")
                else:
                    # 섹션이 없는 경우 전체 종을 하나의 범위로
                    ranges.append(ParsingRange(
                        start_page=type_page,
                        end_page=next_type_page - 1,
                        section_type='전체',
                        insurance_type=type_name
                    ))
                    self.logger.info(f"{type_name} 전체 파싱 범위: {type_page + 1}~{next_type_page}페이지")

        # 파싱 범위를 찾지 못한 경우
        if not ranges:
            ranges.append(ParsingRange(
                start_page=payment_start,
                end_page=len(self.doc) - 1,
                section_type='전체'
            ))
            self.logger.warning("섹션을 찾을 수 없어 전체 문서를 파싱합니다")

        return ranges

    def find_payment_section(self) -> Optional[int]:
        """보험금 지급 섹션의 시작 페이지 찾기"""
        for page_num in range(len(self.doc)):
            if self.markers['payment_section'] in self.doc[page_num].get_text():
                return page_num
        return None

    def find_insurance_types(self) -> List[Tuple[int, str]]:
        """종별 마커가 있는 페이지 찾기"""
        type_pages = []
        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            matches = re.finditer(self.markers['insurance_types'], text)
            for match in matches:
                type_num = match.group(1)
                type_pages.append((page_num, f"[{type_num}종]"))
        return sorted(type_pages)

    def find_section_pages(self) -> Dict[str, int]:
        """각 섹션의 시작 페이지 찾기"""
        section_pages = {}
        for page_num in range(len(self.doc)):
            text = self.doc[page_num].get_text()
            for section, pattern in self.markers['sections'].items():
                if re.search(pattern, text):
                    section_pages[section] = page_num
                    self.logger.info(f"{section} 섹션 발견: {page_num + 1}페이지")
        return section_pages

    def determine_parsing_ranges(self) -> List[ParsingRange]:
        """섹션 기준으로 파싱 범위 결정"""
        ranges = []
        section_pages = self.find_section_pages()
        
        if not section_pages:
            self.logger.error("섹션을 찾을 수 없습니다")
            return ranges
        
        # 섹션별 시작 페이지를 기준으로 정렬
        sorted_sections = sorted(section_pages.items(), key=lambda x: x[1])
        
        # 각 섹션의 범위 설정
        for i, (section_name, start_page) in enumerate(sorted_sections):
            # 다음 섹션의 시작 페이지 - 1 또는 문서 끝까지
            end_page = (sorted_sections[i + 1][1] - 1 
                    if i < len(sorted_sections) - 1 
                    else len(self.doc) - 1)
            
            ranges.append(ParsingRange(
                start_page=start_page,
                end_page=end_page,
                section_type=section_name
            ))
            self.logger.info(f"{section_name} 파싱 범위: {start_page + 1}~{end_page + 1}페이지")

        return ranges

    def extract_tables(self, parsing_range: ParsingRange) -> List[pd.DataFrame]:
        """지정된 범위에서 표 추출"""
        tables = []
        for page_num in range(parsing_range.start_page, parsing_range.end_page + 1):
            try:
                self.logger.info(f"{page_num + 1}페이지 표 추출 중...")
                page_tables = camelot.read_pdf(
                    self.pdf_path,
                    pages=str(page_num + 1),
                    flavor='lattice',
                    line_scale=40,
                    process_background=True
                )
                
                for table in page_tables:
                    df = table.df
                    if len(df.columns) >= 3:  # 최소 컬럼 수 체크
                        df = self.clean_table(df)
                        if not df.empty:
                            df['페이지'] = page_num + 1
                            df['구분'] = parsing_range.section_type
                            if parsing_range.insurance_type:
                                df['보험종류'] = parsing_range.insurance_type
                            tables.append(df)
                            
            except Exception as e:
                self.logger.warning(f"{page_num + 1}페이지 표 추출 중 오류: {e}")
                
        return tables

    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """추출된 표 데이터 정제"""
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # 컬럼명 정제
        df.columns = [str(col).strip() for col in df.columns]
        
        # 셀 내용 정제
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].str.replace('\n', ' ')
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
        
        return df

    def analyze(self) -> str:
        """PDF 분석 실행"""
        try:
            output_dir = Path("data/output")
            output_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = output_dir / f"보험약관분석_{timestamp}.xlsx"

            ranges = self.determine_parsing_ranges()
            if not ranges:
                raise ValueError("유효한 파싱 범위를 찾을 수 없습니다")

            all_tables = []
            for parsing_range in ranges:
                tables = self.extract_tables(parsing_range)
                all_tables.extend(tables)

            if not all_tables:
                raise ValueError("추출된 표가 없습니다")

            # 결과 저장
            with pd.ExcelWriter(output_path) as writer:
                if any(df.get('보험종류') is not None for df in all_tables):
                    # 종별로 시트 생성
                    for insurance_type in sorted(set(df['보험종류'].iloc[0] for df in all_tables if '보험종류' in df)):
                        type_tables = [df for df in all_tables if '보험종류' in df and df['보험종류'].iloc[0] == insurance_type]
                        if type_tables:
                            sheet_name = insurance_type.replace('[', '').replace(']', '')
                            pd.concat(type_tables, ignore_index=True).to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # 단일 시트에 모든 표 저장
                    pd.concat(all_tables, ignore_index=True).to_excel(writer, sheet_name='전체', index=False)

            return str(output_path)

        finally:
            if hasattr(self, 'doc'):
                self.doc.close()