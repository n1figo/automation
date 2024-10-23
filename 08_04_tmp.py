import PyPDF2
import re
import logging
import fitz
import numpy as np
from typing import Dict, List, Tuple, Optional
import os
import pandas as pd
import camelot
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from sentence_transformers import SentenceTransformer

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PageRangeDetector:
    def __init__(self):
        self.section_patterns = {
            "상해관련": r'상해관련\s*특약',
            "질병관련": r'질병관련\s*특약'
        }

    def find_section_pages(self, pdf_path: str) -> Dict[str, Tuple[int, int]]:
        """상해관련 특약과 질병관련 특약의 페이지 범위 찾기"""
        try:
            doc = fitz.open(pdf_path)
            section_starts = {"상해관련": None, "질병관련": None}
            section_ends = {"상해관련": None, "질병관련": None}
            
            for page_num in range(len(doc)):
                text = doc[page_num].get_text()
                
                # 섹션 시작 페이지 찾기
                for section, pattern in self.section_patterns.items():
                    if section_starts[section] is None and re.search(pattern, text):
                        section_starts[section] = page_num
                        logger.info(f"{section} 특약 시작 페이지: {page_num + 1}")
                
                # 섹션 끝 페이지 찾기 (다음 섹션 시작 전 또는 특정 키워드)
                for section in section_starts.keys():
                    if (section_starts[section] is not None and 
                        section_ends[section] is None):
                        next_section_found = any(
                            re.search(pattern, text) 
                            for s, pattern in self.section_patterns.items() 
                            if s != section
                        )
                        if next_section_found or "※" in text:
                            section_ends[section] = page_num - 1
                            logger.info(f"{section} 특약 끝 페이지: {page_num}")

            # 마지막 섹션의 끝 페이지 설정
            for section in section_ends.keys():
                if section_ends[section] is None and section_starts[section] is not None:
                    section_ends[section] = len(doc) - 1

            doc.close()
            
            # 페이지 범위 반환
            return {
                section: (start, section_ends[section])
                for section, start in section_starts.items()
                if start is not None and section_ends[section] is not None
            }
            
        except Exception as e:
            logger.error(f"Error finding section pages: {e}")
            return {}

class TitleExtractor:
    def __init__(self):
        self.model = SentenceTransformer('distiluse-base-multilingual-cased-v1')

    def validate_title_by_rules(self, title_text: str) -> bool:
        """제목 텍스트 규칙 검증"""
        rules = {
            'length': lambda x: 5 < len(x) < 100,
            'forbidden_chars': lambda x: not any(char in x for char in ['※', '►', '▶', '=']),
            'keyword': lambda x: any(keyword in x for keyword in [
                '보장', '특약', '진단', '수술', '입원', '치료'
            ]),
            'number_ratio': lambda x: sum(c.isdigit() for c in x) / len(x) < 0.3
        }
        
        score = sum(rule(title_text) for rule in rules.values())
        return score >= 3

    def validate_position(self, title_block: Dict, table_block: Dict) -> bool:
        """제목 위치 규칙 검증"""
        try:
            title_bottom = title_block['bbox'][3]
            table_top = table_block['bbox'][1]
            distance = table_top - title_bottom
            
            return 0 < distance < 50  # 50포인트 이내의 거리
        except Exception:
            return False

    def extract_title(self, page, table_position) -> str:
        """RAG와 규칙 기반으로 제목 추출"""
        try:
            blocks = page.get_text("dict")["blocks"]
            candidates = []
            
            for block in blocks:
                if "lines" in block:
                    text = " ".join(span["text"] for line in block["lines"] 
                                  for span in line["spans"])
                    
                    if block["bbox"][3] < table_position:  # 표보다 위에 있는 텍스트
                        if self.validate_title_by_rules(text):
                            # 임베딩 생성 및 유사도 계산
                            text_embedding = self.model.encode([text])[0]
                            reference_embedding = self.model.encode(["특약 보장 내용의 제목"])[0]
                            similarity = np.dot(text_embedding, reference_embedding)
                            
                            candidates.append({
                                'text': text,
                                'block': block,
                                'similarity': similarity,
                                'distance': table_position - block["bbox"][3]
                            })
            
            if candidates:
                # 유사도와 거리를 결합하여 최적의 제목 선택
                best_candidate = max(
                    candidates,
                    key=lambda x: x['similarity'] * 0.7 + (1 / (1 + x['distance'])) * 0.3
                )
                return best_candidate['text']
                
            return "Untitled Table"
            
        except Exception as e:
            logger.error(f"Error extracting title: {e}")
            return "Untitled Table"

class TableExtractor:
    def __init__(self):
        self.title_extractor = TitleExtractor()

    def extract_tables_from_section(self, pdf_path: str, start_page: int, end_page: int) -> List[Tuple[str, pd.DataFrame, int]]:
        """섹션 범위 내의 표 추출"""
        try:
            results = []
            doc = fitz.open(pdf_path)
            
            for page_num in range(start_page, end_page + 1):
                page = doc[page_num]
                
                # 표 추출
                tables = self.extract_with_camelot(pdf_path, page_num + 1)
                
                for table in tables:
                    df = self.clean_table(table.df)
                    if not df.empty:
                        # 표의 위치 정보 얻기
                        table_bbox = table.cells[0][0].bbox  # 첫 번째 셀의 bbox
                        table_top = table_bbox[1]
                        
                        # 제목 추출
                        title = self.title_extractor.extract_title(page, table_top)
                        results.append((title, df, page_num + 1))

            doc.close()
            return results
            
        except Exception as e:
            logger.error(f"Error extracting tables from section: {e}")
            return []

    def extract_with_camelot(self, pdf_path: str, page_num: int) -> List:
        """Camelot을 사용한 표 추출"""
        try:
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_num),
                flavor='lattice'
            )
            if not tables:
                tables = camelot.read_pdf(
                    pdf_path,
                    pages=str(page_num),
                    flavor='stream'
                )
            return tables
        except Exception as e:
            logger.error(f"Camelot extraction failed: {str(e)}")
            return []

    def clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """표 데이터 정제"""
        try:
            df = df.dropna(how='all')
            df = df[~df.iloc[:, 0].str.contains("※|주)", regex=False, na=False)]
            return df
        except Exception as e:
            logger.error(f"Error cleaning table: {e}")
            return pd.DataFrame()

class ExcelWriter:
    @staticmethod
    def save_to_excel(sections_data: Dict[str, List[Tuple[str, pd.DataFrame, int]]], output_path: str):
        """섹션별 데이터를 Excel 파일로 저장"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for section, tables in sections_data.items():
                    if not tables:
                        continue
                        
                    current_row = 0
                    sheet_name = "특약표"
                    
                    # 섹션 제목 쓰기
                    section_df = pd.DataFrame([[section]], columns=[''])
                    section_df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        startrow=current_row,
                        index=False,
                        header=False
                    )
                    current_row += 2
                    
                    for title, df, page_num in tables:
                        # 제목 쓰기
                        title_df = pd.DataFrame([[f"{title} (페이지: {page_num})"]], columns=[''])
                        title_df.to_excel(
                            writer,
                            sheet_name=sheet_name,
                            startrow=current_row,
                            index=False,
                            header=False
                        )
                        
                        # 표 데이터 쓰기
                        df.to_excel(
                            writer,
                            sheet_name=sheet_name,
                            startrow=current_row + 2,
                            index=False
                        )
                        
                        # 스타일 적용
                        worksheet = writer.sheets[sheet_name]
                        
                        # 제목 스타일링
                        title_cell = worksheet.cell(row=current_row + 1, column=1)
                        title_cell.font = Font(bold=True, size=12)
                        title_cell.fill = PatternFill(
                            start_color='E6E6E6',
                            end_color='E6E6E6',
                            fill_type='solid'
                        )
                        
                        current_row += len(df) + 5

            logger.info(f"Successfully saved tables to {output_path}")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")

def main():
    try:
        # 파일 경로 설정
        pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
        output_path = "보험특약표.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return

        # 페이지 범위 감지
        page_detector = PageRangeDetector()
        section_ranges = page_detector.find_section_pages(pdf_path)
        
        if not section_ranges:
            logger.error("No sections found in the document")
            return

        # 표 추출
        table_extractor = TableExtractor()
        sections_data = {}
        
        # 각 섹션별 표 추출
        for section, (start_page, end_page) in section_ranges.items():
            logger.info(f"Processing {section} (pages {start_page + 1} to {end_page + 1})")
            tables = table_extractor.extract_tables_from_section(
                pdf_path, start_page, end_page
            )
            if tables:
                sections_data[section] = tables
                logger.info(f"Found {len(tables)} tables in {section}")

        # 결과 저장
        if sections_data:
            ExcelWriter.save_to_excel(sections_data, output_path)
            logger.info("Processing completed successfully")
        else:
            logger.error("No tables extracted from any section")

    except Exception as e:
        logger.error(f"Processing error: {str(e)}")

if __name__ == "__main__":
    main()