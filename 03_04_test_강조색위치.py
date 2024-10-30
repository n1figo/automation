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
from openpyxl.styles import PatternFill, Alignment, Font
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
import cv2
from PIL import Image
from transformers import LayoutLMv3Processor, LayoutLMv3ForSequenceClassification, LayoutLMv3Model
from transformers import LayoutLMv3FeatureExtractor
from PIL import Image, ImageDraw
import torch
from datasets import Dataset
import json

class DocumentAnalyzer:
    def __init__(self):
        self.device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        self.processor = LayoutLMv3Processor.from_pretrained("microsoft/layoutlmv3-base", apply_ocr=True)
        self.model = LayoutLMv3Model.from_pretrained("microsoft/layoutlmv3-base")
        self.model.to(self.device)
        
        # 기존 설정 유지
        self.section_patterns = {
            "종류": r'\[(\d)종\]',
            "특약유형": r'(상해관련|질병관련)\s*특약'
        }
        self.title_patterns = [
            r'기본계약',
            r'의무부가계약',
            r'[\w\s]+관련\s*특약',
            r'[\w\s]+계약',
            r'[\w\s]+보장'
        ]
        
    def process_page(self, pdf_path: str, page_num: int):
        """LayoutLM v3로 페이지 처리"""
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        pix = page.get_pixmap()
        
        # PDF 페이지를 이미지로 변환
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # LayoutLM v3 처리
        encoding = self.processor(img, return_tensors="pt")
        encoding = {k: v.to(self.device) for k, v in encoding.items()}
        
        # 모델 실행
        with torch.no_grad():
            outputs = self.model(**encoding)
        
        # 결과 분석
        last_hidden_states = outputs.last_hidden_state
        
        # OCR 결과와 레이아웃 정보 추출
        words = self.processor.tokenizer.convert_ids_to_tokens(encoding["input_ids"][0])
        boxes = encoding["bbox"][0].cpu().numpy()
        
        return {
            "words": words,
            "boxes": boxes,
            "features": last_hidden_states[0].cpu().numpy(),
            "image": img
        }

    def detect_highlights(self, img, boxes):
        """하이라이트된 영역 감지"""
        img_np = np.array(img)
        hsv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2HSV)
        
        # 초록색 범위 정의
        lower_green = np.array([40, 40, 40])
        upper_green = np.array([80, 255, 255])
        mask = cv2.inRange(hsv, lower_green, upper_green)
        
        highlighted_boxes = []
        for box in boxes:
            x1, y1, x2, y2 = map(int, box)
            region = mask[y1:y2, x1:x2]
            if region.mean() > 30:  # 하이라이트 임계값
                highlighted_boxes.append(box)
        
        return highlighted_boxes

    def extract_table_info(self, page_result, camelot_tables):
        """표 정보 추출 및 하이라이트 매핑"""
        highlighted_boxes = self.detect_highlights(page_result["image"], page_result["boxes"])
        
        tables_info = []
        for table in camelot_tables:
            df = table.df.copy()
            table_area = table._bbox
            
            # 하이라이트된 행 식별
            df['변경사항'] = ''
            for idx, row in df.iterrows():
                row_bbox = self.get_row_bbox(table, idx)
                if self.check_highlight_overlap(row_bbox, highlighted_boxes):
                    df.loc[idx, '변경사항'] = '추가'
            
            # 제목 찾기
            title = self.find_table_title(page_result, table_area)
            
            tables_info.append({
                'title': title,
                'data': df,
                'bbox': table_area
            })
        
        return tables_info

    def get_row_bbox(self, table, row_idx):
        """표의 특정 행의 bbox 계산"""
        y1 = table._bbox[1] + (row_idx * (table._bbox[3] - table._bbox[1]) / len(table.df))
        y2 = y1 + ((table._bbox[3] - table._bbox[1]) / len(table.df))
        return [table._bbox[0], y1, table._bbox[2], y2]

    def check_highlight_overlap(self, row_bbox, highlighted_boxes):
        """행과 하이라이트된 영역의 겹침 확인"""
        def overlap(box1, box2):
            x1, y1, x2, y2 = box1
            x3, y3, x4, y4 = box2
            
            overlap_y = min(y2, y4) - max(y1, y3)
            if overlap_y <= 0:
                return False
                
            box_height = y2 - y1
            overlap_ratio = overlap_y / box_height
            return overlap_ratio > 0.3
        
        return any(overlap(row_bbox, hbox) for hbox in highlighted_boxes)

    def find_table_title(self, page_result, table_area):
        """표 제목 찾기"""
        words = page_result["words"]
        boxes = page_result["boxes"]
        features = page_result["features"]
        
        # 표 위쪽 영역의 텍스트 검사
        title_candidates = []
        for word, box, feature in zip(words, boxes, features):
            if box[1] < table_area[1] and box[3] < table_area[1]:  # 표 위쪽 텍스트
                if any(re.search(pattern, word) for pattern in self.title_patterns):
                    title_candidates.append({
                        'text': word,
                        'distance': table_area[1] - box[3],
                        'feature': feature
                    })
        
        if title_candidates:
            # 가장 가까운 제목 선택
            return min(title_candidates, key=lambda x: x['distance'])['text']
        return "Untitled Table"

class TableExtractor:
    def __init__(self):
        self.document_analyzer = DocumentAnalyzer()
        self.max_distance = 50

    def extract_tables_from_section(self, pdf_path: str, start_page: int, end_page: int) -> List[Tuple[str, pd.DataFrame, int]]:
        try:
            results = []
            for page_num in range(start_page, end_page):
                doc = fitz.open(pdf_path)
                
                # LayoutLM v3로 페이지 분석
                page_result = self.document_analyzer.process_page(pdf_path, page_num)
                
                # Camelot으로 표 추출
                tables = self.extract_with_camelot(pdf_path, page_num + 1)
                
                if tables:
                    # 표 정보 추출 및 하이라이트 매핑
                    tables_info = self.document_analyzer.extract_table_info(page_result, tables)
                    
                    for table_info in tables_info:
                        title = table_info['title']
                        df = table_info['data']
                        
                        if not df.empty:
                            results.append((title, df, page_num + 1))
                            logger.info(f"Found table with title: {title} on page {page_num + 1}")
                            
                doc.close()
                
            return results
            
        except Exception as e:
            logger.error(f"Error extracting tables from section: {e}")
            return []

    def extract_with_camelot(self, pdf_path: str, page_num: int) -> List:
        try:
            # 격자 형식 시도
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_num),
                flavor='lattice'
            )
            # 격자 형식으로 추출 실패시 스트림 형식 시도
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

class ExcelWriter:
    @staticmethod
    def save_to_excel(sections_data: Dict[str, List[Tuple[str, pd.DataFrame, int]]], output_path: str):
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for section, tables in sections_data.items():
                    if not tables:
                        continue
                        
                    sheet_name = section.replace("[", "").replace("]", "")
                    current_row = 0
                    
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
                        
                        # 하이라이트 스타일링
                        yellow_fill = PatternFill(
                            start_color='FFFF00',
                            end_color='FFFF00',
                            fill_type='solid'
                        )
                        
                        if '변경사항' in df.columns:
                            for idx, row in enumerate(df.itertuples(), start=1):
                                if getattr(row, '변경사항') == '추가':
                                    for col in range(1, len(df.columns) + 1):
                                        cell = worksheet.cell(row=current_row + 2 + idx, column=col)
                                        cell.fill = yellow_fill
                        
                        current_row += len(df) + 5

            logger.info(f"Successfully saved tables to {output_path}")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")

def main():
    try:
        pdf_path = "input.pdf"
        output_path = "output.xlsx"
        
        if not os.path.exists(pdf_path):
            logger.error("PDF file not found")
            return

        # 문서 분석기 초기화
        logger.info("문서 분석 시작...")
        document_analyzer = DocumentAnalyzer()
        table_extractor = TableExtractor()
        
        # PDF 파일 열기
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        doc.close()

        # 전체 문서를 하나의 섹션으로 처리
        sections_data = {
            "전체": table_extractor.extract_tables_from_section(pdf_path, 0, total_pages)
        }

        # 결과 저장
        if any(sections_data.values()):
            logger.info("엑셀 파일 생성 중...")
            ExcelWriter.save_to_excel(sections_data, output_path)
            logger.info(f"처리 완료. 결과가 {output_path}에 저장되었습니다.")
        else:
            logger.error("추출된 표가 없습니다.")

    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        main()
        logger.info("프로그램이 성공적으로 완료되었습니다.")
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {str(e)}")