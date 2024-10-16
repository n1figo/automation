import camelot
import pandas as pd
import numpy as np
import cv2
import os
import fitz  # PyMuPDF
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from bs4 import BeautifulSoup
from difflib import SequenceMatcher
import re
import sys
import logging
import argparse
from playwright.sync_api import sync_playwright

def ensure_inspection_upload_folder():
    folder_name = 'inspection upload'
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        logging.info(f"'{folder_name}' 폴더를 생성했습니다.")
    else:
        logging.info(f"'{folder_name}' 폴더가 이미 존재합니다.")
    return folder_name

def ensure_inspection_output_folder():
    folder_name = 'inspection output'
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        logging.info(f"'{folder_name}' 폴더를 생성했습니다.")
    else:
        logging.info(f"'{folder_name}' 폴더가 이미 존재합니다.")
    return folder_name

def find_files_in_folder(folder_path):
    pdf_files = {}
    html_files = {}
    for file_name in os.listdir(folder_path):
        lower_file_name = file_name.lower()
        file_path = os.path.join(folder_path, file_name)
        if lower_file_name.endswith('.pdf'):
            if '요약서' in lower_file_name:
                pdf_files['요약서'] = file_path
            elif '가입예시' in lower_file_name:
                pdf_files['가입예시'] = file_path
        elif (lower_file_name.endswith('.html') or
              lower_file_name.endswith('.htm') or
              lower_file_name.endswith('.mhtml')):
            if '보장내용' in lower_file_name:
                html_files['보장내용'] = file_path
            elif '가입예시' in lower_file_name:
                html_files['가입예시'] = file_path
    return pdf_files, html_files

def extract_html_content_with_playwright(html_path):
    try:
        logging.info(f"Playwright를 사용하여 '{html_path}' 파일을 처리합니다.")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()

            # mhtml 파일을 로드
            page.goto(f'file://{os.path.abspath(html_path)}')

            # 페이지가 로드될 때까지 대기
            page.wait_for_load_state('networkidle')

            html_content = page.content()
            browser.close()

            soup = BeautifulSoup(html_content, 'html.parser')
            logging.info(f"'{html_path}' 파일의 HTML 콘텐츠를 추출했습니다.")
            return soup
    except Exception as e:
        logging.error(f"Playwright를 사용하여 HTML 내용을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def extract_html_content(html_path):
    try:
        if html_path.lower().endswith('.mhtml'):
            # Playwright를 사용하여 mhtml 파일 처리
            soup = extract_html_content_with_playwright(html_path)
            return soup
        else:
            # 일반 HTML 파일 처리
            with open(html_path, 'r', encoding='utf-8') as file:
                html_content = file.read()
            soup = BeautifulSoup(html_content, 'html.parser')
            logging.info(f"'{html_path}' 파일의 HTML 콘텐츠를 추출했습니다.")
            return soup
    except Exception as e:
        logging.error(f"HTML 내용을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def extract_relevant_tables(soup):
    try:
        logging.info("HTML에서 모든 테이블을 추출합니다.")
        tables = soup.find_all('table')
        if not tables:
            logging.error("HTML에서 테이블을 찾을 수 없습니다.")
            sys.exit(1)
        logging.info(f"총 {len(tables)}개의 테이블을 추출했습니다.")
        # 각 테이블의 제목과 테이블을 튜플로 저장
        relevant_tables = []
        for table in tables:
            # 테이블 이전의 제목 추출 (h 태그)
            title_tag = table.find_previous(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
            title = title_tag.get_text(strip=True) if title_tag else "제목 없음"
            relevant_tables.append((title, table))
        return relevant_tables
    except Exception as e:
        logging.error(f"HTML 테이블을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return np.array(img)

def detect_highlights(image):
    hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
    
    # 여러 색상 범위 정의 (HSV)
    color_ranges = [
        ((20, 100, 100), (40, 255, 255)),  # 노란색
        ((100, 100, 100), (140, 255, 255)),  # 파란색
        ((125, 100, 100), (155, 255, 255))  # 보라색
    ]
    
    masks = []
    for lower, upper in color_ranges:
        mask = cv2.inRange(hsv, np.array(lower), np.array(upper))
        masks.append(mask)
    
    # 모든 마스크 결합
    combined_mask = np.zeros_like(masks[0])
    for mask in masks:
        combined_mask = cv2.bitwise_or(combined_mask, mask)
    
    kernel = np.ones((5,5), np.uint8)
    cleaned_mask = cv2.morphologyEx(combined_mask, cv2.MORPH_CLOSE, kernel)
    cleaned_mask = cv2.morphologyEx(cleaned_mask, cv2.MORPH_OPEN, kernel)
    
    contours, _ = cv2.findContours(cleaned_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    return contours

def get_highlight_regions(contours, image_height):
    regions = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        top = image_height - (y + h)
        bottom = image_height - y
        regions.append((top, bottom))
    return regions

def extract_tables_with_titles(pdf_path):
    logging.info("PDF에서 테이블과 제목을 추출합니다.")
    doc = fitz.open(pdf_path)
    tables_with_titles = []
    for page_number in range(len(doc)):
        page = doc[page_number]
        page_text = page.get_text("blocks")
        page_text_sorted = sorted(page_text, key=lambda x: x[1])  # y0를 기준으로 정렬 (위에서부터 아래로)
        titles = []
        for block in page_text_sorted:
            if block[6] == 0:  # 텍스트 블록인 경우
                text = block[4].strip()
                if text:
                    # 제목으로 간주할 만한 길이와 위치의 텍스트인지 판단 (필요에 따라 조건 조정)
                    if len(text) < 50:
                        titles.append((block[1], text))  # (y0, text)
        # 해당 페이지에서 테이블 추출
        tables = camelot.read_pdf(pdf_path, pages=str(page_number+1), flavor='lattice')
        for table in tables:
            # 테이블의 y 좌표 (상단)
            table_y = table._bbox[3]
            # 테이블 상단보다 위에 있는 제목 중 가장 가까운 것 선택
            title = "제목 없음"
            min_distance = float('inf')
            for y0, text in titles:
                distance = table_y - y0
                if 0 < distance < min_distance:
                    min_distance = distance
                    title = text
            tables_with_titles.append((title, table))
    logging.info(f"총 {len(tables_with_titles)}개의 테이블과 제목을 추출했습니다.")
    return tables_with_titles

def preprocess_text(text):
    # 공백 및 특수문자 제거
    text = re.sub(r'[^\w]', '', text)  # 특수문자 제거
    text = re.sub(r'\s+', '', text)    # 공백 제거
    return text

def compare_tables_and_generate_report(html_tables, pdf_tables_with_titles, similarity_threshold):
    total_mismatches = 0
    results = []  # 엑셀 출력을 위한 데이터 저장 리스트

    logging.info(f"총 {len(html_tables)}개의 HTML 테이블과 {len(pdf_tables_with_titles)}개의 PDF 테이블을 비교합니다.")

    for html_idx, (html_title, html_table) in enumerate(html_tables, start=1):
        logging.info(f"HTML 테이블 {html_idx}/{len(html_tables)}: '{html_title}' 비교 시작")
        # HTML 테이블 데이터 추출
        html_rows = html_table.find_all('tr')
        html_data = []
        for tr in html_rows:
            row = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            html_data.append(row)

        # PDF 테이블과 제목 매칭
        best_match_idx = None
        best_similarity = 0
        for pdf_idx, (pdf_title, pdf_table) in enumerate(pdf_tables_with_titles):
            similarity = SequenceMatcher(None, preprocess_text(html_title), preprocess_text(pdf_title)).ratio()
            if similarity > best_similarity:
                best_similarity = similarity
                best_match_idx = pdf_idx

        if best_match_idx is not None and best_similarity >= similarity_threshold:
            pdf_title, pdf_table = pdf_tables_with_titles.pop(best_match_idx)
            pdf_data = pdf_table.df.values.tolist()
            # 컬럼명 유지
            pdf_columns = pdf_data[0]
            pdf_data = pdf_data[1:]
            logging.info(f"'{html_title}'과 '{pdf_title}'를 비교합니다. (유사도: {best_similarity:.2f})")
        else:
            logging.warning(f"'{html_title}'에 매칭되는 PDF 테이블을 찾을 수 없습니다.")
            pdf_data = []
            pdf_columns = []

        # 결과를 저장하기 전에 제목을 추가
        results.append([f"제목: {html_title}"])
        max_rows = max(len(html_data), len(pdf_data))

        # 헤더 작성 (PDF 데이터의 컬럼명을 사용)
        html_header = html_data[0] if html_data else []
        pdf_header = pdf_columns if pdf_columns else []
        header = html_header + pdf_header + ['검수과정']

        # 헤더를 결과에 추가
        results.append(header)

        # 데이터 비교 및 저장
        for i in range(1, max_rows):
            html_row = html_data[i] if i < len(html_data) else [''] * len(html_header)
            pdf_row = pdf_data[i-1] if i-1 < len(pdf_data) else [''] * len(pdf_header)

            # 각 행의 데이터를 병합
            combined_row = html_row + pdf_row

            # 행 단위로 비교
            html_line = ''.join(html_row)
            pdf_line = ''.join(pdf_row)
            similarity = SequenceMatcher(None, preprocess_text(html_line), preprocess_text(pdf_line)).ratio()

            검수과정 = ''
            if similarity < similarity_threshold:
                검수과정 = '불일치'
                total_mismatches += 1

            combined_row.append(검수과정)
            results.append(combined_row)

            if i % 10 == 0:
                logging.debug(f"'{html_title}' 테이블의 {i}/{max_rows}번째 행 비교 완료")

        logging.info(f"테이블 '{html_title}' 비교 완료")

    return results, total_mismatches

def write_results_to_excel(data, output_path):
    try:
        wb = Workbook()
        ws = wb.active

        for row_idx, row in enumerate(data, start=1):
            ws.append(row)
            if row_idx % 100 == 0:
                logging.debug(f"{row_idx}개의 행을 엑셀에 기록했습니다.")

        wb.save(output_path)
        logging.info(f"검수 과정을 '{output_path}' 파일에 저장했습니다.")
    except Exception as e:
        logging.error(f"검수 과정을 엑셀 파일로 저장하는 데 실패했습니다: {e}")

def main(similarity_threshold=0.95, log_level='INFO'):
    numeric_level = getattr(logging, log_level.upper(), None)
    if not isinstance(numeric_level, int):
        print(f"유효하지 않은 로그 레벨입니다: {log_level}")
        sys.exit(1)
    logging.basicConfig(level=numeric_level, format='%(asctime)s - %(levelname)s - %(message)s')

    folder_path = ensure_inspection_upload_folder()
    output_folder_path = ensure_inspection_output_folder()

    # 폴더 내의 파일 리스트 출력
    files_in_folder = os.listdir(folder_path)
    if files_in_folder:
        logging.info("폴더 내의 파일 리스트:")
        for file_name in files_in_folder:
            logging.info(f"- {file_name}")
    else:
        logging.info("폴더가 비어 있습니다.")

    pdf_files, html_files = find_files_in_folder(folder_path)

    # '요약서' PDF와 '보장내용' HTML 파일 비교
    if '요약서' in pdf_files and '보장내용' in html_files:
        pdf_path = pdf_files['요약서']
        html_path = html_files['보장내용']

        logging.info(f"비교할 PDF 파일: {pdf_path}")
        logging.info(f"비교할 HTML 파일: {html_path}")

        # PDF에서 제목 추출
        doc = fitz.open(pdf_path)
        first_page = doc[0]
        page_text = first_page.get_text("text")
        pdf_title = page_text.strip().split('\n')[0]  # 첫 번째 줄을 제목으로 가정
        logging.info(f"PDF에서 제목을 추출했습니다: {pdf_title}")

        # HTML 콘텐츠 추출
        soup = extract_html_content(html_path)

        # HTML에서 모든 테이블 추출
        html_tables = extract_relevant_tables(soup)
        if not html_tables:
            logging.error("HTML에서 테이블을 찾을 수 없습니다.")
            sys.exit(1)

        # PDF에서 테이블과 제목 추출
        pdf_tables_with_titles = extract_tables_with_titles(pdf_path)
        if not pdf_tables_with_titles:
            logging.error("PDF에서 테이블을 찾을 수 없습니다.")
            sys.exit(1)

        # 테이블 비교 및 결과 생성
        comparison_results, total_mismatches = compare_tables_and_generate_report(html_tables, pdf_tables_with_titles, similarity_threshold)

        # 결과를 엑셀 파일로 저장
        output_excel_path = os.path.join(output_folder_path, '검수과정.xlsx')
        write_results_to_excel(comparison_results, output_excel_path)

        if total_mismatches > 0:
            logging.warning(f"{total_mismatches}개의 불일치가 발견되었습니다.")
        else:
            logging.info("모든 테이블이 일치합니다. PASS")
    else:
        logging.error("'요약서' PDF 또는 '보장내용' HTML 파일을 찾을 수 없습니다.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF와 HTML 문서를 비교하여 검수 과정을 자동화합니다.")
    parser.add_argument('--threshold', type=float, default=0.95, help="유사도 임계값 (기본값: 0.95)")
    parser.add_argument('--loglevel', default='INFO', help="로그 레벨 설정 (예: DEBUG, INFO, WARNING, ERROR)")
    args = parser.parse_args()

    main(similarity_threshold=args.threshold, log_level=args.loglevel)
