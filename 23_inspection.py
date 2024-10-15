import fitz  # PyMuPDF
from bs4 import BeautifulSoup
from difflib import SequenceMatcher
import re
import sys
import logging
import argparse
import os
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
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()

            # mhtml 파일을 로드
            page.goto(f'file://{os.path.abspath(html_path)}')

            # 페이지가 로드될 때까지 대기 (필요한 경우)
            # page.wait_for_load_state('networkidle')

            html_content = page.content()
            browser.close()

            soup = BeautifulSoup(html_content, 'html.parser')
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
            return soup
    except Exception as e:
        logging.error(f"HTML 내용을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def extract_relevant_tables(soup):
    try:
        tables = soup.find_all('table')
        logging.info(f"총 {len(tables)}개의 테이블을 발견했습니다.")
        relevant_tables = []

        for idx, table in enumerate(tables):
            logging.debug(f"테이블 {idx+1}/{len(tables)} 처리 중...")
            # 테이블 이전의 모든 요소를 찾습니다.
            previous_elements = table.find_all_previous()
            logging.debug(f"테이블 {idx+1}의 이전 요소 개수: {len(previous_elements)}")
            title_found = False
            for elem in previous_elements:
                text = elem.get_text().strip()
                logging.debug(f"이전 요소 태그: {elem.name}, 텍스트: '{text}'")
                if any(keyword in text for keyword in ['상해관련 특별약관', '상해 관련 특별약관', '상해관련', '상해 관련', '상해',
                                                       '질병관련 특별약관', '질병 관련 특별약관', '질병관련', '질병 관련', '질병']):
                    title = text
                    logging.info(f"테이블 {idx+1}의 제목: '{title}' (태그: {elem.name})")
                    relevant_tables.append((title, table))
                    title_found = True
                    break
            if not title_found:
                logging.debug(f"테이블 {idx+1}에서 관련 제목을 찾지 못했습니다.")

        logging.info(f"관련 테이블 {len(relevant_tables)}개를 추출했습니다.")
        return relevant_tables
    except Exception as e:
        logging.error(f"HTML 테이블을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def extract_relevant_pdf_sections(pdf_path, section_titles):
    try:
        doc = fitz.open(pdf_path)
        content = ''
        for page in doc:
            content += page.get_text()

        logging.debug(f"PDF 전체 내용 길이: {len(content)}")
        sections = {}
        for title in section_titles:
            # 제목을 기준으로 텍스트를 분할합니다.
            pattern = re.escape(title) + r'(.*?)(?=\n[A-Za-z가-힣]+\n|$)'
            logging.debug(f"'{title}'에 대한 패턴: {pattern}")
            matches = re.findall(pattern, content, re.DOTALL)
            if matches:
                sections[title] = matches[0]
                logging.info(f"'{title}' 섹션을 추출했습니다. 길이: {len(matches[0])}")
            else:
                logging.warning(f"PDF에서 '{title}' 섹션을 찾을 수 없습니다.")
                sections[title] = ''
        return sections
    except Exception as e:
        logging.error(f"PDF 내용을 추출하는 데 실패했습니다: {e}")
        sys.exit(1)

def preprocess_text(text):
    # 공백 및 특수문자 제거
    text = re.sub(r'[^\w]', '', text)  # 특수문자 제거
    text = re.sub(r'\s+', '', text)    # 공백 제거
    return text

def compare_tables(soup, html_tables, pdf_sections, similarity_threshold):
    total_mismatches = 0

    for title, table in html_tables:
        html_rows = table.find_all('tr')
        pdf_text = pdf_sections.get(title, '')
        pdf_lines = pdf_text.split('\n')

        logging.debug(f"'{title}' 테이블의 행 개수: {len(html_rows)}")
        logging.debug(f"'{title}' 섹션의 PDF 라인 수: {len(pdf_lines)}")

        # 헤더 행 찾기 또는 생성
        header_row = None
        thead = table.find('thead')
        if thead:
            header_row = thead.find('tr')
        else:
            first_row = html_rows[0]
            th_cells = first_row.find_all('th')
            if th_cells:
                header_row = first_row
            else:
                # 헤더가 없으면 생성
                header_row = soup.new_tag('tr')
                for _ in html_rows[0].find_all('td'):
                    new_th = soup.new_tag('th')
                    header_row.append(new_th)
                table.insert(0, header_row)
                html_rows.insert(0, header_row)

        # '검수결과' 헤더 추가
        new_th = soup.new_tag('th')
        new_th.string = '검수결과'
        header_row.append(new_th)

        for row_idx, tr in enumerate(html_rows[1:], start=1):  # 헤더를 제외하고 시작 인덱스를 1로 설정
            html_cells = tr.find_all(['td', 'th'])
            html_line = ' '.join(tr.stripped_strings)
            is_matched = False

            for pdf_line in pdf_lines:
                similarity = SequenceMatcher(None, preprocess_text(html_line), preprocess_text(pdf_line)).ratio()
                if similarity >= similarity_threshold:
                    is_matched = True
                    break
            if not is_matched:
                total_mismatches += 1
                # 각 셀을 검사하여 어떤 부분이 불일치하는지 확인
                for cell in html_cells:
                    cell_text = ' '.join(cell.stripped_strings)
                    cell_matched = False
                    for pdf_line in pdf_lines:
                        cell_similarity = SequenceMatcher(None, preprocess_text(cell_text), preprocess_text(pdf_line)).ratio()
                        if cell_similarity >= similarity_threshold:
                            cell_matched = True
                            break
                    if not cell_matched:
                        # 불일치하는 셀에 스타일 추가 (초록색 음영)
                        cell['style'] = cell.get('style', '') + 'background-color: lightgreen;'
                # 검수결과 열에 '불일치' 추가
                new_td = soup.new_tag('td')
                new_td.string = '불일치'
                tr.append(new_td)
            else:
                # 검수결과 열에 빈 값 추가
                new_td = soup.new_tag('td')
                new_td.string = ''
                tr.append(new_td)
    return total_mismatches

def main(similarity_threshold=0.95, log_level='INFO'):
    numeric_level = getattr(logging, log_level.upper(), None)
    if not isinstance(numeric_level, int):
        print(f"유효하지 않은 로그 레벨입니다: {log_level}")
        sys.exit(1)
    logging.basicConfig(level=numeric_level)

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

        # HTML 콘텐츠 추출
        soup = extract_html_content(html_path)

        # HTML에서 필요한 테이블 추출
        html_tables = extract_relevant_tables(soup)
        if not html_tables:
            logging.error("HTML에서 필요한 테이블을 찾을 수 없습니다.")
            sys.exit(1)

        # PDF에서 해당하는 섹션 추출
        section_titles = [title for title, _ in html_tables]
        pdf_sections = extract_relevant_pdf_sections(pdf_path, section_titles)

        # 각 라인별로 비교
        total_mismatches = compare_tables(soup, html_tables, pdf_sections, similarity_threshold)

        # 수정된 HTML을 출력 폴더에 저장
        output_html_path = os.path.join(output_folder_path, os.path.basename(html_path))

        with open(output_html_path, 'w', encoding='utf-8') as file:
            file.write(str(soup))
        logging.info(f"수정된 HTML 파일을 '{output_html_path}'에 저장했습니다.")

        if total_mismatches > 0:
            logging.warning(f"{total_mismatches}개의 불일치하는 행이 발견되었습니다.")
        else:
            logging.info("모든 행이 일치합니다. PASS")
    else:
        logging.error("'요약서' PDF 또는 '보장내용' HTML 파일을 찾을 수 없습니다.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF와 HTML 문서를 비교합니다.")
    parser.add_argument('--threshold', type=float, default=0.95, help="유사도 임계값 (기본값: 0.95)")
    parser.add_argument('--loglevel', default='INFO', help="로그 레벨 설정 (예: DEBUG, INFO, WARNING, ERROR)")
    args = parser.parse_args()

    main(similarity_threshold=args.threshold, log_level=args.loglevel)
