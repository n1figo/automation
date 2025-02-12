import os
import requests
import re
import PyPDF2
import camelot
import pandas as pd
import fitz
import logging
import unittest
from concurrent.futures import ProcessPoolExecutor, as_completed
from transformers import AutoModelForCausalLM, AutoTokenizer # transformers 라이브러리 관련 import 추가

# Streamlit UI 관련 라이브러리 (UI 개발 시 활성화)
# import streamlit as st

# 설정 변수 (config.ini 또는 config.yaml 파일로 관리하는 것이 이상적)
MODEL_PATH_MISTRAL = "/workspaces/automation/models/mistral-7b" # 미스트랄 모델 저장 경로 변경
MODEL_URL_MISTRAL = "" # 미스트랄 모델은 Hugging Face Hub에서 직접 다운로드하므로 URL 불필요
MODEL_NAME_MISTRAL = "mistralai/Mistral-7B-v0.1" # Hugging Face Hub 모델 이름
INPUT_PDF_PATH = "/workspaces/automation/data/input/0211/KB Yes!365 건강보험(세만기)(무배당)(25.01)_0214_요약서_v1.1.pdf"
OUTPUT_EXCEL_PATH = "/workspaces/automation/tests/test_data/output/extracted_tables_mistral.xlsx" # 출력 파일명 변경 (Llama 2와 구분)
SEARCH_TERM_INITIAL = "나. 보험금"
SEARCH_TERMS = [
    "상해관련 특별약관",
    "질병관련 특별약관",
    "상해및질병관련특별약관"
]
LLM_CONTEXT_SIZE = 4096
LLM_THREADS = 4
MAX_CHUNK_SIZE = 2000

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 모델 다운로드 및 로드 (LLM 관련 기능 활성화 시 사용)
def load_llm_model(model_path, model_url, model_name="mistralai/Mistral-7B-v0.1"):
    """LLM 모델을 로드하거나 다운로드합니다. (Mistral 모델 지원)"""
    if not os.path.exists(model_path):
        os.makedirs(os.path.dirname(model_path), exist_ok=True)
        logging.info(f"Downloading model from Hugging Face Hub: {model_name} to {model_path} ...")
        try:
            # Hugging Face Transformers를 사용하여 모델 직접 다운로드 및 저장 (모델 파일 직접 다운로드 방식)
            tokenizer = AutoTokenizer.from_pretrained(model_name) # 토크나이저 먼저 다운로드
            model = AutoModelForCausalLM.from_pretrained(model_name) # 모델 다운로드
            tokenizer.save_pretrained(model_path) # 토크나이저 저장
            model.save_pretrained(model_path) # 모델 저장

            logging.info("Model download complete.")
        except requests.exceptions.RequestException as e: # requests 에러 예외 처리 추가 (requests 관련 에러만 catch)
            logging.error(f"모델 다운로드 실패 (Hugging Face Hub): {e}")
            return None
        except Exception as e: # transformers 관련 에러 및 기타 예외 처리
            logging.error(f"모델 로드 실패 (Transformers): {e}")
            return None

    try:
        # 로컬에 저장된 모델 로드 (transformers 사용)
        tokenizer = AutoTokenizer.from_pretrained(model_path)
        model = AutoModelForCausalLM.from_pretrained(model_path)
        logging.info("Mistral 모델 로드 성공 (Transformers).")
        return model, tokenizer # 모델과 토크나이저 함께 반환
    except Exception as e:
        logging.error(f"Mistral 모델 로드 실패 (Transformers): {e}")
        return None, None # 모델, 토크나이저 로드 실패 시 None, None 반환


llm_mistral, tokenizer_mistral = load_llm_model(MODEL_PATH_MISTRAL, MODEL_URL_MISTRAL, MODEL_NAME_MISTRAL) # 모델 로드 및 토크나이저 저장
if llm_mistral: # 모델 로드 성공 시에만 다음 단계 진행
    print("Mistral 모델 로드 완료!")
else:
    print("Mistral 모델 로드 실패!")

class PDFProcessor:
    """PDF 문서 처리 및 분석을 위한 클래스"""

    def __init__(self, file_path):
        """PDFProcessor 초기화"""
        self.file_path = file_path
        self.doc = None # fitz 문서 객체
        self.pdf_reader = None # PyPDF2 PDF reader 객체

    def open_pdf_document(self):
        """PDF 문서를 열고 fitz 및 PyPDF2 객체를 초기화합니다."""
        try:
            self.doc = fitz.open(self.file_path)
            with open(self.file_path, "rb") as pdf_file:
                self.pdf_reader = PyPDF2.PdfReader(pdf_file)
            logging.info(f"PDF 문서 '{self.file_path}' 로드 완료.")
        except Exception as e:
            logging.error(f"PDF 문서 열기 실패: {e}")
            self.doc = None
            self.pdf_reader = None

    def normalize_text(self, text):
        """텍스트 정규화: 공백 문자 제거"""
        if not isinstance(text, str):
            text = str(text)
        return re.sub(r'\s+', '', text)

    def is_header_row(self, row, header=["보장명", "지급사유", "지급금액"]):
        """표 헤더 행 여부 판별"""
        try:
            if len(row) >= 4:
                cells = [self.normalize_text(row[i]) for i in range(1, 4)]
            elif len(row) == 3:
                cells = [self.normalize_text(row[i]) for i in range(0, 3)]
            else:
                return False
            norm_header = [self.normalize_text(h) for h in header]
            logging.debug(f"[DEBUG] Checking row: {cells} against expected: {norm_header}") # 디버깅 로그
            return cells == norm_header
        except Exception as e:
            logging.error(f"[ERROR] is_header_row 검사 중 오류 발생: {e}")
            return False

    def drop_redundant_header(self, df, header=["보장명", "지급사유", "지급금액"]):
        """표 데이터프레임에서 불필요한 헤더 행 제거"""
        keep_rows = []
        for idx, row in df.iterrows():
            if self.is_header_row(row, header):
                logging.info(f"[INFO] 불필요한 헤더 행 {idx} 제거") # 정보 로그
            else:
                keep_rows.append(idx)
        return df.loc[keep_rows]

    def page_has_highlight(self, page_no):
        """페이지에 형광펜 주석이 있는지 확인"""
        try:
            page = self.doc.load_page(page_no)
            annots = page.annots()
            if annots:
                for annot in annots:
                    if annot.type[0] == 8: # Type 8 is highlight annotation
                        return True
        except Exception as ex:
            logging.error(f"[ERROR] 주석 처리 오류: {ex}")
        return False

    def get_page_footer(self, page_no):
        """페이지 하단 텍스트 (페이지 번호 등) 추출"""
        try:
            page = self.doc.load_page(page_no)
            text = page.get_text("text")
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            return lines[-1] if lines else ""
        except Exception as e:
            logging.error(f"[ERROR] 페이지 Footer 추출 오류: {e}")
            return ""

    def split_text_into_chunks(self, text, max_chunk_size=MAX_CHUNK_SIZE):
        """텍스트를 청크로 분할 (LLM 처리 효율 향상)"""
        sentences = text.split('.')
        chunks = []
        current_chunk = []
        current_size = 0

        for sentence in sentences:
            sentence = sentence.strip() + '.'
            sentence_size = len(sentence)

            if current_size + sentence_size > max_chunk_size:
                chunks.append(' '.join(current_chunk))
                current_chunk = [sentence]
                current_size = sentence_size
            else:
                current_chunk.append(sentence)
                current_size += sentence_size

        if current_chunk:
            chunks.append(' '.join(current_chunk))
        return chunks

    def find_section_ranges(self, llm, tokenizer, doc, start_page, end_page, sections):
        """
        LLM (Mistral)을 사용하여 각 섹션의 시작과 끝 페이지를 찾습니다.
        """
        section_ranges = {}
        if llm is None or tokenizer is None: # 모델 또는 토크나이저 None 체크 추가
            logging.warning("LLM 모델 또는 토크나이저가 로드되지 않았습니다. 섹션 범위 식별 기능을 사용할 수 없습니다.")
            return section_ranges # LLM 또는 토크나이저 없을 경우 빈 딕셔너리 반환

        for section in sections:
            section_found = False
            for page_num in range(start_page, end_page + 1):
                page = doc.load_page(page_num)
                text = page.get_text("text")
                chunks = self.split_text_into_chunks(text)

                for chunk in chunks:
                    prompt_text = f"""
                    이 텍스트에서 '{section}'이 시작되거나 끝나는지 확인해주세요:

                    {chunk}

                    다음 중 하나로만 답변해주세요:
                    - 시작
                    - 끝
                    - 해당없음
                    """
                    try:
                        # transformers 파이프라인 대신 직접 추론 코드 사용
                        input_ids = tokenizer.encode(prompt_text, return_tensors="pt") # 텍스트 토큰화
                        output = llm.generate(input_ids, max_length=50, num_return_sequences=1) # 텍스트 생성 (추론)
                        response_text = tokenizer.decode(output[0], skip_special_tokens=True) # 생성된 텍스트 디코딩

                        if "시작" in response_text.lower():
                            if section not in section_ranges:
                                section_ranges[section] = {"start": page_num + 1}
                                section_found = True
                        elif "끝" in response_text.lower() and section_found:
                            section_ranges[section]["end"] = page_num + 1
                    except Exception as e:
                        logging.error(f"[ERROR] LLM 처리 중 오류 발생: {e}")
                        continue
        return section_ranges

    def extract_tables_from_page(self, page_num, term):
        """특정 페이지에서 표 추출 및 데이터프레임 반환 (병렬 처리 단위 작업)"""
        page_str = str(page_num + 1) # Camelot은 1-based 페이지 번호 사용
        try:
            tables = camelot.read_pdf(self.file_path, pages=page_str, flavor="lattice")
            combined_df = pd.DataFrame()

            for idx, table in enumerate(tables):
                suffix = f" - P{page_str}T{idx+1}"
                table_df = table.df.copy()

                footer_text = self.get_page_footer(page_num)

                table_df.insert(0, "PDF페이지", f"{page_str}")
                table_df.insert(0, "출처 페이지", footer_text)
                table_df.insert(0, "Source", term + suffix)

                table_df = self.drop_redundant_header(table_df)
                combined_df = pd.concat([combined_df, table_df], ignore_index=True)
            return combined_df
        except Exception as e:
            logging.error(f"[ERROR] 페이지 {page_num+1} 표 추출 오류: {e}")
            return pd.DataFrame() # 오류 발생 시 빈 데이터프레임 반환

    def extract_tables_parallel(self, term, pages):
        """병렬 처리를 사용하여 여러 페이지에서 표 추출"""
        combined_df = pd.DataFrame()
        page_dfs = [] # 각 페이지의 데이터프레임을 저장할 리스트

        with ProcessPoolExecutor(max_workers=4) as executor: # CPU 코어 수에 맞춰 max_workers 조정
            futures = [executor.submit(self.extract_tables_from_page, page_num, term) for page_num in pages]
            for future in as_completed(futures):
                page_df = future.result()
                if not page_df.empty: # 빈 데이터프레임이 아닌 경우만 추가
                    page_dfs.append(page_df)

        if page_dfs:
            combined_df = pd.concat(page_dfs, ignore_index=True) # 모든 페이지의 데이터프레임 결합
        return combined_df


    def process_pdf(self, llm, tokenizer, search_terms, output_file): # 토크나이저 파라미터 추가
        """PDF 처리 메인 함수"""
        if self.doc is None or self.pdf_reader is None:
            logging.error("PDF 문서가 로드되지 않았습니다. open_pdf_document()를 먼저 호출하세요.")
            return

        total_pages = len(self.pdf_reader.pages)
        start_page = None

        # 1. "나. 보험금" 검색 (기존 코드 유지)
        for i, page in enumerate(self.pdf_reader.pages):
            text = page.extract_text()
            if text and self.normalize_text(SEARCH_TERM_INITIAL) in self.normalize_text(text):
                start_page = i
                logging.info(f"'{SEARCH_TERM_INITIAL}' 용어 페이지 {i+1}에서 발견") # 정보 로그
                break

        if start_page is None:
            logging.warning(f"'{SEARCH_term_initial}' 용어 발견 실패.") # 경고 로그
            return

        # 2. LLM으로 섹션 범위 찾기 (Mistral 모델 사용)
        section_ranges = self.find_section_ranges(llm, tokenizer, doc, start_page, total_pages - 1, search_terms) # 토크나이저 전달

        # 섹션 범위 출력 (로그 레벨 INFO)
        logging.info("\n=== 섹션별 페이지 범위 ===")
        for section, range_info in section_ranges.items():
            start = range_info.get("start", "?")
            end = range_info.get("end", "?")
            logging.info(f"{section}: {start}페이지 ~ {end}페이지")

        # 3. 하이라이트 검색 (전체 섹션 범위 내에서)
        highlight_start = min(range_info["start"] for range_info in section_ranges.values() if "start" in range_info)
        highlight_end = max(range_info["end"] for range_info in section_ranges.values() if "end" in range_info)

        highlight_pages = set()
        for i in range(highlight_start - 1, highlight_end):
            if self.page_has_highlight(i):
                highlight_pages.add(i)
                if i - 1 >= highlight_start - 1:
                    highlight_pages.add(i - 1)
                if i + 1 < highlight_end:
                    highlight_pages.add(i + 1)

        if highlight_pages:
            highlight_pages_sorted = sorted(list(highlight_pages))
            hp_str = ",".join(str(p+1) for p in highlight_pages_sorted)
            logging.info(f"\n하이라이트가 있는 페이지(및 전후): {hp_str}") # 정보 로그

        # 4. 표 추출 및 엑셀 저장 (페이지 번호 포함, 병렬 처리 적용)
        excel_writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        results = {} # 표 추출 페이지 결과를 저장할 딕셔너리 (기존 코드 결과 재사용)
        for term, range_info in section_ranges.items(): # 섹션 범위 정보를 순회
            if "start" in range_info and "end" in range_info: # 섹션 시작과 끝 페이지가 정의된 경우
                pages = range(range_info["start"] - 1, range_info["end"]) # 추출할 페이지 번호 범위 생성 (0-indexed)
                if pages: # 추출할 페이지가 있는 경우
                    combined_df = self.extract_tables_parallel(term, pages) # 병렬 표 추출 함수 호출
                    if not combined_df.empty: # 추출된 표가 있는 경우
                        sheet_name = term.replace(" ", "")[:31]
                        combined_df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
                        results[term] = [p+1 for p in pages] # 결과 딕셔너리에 페이지 번호 저장 (1-indexed)
                    else:
                         logging.warning(f"[{term}] 섹션에서 추출된 표가 없습니다.") # 경고 로그
                else:
                    logging.warning(f"[{term}] 추출할 페이지 범위가 유효하지 않습니다.") # 경고 로그
            else:
                logging.warning(f"[{term}] 섹션 범위 정보가 불완전합니다.") # 경고 로그


        excel_writer.close()
        logging.info(f"\nExcel 파일 저장 완료: {output_file}") # 정보 로그


def main():
    """PDF 처리 자동화 메인 함수"""
    logging.info("PDF 약관 분석 자동화 시작") # 시작 로그

    pdf_processor = PDFProcessor(INPUT_PDF_PATH)
    pdf_processor.open_pdf_document()

    if pdf_processor.doc and pdf_processor.pdf_reader and llm_mistral and tokenizer_mistral: # 모델, 토크나이저 로드 성공 여부 체크 추가
        pdf_processor.process_pdf(llm_mistral, tokenizer_mistral, SEARCH_TERMS, OUTPUT_EXCEL_PATH) # 모델, 토크나이저 전달
    else:
        logging.error("PDF 문서 또는 Mistral 모델 로드 실패, 프로그램 종료.") # 에러 로그

    logging.info("PDF 약관 분석 자동화 완료") # 완료 로그


# Streamlit UI (UI 개발 시 main 함수 대신 streamlit_app 함수 사용)
# def streamlit_app():
#     st.title("보험약관 분석 자동화")
#     uploaded_file = st.file_uploader("PDF 파일 업로드", type="pdf")
#     if uploaded_file is not None:
#         file_path_streamlit = f"/tmp/{uploaded_file.name}" # 임시 파일 경로
#         with open(file_path_streamlit, "wb") as f:
#             f.write(uploaded_file.getbuffer())
#         st.success("파일 업로드 완료!")
#
#         if st.button("분석 시작"):
#             with st.spinner("PDF 분석 중..."):
#                 pdf_processor_streamlit = PDFProcessor(file_path_streamlit)
#                 pdf_processor_streamlit.open_pdf_document()
#                 if pdf_processor_streamlit.doc and pdf_processor_streamlit.pdf_reader and llm_mistral and tokenizer_mistral:
#                     output_excel_streamlit_path = "/tmp/extracted_tables_streamlit_mistral.xlsx" # 임시 출력 경로 (파일명 변경)
#                     pdf_processor_streamlit.process_pdf(llm_mistral, tokenizer_mistral, SEARCH_TERMS, output_excel_streamlit_path)
#                     st.success(f"Excel 파일 생성 완료: {output_excel_streamlit_path}")
#
#                     # 엑셀 파일 다운로드 버튼 (선택 사항)
#                     # with open(output_excel_streamlit_path, "rb") as f:
#                     #     st.download_button(
#                     #         label="엑셀 파일 다운로드",
#                     #         data=f,
#                     #         file_name="extracted_tables.xlsx",
#                     #         mime="application/vnd.ms-excel"
#                     #     )
#                 else:
#                     st.error("PDF 문서 분석 실패.")


if __name__ == "__main__":
    main() # CLI 실행 (UI 개발 시 streamlit_app()으로 변경)
    # streamlit_app() # Streamlit UI 실행 (UI 개발 환경에서 활성화)

