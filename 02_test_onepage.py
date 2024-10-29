import os
import logging
import fitz
import cv2
import numpy as np
from PIL import Image
import pandas as pd
from paddleocr import PPStructure
from langchain_community.llms import LlamaCpp
from langchain.callbacks.manager import CallbackManager
from langchain.callbacks.streaming_stdout import StreamingStdOutCallbackHandler
import time

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    datefmt='%Y/%m/%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class SinglePageAnalyzer:
    def __init__(self, model_path="models/llama-2-7b-chat.gguf"):
        """초기화"""
        logger.info("SinglePageAnalyzer 초기화 시작")
        try:
            # PaddleOCR 초기화 - 최적화된 설정
            self.table_engine = PPStructure(
                show_log=True,
                table=True,
                ocr=True,
                layout=True,
                lang='en',
                layout_model_dir=None,
                det_model_dir=None,
                rec_model_dir=None,
                use_angle_cls=False,  # 방향 분류 비활성화
                cls_model_dir=None,
                recovery=True,
                page_num=0,
                # 성능 최적화 옵션
                use_gpu=False,
                enable_mkldnn=True,
                cpu_threads=4,
                det_db_score_mode='slow',  # 정확도 우선
                det_limit_side_len=2880,  # 고해상도 지원
                det_db_box_thresh=0.5,  # 검출 임계값
                rec_batch_num=1,  # 배치 크기
                # 테이블 관련 설정
                merge_no_span_structure=True,
                table_max_len=488,
                table_algorithm='TableAttn'
            )
            logger.info("PaddleOCR 초기화 완료")

            # LLAMA 초기화
            callback_manager = CallbackManager([StreamingStdOutCallbackHandler()])
            self.llm = LlamaCpp(
                model_path=model_path,
                callback_manager=callback_manager,
                temperature=0.1,
                max_tokens=2000,
                n_ctx=2048,
                n_threads=4  # CPU 스레드 수 지정
            )
            logger.info("LLAMA 모델 초기화 완료")

            # 색상 범위 정의 (HSV) - 민감도 조정
            self.color_ranges = {
                'yellow': [(20, 80, 80), (45, 255, 255)],  # 노란색 범위 확대
                'green': [(35, 80, 80), (85, 255, 255)],   # 녹색 범위 확대
                'blue': [(95, 80, 80), (145, 255, 255)]    # 파란색 범위 확대
            }
            logger.info("색상 범위 정의 완료")
            
        except Exception as e:
            logger.error(f"초기화 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def preprocess_image(self, image):
        """이미지 전처리"""
        try:
            logger.info("이미지 전처리 시작")
            start_time = time.time()
            
            # 이미지 크기 확인
            height, width = image.shape[:2]
            logger.info(f"입력 이미지 크기: {width}x{height}")

            # 노이즈 제거
            denoised = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
            
            # 대비 향상
            lab = cv2.cvtColor(denoised, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
            cl = clahe.apply(l)
            enhanced = cv2.merge((cl,a,b))
            
            result = cv2.cvtColor(enhanced, cv2.COLOR_LAB2BGR)
            
            process_time = time.time() - start_time
            logger.info(f"이미지 전처리 완료 (소요시간: {process_time:.2f}초)")
            return result
            
        except Exception as e:
            logger.error(f"이미지 전처리 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def analyze_color(self, cell_img):
        """셀 이미지의 하이라이트 색상 분석"""
        try:
            hsv = cv2.cvtColor(cell_img, cv2.COLOR_BGR2HSV)
            detected_colors = []
            
            for color_name, (lower, upper) in self.color_ranges.items():
                mask = cv2.inRange(hsv, np.array(lower), np.array(upper))
                pixel_count = np.sum(mask > 0)
                coverage = pixel_count / (mask.shape[0] * mask.shape[1])
                
                # 커버리지 임계값 조정 (25%로 낮춤)
                if coverage > 0.25:
                    detected_colors.append({
                        'color': color_name,
                        'coverage': coverage
                    })
                    logger.debug(f"감지된 색상: {color_name} (커버리지: {coverage:.2f})")
            
            # 커버리지가 가장 높은 색상만 반환
            if detected_colors:
                max_coverage = max(detected_colors, key=lambda x: x['coverage'])
                return [max_coverage['color']]
            return []
            
        except Exception as e:
            logger.error(f"색상 분석 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def extract_page_image(self, pdf_path, page_num):
        """PDF에서 특정 페이지 이미지 추출"""
        try:
            logger.info(f"PDF 페이지 {page_num} 이미지 추출 시작")
            start_time = time.time()
            
            doc = fitz.open(pdf_path)
            page = doc[page_num - 1]
            
            # 고해상도 이미지 추출을 위한 매트릭스 설정
            zoom_x = 2.0
            zoom_y = 2.0
            mat = fitz.Matrix(zoom_x, zoom_y)
            pix = page.get_pixmap(matrix=mat)
            
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            
            process_time = time.time() - start_time
            logger.info(f"PDF 페이지 이미지 추출 완료 (소요시간: {process_time:.2f}초)")
            return np.array(img)
            
        except Exception as e:
            logger.error(f"PDF 페이지 이미지 추출 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def validate_with_llama(self, table_data):
        """LLAMA를 사용하여 결과 검증"""
        try:
            logger.info("LLAMA 검증 시작")
            start_time = time.time()
            
            prompt = f"""
            다음은 보험 약관의 표에서 하이라이트된 셀들의 정보입니다:
            {table_data}

            이 결과를 분석하고 다음을 확인해주세요:
            1. 하이라이트된 셀들이 논리적으로 연결되어 있나요?
            2. 표의 구조상 위치가 적절한가요?
            3. 감지된 색상이 문맥에 맞나요?
            4. 각 셀의 내용이 보험 약관의 맥락에서 의미가 있나요?

            분석 결과를 다음 JSON 형식으로 제공해주세요:
            {
                "logical_connection": true/false,
                "position_appropriate": true/false,
                "color_context_match": true/false,
                "content_relevant": true/false,
                "comments": "상세 설명"
            }
            """
            
            response = self.llm.predict(prompt)
            
            process_time = time.time() - start_time
            logger.info(f"LLAMA 검증 완료 (소요시간: {process_time:.2f}초)")
            return response
            
        except Exception as e:
            logger.error(f"LLAMA 검증 중 오류 발생: {str(e)}", exc_info=True)
            raise

    def process_page(self, pdf_path, page_num=59):
        """특정 페이지 처리"""
        try:
            total_start_time = time.time()
            logger.info(f"페이지 {page_num} 처리 시작")
            
            # 페이지 이미지 추출
            image = self.extract_page_image(pdf_path, page_num)
            processed_image = self.preprocess_image(image)
            
            # PaddleOCR로 표 분석
            logger.info("PaddleOCR 표 분석 시작")
            result = self.table_engine(processed_image)
            tables_data = []
            
            for idx, region in enumerate(result):
                if region['type'] == 'table':
                    logger.info(f"표 {idx+1} 처리 중")
                    table_img = region['img']
                    cells = region['cells']
                    
                    table_data = {
                        'table_index': idx,
                        'highlighted_cells': []
                    }
                    
                    for cell in cells:
                        bbox = cell['bbox']
                        cell_img = table_img[bbox[1]:bbox[3], bbox[0]:bbox[2]]
                        colors = self.analyze_color(cell_img)
                        
                        if colors:
                            cell_info = {
                                'row': cell['row_idx'],
                                'col': cell['col_idx'],
                                'text': cell['text'],
                                'colors': colors
                            }
                            table_data['highlighted_cells'].append(cell_info)
                            logger.info(f"하이라이트 감지: {cell_info}")
                    
                    # LLAMA 검증
                    if table_data['highlighted_cells']:
                        validation = self.validate_with_llama(table_data)
                        table_data['validation'] = validation
                        tables_data.append(table_data)
            
            total_time = time.time() - total_start_time
            logger.info(f"페이지 {page_num} 처리 완료 (총 소요시간: {total_time:.2f}초)")
            return tables_data
            
        except Exception as e:
            logger.error(f"페이지 처리 중 오류 발생: {str(e)}", exc_info=True)
            raise

def save_to_excel(tables_data, output_path):
    """결과를 Excel 파일로 저장"""
    try:
        logger.info(f"Excel 파일 저장 시작: {output_path}")
        start_time = time.time()
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for table_data in tables_data:
                if table_data['highlighted_cells']:
                    df = pd.DataFrame(table_data['highlighted_cells'])
                    sheet_name = f"Table_{table_data['table_index']}"
                    
                    # 데이터 저장
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 검증 결과 저장
                    validation_df = pd.DataFrame([{'validation': table_data['validation']}])
                    validation_df.to_excel(writer, 
                                        sheet_name=sheet_name, 
                                        startrow=len(df) + 2, 
                                        index=False)
        
        process_time = time.time() - start_time
        logger.info(f"Excel 파일 저장 완료: {output_path} (소요시간: {process_time:.2f}초)")
        
    except Exception as e:
        logger.error(f"Excel 저장 중 오류 발생: {str(e)}", exc_info=True)
        raise

def main():
    # 설정
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    output_path = "page_59_analysis.xlsx"
    llama_model_path = "models/llama-2-7b-chat.gguf"
    
    try:
        total_start_time = time.time()
        logger.info("프로그램 시작")
        
        # 분석기 초기화
        analyzer = SinglePageAnalyzer(model_path=llama_model_path)
        
        # 59페이지 분석
        logger.info("59페이지 분석 시작")
        results = analyzer.process_page(pdf_path, page_num=59)
        
        # 결과 저장
        if results:
            save_to_excel(results, output_path)
            logger.info(f"분석 완료. 결과가 {output_path}에 저장되었습니다.")
        else:
            logger.warning("감지된 하이라이트가 없습니다.")
        
        total_time = time.time() - total_start_time
        logger.info(f"프로그램 종료 (총 소요시간: {total_time:.2f}초)")
        
    except Exception as e:
        logger.error(f"실행 중 오류 발생: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()