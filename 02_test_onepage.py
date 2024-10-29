```python
from paddleocr import PPStructure
import cv2
import numpy as np
import fitz
from PIL import Image
import pandas as pd
from langchain.llms import LlamaCpp
from langchain.callbacks.manager import CallbackManager
from langchain.callbacks.streaming_stdout import StreamingStdOutCallbackHandler
import os
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SinglePageAnalyzer:
    def __init__(self, model_path="models/llama-2-7b-chat.gguf"):
        """초기화"""
        # PaddleOCR 초기화
        self.table_engine = PPStructure(
            show_log=False,
            table=True,
            ocr=True,
            layout=True,
            lang='ko'
        )

        # LLAMA 초기화
        callback_manager = CallbackManager([StreamingStdOutCallbackHandler()])
        self.llm = LlamaCpp(
            model_path=model_path,
            callback_manager=callback_manager,
            temperature=0.1,
            max_tokens=2000,
            n_ctx=2048
        )

        # 색상 범위 정의
        self.color_ranges = {
            'yellow': [(20, 100, 100), (40, 255, 255)],
            'green': [(40, 100, 100), (80, 255, 255)],
            'blue': [(100, 100, 100), (140, 255, 255)]
        }

    def preprocess_image(self, image):
        """이미지 전처리"""
        # 노이즈 제거
        denoised = cv2.fastNlMeansDenoisingColored(image, None, 10, 10, 7, 21)
        
        # 대비 향상
        lab = cv2.cvtColor(denoised, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
        cl = clahe.apply(l)
        enhanced = cv2.merge((cl,a,b))
        
        return cv2.cvtColor(enhanced, cv2.COLOR_LAB2BGR)

    def analyze_color(self, cell_img):
        """셀 이미지의 하이라이트 색상 분석"""
        hsv = cv2.cvtColor(cell_img, cv2.COLOR_BGR2HSV)
        detected_colors = []
        
        for color_name, (lower, upper) in self.color_ranges.items():
            mask = cv2.inRange(hsv, np.array(lower), np.array(upper))
            pixel_count = np.sum(mask > 0)
            coverage = pixel_count / (mask.shape[0] * mask.shape[1])
            if coverage > 0.3:  # 30% 이상 커버리지
                detected_colors.append(color_name)
        
        return detected_colors

    def extract_page_image(self, pdf_path, page_num):
        """PDF에서 특정 페이지 이미지 추출"""
        doc = fitz.open(pdf_path)
        page = doc[page_num - 1]
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        return np.array(img)

    def validate_with_llama(self, table_data):
        """LLAMA를 사용하여 결과 검증"""
        prompt = f"""
        다음은 보험 약관의 표에서 하이라이트된 셀들의 정보입니다:
        {table_data}

        이 결과를 분석하고 다음을 확인해주세요:
        1. 하이라이트된 셀들이 논리적으로 연결되어 있나요?
        2. 표의 구조상 위치가 적절한가요?
        3. 감지된 색상이 문맥에 맞나요?

        분석 결과를 JSON 형식으로 제공해주세요.
        """
        
        response = self.llm.predict(prompt)
        logger.info(f"LLAMA 검증 결과: {response}")
        return response

    def process_page(self, pdf_path, page_num=59):
        """특정 페이지 처리"""
        try:
            # 페이지 이미지 추출
            logger.info(f"페이지 {page_num} 처리 시작")
            image = self.extract_page_image(pdf_path, page_num)
            image = self.preprocess_image(image)
            
            # PaddleOCR로 표 분석
            result = self.table_engine(image)
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
                    validation = self.validate_with_llama(table_data)
                    table_data['validation'] = validation
                    tables_data.append(table_data)
            
            return tables_data
            
        except Exception as e:
            logger.error(f"페이지 처리 중 오류 발생: {str(e)}")
            raise

def main():
    # 설정
    pdf_path = "/workspaces/automation/uploads/KB 9회주는 암보험Plus(무배당)(24.05)_요약서_10.1판매_v1.0_앞단.pdf"
    llama_model_path = "models/llama-2-7b-chat.gguf"  # LLAMA 모델 경로 설정
    
    try:
        # 분석기 초기화
        analyzer = SinglePageAnalyzer(model_path=llama_model_path)
        
        # 59페이지 분석
        logger.info("59페이지 분석 시작")
        results = analyzer.process_page(pdf_path, page_num=59)
        
        # 결과 저장
        output_file = "page_59_analysis.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            for table_data in results:
                df = pd.DataFrame(table_data['highlighted_cells'])
                df.to_excel(writer, 
                          sheet_name=f"Table_{table_data['table_index']}",
                          index=False)
        
        logger.info(f"분석 완료. 결과가 {output_file}에 저장되었습니다.")
        
    except Exception as e:
        logger.error(f"실행 중 오류 발생: {str(e)}")
        raise

if __name__ == "__main__":
    main()
```