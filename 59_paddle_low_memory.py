import fitz
import pandas as pd
import numpy as np
import cv2
import os
from PIL import Image
from paddleocr import PaddleOCR
import logging
import traceback
import gc

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

DEBUG_MODE = True
TARGET_HEADERS = ["보장명", "지급사유", "지급금액"]
IMAGE_OUTPUT_DIR = "/workspaces/automation/output/images"
os.makedirs(IMAGE_OUTPUT_DIR, exist_ok=True)

try:
    ocr = PaddleOCR(use_angle_cls=True, lang='korean', use_gpu=False, 
                    enable_mkldnn=False, use_tensorrt=False, 
                    cpu_threads=1, enable_omp=False)
    logging.info("PaddleOCR initialized successfully (CPU mode, memory optimized)")
except Exception as e:
    logging.error(f"Failed to initialize PaddleOCR: {str(e)}")
    raise

def pdf_to_image(page):
    try:
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((1000, 1000))  # 크기를 1000x1000으로 조정
        return np.array(img)
    except Exception as e:
        logging.error(f"Error in pdf_to_image: {str(e)}")
        raise
    finally:
        gc.collect()

def detect_highlights(image, page_num):
    try:
        hsv = cv2.cvtColor(image, cv2.COLOR_RGB2HSV)
        s = hsv[:,:,1]
        v = hsv[:,:,2]

        saturation_threshold = 30
        value_threshold = 200
        highlight_mask = (s > saturation_threshold) & (v > value_threshold)

        kernel = np.ones((5,5), np.uint8)
        highlight_mask = cv2.morphologyEx(highlight_mask.astype(np.uint8), cv2.MORPH_CLOSE, kernel)
        highlight_mask = cv2.morphologyEx(highlight_mask, cv2.MORPH_OPEN, kernel)

        if DEBUG_MODE:
            cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_mask.png'), highlight_mask * 255)

        contours, _ = cv2.findContours(highlight_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        if DEBUG_MODE:
            contour_image = image.copy()
            cv2.drawContours(contour_image, contours, -1, (0, 255, 0), 2)
            cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_num}_contours.png'), cv2.cvtColor(contour_image, cv2.COLOR_RGB2BGR))

        logging.info(f"Highlights detected for page {page_num}")
        return contours
    except Exception as e:
        logging.error(f"Error in detect_highlights: {str(e)}")
        raise
    finally:
        gc.collect()

def get_capture_regions(contours, image_height, image_width):
    try:
        if not contours:
            return []

        capture_height = image_height // 3
        sorted_contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])

        regions = []
        current_region = None

        for contour in sorted_contours:
            x, y, w, h = cv2.boundingRect(contour)

            if current_region is None:
                current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]
            elif y - current_region[1] < capture_height//2:
                current_region[1] = min(image_height, y + h + capture_height//2)
            else:
                regions.append(current_region)
                current_region = [max(0, y - capture_height//2), min(image_height, y + h + capture_height//2)]

        if current_region:
            regions.append(current_region)

        logging.info(f"Capture regions identified: {len(regions)}")
        return regions
    except Exception as e:
        logging.error(f"Error in get_capture_regions: {str(e)}")
        raise
    finally:
        gc.collect()

def process_ocr_result(word_info, highlight_regions):
    text, confidence = word_info[1]
    bbox = word_info[0]
    text_y = bbox[0][1]
    is_highlighted = any(region[0] <= text_y <= region[1] for region in highlight_regions)
    
    for header in TARGET_HEADERS:
        if text.startswith(header):
            return {
                "header": header,
                "value": text[len(header):].strip(),
                "is_highlighted": is_highlighted
            }
    return None

def extract_and_process_text(image, highlight_regions):
    try:
        height, width = image.shape[:2]
        sections = []
        for i in range(0, height, 300):  # 300픽셀 높이의 섹션으로 나누기
            section = image[i:i+300, :]
            sections.append(section)
        
        processed_data = []
        current_row = {}
        
        for section in sections:
            result = ocr.ocr(section, cls=True)
            for line in result:
                for word_info in line:
                    processed = process_ocr_result(word_info, highlight_regions)
                    if processed:
                        if processed["header"] in current_row and current_row[processed["header"]]:
                            processed_data.append(current_row)
                            current_row = {}
                        current_row[processed["header"]] = processed["value"]
                        if processed["is_highlighted"]:
                            current_row["변경사항"] = "추가"
            
            del result
            gc.collect()
        
        if current_row:
            processed_data.append(current_row)
        
        df = pd.DataFrame(processed_data)
        logging.info(f"OCR results processed. Rows extracted: {len(df)}")
        return df
    except Exception as e:
        logging.error(f"Error in extract_and_process_text: {str(e)}")
        raise
    finally:
        gc.collect()

def save_to_excel(df, output_path):
    try:
        df.to_excel(output_path, index=False)
        logging.info(f"File saved successfully to {output_path}")
    except Exception as e:
        logging.error(f"Error saving to Excel: {str(e)}")
        raise
    finally:
        gc.collect()

def main(pdf_path, output_excel_path):
    logging.info("Starting PDF extraction process")
    try:
        doc = fitz.open(pdf_path)
        page_number = 50  # 51페이지 (인덱스는 0부터 시작)
        
        logging.info(f"Processing page {page_number + 1}")
        page = doc[page_number]
        image = pdf_to_image(page)
        
        if DEBUG_MODE:
            cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_original.png'), cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
        
        contours = detect_highlights(image, page_number + 1)
        highlight_regions = get_capture_regions(contours, image.shape[0], image.shape[1])
        
        if DEBUG_MODE:
            highlighted_image = image.copy()
            for region in highlight_regions:
                cv2.rectangle(highlighted_image, (0, region[0]), (image.shape[1], region[1]), (0, 255, 0), 2)
            cv2.imwrite(os.path.join(IMAGE_OUTPUT_DIR, f'page_{page_number + 1}_highlighted.png'), cv2.cvtColor(highlighted_image, cv2.COLOR_RGB2BGR))
        
        processed_df = extract_and_process_text(image, highlight_regions)
        
        if DEBUG_MODE:
            logging.info(f"Page {page_number + 1} processed data:")
            logging.info(processed_df)
        
        save_to_excel(processed_df, output_excel_path)
        
        logging.info(f"51페이지의 처리된 데이터가 {output_excel_path}에 저장되었습니다.")
    except Exception as e:
        logging.error(f"An error occurred in main: {str(e)}")
        logging.error(traceback.format_exc())
    finally:
        gc.collect()

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)