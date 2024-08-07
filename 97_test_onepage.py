import os
import base64
import fitz  # PyMuPDF
from openpyxl import Workbook
from PIL import Image
import io
import json
from dotenv import load_dotenv
import time
import traceback
from openai import OpenAI

print("Script started")

# .env 파일에서 환경 변수 로드
load_dotenv()

# 설정
PDF_PATH = "/workspaces/automation/uploads/1722922992_5._KB_5.10.10_24.05__0801_v1.0.pdf"
OUTPUT_FOLDER = "highlight_images"
START_PAGE = 50
END_PAGE = 52
WAIT_TIME = 30  # 각 API 호출 사이의 대기 시간(초)
MODEL_NAME = "gpt-4-1106-preview"  # 업데이트된 모델 이름

client = OpenAI()

def extract_images_from_pdf(pdf_path, start_page, end_page):
    print(f"Extracting images from PDF: {pdf_path}")
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(start_page - 1, min(end_page, len(doc))):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append((page_num + 1, img))
    doc.close()
    print(f"Extracted {len(images)} images")
    return images

def encode_image(image):
    print("Encoding image")
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode('utf-8')

def detect_highlights(image):
    print("Detecting highlights")
    base64_image = encode_image(image)
    
    try:
        print("Sending request to OpenAI API")
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "이 이미지에서 하이라이트된 텍스트를 찾아 정확히 추출해주세요. 또한 하이라이트된 영역의 좌표(x, y, width, height)도 함께 제공해주세요. JSON 형식으로 응답해주세요. 하이라이트된 텍스트가 없다면 빈 리스트를 반환해주세요."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=300
        )
        
        print(f"API response received")
        content = response.choices[0].message.content
        print(f"API response content: {content}")
        return json.loads(content)
    except Exception as e:
        print(f"API request failed: {str(e)}")
        traceback.print_exc()
        return []

# ... (나머지 함수들은 이전과 동일)

def process_pdf(pdf_path):
    print("Starting PDF processing")
    try:
        print(f"Processing PDF: {pdf_path}")
        
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"The file {pdf_path} does not exist.")
        
        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)
        
        images = extract_images_from_pdf(pdf_path, START_PAGE, END_PAGE)
        print(f"Extracted {len(images)} pages from PDF (pages {START_PAGE}-{END_PAGE})")
        
        highlights = []
        for page_num, image in images:
            print(f"Processing page {page_num}")
            try:
                detected_highlights = detect_highlights(image)
                for j, highlight in enumerate(detected_highlights):
                    if 'text' in highlight and 'coordinates' in highlight:
                        output_image_path = os.path.join(OUTPUT_FOLDER, f"highlight_page{page_num}_{j+1}.png")
                        capture_highlight(image, highlight['coordinates'], output_image_path)
                        highlights.append({
                            'page': page_num,
                            'text': highlight['text'],
                            'image_path': output_image_path
                        })
                        print(f"Detected highlight on page {page_num}: {highlight['text']}")
            except Exception as e:
                print(f"Error processing page {page_num}: {str(e)}")
                traceback.print_exc()
                continue  # 오류가 발생해도 다음 페이지 처리를 계속합니다.
            
            if page_num < END_PAGE:
                print(f"Waiting {WAIT_TIME} seconds before processing the next page...")
                time.sleep(WAIT_TIME)
        
        if highlights:
            excel_filename = "output_highlights.xlsx"
            excel_path = create_excel_with_highlights(highlights, excel_filename)
            print(f"Excel file created: {excel_path}")
            print(f"Total highlighted sections: {len(highlights)}")
        else:
            print("No highlights were detected in the specified pages of the PDF.")
        
    except Exception as e:
        print(f"An error occurred during PDF processing: {str(e)}")
        print(f"Error type: {type(e).__name__}")
        traceback.print_exc()

if __name__ == "__main__":
    print("Main execution started")
    process_pdf(PDF_PATH)
    print("Script completed")


    