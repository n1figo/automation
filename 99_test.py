import os
import base64
import requests
from unstructured.partition.pdf import partition_pdf
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io
import json

# 지정된 PDF 파일 경로
PDF_PATH = "/workspaces/automation/uploads/1722922992_5._KB_5.10.10_24.05__0801_v1.0.pdf"
OUTPUT_FOLDER = "highlight_images"

def get_api_key():
    api_key = os.getenv("OPENAI_API_KEY")
    if api_key:
        print("API 키가 환경 변수에서 발견되었습니다.")
        use_env = input("환경 변수의 API 키를 사용하시겠습니까? (y/n): ").lower()
        if use_env == 'y':
            return api_key

    while True:
        api_key = input("OpenAI API 키를 입력하세요: ").strip()
        if len(api_key) > 30:  # 간단한 유효성 검사
            return api_key
        else:
            print("유효하지 않은 API 키입니다. 다시 시도해주세요.")

def extract_content_from_pdf(pdf_path):
    elements = partition_pdf(pdf_path, extract_images_in_pdf=True)
    return elements

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def detect_highlights(image_path, api_key):
    base64_image = encode_image(image_path)
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    payload = {
        "model": "gpt-4-vision-preview",
        "messages": [
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
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 500
    }

    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    content = response.json()['choices'][0]['message']['content']
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        print(f"Failed to parse JSON from API response: {content}")
        return []

def capture_highlight(image_path, coordinates, output_path):
    with Image.open(image_path) as img:
        x, y, width, height = coordinates
        highlight = img.crop((x, y, x + width, y + height))
        highlight.save(output_path)

def create_excel_with_highlights(highlights, excel_filename):
    wb = Workbook()
    ws = wb.active
    
    ws.cell(row=1, column=1, value="하이라이트된 텍스트")
    ws.cell(row=1, column=2, value="이미지 파일 경로")
    for i, highlight in enumerate(highlights, start=2):
        ws.cell(row=i, column=1, value=highlight['text'])
        ws.cell(row=i, column=2, value=highlight['image_path'])
    
    wb.save(excel_filename)
    return excel_filename

def process_pdf(pdf_path, api_key):
    try:
        print(f"Processing PDF: {pdf_path}")
        
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"The file {pdf_path} does not exist.")
        
        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)
        
        elements = extract_content_from_pdf(pdf_path)
        print(f"Extracted {len(elements)} elements from PDF")
        
        highlights = []
        for i, element in enumerate(elements):
            if hasattr(element, 'type') and element.type == 'Image':
                image_path = element.image_path
                print(f"Processing image: {image_path}")
                detected_highlights = detect_highlights(image_path, api_key)
                for j, highlight in enumerate(detected_highlights):
                    output_image_path = os.path.join(OUTPUT_FOLDER, f"highlight_{i}_{j}.png")
                    capture_highlight(image_path, highlight['coordinates'], output_image_path)
                    highlights.append({
                        'text': highlight['text'],
                        'image_path': output_image_path
                    })
                    print(f"Detected highlight: {highlight['text']}")
        
        excel_filename = "output_highlights.xlsx"
        excel_path = create_excel_with_highlights(highlights, excel_filename)
        print(f"Excel file created: {excel_path}")
        print(f"Total highlighted sections: {len(highlights)}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    api_key = get_api_key()
    process_pdf(PDF_PATH, api_key)