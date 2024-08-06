import os
import subprocess
from flask import Flask, render_template, request, send_file
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import PatternFill
import pandas as pd
import io
import olefile
import zlib
import struct
import torch
from transformers import AutoProcessor, AutoModel
from PIL import Image
from werkzeug.utils import secure_filename
import time
from pdf2image import convert_from_path

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

ALLOWED_EXTENSIONS = {'hwp'}

# Create uploads directory if it doesn't exist
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# CLIP model and processor loading
processor = AutoProcessor.from_pretrained("openai/clip-vit-base-patch32")
model = AutoModel.from_pretrained("openai/clip-vit-base-patch32")

def setup_browser():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, page

def capture_full_page(url):
    playwright, browser, page = setup_browser()
    try:
        page.goto(url)
        page.wait_for_timeout(5000)  # Wait for the page to load
        screenshot_path = "full_page.png"
        page.screenshot(path=screenshot_path, full_page=True)
        return screenshot_path
    finally:
        browser.close()
        playwright.stop()

def process_hwp_file(filepath):
    extracted_text = []
    with olefile.OleFileIO(filepath) as ole:
        for stream in ole.listdir():
            if stream[-1] == 'BinData':
                continue
            
            try:
                encoded_data = ole.openstream(stream).read()
                
                # Try different decompression methods
                try:
                    decompressed = zlib.decompress(encoded_data, -15)
                except zlib.error:
                    try:
                        decompressed = zlib.decompress(encoded_data)
                    except zlib.error:
                        print(f"Failed to decompress stream {stream}")
                        continue
                
                i = 0
                size = len(decompressed)
                while i < size:
                    header = struct.unpack('<I', decompressed[i:i+4])[0]
                    rec_type = header & 0x3ff
                    rec_len = (header >> 20) & 0xfff
                    
                    if rec_type == 67:  # Text record
                        rec_data = decompressed[i+4:i+4+rec_len]
                        try:
                            text = rec_data.decode('utf-16')
                            extracted_text.append(text)
                        except UnicodeDecodeError:
                            print(f"Failed to decode text in stream {stream}")
                    
                    i += 4 + rec_len
            except Exception as e:
                print(f"Error processing stream {stream}: {str(e)}")
                continue

    return ' '.join(extracted_text)

def hwp_to_pdf(hwp_path, pdf_path):
    # 이 부분은 시스템에 따라 다를 수 있습니다.
    # 예: Windows에서 한컴오피스 설치 경로
    hwp_converter = r"C:\Program Files (x86)\Hnc\Office2020\Hwp2Pdf.exe"
    subprocess.run([hwp_converter, hwp_path, pdf_path])

def pdf_to_image(pdf_path, output_path):
    images = convert_from_path(pdf_path)
    images[0].save(output_path, 'PNG')

def hwp_to_image(hwp_path, output_path):
    pdf_path = hwp_path.replace('.hwp', '.pdf')
    hwp_to_pdf(hwp_path, pdf_path)
    pdf_to_image(pdf_path, output_path)
    os.remove(pdf_path)  # Remove the temporary PDF file

def detect_highlights(image_path):
    image = Image.open(image_path)
    
    width, height = image.size
    sections = []
    for i in range(0, height, 100):
        for j in range(0, width, 100):
            section = image.crop((j, i, j+100, i+100))
            sections.append(section)
    
    section_features = []
    for section in sections:
        inputs = processor(images=section, return_tensors="pt")
        with torch.no_grad():
            features = model.get_image_features(**inputs)
        section_features.append(features)
    
    text_queries = ["노란색 하이라이트", "강조된 텍스트"]
    text_features = []
    for query in text_queries:
        text_inputs = processor(text=query, return_tensors="pt", padding=True)
        with torch.no_grad():
            features = model.get_text_features(**text_inputs)
        text_features.append(features)
    
    highlighted_sections = []
    for i, section_feature in enumerate(section_features):
        for text_feature in text_features:
            similarity = torch.nn.functional.cosine_similarity(section_feature, text_feature)
            if similarity > 0.5:  # Threshold
                highlighted_sections.append(i)
                break
    
    return highlighted_sections

def create_excel_with_data(image_path, excel_filename, data, hwp_text, highlighted_sections):
    wb = Workbook()
    ws = wb.active
    
    img = XLImage(image_path)
    img.width = 800
    img.height = 600
    
    ws.add_image(img, 'A1')
    
    # Add scraped data
    info_df = pd.DataFrame([
        ['상품명', data.get('상품명', '')],
        ['보장내용', data.get('보장내용', '')],
        ['보험기간', data.get('보험기간', '')]
    ])
    info_df.to_excel(ws, startrow=ws.max_row + 2, index=False, header=False)
    
    # Add table data
    if data.get('테이블 데이터'):
        table_df = pd.DataFrame(data['테이블 데이터'])
        table_df.to_excel(ws, startrow=ws.max_row + 2, index=False)
    
    # Add HWP text and highlight information
    ws.cell(row=ws.max_row + 2, column=1, value="HWP 내용")
    ws.cell(row=ws.max_row + 1, column=1, value=hwp_text)
    
    ws.cell(row=ws.max_row + 2, column=1, value="하이라이트된 섹션")
    for i, section in enumerate(highlighted_sections):
        ws.cell(row=ws.max_row + 1, column=1, value=f"섹션 {section}")
    
    # Highlight sections
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    for section in highlighted_sections:
        cell = ws.cell(row=ws.max_row - len(highlighted_sections) + section + 1, column=1)
        cell.fill = yellow_fill
    
    output_dir = 'output/excel'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    excel_path = os.path.join(output_dir, excel_filename)
    wb.save(excel_path)
    return excel_path

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def secure_file_name(filename):
    # Step 1: Use secure_filename
    secure_name = secure_filename(filename)
    
    # Step 2: Check for allowed extensions
    if not allowed_file(secure_name):
        return None
    
    # Step 3: Limit filename length (e.g., to 50 characters)
    name, ext = os.path.splitext(secure_name)
    secure_name = name[:50] + ext
    
    # Step 4 & 5: Add timestamp to ensure uniqueness
    timestamp = int(time.time())
    final_name = f"{timestamp}_{secure_name}"
    
    return final_name

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        url = request.form['url']
        image_path = capture_full_page(url)
        
        # Here you should implement the web scraping function to get the data
        # For now, I'll use a dummy data structure
        data = {
            '상품명': 'Sample Product',
            '보장내용': 'Sample Coverage',
            '보험기간': 'Sample Period',
            '테이블 데이터': [{'Column1': 'Value1', 'Column2': 'Value2'}]
        }
        
        hwp_text = ""
        highlighted_sections = []
        
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename:
                filename = secure_file_name(file.filename)
                if filename:
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    
                    # Ensure the directory exists before saving the file
                    os.makedirs(os.path.dirname(filepath), exist_ok=True)
                    
                    file.save(filepath)
                    
                    hwp_text = process_hwp_file(filepath)
                    
                    # Convert HWP to image
                    hwp_image_path = filepath.replace('.hwp', '.png')
                    hwp_to_image(filepath, hwp_image_path)
                    
                    highlighted_sections = detect_highlights(hwp_image_path)
                    
                    # Clean up temporary files
                    os.remove(filepath)
                    os.remove(hwp_image_path)
                else:
                    return "Invalid file type. Please upload only HWP files.", 400
        
        excel_filename = "output.xlsx"
        excel_path = create_excel_with_data(image_path, excel_filename, data, hwp_text, highlighted_sections)
        return send_file(excel_path, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)