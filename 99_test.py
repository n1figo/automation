import os
import torch
from transformers import AutoProcessor, AutoModel
from PIL import Image
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import PatternFill

# 지정된 PDF 파일 경로
PDF_PATH = "/workspaces/automation/uploads/1722922992_5._KB_5.10.10_24.05__0801_v1.0.pdf"

# CLIP model and processor loading
processor = AutoProcessor.from_pretrained("openai/clip-vit-base-patch32")
model = AutoModel.from_pretrained("openai/clip-vit-base-patch32")

def pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    for page in doc:
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images

def detect_highlights(image):
    width, height = image.size
    sections = []
    for i in range(0, height, 100):
        for j in range(0, width, 100):
            section = image.crop((j, i, j+100, i+100))
            sections.append((section, (j, i, j+100, i+100)))
    
    section_features = []
    for section, _ in sections:
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
                highlighted_sections.append(sections[i][1])
                break
    
    return highlighted_sections

def create_excel_with_highlights(pdf_path, excel_filename, highlighted_sections):
    wb = Workbook()
    ws = wb.active
    
    # Add highlighted sections
    ws.cell(row=1, column=1, value="하이라이트된 섹션")
    for i, section in enumerate(highlighted_sections):
        img = Image.open(pdf_path)
        highlight = img.crop(section)
        
        # Save highlight as image
        highlight_path = f"highlight_{i}.png"
        highlight.save(highlight_path)
        
        # Add highlight image to Excel
        img = XLImage(highlight_path)
        ws.add_image(img, f'A{i+2}')
        ws.row_dimensions[i+2].height = 75  # Adjust row height
        
        # Clean up temporary image file
        os.remove(highlight_path)
    
    wb.save(excel_filename)
    return excel_filename

def process_pdf(pdf_path):
    try:
        print(f"Processing PDF: {pdf_path}")
        
        # Check if file exists
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"The file {pdf_path} does not exist.")
        
        # Convert PDF to images
        pdf_images = pdf_to_images(pdf_path)
        print(f"Number of pages in PDF: {len(pdf_images)}")
        
        # Detect highlights in all pages
        all_highlighted_sections = []
        for i, img in enumerate(pdf_images):
            highlighted_sections = detect_highlights(img)
            all_highlighted_sections.extend(highlighted_sections)
            print(f"Page {i+1}: {len(highlighted_sections)} highlighted sections found")
        
        # Create Excel file with highlights
        excel_filename = "output.xlsx"
        excel_path = create_excel_with_highlights(pdf_path, excel_filename, all_highlighted_sections)
        print(f"Excel file created: {excel_path}")
        print(f"Total highlighted sections: {len(all_highlighted_sections)}")
        


        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    process_pdf(PDF_PATH)