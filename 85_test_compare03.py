import fitz  # PyMuPDF
from PIL import Image
import io
import torch
from transformers import AutoProcessor, AutoModel

# CLIP model and processor loading
processor = AutoProcessor.from_pretrained("openai/clip-vit-base-patch32")
model = AutoModel.from_pretrained("openai/clip-vit-base-patch32")

def is_color_highlighted(color):
    if isinstance(color, (tuple, list)) and len(color) == 3:
        return color not in [(1, 1, 1), (0.9, 0.9, 0.9)] and any(c < 0.9 for c in color)
    elif isinstance(color, int):
        return 0 < color < 230
    else:
        return False

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

def extract_highlighted_text_with_context(pdf_path):
    print("PDF에서 음영 처리된 텍스트 추출 시작...")
    doc = fitz.open(pdf_path)
    highlighted_texts_with_context = []
    
    for page_num, page in enumerate(doc):
        # Text-based extraction
        blocks = page.get_text("dict")["blocks"]
        lines = page.get_text("text").split('\n')
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if "color" in span and is_color_highlighted(span["color"]):
                            highlighted_text = span["text"]
                            line_index = lines.index(highlighted_text) if highlighted_text in lines else -1
                            if line_index != -1:
                                context = '\n'.join(lines[max(0, line_index-5):line_index+5])
                                highlighted_texts_with_context.append((context, highlighted_text, page_num))
        
        # Image-based extraction
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        highlighted_sections = detect_highlights(img)
        
        for section in highlighted_sections:
            text = page.get_text("text", clip=section)
            if text.strip():
                context = page.get_text("text", clip=(section[0]-50, section[1]-50, section[2]+50, section[3]+50))
                highlighted_texts_with_context.append((context, text, page_num))
    
    doc.close()
    print("PDF에서 음영 처리된 텍스트 추출 완료")
    return highlighted_texts_with_context

# Example usage
if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    results = extract_highlighted_text_with_context(pdf_path)
    for context, highlighted_text, page_num in results:
        print(f"Page {page_num + 1}:")
        print(f"Highlighted text: {highlighted_text}")
        print(f"Context: {context}")
        print("---")