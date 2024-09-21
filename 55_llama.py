import torch
from transformers import AutoTokenizer, AutoModelForCausalLM
from PIL import Image
import fitz
import camelot
import pandas as pd
import os
import cv2
import numpy as np

# 모델 및 토크나이저 로드
model_name = "Bllossom/llama-2-ko-7b"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForCausalLM.from_pretrained(model_name, torch_dtype=torch.float16)

def pdf_to_image(page):
    pix = page.get_pixmap()
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def extract_tables_with_camelot(pdf_path, page_number):
    print(f"Extracting tables from page {page_number} using Camelot...")
    tables = camelot.read_pdf(pdf_path, pages=str(page_number), flavor='lattice')
    print(f"Found {len(tables)} tables on page {page_number}")
    return tables

def process_text(text):
    inputs = tokenizer(text, return_tensors="pt", max_length=512, truncation=True)
    
    with torch.no_grad():
        outputs = model.generate(**inputs, max_new_tokens=50, do_sample=True, top_k=50, top_p=0.95)
    
    generated_text = tokenizer.decode(outputs[0], skip_special_tokens=True)
    return generated_text

def main(pdf_path, output_excel_path):
    print("Extracting tables and analyzing content from PDF...")

    page_number = 50  # 51페이지 (0-based index)

    # PyMuPDF로 PDF 열기 및 이미지 변환
    doc = fitz.open(pdf_path)
    page = doc[page_number]
    image = pdf_to_image(page)

    # Camelot을 사용하여 표 추출
    tables = extract_tables_with_camelot(pdf_path, page_number + 1)

    processed_data = []
    for i, table in enumerate(tables):
        df = table.df
        df['Table_Number'] = i + 1
        df['변경사항'] = ""

        for idx, row in df.iterrows():
            row_text = " ".join(row.astype(str))
            prompt = f"다음 표의 행에서 변경된 내용이 있는지 분석해주세요: {row_text}"
            
            analysis = process_text(prompt)
            
            if "변경" in analysis or "추가" in analysis:
                df.at[idx, '변경사항'] = "추가"
        
        processed_data.append(df)

    final_df = pd.concat(processed_data, ignore_index=True)
    final_df.to_excel(output_excel_path, index=False)
    print(f"Processed data saved to {output_excel_path}")

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables_llama.xlsx"
    main(pdf_path, output_excel_path)