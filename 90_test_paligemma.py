import os
from dotenv import load_dotenv
import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
from PIL import Image
import io
import torch
from transformers import AutoProcessor, PaliGemmaForConditionalGeneration
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from huggingface_hub import login

# .env 파일 로드
load_dotenv()

# Hugging Face API 토큰 설정
hf_token = os.getenv('HUGGINGFACE_TOKEN')
if not hf_token:
    raise ValueError("HUGGINGFACE_TOKEN이 .env 파일에 설정되지 않았습니다.")

# Hugging Face에 로그인
login(hf_token)

# PaLIGEMMA 모델 및 프로세서 로드
model_id = "google/paligemma-3b-mix-224"
try:
    model = PaliGemmaForConditionalGeneration.from_pretrained(model_id).eval()
    processor = AutoProcessor.from_pretrained(model_id)
except Exception as e:
    raise ValueError(f"모델 로딩 중 오류 발생: {str(e)}")

def extract_content_from_url(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    driver.get(url)
    
    # 스크린샷 캡처
    screenshot = driver.get_screenshot_as_png()
    screenshot_image = Image.open(io.BytesIO(screenshot))
    
    # HTML 내용 추출
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    text_content = soup.get_text()
    
    driver.quit()
    return text_content, screenshot_image

def extract_content_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text_content = ""
    images = []
    
    for page in doc:
        text_content += page.get_text()
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    
    doc.close()
    return text_content, images[0] if images else None  # 첫 페이지 이미지만 사용

def analyze_content(text, image):
    prompt = f"Analyze the following content and describe any changes or updates: {text[:500]}"  # 텍스트 길이 제한
    model_inputs = processor(text=prompt, images=image, return_tensors="pt")
    input_len = model_inputs["input_ids"].shape[-1]
    
    with torch.inference_mode():
        generation = model.generate(**model_inputs, max_new_tokens=200, do_sample=False)
        generation = generation[0][input_len:]
        analysis = processor.decode(generation, skip_special_tokens=True)
    
    return analysis

def compare_analyses(web_analysis, pdf_analysis):
    prompt = f"Compare these two analyses and highlight the main differences:\n\nWeb analysis: {web_analysis}\n\nPDF analysis: {pdf_analysis}"
    model_inputs = processor(text=prompt, return_tensors="pt")
    input_len = model_inputs["input_ids"].shape[-1]
    
    with torch.inference_mode():
        generation = model.generate(**model_inputs, max_new_tokens=300, do_sample=False)
        generation = generation[0][input_len:]
        comparison = processor.decode(generation, skip_special_tokens=True)
    
    return comparison

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    
    print("웹페이지 콘텐츠 추출 중...")
    web_text, web_image = extract_content_from_url(url)
    
    print("PDF 문서 콘텐츠 추출 중...")
    pdf_text, pdf_image = extract_content_from_pdf(pdf_path)
    
    print("웹페이지 콘텐츠 분석 중...")
    web_analysis = analyze_content(web_text, web_image)
    
    print("PDF 문서 콘텐츠 분석 중...")
    pdf_analysis = analyze_content(pdf_text, pdf_image)
    
    print("변경 사항 비교 분석 중...")
    comparison = compare_analyses(web_analysis, pdf_analysis)
    
    print("\n변경 사항 분석 결과:")
    print(comparison)

    # 결과를 파일로 저장
    with open("change_analysis_report.txt", "w", encoding="utf-8") as f:
        f.write(comparison)
    print("\n분석 보고서가 'change_analysis_report.txt' 파일로 저장되었습니다.")

if __name__ == "__main__":
    main()
