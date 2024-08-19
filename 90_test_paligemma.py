import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
from selenium import webdriver
from PIL import Image
import io
import os
import torch
from transformers import AutoProcessor, AutoModel
from selenium.webdriver.chrome.options import Options

# PaLM-E 모델 및 프로세서 로드
processor = AutoProcessor.from_pretrained("google/paligemma-3b-pt-224")
model = AutoModel.from_pretrained("google/paligemma-3b-pt-224")

def extract_content_from_url(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
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
    return text_content, images

def analyze_changes_with_palm_e(web_text, web_image, pdf_text, pdf_images):
    # 웹 콘텐츠 분석
    web_inputs = processor(text=web_text, images=web_image, return_tensors="pt")
    web_outputs = model(**web_inputs)
    web_embedding = web_outputs.image_embeds

    # PDF 콘텐츠 분석
    pdf_inputs = processor(text=pdf_text, images=pdf_images[0], return_tensors="pt")  # 첫 번째 페이지만 사용
    pdf_outputs = model(**pdf_inputs)
    pdf_embedding = pdf_outputs.image_embeds

    # 변경 사항 분석을 위한 임베딩 비교
    similarity = torch.nn.functional.cosine_similarity(web_embedding, pdf_embedding)
    
    # 분석 결과 생성
    if similarity.item() > 0.8:  # 임계값 설정
        analysis = "웹페이지와 PDF 문서 사이에 중요한 변경 사항이 감지되지 않았습니다."
    else:
        analysis = "웹페이지와 PDF 문서 사이에 중요한 변경 사항이 감지되었습니다. 자세한 검토가 필요합니다."
    
    return analysis, similarity.item()

def create_report(analysis, similarity):
    report = f"""
    변경 사항 분석 보고서
    ----------------------
    분석 결과: {analysis}
    유사도 점수: {similarity:.2f}

    추천 사항:
    1. 유사도 점수가 낮은 경우 (0.8 미만), 문서 전체를 상세히 검토하세요.
    2. 특히 이미지, 표, 그래프 등의 시각적 요소의 변경 여부를 확인하세요.
    3. 텍스트 내용의 추가, 삭제, 수정 사항을 꼼꼼히 검토하세요.
    """
    return report

def main():
    url = "https://www.kbinsure.co.kr/CG302120001.ec"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    
    print("웹페이지 콘텐츠 추출 중...")
    web_text, web_image = extract_content_from_url(url)
    
    print("PDF 문서 콘텐츠 추출 중...")
    pdf_text, pdf_images = extract_content_from_pdf(pdf_path)
    
    print("변경 사항 분석 중...")
    analysis, similarity = analyze_changes_with_palm_e(web_text, web_image, pdf_text, pdf_images)
    
    print("보고서 생성 중...")
    report = create_report(analysis, similarity)
    
    print("\n" + report)

    # 보고서를 파일로 저장
    with open("change_analysis_report.txt", "w", encoding="utf-8") as f:
        f.write(report)
    print("보고서가 'change_analysis_report.txt' 파일로 저장되었습니다.")

if __name__ == "__main__":
    main()