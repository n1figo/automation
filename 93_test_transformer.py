from sentence_transformers import SentenceTransformer, util
from PIL import Image, ImageDraw, ImageFont
import pytesseract
import fitz  # PyMuPDF
import io

def extract_text_from_image(image_path):
    print(f"이미지 파일 '{image_path}'에서 텍스트 추출 중...")
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image, lang='kor+eng')
    print("이미지에서 텍스트 추출 완료")
    return text

def extract_text_from_pdf(pdf_path):
    print(f"PDF 파일 '{pdf_path}'에서 텍스트 추출 중...")
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    print("PDF에서 텍스트 추출 완료")
    return text

def extract_sentences(text):
    return [sent.strip() for sent in text.split('.') if sent.strip()]

def compare_documents(text1, text2):
    print("문서 비교 시작...")
    model = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    sentences1 = extract_sentences(text1)
    sentences2 = extract_sentences(text2)
    
    print("문장 임베딩 생성 중...")
    embeddings1 = model.encode(sentences1, convert_to_tensor=True)
    embeddings2 = model.encode(sentences2, convert_to_tensor=True)
    
    print("문장 유사도 계산 중...")
    cosine_scores = util.pytorch_cos_sim(embeddings1, embeddings2)
    
    changes = []
    for i, sent1 in enumerate(sentences1):
        max_score = max(cosine_scores[i])
        if max_score < 0.8:  # Threshold for considering as changed
            changes.append((sent1, sentences2[cosine_scores[i].argmax()]))
    
    print(f"문서 비교 완료. {len(changes)}개의 변경 사항 발견")
    return changes

def highlight_changes_on_image(image_path, changes):
    print("변경 사항을 이미지에 하이라이트 표시 중...")
    image = Image.open(image_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", 20)  # Adjust path as needed

    y_position = 10
    for old, new in changes:
        draw.rectangle([10, y_position, image.width - 10, y_position + 80], 
                       fill=(255, 255, 0, 128))  # Semi-transparent yellow
        draw.text((15, y_position), f"변경 전: {old[:50]}...", fill=(255, 0, 0), font=font)
        draw.text((15, y_position + 40), f"변경 후: {new[:50]}...", fill=(0, 255, 0), font=font)
        y_position += 85

    print("이미지 하이라이트 표시 완료")
    return image

def main():
    print("프로그램 시작")
    original_image_path = "/workspaces/automation/uploads/변경전.jpeg"
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"

    print("원본 이미지에서 텍스트 추출 시작")
    original_text = extract_text_from_image(original_image_path)

    print("PDF에서 텍스트 추출 시작")
    pdf_text = extract_text_from_pdf(pdf_path)

    print("문서 비교 시작")
    changes = compare_documents(original_text, pdf_text)

    print("변경 사항 하이라이트 시작")
    highlighted_image = highlight_changes_on_image(original_image_path, changes)

    print("변경 사항이 하이라이트된 이미지 저장 중...")
    highlighted_image.save("highlighted_changes.png")
    print("변경 사항이 하이라이트된 이미지가 'highlighted_changes.png'로 저장되었습니다.")

    print("\n감지된 변경 사항:")
    for i, (old, new) in enumerate(changes, 1):
        print(f"변경 사항 {i}:")
        print(f"변경 전: {old}")
        print(f"변경 후: {new}")
        print("---")

    print("프로그램 종료")

if __name__ == "__main__":
    main()