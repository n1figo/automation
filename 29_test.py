import cv2
import pytesseract
from PIL import Image
import numpy as np

def ocr_image_tesseract(image_array):
    """
    Perform OCR on an image array using Tesseract with Korean language support.

    Args:
        image_array (numpy.ndarray): The image array in BGR format (as read by OpenCV).

    Returns:
        str: The extracted text from the image.
    """
    try:
        # 이미지 전처리: 그레이스케일 변환 및 이진화
        gray = cv2.cvtColor(image_array, cv2.COLOR_BGR2GRAY)
        # 노이즈 제거
        gray = cv2.medianBlur(gray, 3)
        # 이진화 (Thresholding)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # PIL 이미지로 변환
        pil_image = Image.fromarray(thresh)

        # Tesseract OCR 수행 (한글 언어 설정: 'kor')
        ocr_text = pytesseract.image_to_string(pil_image, lang='kor')

        return ocr_text.strip()

    except Exception as e:
        print(f"OCR 처리 중 오류 발생: {e}")
        return ""

def test_ocr():
    sample_image_path = "/workspaces/automation/output/test.png"
    sample_image = cv2.imread(sample_image_path)
    
    if sample_image is None:
        print(f"샘플 이미지를 찾을 수 없습니다: {sample_image_path}")
        return
    
    ocr_text = ocr_image_tesseract(sample_image)
    print(f"샘플 OCR 결과: {ocr_text}")

if __name__ == "__main__":
    test_ocr()
