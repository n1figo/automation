import cv2
import pytesseract
from PIL import Image
import numpy as np

def resize_image(image, scale_factor=2):
    """
    이미지의 해상도를 증가시키는 함수.

    Args:
        image (numpy.ndarray): 원본 이미지.
        scale_factor (float): 이미지 크기 증가 배율.

    Returns:
        numpy.ndarray: 크기가 증가된 이미지.
    """
    width = int(image.shape[1] * scale_factor)
    height = int(image.shape[0] * scale_factor)
    dim = (width, height)
    resized = cv2.resize(image, dim, interpolation=cv2.INTER_LINEAR)
    return resized

def ocr_image_tesseract_no_threshold_no_deskew(image_array, debug_mode=False):
    """
    Perform OCR on an image array using Tesseract with Korean language support.
    Applies grayscale conversion, image resizing, noise removal, contrast enhancement,
    and dilation to improve OCR accuracy. Excludes thresholding and deskewing.

    Args:
        image_array (numpy.ndarray): The image array in BGR format (as read by OpenCV).
        debug_mode (bool): If True, saves intermediate preprocessing images for debugging.

    Returns:
        str: The extracted text from the image.
    """
    try:
        # 1. 그레이스케일 변환
        gray = cv2.cvtColor(image_array, cv2.COLOR_BGR2GRAY)
        if debug_mode:
            cv2.imwrite('debug_grayscale.png', gray)
            print("그레이스케일 변환 완료: debug_grayscale.png")

        # 2. 이미지 해상도 향상 (스케일 팩터 2배)
        resized = resize_image(gray, scale_factor=2)
        if debug_mode:
            cv2.imwrite('debug_resized.png', resized)
            print("이미지 해상도 향상 완료: debug_resized.png")

        # 3. 노이즈 제거 (Median Blur)
        denoised = cv2.medianBlur(resized, 3)
        if debug_mode:
            cv2.imwrite('debug_denoised.png', denoised)
            print("노이즈 제거 완료: debug_denoised.png")

        # 4. 대비 향상 (CLAHE 적용)
        clahe = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(8,8))  # clipLimit을 낮추어 대비 향상 제한
        enhanced = clahe.apply(denoised)
        if debug_mode:
            cv2.imwrite('debug_clahe.png', enhanced)
            print("대비 향상 완료: debug_clahe.png")

        # 5. 팽창(Dilation) 적용하여 텍스트 강조
        kernel = np.ones((1,1), np.uint8)
        dilated = cv2.dilate(enhanced, kernel, iterations=1)
        if debug_mode:
            cv2.imwrite('debug_dilate.png', dilated)
            print("팽창(Dilation) 적용 완료: debug_dilate.png")

        # PIL 이미지로 변환
        pil_image = Image.fromarray(dilated)

        # Tesseract OCR 수행 (한글 언어 설정: 'kor') 및 --oem 3 설정
        custom_config = r'--oem 3 --psm 6'
        ocr_text = pytesseract.image_to_string(pil_image, lang='kor', config=custom_config)

        return ocr_text.strip()

    except Exception as e:
        print(f"OCR 처리 중 오류 발생: {e}")
        return ""

def test_ocr_no_threshold_no_deskew():
    sample_image_path = "/workspaces/automation/output/highlight_51_0.png"
    sample_image = cv2.imread(sample_image_path)
    
    if sample_image is None:
        print(f"샘플 이미지를 찾을 수 없습니다: {sample_image_path}")
        return
    
    ocr_text = ocr_image_tesseract_no_threshold_no_deskew(sample_image, debug_mode=True)
    print(f"\n샘플 OCR 결과:\n{ocr_text}")

if __name__ == "__main__":
    test_ocr_no_threshold_no_deskew()
