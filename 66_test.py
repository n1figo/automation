from PIL import Image, ImageDraw, ImageFont

def verify_tesseract():
    """
    Tesseract OCR가 정상적으로 작동하는지 확인하기 위한 함수.
    간단한 텍스트가 포함된 이미지를 생성하여 OCR을 수행합니다.
    """
    sample_text = "테스트"
    
    # 이미지 생성
    img_width, img_height = 400, 200
    img = Image.new('RGB', (img_width, img_height), color=(255, 255, 255))  # 흰색 배경
    draw = ImageDraw.Draw(img)
    
    # 한글 폰트 경로 설정
    # Ubuntu 예시:
    font_path = "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
    # macOS 예시:
    # font_path = "/Library/Fonts/NotoSansCJK-Regular.ttc"
    # Windows 예시:
    font_path = r"C:\Windows\Fonts\malgun.ttf"  # 맑은 고딕 폰트 경로
    
    # 폰트 로드
    try:
        font = ImageFont.truetype(font_path, 48)
    except IOError:
        print(f"폰트를 찾을 수 없습니다: {font_path}. 폰트 경로를 확인하세요.")
        return False
    
    # 텍스트 그리기
    text_width, text_height = draw.textsize(sample_text, font=font)
    position = ((img_width - text_width) // 2, (img_height - text_height) // 2)
    draw.text(position, sample_text, font=font, fill=(0, 0, 0))  # 검은색 텍스트
    
    # 이미지 저장 (디버깅용)
    test_image_path = os.path.join(TEXT_OUTPUT_DIR, "test_ocr.png")
    img.save(test_image_path)
    if DEBUG_MODE:
        print(f"테스트 이미지가 '{test_image_path}'에 저장되었습니다.")
    
    # OCR 수행
    extracted_text = pytesseract.image_to_string(img, lang='kor')
    print(f"OCR Extracted Text: '{extracted_text.strip()}'")  # 추가된 로그
    
    if sample_text in extracted_text:
        print("Tesseract OCR 검증 성공: '테스트'가 인식되었습니다.")
        return True
    else:
        print("Tesseract OCR 검증 실패: '테스트'가 인식되지 않았습니다.")
        return False
