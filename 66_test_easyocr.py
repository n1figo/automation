import easyocr
import os

# 디버깅 모드 설정
DEBUG_MODE = True

def verify_easyocr():
    """
    EasyOCR가 정상적으로 작동하는지 확인하기 위한 함수.
    기존의 'test.png' 이미지를 사용하여 OCR을 수행합니다.
    """
    sample_text = "테스트"
    
    # EasyOCR 리더 초기화 (한국어 지원)
    try:
        reader = easyocr.Reader(['ko'], gpu=False)  # GPU 사용 시 gpu=True로 설정
        if DEBUG_MODE:
            print("EasyOCR 리더 초기화 완료.")
    except Exception as e:
        print(f"EasyOCR 리더 초기화 중 오류 발생: {e}")
        return False
    
    # OCR 수행할 이미지 경로 설정
    test_image_path = os.path.join(TEXT_OUTPUT_DIR, "test.png")  # 'test.png'의 실제 경로로 변경하세요
    
    if not os.path.isfile(test_image_path):
        print(f"테스트 이미지 파일을 찾을 수 없습니다: {test_image_path}")
        return False
    
    if DEBUG_MODE:
        print(f"OCR 수행할 이미지 경로: {test_image_path}")
    
    # OCR 수행
    try:
        result = reader.readtext(test_image_path, detail=0, paragraph=False)
        if DEBUG_MODE:
            print(f"EasyOCR OCR 결과: {result}")
    except Exception as e:
        print(f"OCR 수행 중 오류 발생: {e}")
        return False
    
    # 추출된 텍스트 결합
    extracted_text = ' '.join(result).strip()
    print(f"OCR Extracted Text: '{extracted_text}'")  # 추가된 로그
    
    # 검증
    if sample_text in extracted_text:
        print("EasyOCR 검증 성공: '테스트'가 인식되었습니다.")
        return True
    else:
        print("EasyOCR 검증 실패: '테스트'가 인식되지 않았습니다.")
        return False

if __name__ == "__main__":
    # 텍스트 출력 디렉토리 설정 (이미 존재한다고 가정)
    TEXT_OUTPUT_DIR = "/workspaces/automation"
    
    # EasyOCR 검증 함수 실행
    success = verify_easyocr()
    if success:
        print("EasyOCR이 정상적으로 작동하고 있습니다.")
    else:
        print("EasyOCR이 정상적으로 작동하지 않습니다.")
