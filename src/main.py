from capture import capture_full_page, save_image
from excel_handler import create_excel_with_image
import os

def main():
    url = "https://www.kbinsure.co.kr/CG302130001.ec"
    
    # 필요한 디렉토리 생성
    os.makedirs("/workspaces/automation/data/input", exist_ok=True)
    os.makedirs("/workspaces/automation/data/output", exist_ok=True)
    
    # 웹페이지 캡처
    full_page_image = capture_full_page(url)
    
    # 이미지 저장
    image_filename = "full_page.png"
    save_image(full_page_image, image_filename)
    
    # 엑셀 파일 생성 및 이미지 삽입
    excel_filename = "output.xlsx"
    create_excel_with_image(f"output/images/{image_filename}", excel_filename)
    
    print("Process completed successfully!")

if __name__ == "__main__":
    main()