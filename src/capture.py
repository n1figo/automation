from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import time
import io

def capture_full_page(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # 브라우저를 띄우지 않고 실행
    driver = webdriver.Chrome(options=chrome_options)
    
    driver.get(url)
    time.sleep(5)  # 페이지 로딩 대기

    # 페이지 전체 높이 구하기
    total_height = driver.execute_script("return document.body.scrollHeight")
    
    # 브라우저 창 크기 설정
    driver.set_window_size(1200, total_height)
    
    # 스크롤 다운하면서 캡처
    image_parts = []
    offset = 0
    while offset < total_height:
        driver.execute_script(f"window.scrollTo(0, {offset});")
        time.sleep(0.5)  # 스크롤 후 잠시 대기
        
        # 현재 화면 캡처
        png = driver.get_screenshot_as_png()
        image_parts.append(Image.open(io.BytesIO(png)))
        
        offset += 900  # 스크롤 간격 조정 (브라우저 높이에 맞게)

    # 이미지 합치기
    total_image = Image.new('RGB', (1200, total_height))
    offset = 0
    for image in image_parts:
        total_image.paste(image, (0, offset))
        offset += image.size[1]

    driver.quit()
    return total_image

def save_image(image, filename):
    image.save(f"output/images/{filename}")