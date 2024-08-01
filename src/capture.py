from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
import time
import io

def capture_full_page(url, max_retries=3):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    service = Service(ChromeDriverManager().install())

    for attempt in range(max_retries):
        try:
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get(url)

            # 명시적 대기 추가
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # 페이지 전체 높이 구하기
            total_height = driver.execute_script("return document.body.scrollHeight")
            
            # 브라우저 창 크기 설정
            driver.set_window_size(1920, total_height)
            
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
            total_image = Image.new('RGB', (1920, total_height))
            offset = 0
            for image in image_parts:
                total_image.paste(image, (0, offset))
                offset += image.size[1]

            driver.quit()
            return total_image

        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {str(e)}")
            if attempt + 1 == max_retries:
                raise
            time.sleep(5)  # 재시도 전 5초 대기

def save_image(image, filename):
    image.save(f"output/images/{filename}")