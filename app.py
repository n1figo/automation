from flask import Flask, render_template, request, send_file
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import os

app = Flask(__name__)

def setup_browser():
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    return playwright, browser, page

def capture_full_page(url):
    playwright, browser, page = setup_browser()
    try:
        page.goto(url)
        page.wait_for_timeout(5000)  # Wait for the page to load
        screenshot_path = "full_page.png"
        page.screenshot(path=screenshot_path, full_page=True)
        return screenshot_path
    finally:
        browser.close()
        playwright.stop()

def create_excel_with_image(image_path, excel_filename):
    wb = Workbook()
    ws = wb.active
    
    img = XLImage(image_path)
    img.width = 800
    img.height = 600
    
    ws.add_image(img, 'A1')
    
    output_dir = 'output/excel'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    excel_path = os.path.join(output_dir, excel_filename)
    wb.save(excel_path)
    return excel_path

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        url = request.form['url']
        image_path = capture_full_page(url)
        excel_filename = "output.xlsx"
        excel_path = create_excel_with_image(image_path, excel_filename)
        return send_file(excel_path, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
