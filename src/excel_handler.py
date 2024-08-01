from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

def create_excel_with_image(image_path, excel_filename):
    wb = Workbook()
    ws = wb.active
    
    img = XLImage(image_path)
    
    # 이미지 크기 조정 (필요시)
    img.width = 800
    img.height = 600
    
    # A1 셀에 이미지 삽입
    ws.add_image(img, 'A1')
    
    wb.save(f"output/excel/{excel_filename}")