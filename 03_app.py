def create_excel_with_data(image_path, excel_filename, data, pdf_text, highlighted_sections):
    wb = Workbook()
    ws = wb.active
    
    img = XLImage(image_path)
    img.width = 800
    img.height = 600
    
    ws.add_image(img, 'A1')
    
    # Add scraped data
    info_data = [
        ['상품명', data.get('상품명', '')],
        ['보장내용', data.get('보장내용', '')],
        ['보험기간', data.get('보험기간', '')]
    ]
    for row, (key, value) in enumerate(info_data, start=ws.max_row + 2):
        ws.cell(row=row, column=1, value=key)
        ws.cell(row=row, column=2, value=value)
    
    # Add table data
    if data.get('테이블 데이터'):
        table_data = data['테이블 데이터']
        headers = list(table_data[0].keys())
        ws.append(headers)
        for row in table_data:
            ws.append([row.get(header, '') for header in headers])
    
    # Add PDF text and highlight information
    ws.cell(row=ws.max_row + 2, column=1, value="PDF 내용")
    ws.cell(row=ws.max_row + 1, column=1, value=pdf_text)
    
    ws.cell(row=ws.max_row + 2, column=1, value="하이라이트된 섹션")
    for i, section in enumerate(highlighted_sections):
        ws.cell(row=ws.max_row + 1, column=1, value=f"섹션 {section}")
    
    # Highlight sections
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    for section in highlighted_sections:
        cell = ws.cell(row=ws.max_row - len(highlighted_sections) + section + 1, column=1)
        cell.fill = yellow_fill
    
    output_dir = 'output/excel'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    excel_path = os.path.join(output_dir, excel_filename)
    wb.save(excel_path)
    return excel_path