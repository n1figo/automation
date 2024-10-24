def process_page(pdf_path, page_num):
    """페이지의 표와 제목을 처리"""
    print(f"\n=== {page_num}페이지 처리 시작 ===")
    
    # 제목 추출
    titles = get_titles_with_positions(pdf_path, page_num)
    print("\n발견된 제목들:")
    for title in titles:
        print(f"- {title['text']} (y: {title['y_top']:.1f} - {title['y_bottom']:.1f})")
    
    # 표 추출
    tables = camelot.read_pdf(
        pdf_path,
        pages=str(page_num),
        flavor='lattice'
    )
    if not tables:
        tables = camelot.read_pdf(
            pdf_path,
            pages=str(page_num),
            flavor='stream'
        )
    print(f"\n추출된 표 수: {len(tables)}")
    
    # 페이지 크기 정보 가져오기
    doc = fitz.open(pdf_path)
    page = doc[page_num - 1]
    page_height = page.rect.height
    doc.close()
    
    # 표 위치 정보 추출
    table_positions = [get_table_positions(table, page_height) for table in tables]
    
    # 제목과 표 매칭
    matches = match_titles_to_tables(titles, table_positions)
    
    # 결과 생성
    results = []
    for i, (match, table) in enumerate(zip(matches, tables)):
        title = match["title"] if match["title"] else f"표 {i+1} (제목 없음)"
        
        # 거리 출력 수정
        distance_str = f"{match['distance']:.1f}" if match['distance'] is not None else "N/A"
        print(f"\n표 {i+1}:")
        print(f"- 제목: {title}")
        print(f"- 거리: {distance_str}")
        
        results.append({
            'title': title,
            'table': table.df,
            'page': page_num,
            'title_bbox': match["title_bbox"],
            'table_bbox': match["table_bbox"],
            'distance': match["distance"]
        })
    
    return results

def save_to_excel(results, output_path):
    """추출된 표와 제목을 Excel 파일로 저장"""
    wb = Workbook()
    ws = wb.active
    current_row = 1

    for i, item in enumerate(results, 1):
        # 제목과 위치 정보
        title_cell = ws.cell(row=current_row, column=1, 
                           value=f"{item['title']} (Page: {item['page']})")
        title_cell.font = Font(bold=True, size=12)
        title_cell.fill = PatternFill(start_color='E6E6E6', 
                                    end_color='E6E6E6', 
                                    fill_type='solid')
        
        # 위치 정보 추가 - 수정된 부분
        if item['distance'] is not None:
            distance_str = f"{item['distance']:.1f}"
            ws.cell(row=current_row, column=2,
                   value=f"Distance: {distance_str}")
        else:
            ws.cell(row=current_row, column=2,
                   value="Distance: N/A")
        
        current_row += 2

        # 표 데이터
        df = item['table']
        for r_idx, row in enumerate(df.values):
            for c_idx, value in enumerate(row):
                cell = ws.cell(row=current_row + r_idx, 
                             column=c_idx + 1, 
                             value=value)
                cell.alignment = Alignment(wrap_text=True)

        current_row += len(df) + 3

    wb.save(output_path)
    print(f"\n결과가 {output_path}에 저장되었습니다.")