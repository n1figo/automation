def process_pdf_and_save_tables(pdf_path: str, output_path: str):
    """
    PDF에서 표를 추출하고 Excel 파일로 저장하는 메인 함수
    """
    try:
        # 섹션 탐지
        section_detector = SectionDetector()
        section_ranges = section_detector.find_section_ranges(pdf_path)
        
        if not section_ranges:
            logger.error("No sections found in PDF")
            return
            
        # 표 추출
        table_extractor = TableExtractor()
        
        injury_df = pd.DataFrame()
        disease_df = pd.DataFrame()
        
        # 상해관련 특약 표 추출
        injury_range = section_ranges.get('상해관련')
        if injury_range:
            logger.info(f"Extracting injury section tables from pages {injury_range[0]} to {injury_range[1]}")
            injury_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                injury_range[0], 
                injury_range[1]
            )
            
        # 질병관련 특약 표 추출
        disease_range = section_ranges.get('질병관련')
        if disease_range:
            logger.info(f"Extracting disease section tables from pages {disease_range[0]} to {disease_range[1]}")
            disease_df = table_extractor.extract_tables_from_range(
                pdf_path, 
                disease_range[0], 
                disease_range[1]
            )
        
        # Excel 파일로 저장
        if injury_df.empty and disease_df.empty:
            logger.error("Both DataFrames are empty, no data to save.")
            return

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if not injury_df.empty:
                injury_df.to_excel(
                    writer, 
                    sheet_name='특약표', 
                    index=False, 
                    startrow=0,
                    startcol=0
                )
                logger.info("Saved injury section tables")
                
            if not disease_df.empty:
                # 빈 줄 추가를 위한 시작 행 계산
                start_row = len(injury_df) + 3 if not injury_df.empty else 0
                disease_df.to_excel(
                    writer, 
                    sheet_name='특약표',
                    index=False,
                    startrow=start_row,
                    startcol=0
                )
                logger.info("Saved disease section tables")
        
        logger.info(f"Successfully saved tables to {output_path}")
        
    except Exception as e:
        logger.error(f"Error processing PDF and saving tables: {str(e)}")
