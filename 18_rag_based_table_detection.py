import PyPDF2
import camelot
import pandas as pd
import fitz
import os
import re
from sentence_transformers import SentenceTransformer
import faiss
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

def extract_text_with_positions(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text_chunks = []
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            lines = text.split('\n')
            for line_num, line in enumerate(lines):
                text_chunks.append({
                    'text': line,
                    'page': page_num + 1,
                    'line': line_num + 1
                })
    return text_chunks

def create_embeddings(text_chunks, model):
    texts = [chunk['text'] for chunk in text_chunks]
    embeddings = model.encode(texts)
    return embeddings

def detect_table_boundaries(text_chunks, embeddings, model):
    start_patterns = [r'표\s*\d+', r'선택특약\s*내용', r'상해관련\s*특약', r'질병관련\s*특약']
    end_patterns = [r'합\s*계', r'총\s*계', r'주\s*\)', r'※']

    index = faiss.IndexFlatL2(embeddings.shape[1])
    index.add(embeddings.astype('float32'))

    tables = []
    table_start = None

    for i, chunk in enumerate(text_chunks):
        text = chunk['text']

        if any(re.search(pattern, text) for pattern in start_patterns) and table_start is None:
            table_start = i
            continue

        if table_start is not None and any(re.search(pattern, text) for pattern in end_patterns):
            table_end = i
            
            context_start = max(0, table_start - 5)
            context_end = min(len(text_chunks), table_end + 5)
            context = text_chunks[context_start:context_end]
            
            tables.append({
                'start': text_chunks[table_start],
                'end': text_chunks[table_end],
                'context': context
            })
            table_start = None

    return tables

def extract_tables_with_camelot(pdf_path, page_numbers):
    all_tables = []
    for page in page_numbers:
        print(f"Extracting tables from page {page} using Camelot...")
        tables = camelot.read_pdf(pdf_path, pages=str(page), flavor='lattice')
        all_tables.extend(tables)
    print(f"Found {len(all_tables)} tables in total")
    return all_tables

def process_tables(tables, table_info):
    processed_data = []
    for i, table in enumerate(tables):
        df = table.df
        title = table_info[i]['context'][0]['text'] if i < len(table_info) else "Unknown"
        for row_index in range(len(df)):
            row_data = df.iloc[row_index].copy()
            row_data["Table_Number"] = i + 1
            row_data["Table_Title"] = title
            processed_data.append(row_data)
    return pd.DataFrame(processed_data)

def save_to_excel(df_dict, output_path, title=None):
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    for sheet_name, df in df_dict.items():
        ws = wb.create_sheet(title=sheet_name)

        if title:
            ws.cell(row=1, column=1, value=f"{title} - {sheet_name}")
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
            title_cell = ws.cell(row=1, column=1)
            title_cell.font = Font(size=20, bold=True)
            ws.row_dimensions[1].height = 30

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.alignment = Alignment(wrap_text=True, vertical='top')

        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output_path)
    print(f"Data saved to '{output_path}'")

def main():
    uploads_folder = "/workspaces/automation/uploads"
    output_folder = "/workspaces/automation/output"
    
    os.makedirs(output_folder, exist_ok=True)

    pdf_files = [f for f in os.listdir(uploads_folder) if f.endswith('.pdf')]
    if not pdf_files:
        print("No PDF files found in the uploads folder.")
        return

    pdf_file = pdf_files[0]
    pdf_path = os.path.join(uploads_folder, pdf_file)
    output_excel_path = os.path.join(output_folder, f"{os.path.splitext(pdf_file)[0]}_analysis.xlsx")

    print(f"Processing PDF file: {pdf_file}")

    model = SentenceTransformer('distiluse-base-multilingual-cased-v1')
    text_chunks = extract_text_with_positions(pdf_path)
    embeddings = create_embeddings(text_chunks, model)
    tables = detect_table_boundaries(text_chunks, embeddings, model)

    # Extract pages for each type
    types = ["[1종]", "[2종]", "[3종]", "선택특약"]
    type_pages = {t: [chunk['page'] for chunk in text_chunks if t in chunk['text']] for t in types}

    df_dict = {}
    for insurance_type in ["[1종]", "[2종]", "[3종]"]:
        type_tables = [t for t in tables if t['start']['page'] in type_pages[insurance_type]]
        type_table_pages = list(set([t['start']['page'] for t in type_tables]))
        camelot_tables = extract_tables_with_camelot(pdf_path, type_table_pages)
        df = process_tables(camelot_tables, type_tables)
        df_dict[insurance_type.strip('[]')] = df

    # Extract title from the first page
    doc = fitz.open(pdf_path)
    first_page = doc[0]
    page_text = first_page.get_text("text")
    title = page_text.strip().split('\n')[0]
    print(f"Extracted title: {title}")

    # Save results to Excel
    save_to_excel(df_dict, output_excel_path, title=title)

    print(f"All processed data has been saved to {output_excel_path}")

if __name__ == "__main__":
    main()