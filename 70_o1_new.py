import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# 기본 색상 정의 (RGB 값은 0~1 범위)
DEFAULT_FONT_COLOR = (0, 0, 0)  # 검정색
DEFAULT_BG_COLOR = (1, 1, 1)    # 흰색

# 색상 비교를 위한 허용 오차
COLOR_TOLERANCE = 0.01

# 색상이 기본 색상인지 판단하는 함수
def is_default_color(color, default_color, tolerance=COLOR_TOLERANCE):
    return all(abs(color[i] - default_color[i]) <= tolerance for i in range(3))

# 페이지에서 텍스트와 스타일 정보를 추출하는 함수
def extract_revisions_from_page(page, page_number):
    revisions = []
    blocks = page.get_text("dict")["blocks"]
    for block in blocks:
        if block['type'] == 0:  # 텍스트 블록인 경우
            for line in block['lines']:
                for span in line['spans']:
                    text = span['text'].strip()
                    if not text:
                        continue

                    font_color = span.get('color', DEFAULT_FONT_COLOR)
                    bg_color = span.get('bgcolor', DEFAULT_BG_COLOR)

                    # 글꼴 색상이 기본 색상이 아닌 경우
                    is_font_color_changed = not is_default_color(font_color, DEFAULT_FONT_COLOR)
                    # 배경색이 기본 색상이 아닌 경우
                    is_bg_color_changed = not is_default_color(bg_color, DEFAULT_BG_COLOR)

                    if is_font_color_changed or is_bg_color_changed:
                        status = "추가"
                    else:
                        status = "유지"

                    revisions.append({
                        "페이지": page_number + 1,
                        "텍스트": text,
                        "글꼴 색상": font_color,
                        "배경색": bg_color,
                        "변경사항": status
                    })
    return revisions

# 전체 PDF에서 개정된 부분을 추출하는 함수
def extract_revisions_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_revisions = []
    for page_number in range(len(doc)):
        page = doc[page_number]
        revisions = extract_revisions_from_page(page, page_number)
        all_revisions.extend(revisions)
    return all_revisions

# 엑셀 파일로 저장하는 함수
def save_revisions_to_excel(revisions, output_excel_path):
    df = pd.DataFrame(revisions)

    # 엑셀 파일 생성
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "개정된 부분"

    # 노란색 강조 표시
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 데이터프레임을 엑셀로 작성
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        sheet.append(row)
        if r_idx == 1:
            continue  # 헤더는 제외
        # 변경사항이 "추가"인 경우 강조 표시
        if row[df.columns.get_loc("변경사항")] == "추가":
            for cell in sheet[r_idx]:
                cell.fill = yellow_fill
                cell.alignment = Alignment(wrap_text=True)

    workbook.save(output_excel_path)
    print(f"개정된 부분이 '{output_excel_path}'에 저장되었습니다.")

# 메인 함수
def main(pdf_path, output_excel_path):
    revisions = extract_revisions_from_pdf(pdf_path)
    save_revisions_to_excel(revisions, output_excel_path)

if __name__ == "__main__":
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_excel_path = "/workspaces/automation/output/extracted_tables.xlsx"
    main(pdf_path, output_excel_path)
