import requests
from bs4 import BeautifulSoup
import pandas as pd
import pdfplumber
import os

def extract_table_from_html(url):
    print("HTML에서 표 추출 시작...")
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # 표 찾기 (이 부분은 실제 HTML 구조에 따라 조정이 필요할 수 있습니다)
    table = soup.find('table', {'class': 'table_data'})  # 클래스명은 예시입니다
    
    if table:
        df = pd.read_html(str(table))[0]
        print("HTML에서 표 추출 완료")
        return df
    else:
        print("HTML에서 표를 찾을 수 없습니다.")
        return None

def extract_table_from_pdf(pdf_path):
    print("PDF에서 표 추출 시작...")
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    tables.append(pd.DataFrame(table[1:], columns=table[0]))
        
        if tables:
            combined_table = pd.concat(tables, ignore_index=True)
            print("PDF에서 표 추출 완료")
            return combined_table
        else:
            print("PDF에서 표를 찾을 수 없습니다.")
            return None
    except Exception as e:
        print(f"PDF 처리 중 오류 발생: {e}")
        return None

def save_dataframe_to_csv(df, output_path):
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"CSV 파일이 성공적으로 저장되었습니다: {output_path}")
    except Exception as e:
        print(f"CSV 파일 저장 중 오류 발생: {e}")

def main():
    print("프로그램 시작")
    url = "https://www.kbinsure.co.kr/CG302120001.ec"  # 실제 URL로 변경해주세요
    pdf_path = "/workspaces/automation/uploads/5. KB 5.10.10 플러스 건강보험(무배당)(24.05)_요약서_0801_v1.0.pdf"
    output_dir = "/workspaces/automation/output"
    os.makedirs(output_dir, exist_ok=True)

    # HTML에서 변경 전 표 추출
    df_html = extract_table_from_html(url)
    if df_html is not None:
        html_csv_path = os.path.join(output_dir, "변경전.csv")
        save_dataframe_to_csv(df_html, html_csv_path)
    else:
        print("HTML에서 표 추출에 실패했습니다.")

    # PDF에서 변경 후 표 추출
    df_pdf = extract_table_from_pdf(pdf_path)
    if df_pdf is not None:
        pdf_csv_path = os.path.join(output_dir, "변경후.csv")
        save_dataframe_to_csv(df_pdf, pdf_csv_path)
    else:
        print("PDF에서 표 추출에 실패했습니다.")

    print("프로그램 종료")

if __name__ == "__main__":
    main()