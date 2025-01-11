import os
from pathlib import Path
import pandas as pd
import json
from datetime import datetime
from llama_parse import LlamaParse
import nest_asyncio
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# 환경 변수에서 API 키 가져오기
LLAMA_PARSE_API_KEY = os.getenv('LLAMA_PARSE_API_KEY')
if not LLAMA_PARSE_API_KEY:
    raise ValueError("LLAMA_PARSE_API_KEY 환경 변수가 설정되지 않았습니다.")

nest_asyncio.apply()

class PDFAnalyzer:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.parser = LlamaParse(
            api_key=LLAMA_PARSE_API_KEY,  # API 키 추가
            result_type="markdown",
            use_vendor_multimodal_model=True,
            vendor_multimodal_model_name="anthropic-sonnet-3.5",
        )
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)
        
    def parse_pdf(self):
        """PDF 파싱 및 결과 출력"""
        print(f"\n=== PDF 파싱 시작: {self.pdf_path} ===")
        
        # 1. PDF 파싱
        md_json_objs = self.parser.get_json_result(self.pdf_path)
        md_json_list = md_json_objs[0]["pages"]
        
        # 2. 파싱 결과 출력 (67, 68 페이지)
        test_pages = [67, 68]
        filtered_pages = [page for idx, page in enumerate(md_json_list, 1) if idx in test_pages]
        for idx, page in zip(test_pages, filtered_pages):
            print(f"\n=== {idx}페이지 파싱 결과 ===")
            print(page["md"])
            
            if "tables" in page:
                print(f"\n테이블 {len(page['tables'])}개 발견")
                for t_idx, table in enumerate(page["tables"], 1):
                    print(f"\n테이블 {t_idx}:")
                    print(table)

        # 3. 결과를 DataFrame으로 변환
        parsed_data = []
        for page in md_json_list:
            page_data = {
                "page_number": page.get("page", ""),
                "content": page.get("md", ""),
                "num_tables": len(page.get("tables", [])),
                "tables": page.get("tables", [])
            }
            parsed_data.append(page_data)
            
        df = pd.DataFrame(parsed_data)
        
        # 4. 결과 저장
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Excel 저장
        excel_path = self.output_dir / f"parsed_results_{timestamp}.xlsx"
        df.to_excel(excel_path, index=False)
        print(f"\nExcel 파일 저장 완료: {excel_path}")
        
        # JSON 저장
        json_path = self.output_dir / f"parsed_results_{timestamp}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump({
                "pdf_path": self.pdf_path,
                "timestamp": timestamp,
                "pages": md_json_list
            }, f, ensure_ascii=False, indent=2)
        print(f"JSON 파일 저장 완료: {json_path}")
        
        return df, md_json_list

def main():
    # PDF 경로 설정
    pdf_path = "/workspaces/automation/tests/test_data/ㅇKB+9회주는+암보험Plus(무배당)(25.01)_요약서_v1.0.hwp.pdf"
    
    analyzer = PDFAnalyzer(pdf_path)
    df, json_data = analyzer.parse_pdf()

if __name__ == "__main__":
    main()