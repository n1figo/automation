import pandas as pd
from pathlib import Path

class ImprovedTableParser:
    def parse_table(self, pdf_path: str, page_number: int = 1) -> pd.DataFrame:
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
            
        # 여기서는 테스트를 위한 더미 데이터를 반환합니다
        # 실제 구현시에는 PDF 파싱 로직을 구현해야 합니다
        data = {
            '담보명': ['상해치료비', '암진단금'],
            '보험금액': ['1,000만원', '5,000만원'],
            '보험료': ['10,000원', '50,000원']
        }
        return pd.DataFrame(data)
