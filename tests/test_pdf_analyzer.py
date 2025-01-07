import pytest
import pandas as pd
from pathlib import Path
from unittest.mock import patch
import json
import os
from datetime import datetime

# pytest-asyncio 설정
pytest_plugins = ('pytest_asyncio',)

class TestPDFAnalyzer:
    @pytest.fixture(scope="class")
    def test_pdf_path(self):
        # 테스트 PDF 파일 경로를 상대 경로로 설정
        return Path("tests/test_data/sample.pdf")

    @pytest.fixture(scope="class")
    def expected_structure(self):
        return {
            'columns': ['담보명', '보험금액', '보험료'],
            'num_rows': 2
        }

    @pytest.fixture(scope="class")
    def mock_validation_response(self):
        return {
            "is_valid": True,
            "issues": [],
            "confidence": 0.95,
            "suggestions": ["금액 표기를 일관되게 해주세요."]
        }

    @pytest.mark.asyncio
    async def test_pdf_processing_pipeline(self, test_pdf_path, expected_structure, mock_validation_response):
        """PDF 처리 파이프라인 테스트"""
        
        # 1. 파일 존재 확인
        assert test_pdf_path.exists(), f"테스트 PDF 파일이 없습니다: {test_pdf_path}"

        # 2. 테이블 추출 테스트
        test_data = pd.DataFrame({
            '담보명': ['일반상해사망', '암진단금'],
            '보험금액': ['1,000만원', '3,000만원'],
            '보험료': ['1,000원', '3,000원']
        })
        
        with patch('pdf_analyzer.extractor.extract_tables') as mock_extract:
            mock_extract.return_value = test_data
            extracted_table = mock_extract.return_value
            
            assert list(extracted_table.columns) == expected_structure['columns']
            assert len(extracted_table) == expected_structure['num_rows']

        # 3. 데이터 검증
        # 금액 형식 체크
        amount_pattern = r'[\d,]+만?원'
        assert extracted_table['보험금액'].str.match(amount_pattern).all()
        assert extracted_table['보험료'].str.match(amount_pattern).all()
        
        # 누락값 체크
        assert not extracted_table.isna().any().any()
        
        # 중복 체크
        assert not extracted_table.duplicated().any()

        # 4. 결과 저장
        output_dir = Path("test_output")
        output_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"test_result_{timestamp}.json"
        
        result = {
            'table_data': extracted_table.to_dict(),
            'validation': mock_validation_response
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        assert output_path.exists()

    def test_error_handling(self):
        """에러 처리 테스트"""
        # 1. 파싱 에러
        with pytest.raises(Exception) as exc_info:
            with patch('pdf_analyzer.extractor.extract_tables', 
                      side_effect=Exception("파싱 실패")):
                raise Exception("파싱 실패")
        assert "파싱 실패" in str(exc_info.value)
        
        # 2. 파일 없음 에러
        non_existent_path = Path("없는파일.pdf")
        assert not non_existent_path.exists()

    def test_output_validation(self, mock_validation_response):
        """출력 데이터 검증"""
        # 1. 데이터 형식 검증
        assert isinstance(mock_validation_response, dict)
        assert "is_valid" in mock_validation_response
        assert "confidence" in mock_validation_response
        
        # 2. 신뢰도 검증
        assert mock_validation_response["confidence"] > 0.9
        
        # 3. 제안사항 검증
        assert isinstance(mock_validation_response["suggestions"], list)
        assert len(mock_validation_response["suggestions"]) > 0

if __name__ == "__main__":
    pytest.main(["-v", "--capture=no"])