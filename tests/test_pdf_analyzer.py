import sys
from pathlib import Path

# 프로젝트 루트 디렉토리를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

import pytest
import pandas as pd
from unittest.mock import patch
import json
import os
from datetime import datetime
from pdf_analyzer.validators import PDFTableValidator
from pdf_analyzer.parsers import ImprovedTableParser

# pytest-asyncio 설정
pytest_plugins = ('pytest_asyncio',)

class TestPDFAnalyzer:
    @pytest.fixture(scope="class")
    def table_parser(self):
        return ImprovedTableParser()
    
    @pytest.fixture(scope="class")
    def table_validator(self):
        return PDFTableValidator(accuracy_threshold=0.8)

    @pytest.fixture(scope="class")
    def test_pdf_path(self):
        # 테스트 PDF 파일 경로를 상대 경로로 설정
        return Path("/workspaces/automation/tests/test_data/ㅇKB+9회주는+암보험Plus(무배당)(25.01)_요약서_v1.0.hwp.pdf")

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
    async def test_pdf_processing_pipeline(self, test_pdf_path, expected_structure, 
                                         mock_validation_response, table_parser, table_validator):
        """PDF 처리 파이프라인 테스트"""
        
        # 1. 파일 존재 확인
        assert test_pdf_path.exists(), f"테스트 PDF 파일이 없습니다: {test_pdf_path}"

        # 2. 테이블 추출 테스트
        extracted_table = table_parser.parse_table(str(test_pdf_path), page_number=1)
        
        assert list(extracted_table.columns) == expected_structure['columns']
        assert len(extracted_table) == expected_structure['num_rows']

        # 3. 데이터 검증
        validation_result = table_validator.validate(extracted_table)
        
        assert validation_result["is_valid"]
        assert validation_result["confidence"] >= table_validator.accuracy_threshold
        
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
            'validation': validation_result
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        assert output_path.exists()

    def test_error_handling(self, table_parser):
        """에러 처리 테스트"""
        # 1. 파싱 에러
        with pytest.raises(Exception):
            table_parser.parse_table("없는파일.pdf", page_number=1)
        
        # 2. 파일 없음 에러
        non_existent_path = Path("없는파일.pdf")
        assert not non_existent_path.exists()

    def test_output_validation(self, table_validator, table_parser, test_pdf_path):
        """출력 데이터 검증"""
        # 실제 데이터로 검증
        extracted_table = table_parser.parse_table(str(test_pdf_path), page_number=1)
        validation_result = table_validator.validate(extracted_table)
        
        # 1. 데이터 형식 검증
        assert isinstance(validation_result, dict)
        assert "is_valid" in validation_result
        assert "confidence" in validation_result
        
        # 2. 신뢰도 검증
        assert validation_result["confidence"] >= table_validator.accuracy_threshold
        
        # 3. 제안사항 검증
        assert isinstance(validation_result["suggestions"], list)

if __name__ == "__main__":
    pytest.main(["-v", "--capture=no"])