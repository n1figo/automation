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

    @pytest.fixture(scope="class")
    def test_pages(self):
        return [67, 68]

    @pytest.mark.asyncio
    async def test_pdf_processing_pipeline(self, test_pdf_path, expected_structure, 
                                         mock_validation_response, table_parser, 
                                         table_validator, test_pages):
        """PDF 처리 파이프라인 테스트"""
        
        # 1. 파일 존재 확인
        assert test_pdf_path.exists(), f"테스트 PDF 파일이 없습니다: {test_pdf_path}"

        # output 폴더 생성
        output_dir = Path("/workspaces/automation/tests/test_data/output")
        output_dir.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        all_extracted_tables = []
        all_validated_tables = []

        # 2. 지정된 페이지에서 테이블 추출 테스트
        for page in test_pages:
            # 1차 파싱
            extracted_table = table_parser.parse_table(str(test_pdf_path), page_number=page)
            all_extracted_tables.append(extracted_table)
            
            # 2차 AI 검수 및 수정
            validation_result = table_validator.validate(extracted_table)
            validated_table = table_validator.correct_table(extracted_table)  # AI 교정 가정
            all_validated_tables.append(validated_table)
            
            assert validation_result["is_valid"]
            assert validation_result["confidence"] >= table_validator.accuracy_threshold

        # 3. 결과 저장 - Excel
        # 1차 파싱 결과
        raw_excel_path = output_dir / f"raw_parsed_tables_{timestamp}.xlsx"
        with pd.ExcelWriter(raw_excel_path) as writer:
            for i, table in enumerate(all_extracted_tables):
                table.to_excel(writer, sheet_name=f"Page_{test_pages[i]}", index=False)

        # 2차 AI 검수 결과
        validated_excel_path = output_dir / f"ai_validated_tables_{timestamp}.xlsx"
        with pd.ExcelWriter(validated_excel_path) as writer:
            for i, table in enumerate(all_validated_tables):
                table.to_excel(writer, sheet_name=f"Page_{test_pages[i]}", index=False)

        # JSON 결과도 함께 저장
        json_path = output_dir / f"validation_results_{timestamp}.json"
        
        result = {
            'raw_tables': [df.to_dict() for df in all_extracted_tables],
            'validated_tables': [df.to_dict() for df in all_validated_tables],
            'validation': validation_result
        }
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        assert raw_excel_path.exists()
        assert validated_excel_path.exists()
        assert json_path.exists()

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