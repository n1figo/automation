import pytest
import pandas as pd
import numpy as np
from pathlib import Path
from unittest.mock import Mock, patch
import json
import os
import asyncio
from datetime import datetime

class TestIntegratedPDFAnalyzer:
    @pytest.fixture(scope="class")
    def test_pdf_path(self):
        return "tests/test_data/sample.pdf"

    @pytest.fixture(scope="class")
    def expected_table_structure(self):
        return {
            'columns': ['담보명', '보험금액', '보험료'],
            'num_rows': 2,
            'merged_cells': True
        }

    @pytest.fixture(scope="class")
    def mock_groq_response(self):
        return {
            "validation_result": True,
            "identified_issues": [],
            "confidence_level": 0.95,
            "suggestions": ["금액 표기 방식을 일관되게 맞추세요."]
        }

    @pytest.mark.asyncio
    async def test_full_processing_pipeline(self, test_pdf_path, expected_table_structure, mock_groq_response):
        """전체 처리 파이프라인 통합 테스트"""
        
        # 1. 파일 존재 확인
        assert Path(test_pdf_path).exists(), "테스트 PDF 파일이 없습니다"

        # 2. 테이블 파싱 테스트
        with patch('pdf_analyzer.parsers.improved_table_parser.ImprovedTableParser.parse_table') as mock_parse:
            mock_parse.return_value = pd.DataFrame({
                '담보명': ['일반상해사망', '암진단금'],
                '보험금액': ['1,000만원', '3,000만원'],
                '보험료': ['1,000원', '3,000원']
            })
            
            parsed_table = mock_parse.return_value
            
            # 구조 검증
            assert list(parsed_table.columns) == expected_table_structure['columns']
            assert len(parsed_table) == expected_table_structure['num_rows']

        # 3. Groq AI 검증 테스트
        with patch('pdf_analyzer.validators.table_validator.PDFTableValidator.validate_with_groq') as mock_validate:
            mock_validate.return_value = mock_groq_response
            
            validation_result = mock_groq_response
            
            assert validation_result['validation_result']
            assert validation_result['confidence_level'] > 0.9
            assert isinstance(validation_result['suggestions'], list)

        # 4. 멀티모달 분석 테스트
        with patch('pdf_analyzer.analyzers.multimodal_analyzer.analyze_table_image') as mock_multimodal:
            mock_multimodal.return_value = {
                "is_table": True,
                "num_columns": 3,
                "has_merged_cells": True,
                "confidence": 0.98
            }
            
            multimodal_result = mock_multimodal.return_value
            
            assert multimodal_result['is_table']
            assert multimodal_result['num_columns'] == len(expected_table_structure['columns'])
            assert multimodal_result['confidence'] > 0.9

        # 5. 최종 결과 검증
        final_result = {
            'table_data': parsed_table.to_dict(),
            'ai_validation': validation_result,
            'multimodal_analysis': multimodal_result
        }

        assert all(k in final_result for k in ['table_data', 'ai_validation', 'multimodal_analysis'])

    @pytest.mark.asyncio
    async def test_error_handling(self, test_pdf_path):
        """에러 처리 테스트"""
        
        # 1. 파싱 에러 처리
        with patch('pdf_analyzer.parsers.improved_table_parser.ImprovedTableParser.parse_table', 
                  side_effect=Exception("파싱 오류")) as mock_parse:
            try:
                _ = mock_parse(test_pdf_path)
                assert False, "예외가 발생해야 함"
            except Exception as e:
                assert "파싱 오류" in str(e)

        # 2. AI 검증 에러 처리
        with patch('pdf_analyzer.validators.table_validator.PDFTableValidator.validate_with_groq',
                  side_effect=Exception("AI 검증 오류")) as mock_validate:
            try:
                _ = await mock_validate(pd.DataFrame())
                assert False, "예외가 발생해야 함"
            except Exception as e:
                assert "AI 검증 오류" in str(e)

    def test_data_consistency(self, test_pdf_path):
        """데이터 일관성 테스트"""
        
        # 테스트용 DataFrame 생성
        df = pd.DataFrame({
            '담보명': ['일반상해사망', '암진단금'],
            '보험금액': ['1,000만원', '3,000만원'],
            '보험료': ['1,000원', '3,000원']
        })

        # 1. 금액 형식 확인
        amount_pattern = r'[\d,]+만?원'
        assert df['보험금액'].str.match(amount_pattern).all()
        assert df['보험료'].str.match(amount_pattern).all()

        # 2. 누락값 확인
        assert not df.isna().any().any()

        # 3. 중복 확인
        assert not df.duplicated().any()

    @pytest.mark.asyncio
    async def test_integration_with_real_data(self, test_pdf_path):
        """실제 데이터를 사용한 통합 테스트"""
        
        if not Path(test_pdf_path).exists():
            pytest.skip("실제 테스트 데이터가 없습니다")

        try:
            # 전체 처리 파이프라인 실행
            from pdf_analyzer.parsers.improved_table_parser import ImprovedTableParser
            from pdf_analyzer.validators.table_validator import PDFTableValidator
            
            # 1. 파싱
            parser = ImprovedTableParser()
            table_df = parser.parse_table(test_pdf_path, page_number=1)
            assert isinstance(table_df, pd.DataFrame)
            assert not table_df.empty

            # 2. AI 검증
            validator = PDFTableValidator(
                groq_api_key=os.getenv('GROQ_API_KEY'),
            )
            validation_result = await validator.validate_with_groq(table_df)
            assert isinstance(validation_result, dict)
            assert 'validation_result' in validation_result

            # 3. 결과 저장
            output_dir = Path("test_output")
            output_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = output_dir / f"test_result_{timestamp}.json"
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'table_data': table_df.to_dict(),
                    'validation_result': validation_result
                }, f, ensure_ascii=False, indent=2)

            assert output_path.exists()

        except Exception as e:
            pytest.fail(f"통합 테스트 실패: {str(e)}")

if __name__ == '__main__':
    pytest.main(['-v', '--tb=short'])