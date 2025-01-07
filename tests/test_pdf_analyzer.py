import pytest
import pandas as pd
import numpy as np
from unittest.mock import Mock, patch
import json
import os
from pathlib import Path
from pdf_analyzer.parsers.improved_table_parser import ImprovedTableParser
from pdf_analyzer.validators.table_validator import PDFTableValidator
# from improved_table_parser import ImprovedTableParser
# from table_validator import PDFTableValidator

class TestPDFAnalyzer:
    @pytest.fixture
    def parser(self):
        return ImprovedTableParser()

    @pytest.fixture
    def validator(self):
        return PDFTableValidator(
            llama_model_path="test_model.gguf",
            groq_api_key="test-key",
            accuracy_threshold=0.8
        )

    @pytest.fixture
    def sample_table(self):
        return pd.DataFrame({
            '담보명': ['일반상해사망', '암진단금'],
            '보험금액': ['1,000만원', '3,000만원'],
            '보험료': ['1,000원', '3,000원']
        })

    @pytest.fixture
    def complex_table(self):
        return pd.DataFrame({
            '구분': ['기본계약', '기본계약', '선택특약'],
            '담보명': ['일반상해사망', '암진단금', '특정질병수술비'],
            '보장내용': ['일반상해로 사망시 지급', '암진단시 지급', '수술시 지급'],
            '보험금액': ['1,000만원', '3,000만원', '500만원']
        })

    def test_clean_cell_content(self, parser):
        assert parser._clean_cell_content("  test  ") == "test"
        assert parser._clean_cell_content("test\ntest") == "test test"
        assert parser._clean_cell_content("test    test") == "test test"
        assert parser._clean_cell_content("") == ""
        assert parser._clean_cell_content(np.nan) == ""

    def test_clean_table(self, parser, sample_table):
        dirty_df = sample_table.copy()
        dirty_df.iloc[0, 0] = "  일반상해사망\n"
        dirty_df.iloc[1, 1] = "3,000만원  "
        
        clean_df = parser._clean_table(dirty_df)
        
        assert clean_df.iloc[0, 0] == "일반상해사망"
        assert clean_df.iloc[1, 1] == "3,000만원"
        assert clean_df.shape == sample_table.shape

    def test_extract_tables_with_camelot(self, validator):
        with patch('camelot.read_pdf') as mock_read_pdf:
            mock_table = Mock()
            mock_table.df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
            mock_table.parsing_report = {'accuracy': 0.9}
            mock_read_pdf.return_value = [mock_table]
            
            result = validator.extract_tables('test.pdf')
            
            assert len(result) == 1
            assert isinstance(result[0], pd.DataFrame)
            assert mock_read_pdf.called

    @pytest.mark.asyncio
    async def test_validate_with_llama(self, validator, sample_table):
        with patch.object(validator.llm, 'create_completion') as mock_completion:
            mock_response = Mock()
            mock_response.choices = [Mock(text=json.dumps({
                "is_valid": True,
                "issues": [],
                "confidence_score": 0.95
            }))]
            mock_completion.return_value = mock_response
            
            result = validator.validate_with_llama(sample_table)
            
            assert result['is_valid']
            assert isinstance(result['confidence_score'], float)
            assert isinstance(result['issues'], list)

    @pytest.mark.asyncio
    async def test_validate_with_groq(self, validator, sample_table):
        with patch.object(validator.groq_client.chat.completions, 'create') as mock_create:
            mock_response = Mock()
            mock_response.choices = [Mock(
                message=Mock(content=json.dumps({
                    "validation_result": True,
                    "identified_issues": [],
                    "confidence_level": 0.9,
                    "suggestions": ["Format money values consistently"]
                }))
            )]
            mock_create.return_value = mock_response
            
            result = await validator.validate_with_groq(sample_table)
            
            assert result['validation_result']
            assert isinstance(result['confidence_level'], float)
            assert isinstance(result['suggestions'], list)

    def test_table_structure_detection(self, parser, complex_table):
        with patch.object(parser, 'extract_with_pdfplumber') as mock_pdfplumber:
            mock_pdfplumber.return_value = [complex_table]
            
            # 병합된 셀이 있는 표 테스트
            result = parser.parse_table("test.pdf", 1)
            assert result is not None
            assert result.shape == complex_table.shape
            
            # 컬럼 구조 체크
            assert all(col in result.columns for col in ['구분', '담보명', '보장내용', '보험금액'])

    def test_korean_text_handling(self, parser, sample_table):
        """한글 텍스트 처리 테스트"""
        result = parser._clean_table(sample_table)
        
        # 한글 텍스트가 깨지지 않는지 확인
        assert '일반상해사망' in result['담보명'].values
        assert '암진단금' in result['담보명'].values
        
        # 금액 형식이 유지되는지 확인
        assert '1,000만원' in result['보험금액'].values
        assert '3,000만원' in result['보험금액'].values

    @pytest.mark.asyncio
    async def test_validation_integration(self, validator, complex_table):
        """파싱과 검증 통합 테스트"""
        with patch.object(validator, 'extract_tables') as mock_extract, \
             patch.object(validator, 'validate_table') as mock_validate:
            
            mock_extract.return_value = [complex_table]
            mock_validate.return_value = {
                "is_valid": True,
                "confidence": 0.9,
                "issues": [],
                "suggestions": [],
                "llama_result": {},
                "groq_result": {}
            }
            
            results = await validator.process_pdf('/workspaces/automation/test/test_data/KB 금쪽같은 자녀보험Plus(무배당)(24.05)_11월11일판매_요약서_v1.1.pdf')
            
            assert len(results) == 1
            assert results[0]['validation_result']['is_valid']
            assert isinstance(results[0]['table_data'], dict)

    def test_error_recovery(self, parser):
        """에러 복구 기능 테스트"""
        # Camelot 실패 시 pdfplumber로 대체
        with patch.object(parser, 'extract_with_camelot', return_value=[]), \
             patch.object(parser, 'extract_with_pdfplumber') as mock_pdfplumber:
            
            mock_pdfplumber.return_value = [pd.DataFrame({'A': [1, 2]})]
            result = parser.parse_table("test.pdf", 1)
            
            assert result is not None
            assert isinstance(result, pd.DataFrame)

    @pytest.mark.asyncio
    async def test_error_handling(self, validator, sample_table):
        # LLaMA 에러 테스트
        with patch.object(validator.llm, 'create_completion', 
                         side_effect=Exception("LLaMA error")):
            result = validator.validate_with_llama(sample_table)
            assert not result['is_valid']
            assert "LLaMA error" in result['issues'][0]

        # Groq 에러 테스트
        with patch.object(validator.groq_client.chat.completions, 'create', 
                         side_effect=Exception("Groq error")):
            result = await validator.validate_with_groq(sample_table)
            assert not result['validation_result']
            assert "Groq error" in result['identified_issues'][0]

    @pytest.mark.integration
    def test_real_pdf_files(self, parser):
        """실제 PDF 파일을 사용한 통합 테스트"""
        test_files_dir = Path(__file__).parent / "test_data"
        
        # 다양한 케이스의 PDF 파일들 테스트
        test_cases = [
            "basic_table.pdf",
            "complex_table.pdf",
            "merged_cells.pdf",
            "korean_text.pdf"
        ]
        
        for test_file in test_cases:
            pdf_path = test_files_dir / test_file
            if not pdf_path.exists():
                continue
                
            result = parser.parse_table(str(pdf_path), 1)
            assert result is not None
            assert isinstance(result, pd.DataFrame)
            assert not result.empty

if __name__ == '__main__':
    pytest.main(['-v', '--tb=short'])