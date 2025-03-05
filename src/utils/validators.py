from dataclasses import dataclass
from typing import List, Dict, Any, Optional
import pandas as pd
import logging
import xml.etree.ElementTree as ET

logger = logging.getLogger(__name__)

@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str] = None

    def __post_init__(self):
        if self.errors is None:
            self.errors = []

def validate_table_structure(df: pd.DataFrame) -> ValidationResult:
    """테이블 구조 검증"""
    errors = []
    
    try:
        # 1. DataFrame이 비어있는지 확인
        if df.empty:
            errors.append("Empty table")
            return ValidationResult(False, errors)

        # 2. 필수 컬럼 확인
        required_cols = ['담보명', '지급사유', '지급금액']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            errors.append(f"Missing required columns: {', '.join(missing_cols)}")

        # 3. 모든 행의 컬럼 수가 동일한지 확인
        if not all(len(row) == len(df.columns) for _, row in df.iterrows()):
            errors.append("Inconsistent number of columns across rows")

        return ValidationResult(len(errors) == 0, errors)

    except Exception as e:
        logger.error(f"Table validation failed: {str(e)}")
        return ValidationResult(False, [f"Validation error: {str(e)}"])

def validate_xml_structure(xml_content: str) -> ValidationResult:
    """XML 구조 검증"""
    try:
        if not xml_content:
            return ValidationResult(False, ["Empty XML content"])

        # XML 파싱
        root = ET.fromstring(xml_content)
        errors = []

        # 필수 요소 확인
        required_elements = ['table', 'header', 'rows']
        for element in required_elements:
            if root.find(f".//{element}") is None:
                errors.append(f"Missing required element: {element}")

        # 테이블 구조 검증
        tables = root.findall(".//table")
        for table in tables:
            # 헤더 검증
            header = table.find("header")
            if header is None or len(list(header)) == 0:
                errors.append("Table missing header or empty header")

            # 행 검증
            rows = table.find("rows")
            if rows is None or len(list(rows)) == 0:
                errors.append("Table missing rows or empty rows")

            # 페이지 번호 속성 검증
            if 'page' not in table.attrib:
                errors.append("Table missing page number attribute")

        return ValidationResult(len(errors) == 0, errors)

    except ET.ParseError as e:
        return ValidationResult(False, [f"XML parsing error: {str(e)}"])
    except Exception as e:
        logger.error(f"XML validation failed: {str(e)}")
        return ValidationResult(False, [f"Validation error: {str(e)}"])

def validate_table_data(df: pd.DataFrame) -> Dict[str, List[str]]:
    """테이블 데이터 검증"""
    validation_results = {
        'structure': [],
        'content': [],
        'format': []
    }

    # 구조 검증
    structure_result = validate_table_structure(df)
    if not structure_result.is_valid:
        validation_results['structure'].extend(structure_result.errors)

    # 데이터 타입 및 내용 검증
    for col in df.columns:
        # 숫자여야 하는 컬럼 검증
        if col in ['페이지']:
            numeric_check = pd.to_numeric(df[col], errors='coerce')
            if numeric_check.isna().any():
                validation_results['format'].append(f"Invalid numeric values in column {col}")

        # 빈 값 검증
        if df[col].isna().any():
            validation_results['content'].append(f"Missing values in column {col}")

    # 중복 행 검증
    if df.duplicated().any():
        validation_results['content'].append("Duplicate rows found")

    return validation_results

def validate_xml_content(root: ET.Element) -> ValidationResult:
    """XML 내용 검증"""
    try:
        errors = []

        # 테이블 내용 검증
        for table in root.findall(".//table"):
            # 헤더 컬럼 검증
            header = table.find("header")
            if header is not None:
                columns = [col.text for col in header.findall("column")]
                if len(columns) != len(set(columns)):
                    errors.append("Duplicate column names in header")
                if '' in columns or None in columns:
                    errors.append("Empty column names in header")

            # 행 데이터 검증
            rows = table.find("rows")
            if rows is not None:
                for row in rows.findall("row"):
                    cells = row.findall("cell")
                    # 모든 행의 셀 개수가 헤더 컬럼 수와 일치하는지 확인
                    if len(cells) != len(columns):
                        errors.append("Row cell count does not match header column count")
                    # 빈 셀 체크
                    for cell in cells:
                        if not cell.text or cell.text.isspace():
                            errors.append("Empty cell found")

        return ValidationResult(len(errors) == 0, errors)

    except Exception as e:
        logger.error(f"XML content validation failed: {str(e)}")
        return ValidationResult(False, [f"Validation error: {str(e)}"])