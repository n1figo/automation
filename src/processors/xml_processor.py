import xml.etree.ElementTree as ET
from xml.dom import minidom
import logging
from typing import List, Dict, Optional
import pandas as pd
from pathlib import Path
from ..utils.validators import validate_xml_structure, validate_xml_content, ValidationResult

logger = logging.getLogger(__name__)

class XMLProcessor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def df_to_xml(self, df: pd.DataFrame, page_number: int) -> ET.Element:
        """DataFrame을 XML로 변환"""
        try:
            # 테이블 요소 생성
            table = ET.Element('table')
            table.set('page', str(page_number))
            
            # 헤더 추가
            header = ET.SubElement(table, 'header')
            for col in df.columns:
                column = ET.SubElement(header, 'column')
                column.text = str(col)
            
            # 데이터 행 추가
            rows = ET.SubElement(table, 'rows')
            for _, row in df.iterrows():
                row_elem = ET.SubElement(rows, 'row')
                for col_name, value in row.items():
                    cell = ET.SubElement(row_elem, 'cell')
                    cell.set('column', col_name)
                    cell.text = str(value) if pd.notna(value) else ''
            
            return table

        except Exception as e:
            self.logger.error(f"DataFrame to XML conversion failed: {str(e)}")
            raise

    def xml_to_df(self, xml_element: ET.Element) -> pd.DataFrame:
        """XML을 DataFrame으로 변환"""
        try:
            # XML 구조 검증
            xml_str = ET.tostring(xml_element, encoding='unicode')
            validation_result = validate_xml_structure(xml_str)
            if not validation_result.is_valid:
                raise ValueError(f"Invalid XML structure: {validation_result.errors}")

            # 헤더 추출
            columns = [col.text for col in xml_element.find('header').findall('column')]
            
            # 데이터 추출
            data = []
            for row in xml_element.find('rows').findall('row'):
                row_data = []
                for cell in row.findall('cell'):
                    row_data.append(cell.text or '')
                data.append(row_data)
            
            # DataFrame 생성
            df = pd.DataFrame(data, columns=columns)
            
            # 페이지 번호 추가
            page_num = xml_element.get('page')
            if page_num:
                df['페이지'] = int(page_num)
            
            return df

        except Exception as e:
            self.logger.error(f"XML to DataFrame conversion failed: {str(e)}")
            raise

    def save_xml(self, root: ET.Element, output_path: str):
        """XML 파일로 저장"""
        try:
            # XML 문자열 생성 및 검증
            xml_str = ET.tostring(root, encoding='unicode')
            validation_result = validate_xml_structure(xml_str)
            if not validation_result.is_valid:
                raise ValueError(f"Invalid XML structure: {validation_result.errors}")
            
            # 내용 검증
            content_result = validate_xml_content(root)
            if not content_result.is_valid:
                raise ValueError(f"Invalid XML content: {content_result.errors}")
            
            # XML 포맷팅
            pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")
            
            # 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(pretty_xml)
            
            self.logger.info(f"XML file saved successfully: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to save XML file: {str(e)}")
            raise

    def load_xml(self, xml_path: str) -> ET.Element:
        """XML 파일 로드"""
        try:
            # 파일 읽기
            with open(xml_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            
            # XML 구조 검증
            validation_result = validate_xml_structure(xml_content)
            if not validation_result.is_valid:
                raise ValueError(f"Invalid XML structure: {validation_result.errors}")
            
            root = ET.fromstring(xml_content)
            
            # 내용 검증
            content_result = validate_xml_content(root)
            if not content_result.is_valid:
                raise ValueError(f"Invalid XML content: {content_result.errors}")
            
            return root

        except Exception as e:
            self.logger.error(f"Failed to load XML file: {str(e)}")
            raise

    def combine_xmls(self, xml_files: List[str]) -> ET.Element:
        """여러 XML 파일 병합"""
        try:
            # 루트 요소 생성
            root = ET.Element('document')
            
            # 각 XML 파일 처리
            for xml_file in xml_files:
                file_root = self.load_xml(xml_file)
                for table in file_root.findall('.//table'):
                    root.append(table)
            
            return root

        except Exception as e:
            self.logger.error(f"Failed to combine XML files: {str(e)}")
            raise