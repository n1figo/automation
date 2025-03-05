### main.py
import os
import sys
import logging
# 현재 디렉토리를 Python 경로에 추가하여 모듈을 찾을 수 있게 합니다
sys.path.append('.')

# 스크립트 실행 위치의 절대 경로를 구하기
current_dir = os.path.dirname(os.path.abspath(__file__))
# 해당 경로를 sys.path에 추가
sys.path.append(current_dir)

# 그 후에 모듈을 임포트합니다
from src.gui.app import PDFAnalyzerGUI
import pandas as pd
import xml.etree.ElementTree as ET
import argparse


def setup_logging():
    # 로깅 설정
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('pdf_analyzer.log')
        ]
    )
    return logging.getLogger(__name__)

def main():
    try:
        # 로거 설정
        logger = setup_logging()
        
        # CLI 모드인지 GUI 모드인지 확인하는 코드 추가
        parser = argparse.ArgumentParser(description='Mindful Revision Helper')
        parser.add_argument('--cli', action='store_true', help='CLI 모드로 실행 (GUI 없음)')
        args = parser.parse_args()
        
        if args.cli:
            # CLI 모드 실행 (기존 코드)
            # ... 기존 코드 ...
            pass
        else:
            # GUI 모드 실행
            app = PDFAnalyzerGUI()
            app.run()  # mainloop() 대신 run() 메서드 사용
        
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {str(e)}")
        sys.exit(1)

def convert_xml_to_dataframe(self, xml_path: str) -> pd.DataFrame:
    """XML 파일을 DataFrame으로 변환"""
    try:
        # XML 파일 읽기
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # 데이터프레임 생성
        data = []
        for table in root.findall('table'):
            row_data = []
            for cell in table.findall('cell'):
                row_data.append(cell.text)
            data.append(row_data)
        
        df = pd.DataFrame(data)
        logger.info(f"XML 파일에서 {len(data)}개의 테이블을 성공적으로 변환했습니다.")
        return df
        
    except Exception as e:
        logger.error(f"XML 파일 변환 중 오류 발생: {str(e)}")
        raise

if __name__ == "__main__":
    main()