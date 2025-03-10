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
        parser.add_argument('--streamlit', action='store_true', help='Streamlit 모드로 실행')
        args = parser.parse_args()
        
        if args.cli:
            # CLI 모드 실행
            logger.info("CLI 모드로 실행합니다.")
            # CLI 관련 코드...
            pass
        elif args.streamlit:
            # Streamlit 모드 실행
            logger.info("Streamlit 모드로 실행합니다.")
            import subprocess
            import sys
            
            # Streamlit 앱 실행
            streamlit_path = os.path.join(current_dir, "test_app_v7.py")
            subprocess.run([sys.executable, "-m", "streamlit", "run", streamlit_path])
        else:
            try:
                # GUI 모드 실행 시도
                logger.info("GUI 모드로 실행을 시도합니다.")
                app = PDFAnalyzerGUI()
                app.run()
            except Exception as e:
                # GUI 실행 실패 시 대안 제시
                logger.error(f"GUI 모드 실행 실패: {str(e)}")
                print("GUI를 실행할 수 없습니다. 다음 명령어로 실행해보세요:")
                print(f"  {sys.executable} {__file__} --cli  # CLI 모드")
                print(f"  {sys.executable} {__file__} --streamlit  # Streamlit 웹 인터페이스")
                sys.exit(1)
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {str(e)}")
        sys.exit(1)

def convert_xml_to_dataframe(xml_path: str) -> pd.DataFrame:
    """XML 파일을 DataFrame으로 변환"""
    logger = logging.getLogger(__name__)  # 로거 가져오기
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