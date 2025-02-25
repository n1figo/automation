from playwright.sync_api import Page
import pandas as pd
import logging
from bs4 import BeautifulSoup
from typing import List, Dict
import re
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

class WebExtractor:
    def __init__(self, page: Page):
        self.page = page
        self.extracted_tables: List[pd.DataFrame] = []
        
    async def extract_example_tables(self) -> List[pd.DataFrame]:
        """가입예시 테이블 추출"""
        try:
            # 가입예시 탭 클릭
            await self._click_example_tab()
            await self.page.wait_for_load_state('networkidle')
            
            # 페이지 스크롤
            await self._scroll_page()
            
            # HTML 파싱
            html_content = await self.page.content()
            soup = BeautifulSoup(html_content, 'html.parser')
            
            tables = []
            for table in soup.find_all('table'):
                try:
                    df = pd.read_html(str(table))[0]
                    if not df.empty and self._is_valid_table(df):
                        df = self._clean_table(df)
                        tables.append(df)
                except Exception as e:
                    logger.debug(f"테이블 파싱 실패: {str(e)}")
                    continue
                    
            self.extracted_tables = tables
            return tables
            
        except Exception as e:
            logger.error(f"테이블 추출 실패: {str(e)}")
            return []

    async def _click_example_tab(self):
        """가입예시 탭 클릭"""
        selectors = [
            '#tabexmpl',
            'a:has-text("가입예시")',
            '[role="tab"]:has-text("가입예시")'
        ]
        
        for selector in selectors:
            try:
                await self.page.click(selector)
                logger.info("가입예시 탭 클릭 성공")
                return
            except:
                continue
                
        logger.warning("가입예시 탭을 찾을 수 없음")

    async def _scroll_page(self):
        """페이지 스크롤"""
        try:
            # 초기 높이
            last_height = await self.page.evaluate('document.body.scrollHeight')
            
            while True:
                # 페이지 맨 아래로 스크롤
                await self.page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
                await self.page.wait_for_timeout(2000)
                
                # 새로운 높이 확인
                new_height = await self.page.evaluate('document.body.scrollHeight')
                if new_height == last_height:
                    break
                last_height = new_height
                
        except Exception as e:
            logger.error(f"스크롤 중 오류: {str(e)}")

    def _is_valid_table(self, df: pd.DataFrame) -> bool:
        """유효한 테이블인지 검증"""
        # 최소 크기 검증
        if len(df) < 2 or len(df.columns) < 3:
            return False
            
        # 컬럼명 검증
        columns = [str(col).lower() for col in df.columns]
        keywords = ['가입', '보험', '보장', '금액', '나이']
        if not any(keyword in ' '.join(columns) for keyword in keywords):
            return False
            
        return True

    def _clean_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """테이블 데이터 정제"""
        try:
            # 컬럼명 정리
            df.columns = [str(col).strip() for col in df.columns]
            
            # 데이터 정제
            for col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            
            # 빈 행 제거
            df = df.dropna(how='all')
            
            return df
            
        except Exception as e:
            logger.error(f"테이블 정제 중 오류: {str(e)}")
            return df
