import pytest
from pathlib import Path
import json
import pandas as pd
import camelot
import numpy as np
import fitz
from difflib import SequenceMatcher
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import logging
import argparse
import sys
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from datetime import datetime

@dataclass
class TableQualityMetrics:
    """표 파싱 품질 메트릭"""
    empty_cells_ratio: float
    suspicious_cells_ratio: float
    structure_score: float
    consistency_score: float
    merged_cell_detection_score: float
    text_accuracy_score: float
    overall_score: float
    issues: List[str]
    merged_cell_details: Dict[str, any]
    text_comparison_details: Dict[str, any]

class TableAnalysisEvaluator:
    def __init__(self):
        self.suspicious_patterns = [
            r'^\s*$',              # 빈 셀
            r'^[^\w\s가-힣]+$',    # 특수문자만 있는 셀
            r'^\d{1,2}$',         # 1-2자리 숫자만 있는 셀
            r'^[.]{2,}$'          # 점만 여러 개 있는 셀
        ]

    def _evaluate_text_accuracy(self, extracted_df: pd.DataFrame, pdf_path: str, page_num: int) -> Tuple[float, Dict, List[str]]:
        """텍스트 추출 정확도 평가"""
        issues = []
        text_details = {
            'total_cells': 0,
            'exact_matches': 0,
            'close_matches': 0,
            'mismatches': 0,
            'sample_errors': []
        }

        try:
            # PDF에서 원본 텍스트 추출
            doc = fitz.open(pdf_path)
            page = doc[page_num - 1]
            
            # 표 영역 찾기
            tables = page.find_tables()
            if not tables:
                issues.append("원본 PDF에서 표를 찾을 수 없음")
                return 0.0, text_details, issues

            original_table = tables[0]
            
            # 셀별 비교
            total_similarity = 0
            cell_count = 0
            
            for i, row in extracted_df.iterrows():
                for j, extracted_text in enumerate(row):
                    text_details['total_cells'] += 1
                    
                    try:
                        original_text = original_table.cells[i][j].text.strip()
                        extracted_text = str(extracted_text).strip()
                        
                        similarity = SequenceMatcher(None, original_text, extracted_text).ratio()
                        
                        if similarity == 1.0:
                            text_details['exact_matches'] += 1
                        elif similarity >= 0.8:
                            text_details['close_matches'] += 1
                        else:
                            text_details['mismatches'] += 1
                            
                            if len(text_details['sample_errors']) < 5:
                                text_details['sample_errors'].append({
                                    'original': original_text,
                                    'extracted': extracted_text,
                                    'similarity': similarity,
                                    'position': f'({i+1}, {j+1})'
                                })
                        
                        total_similarity += similarity
                        cell_count += 1
                        
                    except IndexError:
                        issues.append(f"셀 위치 불일치: ({i+1}, {j+1})")
                        continue

            doc.close()
            accuracy_score = total_similarity / cell_count if cell_count > 0 else 0.0
            
            if text_details['mismatches'] / text_details['total_cells'] > 0.3:
                issues.append("30% 이상의 셀에서 텍스트 불일치 발견")
            
            return accuracy_score, text_details, issues
            
        except Exception as e:
            issues.append(f"텍스트 정확도 평가 중 오류: {str(e)}")
            return 0.0, text_details, issues

    def _evaluate_merged_cells(self, df: pd.DataFrame, merged_cells: List) -> Tuple[float, Dict, List[str]]:
        """병합된 셀 검출 품질 평가"""
        issues = []
        merged_cell_details = {
            'total_merged_areas': len(merged_cells),
            'problematic_merges': 0,
            'irregular_patterns': 0
        }

        if not merged_cells:
            return 1.0, merged_cell_details, []

        # 병합된 영역 검증
        for mc in merged_cells:
            if mc.end_row - mc.start_row < 0 or mc.end_col - mc.start_col < 0:
                merged_cell_details['problematic_merges'] += 1
                issues.append(f"잘못된 병합 범위: ({mc.start_row}, {mc.start_col}) -> ({mc.end_row}, {mc.end_col})")

            try:
                merged_area = df.iloc[mc.start_row-1:mc.end_row, mc.start_col-1:mc.end_col]
                non_empty_cells = merged_area.notna().sum().sum()
                if non_empty_cells > 1:
                    merged_cell_details['irregular_patterns'] += 1
                    issues.append(f"병합 영역 내 다중 값 감지: ({mc.start_row}, {mc.start_col})")
            except IndexError:
                merged_cell_details['problematic_merges'] += 1
                issues.append("병합 영역이 표 범위를 벗어남")

        pattern_score = self._analyze_merge_patterns(merged_cells)
        
        problematic_ratio = merged_cell_details['problematic_merges'] / len(merged_cells)
        irregular_ratio = merged_cell_details['irregular_patterns'] / len(merged_cells)
        
        score = (
            (1 - problematic_ratio) * 0.4 +
            (1 - irregular_ratio) * 0.3 +
            pattern_score * 0.3
        )

        return score, merged_cell_details, issues

    def _analyze_merge_patterns(self, merged_cells: List) -> float:
        """병합 패턴의 규칙성 분석"""
        if not merged_cells:
            return 1.0

        pattern_scores = []

        # 수평 병합 패턴 분석
        horizontal_merges = [mc for mc in merged_cells if mc.end_col - mc.start_col > 0]
        if horizontal_merges:
            h_widths = [mc.end_col - mc.start_col + 1 for mc in horizontal_merges]
            h_score = 1 - (np.std(h_widths) / max(np.mean(h_widths), 1))
            pattern_scores.append(h_score)

        # 수직 병합 패턴 분석
        vertical_merges = [mc for mc in merged_cells if mc.end_row - mc.start_row > 0]
        if vertical_merges:
            v_heights = [mc.end_row - mc.start_row + 1 for mc in vertical_merges]
            v_score = 1 - (np.std(v_heights) / max(np.mean(v_heights), 1))
            pattern_scores.append(v_score)

        # 병합 위치의 규칙성 분석
        start_positions = [(mc.start_row, mc.start_col) for mc in merged_cells]
        if len(start_positions) > 1:
            row_diffs = np.diff([pos[0] for pos in start_positions])
            col_diffs = np.diff([pos[1] for pos in start_positions])
            
            position_score = 1 - (
                (np.std(row_diffs) / max(np.mean(abs(row_diffs)), 1) +
                 np.std(col_diffs) / max(np.mean(abs(col_diffs)), 1)
                ) / 2
            )
            pattern_scores.append(position_score)

        return np.mean(pattern_scores) if pattern_scores else 0.0

    def evaluate_table(self, df: pd.DataFrame, pdf_path: str, page_num: int, 
                      merged_cells: List = None) -> TableQualityMetrics:
        """표 품질 평가"""
        empty_ratio = df.isna().sum().sum() / df.size
        suspicious_ratio = self._evaluate_suspicious_cells(df)
        structure_score = self._evaluate_structure(df)
        consistency_score = self._evaluate_consistency(df)
        
        merged_score, merged_details, merged_issues = self._evaluate_merged_cells(
            df, merged_cells or []
        )
        
        text_score, text_details, text_issues = self._evaluate_text_accuracy(
            df, pdf_path, page_num
        )
        
        overall_score = self._calculate_overall_score(
            empty_ratio, suspicious_ratio, structure_score,
            consistency_score, merged_score, text_score
        )
        
        issues = []
        if merged_issues:
            issues.extend(merged_issues)
        if text_issues:
            issues.extend(text_issues)
        
        return TableQualityMetrics(
            empty_cells_ratio=empty_ratio,
            suspicious_cells_ratio=suspicious_ratio,
            structure_score=structure_score,
            consistency_score=consistency_score,
            merged_cell_detection_score=merged_score,
            text_accuracy_score=text_score,
            overall_score=overall_score,
            issues=issues,
            merged_cell_details=merged_details,
            text_comparison_details=text_details
        )

    def _evaluate_suspicious_cells(self, df: pd.DataFrame) -> float:
        """의심스러운 셀 비율 계산"""
        suspicious_count = 0
        total_cells = df.size
        
        for col in df.columns:
            values = df[col].astype(str)
            for pattern in self.suspicious_patterns:
                suspicious_count += values.str.match(pattern).sum()
                
        return suspicious_count / total_cells

    def _evaluate_structure(self, df: pd.DataFrame) -> float:
        """표 구조 평가"""
        scores = []
        
        # 컬럼명 평가
        has_valid_columns = all(isinstance(col, str) and len(col.strip()) > 0 
                              for col in df.columns)
        scores.append(1.0 if has_valid_columns else 0.5)
        
        # 데이터 타입 일관성 평가
        for col in df.columns:
            unique_types = df[col].apply(type).unique()
            scores.append(1.0 if len(unique_types) == 1 else 
                        0.7 if len(unique_types) == 2 else 0.3)
        
        # 행 길이 일관성 평가
        row_lengths = df.apply(lambda x: len(x.dropna()), axis=1)
        length_consistency = 1.0 - (row_lengths.std() / row_lengths.mean() 
                                  if row_lengths.mean() > 0 else 1.0)
        scores.append(length_consistency)
        
        return np.mean(scores)

    def _evaluate_consistency(self, df: pd.DataFrame) -> float:
        """데이터 일관성 평가"""
        scores = []
        
        # 값의 길이 일관성
        for col in df.columns:
            values = df[col].astype(str)
            lengths = values.str.len()
            length_variance = lengths.std() / lengths.mean() if lengths.mean() > 0 else 1.0
            scores.append(1.0 - min(length_variance, 1.0))
        
        # 형식 일관성
        for col in df.columns:
            values = df[col].astype(str)
            has_numbers = values.str.contains(r'\d').sum()
            if has_numbers > 0:
                number_format_consistency = values.str.match(r'^\d+([,.]\d+)?$').sum() / has_numbers
                scores.append(number_format_consistency)
        
        return np.mean(scores) if scores else 0.0

    def _calculate_overall_score(self, empty_ratio: float, suspicious_ratio: float,
                               structure_score: float, consistency_score: float,
                               merged_score: float, text_score: float) -> float:
        """전체 품질 점수 계산"""
        weights = {
            'empty_ratio': 0.15,
            'suspicious_ratio': 0.15,
            'structure_score': 0.15,
            'consistency_score': 0.15,
            'merged_score': 0.2,
            'text_score': 0.2
        }
        
        score = (
            (1 - empty_ratio) * weights['empty_ratio'] +
            (1 - suspicious_ratio) * weights['suspicious_ratio'] +
            structure_score * weights['structure_score'] +
            consistency_score * weights['consistency_score'] +
            merged_score * weights['merged_score'] +
            text_score * weights['text_score']
        )
        
        return round(score * 100, 2)

    def format_evaluation_report(self, metrics: TableQualityMetrics) -> str:
        """평가 보고서 포맷팅"""
        report = [
            "=== 표 품질 평가 보고서 ===",
            f"전체 품질 점수: {metrics.overall_score:.1f}/100점",
            "",
            "세부 메트릭:",
            f"- 빈 셀 비율: {metrics.empty_cells_ratio:.1%}",
            f"- 의심 셀 비율: {metrics.suspicious_cells_ratio:.1%}",
            f"- 구조 점수: {metrics.structure_score:.2f}",
            f"- 일관성 점수: {metrics.consistency_score:.2f}",
            f"- 병합 셀 검출 점수: {metrics.merged_cell_detection_score:.2f}",
            f"- 텍스트 정확도 점수: {metrics.text_accuracy_score:.2f}",
            "",
            "병합 셀 분석:",
            f"- 총 병합 영역 수: {metrics.merged_cell_details['total_merged_areas']}",
            f"- 문제있는 병합 수: {metrics.merged_cell_details['problematic_merges']}",
            f"- 불규칙한 패턴 수: {metrics.merged_cell_details['irregular_patterns']}",
            "",
            "텍스트 정확도 분석:",
            f"- 총 셀 수: {metrics.text_comparison_details['total_cells']}",
            f"- 정확히 일치: {metrics.text_comparison_details['exact_matches']}",
            f"- 유사 일치 (80% 이상): {metrics.text_comparison_details['close_matches']}",
            f"- 불일치: {metrics.text_comparison_details['mismatches']}"
        ]

        if metrics.text_comparison_details['sample_errors']:
            report.extend([
                "",
                "텍스트 오류 예시:"
            ])
            for error in metrics.text_comparison_details['sample_errors']:
                report.extend([
                    f"- 위치 {error['position']}:",
                    f"  원본: '{error['original']}'",
                    f"  추출: '{error['extracted']}'",
                    f"  유사도: {error['similarity']:.2f}"
                ])

        if metrics.issues:
            report.extend([
                "",
                "발견된 문제:",
                *[f"- {issue}" for issue in metrics.issues]
            ])

        return "\n".join(report)

class EnhancedTableAnalyzer:
    def __init__(self):
        self._setup_logging()
        self.evaluator = TableAnalysisEvaluator()

    def _setup_logging(self):
        """로깅 설정"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler(
                    f'table_analysis_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
                )
            ]
        )
        self.logger = logging.getLogger(__name__)

    def basic_camelot_parse(self, pdf_path: str, page_num: int) -> pd.DataFrame:
        """기본 Camelot 파싱"""
        try:
            tables = camelot.read_pdf(
                pdf_path,
                pages=str(page_num),
                flavor='lattice',
                line_scale=40
            )
            if len(tables) > 0:
                self.logger.info(f"페이지 {page_num}: Camelot 파싱 성공")
                return tables[0].df
            else:
                self.logger.warning(f"페이지 {page_num}: 표를 찾을 수 없음")
                return pd.DataFrame()
        except Exception as e:
            self.logger.error(f"페이지 {page_num} Camelot 파싱 오류: {str(e)}")
            return pd.DataFrame()

    def analyze_pages(self, pdf_path: str, pages: List[int], output_dir: str) -> str:
        """여러 페이지 분석"""
        try:
            wb = Workbook()
            wb.remove(wb.active)
            
            for page_num in pages:
                self.logger.info(f"\n=== 페이지 {page_num} 분석 시작 ===")
                
                # 1. 기본 Camelot 파싱
                basic_df = self.basic_camelot_parse(pdf_path, page_num)
                basic_metrics = self.evaluator.evaluate_table(basic_df, pdf_path, page_num)
                
                self.logger.info(f"\n[기본 Camelot 파싱 평가 - 페이지 {page_num}]")
                self.logger.info(self.evaluator.format_evaluation_report(basic_metrics))
                
                ws_basic = wb.create_sheet(f"Page{page_num}_Basic")
                self._write_dataframe_to_sheet(
                    ws_basic, 
                    basic_df, 
                    f"페이지 {page_num} - 기본 Camelot 파싱"
                )
                
                # 2. TableAnalyzer 분석
                try:
                    analyzer = TableAnalyzer()
                    result = analyzer.analyze_page(pdf_path, page_num)
                    enhanced_metrics = self.evaluator.evaluate_table(
                        result['data'], 
                        pdf_path, 
                        page_num,
                        result.get('merged_cells', [])
                    )
                    
                    self.logger.info(f"\n[향상된 분석 평가 - 페이지 {page_num}]")
                    self.logger.info(self.evaluator.format_evaluation_report(enhanced_metrics))
                    
                    ws_enhanced = wb.create_sheet(f"Page{page_num}_Enhanced")
                    self._write_analysis_result_to_sheet(ws_enhanced, result)
                    
                except Exception as e:
                    self.logger.error(f"TableAnalyzer 분석 오류 (페이지 {page_num}): {str(e)}")

                self.logger.info(f"페이지 {page_num} 분석 완료")

            output_path = Path(output_dir) / f"table_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            self._apply_styles_to_workbook(wb)
            wb.save(str(output_path))
            
            self.logger.info(f"\n분석 결과 저장됨: {output_path}")
            return str(output_path)
            
        except Exception as e:
            self.logger.error(f"분석 중 오류 발생: {str(e)}")
            raise

    # ... (이전의 _write_dataframe_to_sheet, _write_analysis_result_to_sheet, _apply_styles_to_workbook 메서드들은 동일)

def main():
    parser = argparse.ArgumentParser(description='PDF 표 분석기')
    parser.add_argument('pdf_path', help='PDF 파일 경로')
    parser.add_argument('pages', help='분석할 페이지 (예: 1,3-5,7)')
    parser.add_argument('--output', '-o', help='출력 디렉토리', default='output')
    
    args = parser.parse_args()
    
    try:
        # 페이지 번호 파싱
        pages = parse_page_numbers(args.pages)
        print(f"\n분석할 페이지: {pages}")
        
        # PDF 파일 존재 확인
        if not Path(args.pdf_path).exists():
            print(f"\n오류: PDF 파일을 찾을 수 없습니다: {args.pdf_path}")
            return 1
            
        # 분석 실행
        analyzer = EnhancedTableAnalyzer()
        output_path = analyzer.analyze_pages(args.pdf_path, pages, args.output)
        print(f"\n분석이 완료되었습니다. 결과 파일: {output_path}")
        return 0
        
    except Exception as e:
        print(f"\n오류 발생: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())