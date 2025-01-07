import pandas as pd

class PDFTableValidator:
    def __init__(self, accuracy_threshold=0.8):
        self.accuracy_threshold = accuracy_threshold

    def validate(self, table: pd.DataFrame) -> dict:
        is_valid = True
        issues = []
        confidence = 0.95  # 기본 신뢰도
        suggestions = []

        # 기본 검증 로직
        if table.empty:
            is_valid = False
            issues.append("테이블이 비어있습니다.")
            confidence = 0.0

        if table.isna().any().any():
            issues.append("누락된 값이 있습니다.")
            confidence *= 0.8

        # 금액 형식 검증
        if '보험금액' in table.columns:
            if not table['보험금액'].str.match(r'[\d,]+만?원').all():
                suggestions.append("금액 표기를 일관되게 해주세요.")
                confidence *= 0.9

        return {
            "is_valid": is_valid,
            "issues": issues,
            "confidence": confidence,
            "suggestions": suggestions
        }
