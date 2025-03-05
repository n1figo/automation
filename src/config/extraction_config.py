from dataclasses import dataclass
from typing import Dict, Any

@dataclass
class TableExtractionConfig:
    """PDF 테이블 추출 설정"""
    
    # Lattice 모드 기본 설정
    LATTICE_CONFIG = {
        # 기존 라인 인식 설정
        'line_scale': 40,
        'process_background': True,
        
        # 텍스트 처리 설정
        'copy_text': ['v'],      # 세로 방향 텍스트 복사
        'split_text': False,      # 긴 텍스트 자동 분할
        'strip_text': '\n',      # 줄바꿈 제거
        
        # 라인 감지 미세 조정
        'line_tol': 5,          # 라인 허용 오차
        'joint_tol': 5      # 라인 교차점 허용 오차
        
        # 이미지 처리 설정
        # 'threshold_blocksize': 15,  # 이미지 이진화 블록 크기
        # 'threshold_constant': -2    # 이진화 임계값 조정
    }   
    
    # Stream 모드 기본 설정
    STREAM_CONFIG = {
        'row_tol': 3,
        'col_tol': 3,
        'edge_tol': 50,
        'split_text': False,
        'strip_text': '\n',     # 줄바꿈 제거
        'flag_size': True,      # 셀 크기 고려
        'edge_segments': True   # 정교한 경계 감지
    }
    
    # 테이블 검증 설정
    VALIDATION_CONFIG = {
        'min_columns': 3,
        'min_rows': 2,
        'similarity_threshold': 0.9
    }
    
    # 정제 설정
    CLEANING_CONFIG = {
        'remove_empty': True,
        'merge_similar_rows': True,
        'normalize_whitespace': True,
        'similarity_threshold': 0.9
    }
    
    @classmethod
    def get_lattice_config(cls, **overrides) -> Dict[str, Any]:
        """Lattice 모드 설정 반환 (선택적 오버라이드 가능)"""
        config = cls.LATTICE_CONFIG.copy()
        config.update(overrides)
        return config
    
    # TableExtractionConfig 클래스에 추가 (테이블 영역 보정)
    @classmethod
    def get_stream_config(cls):
        return {
            'table_areas': ['0,500,600,0'],  # 좌표 조정 (상황에 맞게 수정)
            'split_text': True,
            'suppress_warnings': True
        }
    
    @classmethod
    def get_table_validation_config(cls, **overrides) -> Dict[str, Any]:
        """테이블 검증 설정 반환"""
        config = cls.VALIDATION_CONFIG.copy()
        config.update(overrides)
        return config
    
    @classmethod
    def get_cleaning_config(cls, **overrides) -> Dict[str, Any]:
        """테이블 정제 설정 반환"""
        config = cls.CLEANING_CONFIG.copy()
        config.update(overrides)
        return config