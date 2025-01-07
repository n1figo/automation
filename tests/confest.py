import pytest
import pandas as pd
from pathlib import Path

@pytest.fixture(scope="session")
def test_data_dir():
    return Path(__file__).parent / "test_data"

@pytest.fixture
def sample_pdf_path(test_data_dir):
    return test_data_dir / "KB 금쪽같은 자녀보험Plus(무배당)(24.05)_11월11일판매_요약서_v1.1.pdf"