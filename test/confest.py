import pytest
import pandas as pd
from pathlib import Path

@pytest.fixture(scope="session")
def test_data_dir():
    return Path(__file__).parent / "test_data"

@pytest.fixture
def sample_pdf_path(test_data_dir):
    return test_data_dir / "basic_table.pdf"