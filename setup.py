# setup.py
from setuptools import setup, find_packages

setup(
    name="pdf-processor",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "PyMuPDF",
        "PyPDF2",
        "camelot-py",
        "pandas",
        "llama-cpp-python",
    ],
)