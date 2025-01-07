from setuptools import setup, find_packages

setup(
    name="pdf_analyzer",
    version="0.1.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas",
        "numpy",
        "camelot-py[cv]",
        "pdfplumber",
        "llama-cpp-python",
        "groq",
    ],
    extras_require={
        "test": ["pytest", "pytest-asyncio"],
    },
)