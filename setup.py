from setuptools import setup, find_packages

setup(
    name="pdf_analyzer",
    version="0.1.0",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.8",
    install_requires=[
        "pandas",
        "numpy",
        "camelot-py[cv]",
        "pdfplumber",
        "pycryptodome",  # pycrypto 대신 pycryptodome 사용
        "llama-cpp-python",
        "groq",
    ],
    extras_require={
        "test": [
            "pytest",
            "pytest-asyncio",
        ],
    },
)