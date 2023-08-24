from setuptools import find_packages, setup

PACKAGE = "excel_writer"

setup(
    name="excel-writer",
    version="0.0.1",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "numpy",
        "openpyxl",
        "loguru",
        "pypdf",
    ],
    extras_require={
        "dev": [
            "black",
            "isort",
            "pycln",
            "pytest",
            "pytest-cov",
            "pytest-mock",
            "radon",
            "codespell",
            "pre-commit",
            "pyright",
            "pylint",
            "pyinstaller",
            "pandas-stubs",
        ]
    },
)
