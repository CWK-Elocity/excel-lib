from setuptools import setup, find_packages

setup(
    name="excel-lib",  
    version="0.1.0",  
    packages=find_packages(),
    install_requires=[
        "openpyxl>=3.0.0",
        "pandas>=1.0.0",
    ],
    description="Library for working with Excel files using pandas and openpyxl",
    author="Patryk Skibniewski",
    author_email="patrykski07@gmail.com",
    url="https://github.com/RyKaT07/excel-lib.git",
    python_requires=">=3.7"
)
