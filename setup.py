from setuptools import find_packages, setup

setup(
    name="trello-csv",
    version="1.0.0",
    description="CLI tool to export Trello board data to a CSV file.",
    author="Mattias Holmgren",
    author_email="me@mattjh.sh",
    url="https://github.com/mattjh1/trello-csv-exporter",
    packages=find_packages(),
    package_data={
        "trello_csv": [
            "csv/*",
            "trello_template.xlsx",
        ],
    },
    install_requires=[
        "python-dotenv>=1.0.0",
        "requests>=2.20,<3.0",
        "openpyxl>=3.1,<4.0",
        "pandas>=1.0,<3.0",
        "loguru>=0.7,<1.0",
        "colorama>=0.4.0",
        "boto3",
    ],
    entry_points={
        "console_scripts": [
            "trello-csv=trello_csv.__main__:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.10",
)
