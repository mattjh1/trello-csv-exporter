from setuptools import find_packages, setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="trello-csv",
    version="1.0.3",
    description="CLI tool to export Trello board data to a CSV file.",
    long_description=long_description,
    long_description_content_type="text/markdown",
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
        "setuptools",
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
    project_urls={
        "Personal Website": "https://www.mattjh.sh",
        "Bug Tracker": "https://github.com/mattjh1/trello-csv-exporter/issues",
    },
)
