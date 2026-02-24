from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="combinefilewithsamename",
    version="1.0.0",
    author="jeet071992-png",
    description="Combine multiple Excel workbooks with a popup file selector",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/jeet071992-png/combinefilewithsamename",
    packages=find_packages(),
    install_requires=[
        "openpyxl>=3.0.0",
    ],
    extras_require={
        "clipboard": ["pyperclip"],
    },
    entry_points={
        "console_scripts": [
            "combinefiles=combinefilewithsamename.core:run",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
    ],
    python_requires=">=3.7",
)
