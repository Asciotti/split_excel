import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="split_excel",
    version="0.0.1",
    author="Andrew Sciotti",
    author_email="andrew.sciotti@gmail.com",
    description="A simple script to split an excel sheet by rows",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Asciotti/split_excel",
    packages=['openpyxl','argpase'],
    classifiers=[
        "Programming Language :: Python :: 3.6",
        "License :: OSI Approved :: MIT License",
        "Development Status :: 3 - Alpha"
    ],
)