from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="word-local-autosave",
    version="0.1.0",
    author="Word Local Autosave Contributors",
    author_email="example@example.com",
    description="A local autosave utility for Microsoft Word documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/word-local-autosave",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Topic :: Office/Business :: Office Suites",
        "Development Status :: 4 - Beta",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pywin32>=223",
    ],
    entry_points={
        "console_scripts": [
            "word-autosave=main:main",
        ],
    },
) 