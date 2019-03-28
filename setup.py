import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="uipath",
    version="0.0.2",
    author="Christian Blandford",
    author_email="christianblandford@me.com",
    description="A python API for editing UiPath .xaml files programatically.",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/christianblandford/UiPath-Python",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=[
        "beautifulsoup4 >= 4.7.1"
    ],
)
