import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="uipath_scaffold_generator",
    version="0.0.2",
    author="Christian Blandford",
    author_email="christianblandford@me.com",
    description="Generator system that provides scaffolding for your UiPath projects/sequences.",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/christianblandford/UiPath-Scaffold-Generator",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=[
        "beautifulsoup4 >= 4.7.1",
        "uipath >= 0.0.1",
        "openpyxl >= 2.6.1"
    ],
)
