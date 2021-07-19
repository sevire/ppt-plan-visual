import pathlib
from setuptools import setup, find_packages

HERE = pathlib.Path(__file__).parent
VERSION = '0.0.10'

INSTALL_REQUIRES = [
      'numpy',
      'pandas',
      'python-pptx'
]

setup(
      version=VERSION,
      install_requires=INSTALL_REQUIRES,
      packages=find_packages()
      )