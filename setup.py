#!/usr/bin/env python
import sys, os
#from distutils.core import setup
from setuptools import setup


setup(name="odt2rst",
    version="1.0",
    description="convert odt to rst.",
    author="Vivian De Smedt",
    author_email="vivian@vdesmedt.com",
    scripts=["odt2rst.py"],
    )
