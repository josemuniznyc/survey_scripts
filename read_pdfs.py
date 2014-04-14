#!/usr/bin/env python
# -- coding: utf-8 --
"""
Given a PDF, return a CSV file with the appropriate data
"""
import csv
from pyPDF2 import PdfFileReader

OUTPUT = "/Users/jmuniz/Desktop/test.py"
INPUT = "/Users/jmuniz/Desktop/test.pdf"
PDF = PdfFileReader(file(INPUT,'rb'))

class Pages()
