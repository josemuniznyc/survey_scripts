#!/usr/bin/env python
# -- coding: utf-8 --
"""
Given a dir of mail merged pdf files, rename using course/instructor/term
"""
import os
from PyPDF2 import PdfFileReader

CONFIG = [ '/Users/jmuniz/Dropbox/Documents/SPS/F13/Online/PDF/', ]

def get_instructor_name(index, pdf_text):
    """Parse out instructor name, testing for multi-word last names"""
    test_text = pdf_text[index + 1]
    if test_text[-1] == ",":
        last = test_text
        first = pdf_text[index + 2]
    else:
        last = test_text + " " + pdf_text[index + 2]
        first = pdf_text[index + 3]
    return (last, first)


class EvaluationReport:
    """Represents an individual evaluation report"""
    def __init__(self, filename):
        self.evaluation_info = {}
        self.evaluation_info['filename'] = filename
        self.report = PdfFileReader(file(filename, "r"))
    
    def extract_text(self):
        """Grab text out of PDF and parse for course/instructor/term info"""
        text = self.report.getPage(0).extractText()
        text = text.split(" ")
        self.evaluation_info['C_prefix'] = text[5]
        self.evaluation_info['C_number_section'] = text[6]
        instructor_index = text.index("Instructor:")
        instructor_name = get_instructor_name(instructor_index, text)
        last_name = instructor_name[0]
        first_name = instructor_name[1]
        term_index = text.index("Term:")
        semester_index = term_index + 1
        year_index = term_index + 2
        self.evaluation_info['I_firstname'] = first_name 
        self.evaluation_info['I_lastname'] = last_name 
        self.evaluation_info['Semester'] = text[semester_index]
        self.evaluation_info['Year'] = text[year_index]

    def create_new_filename(self):
        """Create new filename from the evaluation info parsed out of PDF"""
        file_suffix = ".pdf"
        e_id = self.evaluation_info
        string = e_id['C_prefix'] + " " + e_id['C_number_section'] + " - " +  \
             e_id['I_lastname'] + " " + e_id['I_firstname'] + " - " + \
             e_id['Semester'] + " " + e_id['Year'] + file_suffix
        self.evaluation_info['new_filename'] = string 

    def rename_pdfs(self):
        """Apply new filename to PDF"""
        old = self.evaluation_info['filename']
        new = self.evaluation_info['new_filename']
        os.rename(old, new)
        print old + " has been renamed as " + new

os.chdir(CONFIG[0])
PDFS = os.listdir(CONFIG[0])

print "Rename evaluation PDFs. \nAs you wish."

COUNT = 0

for i in PDFS:
    if i[-4:] == '.pdf':
        x = EvaluationReport(i)
        x.extract_text()
        x.create_new_filename()
        x.rename_pdfs()
        COUNT += 1
    else:
        pass

print "\n" + str(COUNT) + " PDFs have been renamed.\nDone & Done."
