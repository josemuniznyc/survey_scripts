#!/usr/bin/env python
# -- coding: utf-8 --
"""
Renames a directory of evaluation surveys downloaded from surveymonkey with cour
se, instructor, and term information.
"""
import os, xlrd

DATA = "/Users/jmuniz/Dropbox/Statistics/F13 Data/"

def get_data(data):
    """Returns a list of evaluation surveys to be renamed"""
    os.chdir(data)
    print "Moved to "+data
    evals = os.listdir(data)
    print "Loaded excel files into queue."
    return evals

def get_term(semester, year):
    """Parses the semester and year and converts to a term id"""
    if semester == u'Fall':
        semester = u'F'
    elif semester == u'Spring':
        semester = u'S'
    elif semester == u'Summer':
        semester = u'M'
    else:
        pass 
    term = semester+year[-2:]
    return term

def get_instructor(last_name, first_name):
    """Parses, cleans, and joins instructor information""" 
    last_name = " ".join(last_name)
    instructor = "-"+last_name+" "+first_name+"-"
    instructor = instructor.title()
    return instructor

def rename(evaluation):
    """opens evaluation survey, generates a new filename, 
    and then renames the excel file."""
    workbook = xlrd.open_workbook(evaluation)
    worksheet = workbook.sheet_by_name(u'Questions')
    title = worksheet.row_values(0, end_colx=1)
    e = title[0].rsplit()
    prefix = " ".join((e[6], e[7],))
    instructor = get_instructor(e[9:-1], e[-1])    
    term = get_term(e[3], e[4])
    suffix = ".xls"
    file_name = prefix+instructor+term+suffix
    os.rename(evaluation, file_name)
    return file_name

# Iterate through excel files

surveys = get_data(DATA)
for survey in surveys:
    new_name = rename(survey)
    print "File "+survey+" renamed to "+new_name