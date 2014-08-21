#!/usr/bin/env python
# -- coding: utf-8 --

import pandas as pd
import os

SCAN_DIRECTORY = "/Users/jmuniz/Dropbox/SPS/scans_sce/current"
ROSTER = "/Users/jmuniz/Desktop/Roster_to_m14.xlsx"

def get_term(term_semester, term_year):
    term_year = str(term_year)[-2:]
    if term_semester == "Fall":
        term_semester = "F"
    elif term_semester == "Spring":
        term_semester = "S"
    elif term_semester == "Summer":
        term_semester = "M"
    else:
        term_semester = "X" #Catch all 
    term = term_semester + term_year
    return term    

def get_new_name(file_name):
    course_id  = int(file_name[:-4])
    course = course_roster.loc[[course_id],:].squeeze()
    
    new_name = []
    
    new_name.append(course[1]) # course_string
    new_name.append(course[5]) # instructor_string
    new_name.append(str(course_id)) # id_string
    new_name.append(get_term(course[7], course[8]))# term_string
    
    return [file_name, new_name]

def clean_names(list_of_names):
    cleaned_name = list_of_names[1]
    cleaned_name = "-".join(cleaned_name)
    cleaned_name = cleaned_name.replace(",", "")
    cleaned_name = cleaned_name.replace(" ", "_")
    cleaned_name = cleaned_name + ".pdf"
    return [list_of_names[0], cleaned_name]

def rename_scans(list_of_names):
    old_file = os.path.join(SCAN_DIRECTORY, list_of_names[0])
    new_name = os.path.join(SCAN_DIRECTORY, list_of_names[1])
    os.rename(old_file, new_name)

scan_names = [x for x in os.listdir(SCAN_DIRECTORY) if x[-4:] == ".pdf"]
course_roster = pd.read_excel(ROSTER, index_col=0)
name_bridge = [get_new_name(x) for x in scan_names]
cleaned_names = [clean_names(x) for x in name_bridge]
log_events = [rename_scans(x) for x in cleaned_names]
print log_events