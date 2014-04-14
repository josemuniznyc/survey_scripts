#!/usr/bin/env python
# -- coding: utf-8 --
"""
Analysis of evaluation data
"""

import pandas as pd
import csv

CONFIG = ['/Users/jmuniz/Dropbox/Documents/SPS/F13-Evaluation data-1312.xlsx',
          '/Users/jmuniz/Desktop/median_numbers.csv'  
        ]

def get_median_data(reference):
    data = [
    round(reference.R_rate.mean(),3),
    round(reference.I_sat_ra.median(),1),
    round(reference.C_sat_ra.median(),1),
    round(reference.q1_ra.median(),1),
    round(reference.q2_ra.median(),1),
    round(reference.q3_ra.median(),1),
    round(reference.q4_ra.median(),1),
    round(reference.q5_ra.median(),1),
    round(reference.q6_ra.median(),1),
    round(reference.q7_ra.median(),1),
    round(reference.q8_ra.median(),1),
    round(reference.q9_ra.median(),1),
    round(reference.q10_ra.median(),1),
    round(reference.q11_ra.median(),1),
    round(reference.q12_ra.median(),1),
    round(reference.q13_ra.median(),1),
    round(reference.q14_ra.median(),1),
    round(reference.q15_ra.median(),1),
    ]
    return data 

def get_academic_programs():
    counter = []
    programs = d.Dep_code.unique()
    for program in programs:
        program = int(program)
        counter.append(program)
    return counter

def make_header():
    header = [ 'Program',
            'R_rate',
            'I_sat_ra',
            'C_sat_ra',
            'q1_ra',
            'q2_ra',
            'q3_ra',
            'q4_ra',
            'q5_ra',
            'q6_ra',
            'q7_ra',
            'q8_ra',
            'q9_ra',
            'q10_ra',
            'q11_ra',
            'q12_ra',
            'q13_ra',
            'q14_ra',
            'q15_ra',
            ]
    write_row(header)
    return

def write_row(row):
    with open(CONFIG[1],'a') as csvfile:
        w = csv.writer(csvfile)
        w.writerow(row)
    return

f = pd.ExcelFile(CONFIG[0])
d = f.parse('Online')

make_header()

all_online_courses = get_median_data(d)
all_online_courses.insert(0, 'All Online Courses')
write_row(all_online_courses)

academic_programs = get_academic_programs()
for program in academic_programs:
    by_program = d.ix[d.Dep_code == program]
    by_program_data = get_median_data(by_program)
    by_program_data.insert(0, program)
    write_row(by_program_data)
