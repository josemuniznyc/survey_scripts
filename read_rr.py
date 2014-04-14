#!/usr/bin/env python
# -- coding: utf-8 --
"""
Read and parse the registrar's report
"""

import csv

def load_report():
    report = open('/Users/jmuniz/Dropbox/Dev/survey_scripts/updated_rr.txt',
            'r')
    lines = report.readlines()
    report.close()
    return lines

def remove_headers_n_totals(lines):
    holder = []
    for line in lines:
        if line[0].isdigit():
            holder.append(line)
        else:
            pass
    return holder

def split(lines):
    holder = []
    for line in lines:
        holder.append(line.split('\t'))
    return  holder

def remove_non_course_lines(lines):
    holder = []
    for line in lines:
        if len(line) > 13:
            holder.append(line)
    return holder

def clean_lines(lines):
    holder = []
    for line in lines:
        course = str(line[3]) + " " + str(line[4]) + "." + str(line[6])
        seats = line[-4]
        values = [course, seats]
        holder.append(values)
    return holder

def write_rows(rows):
    with open('/Users/jmuniz/Desktop/updated_seat_counts.csv', 'w') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(rows)
    return
    
text = load_report()
text = remove_headers_n_totals(text)
text = split(text)
text = remove_non_course_lines(text)
text = clean_lines(text)
write_rows(text)
print "Done & done."
