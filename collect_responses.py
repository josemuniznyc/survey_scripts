#!/usr/bin/env python
# -- coding: utf-8 --
"""
Given a directory of evaluations, return a csv file that includes the responses
in those surveys (one survey per line)

TO DO: aggregate comments
       clean data after calling write

"""
import os
import xlrd
import csv
import sys

CONFIG = {
    "data_directory": "/Users/jmuniz/Dropbox/Statistics/Evals",
    "csv_filename": "evaluation_data.csv", }


def write(data):
    """Write collected data to csv file"""
    with open(CONFIG['csv_filename'], 'a') as csvfile:
        w = csv.writer(csvfile, delimiter=',', quotechar='|')
        w.writerow(data)


def load_surveys():
    """Return a list with surveys in data directory"""
    os.chdir(CONFIG['data_directory'])
    surveys = os.listdir(CONFIG['data_directory'])
    return surveys


def parse_filename(filename):
    """Given a filename, generate metadata for survey"""
    filename = filename[:-4]
    filename = filename.split("-", 1)
    course_prefix = filename[0]
    course_number = course_prefix[:-3]
    course_section = course_prefix[-2:]
    temp = filename[1].rsplit("-", 1)
    term = temp[1]
    semester = term[:1]
    if semester == 'F':
        semester = 'Fall'
    elif semester == 'S':
        semester = 'Spring'
    else:
        pass
    year = "20" + term[-2:]
    instructor = temp[0]
    instructor = instructor.split(",")
    instructor_last = instructor[0].title()
    instructor_first = instructor[1].strip()
    instructor_first = instructor_first.title()
    course_information = [course_number, course_section, instructor_first,
                          instructor_last, semester, year]
    return course_information


def grab_likert_scores(sheet):
    """Grab all likert scale data return as a list of lists"""

    scores = [
        sheet.row_values(3, 2, 9),
        sheet.row_values(4, 2, 9),
        sheet.row_values(5, 2, 9),
        sheet.row_values(6, 2, 9),
        sheet.row_values(7, 2, 9),
        sheet.row_values(8, 2, 9),
        sheet.row_values(9, 2, 9),
        sheet.row_values(16, 2, 9),
        sheet.row_values(17, 2, 9),
        sheet.row_values(18, 2, 9),
        sheet.row_values(19, 2, 9),
        sheet.row_values(20, 2, 9),
        sheet.row_values(27, 2, 9),
        sheet.row_values(28, 2, 9),
        sheet.row_values(29, 2, 9),
        sheet.row_values(36, 2, 8),
        sheet.row_values(43, 2, 8),
    ]
    return scores


def grab_comments(number_and_location):
    """ Given a tuple of response counts and first lines, pull out the comments
    """
    comments = []
    for c, x in number_and_location:
        if c == 0:
            comments.append((0, "No responses."))
        else:
            comments.append((c, worksheet.col_values(2, x, x + c)))
    return comments


def get_response_rows(data):
    """ Given a slice of column data, find response rows """
    responses = []
    for c, x in enumerate(data):
        if x == "Response Count":
            count = int(worksheet.cell_value(c + 1, 2))  # how many responses
            first_row = c + 6  # Row offset to the first response
            responses.append((count, first_row))
    response_by_question = grab_comments(responses)
    return response_by_question


def collect_comments(sheet):
    """ Grab all rows in column and return as a list of strings """
    collected_rows = sheet.col_values(2)
    survey_responses = get_response_rows(collected_rows)
    return survey_responses

if len(sys.argv) > 1:
    CONFIG['data_directory'] = sys.argv[1]
else:
    pass

EVALS = load_surveys()

print("Checking ", CONFIG['data_directory'], " for evaluations.")

for i in EVALS:
    if i[-4:] != ".xls":
        pass
    else:
        course = parse_filename(i)
        workbook = xlrd.open_workbook(i)
        worksheet = workbook.sheet_by_name(u"Questions")
        likert_scores = grab_likert_scores(worksheet)
        open_end_responses = collect_comments(worksheet)
        line = [course, likert_scores, open_end_responses]
        write(line)
        print(course[0], ".", course[1], " completed.")

os.rename(CONFIG['csv_filename'], "Current_" + CONFIG['csv_filename'])
print("Evaluation data written to ", "Current_", CONFIG['csv_filename'])
