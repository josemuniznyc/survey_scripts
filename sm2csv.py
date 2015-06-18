#!/usr/bin/env python
"""
sm2csv.py

Convert a directory of SM excel files into a CSV of student course evals
"""

import os
import xlrd
import csv
from collections import OrderedDict

CONFIG = {}
CONFIG['working_dir'] = 'Dropbox/Documents/SPS/evals/s15/online_courses/sm/'
CONFIG['course_roster'] = 'byTerm_courses.xlsx'
CONFIG['excel_dir'] = 'excel_files/'
CONFIG['output_file'] = 'sm_online_evals_responses.csv'


def collect_data(excel_file):
    """
    Establish connection to excel file, then parse quantitative data.
    """
    print('Opening {file} for processing'.format(file=excel_file))
    wk = xlrd.open_workbook(os.path.join(CONFIG["excel_dir"], excel_file))
    ws = wk.sheet_by_index(0)

    # create dicts to hold values
    survey_values = OrderedDict()
    lickert_scores = OrderedDict()
    satisfaction_ratings = OrderedDict()

    def get_lickert_scores():
        start_col = 2
        end_col = 9

        # grab the question values
        # "1. In this course."
        lickert_scores["q1"] = ws.row_values(3, start_col, end_col)
        lickert_scores["q2"] = ws.row_values(4, start_col, end_col)
        lickert_scores["q3"] = ws.row_values(5, start_col, end_col)
        lickert_scores["q4"] = ws.row_values(6, start_col, end_col)
        lickert_scores["q5"] = ws.row_values(7, start_col, end_col)
        lickert_scores["q6"] = ws.row_values(8, start_col, end_col)
        lickert_scores["q7"] = ws.row_values(9, start_col, end_col)

        # "2. The instructor in the course"
        lickert_scores["q8"] = ws.row_values(16, start_col, end_col)
        lickert_scores["q9"] = ws.row_values(17, start_col, end_col)
        lickert_scores["q10"] = ws.row_values(18, start_col, end_col)
        lickert_scores["q11"] = ws.row_values(19, start_col, end_col)
        lickert_scores["q12"] = ws.row_values(20, start_col, end_col)

        # "3. The instructor in the course:"
        lickert_scores["q13"] = ws.row_values(27, start_col, end_col)
        lickert_scores["q14"] = ws.row_values(28, start_col, end_col)
        lickert_scores["q15"] = ws.row_values(29, start_col, end_col)

    def get_satisfaction_ratings():
        start_col = 2
        end_col = 8

        satisfaction_ratings["q16"] = ws.row_values(36, start_col, end_col)
        satisfaction_ratings["q17"] = ws.row_values(43, start_col, end_col)

    def get_open_comments():
        open_comment_col = ws.col_values(2)
        comment_locs = []
        for c, x in enumerate(open_comment_col):
            if x == "Response Count":
                count = int(ws.cell_value(c + 1, 2))  # how many responses
                first_row = c + 6  # Row offset to the first response
                comment_locs.append((count, first_row))
        survey_values['c1'] = extract_response(
            comment_locs[0], open_comment_col)
        survey_values['c2'] = extract_response(
            comment_locs[1], open_comment_col)
        survey_values['c3'] = extract_response(
            comment_locs[2], open_comment_col)
        survey_values['c4'] = extract_response(
            comment_locs[3], open_comment_col)
        survey_values['c5'] = extract_response(
            comment_locs[4], open_comment_col)

    def extract_response(location, data):
        if location[0] == 0:
            open_comment_response = "No responses."
        else:
            """
            * count to where comments start
            * up to the number of lines (one per comment)
            * then format the reponses before returning to dictionary
            """
            open_comment_response = data[
                location[1]: location[1] + location[0]
            ]
            temp_holder = []
            for count, statement in enumerate(open_comment_response, 1):
                # count_string = str(count)
                # temp = count_string + ": " + statement
                temp = "{counter}: {survey_response}".format(
                    counter=count,
                    survey_response=statement
                )
                temp_holder.append(temp)
            open_comment_response = "\n".join(temp_holder)
        return open_comment_response

    def get_course_data():
        survey_title = ws.row_values(0, 0, 1)
        survey_title = survey_title[0].split()  # ST is in a list,have to break
        survey_values["Term"] = survey_title[0]
        survey_values["Year"] = survey_title[1]
        survey_values["Course"] = "{} {}".format(
            survey_title[2], survey_title[3])
        response_col = ws.col_values(9)
        response_nums = [x for x in response_col if isinstance(x, float)]
        survey_values["Response count"] = int(max(response_nums))

    def break_list_values(key, value, q_type):
        if q_type == "lickert":
            survey_values[key + "_SA"] = value[0]
            survey_values[key + "_A"] = value[1]
            survey_values[key + "_N"] = value[2]
            survey_values[key + "_D"] = value[3]
            survey_values[key + "_SD"] = value[4]
            survey_values[key + "_NA"] = value[5]
            survey_values[key + "_RA"] = round(value[6], 1)
        elif q_type == "satisfaction":
            survey_values[key + "_E"] = value[0]
            survey_values[key + "_A"] = value[1]
            survey_values[key + "_AA"] = value[2]
            survey_values[key + "_BA"] = value[3]
            survey_values[key + "_NA"] = value[4]
            survey_values[key + "_RA"] = round(value[5], 1)
        else:
            pass

    # Release the funcs!

    get_lickert_scores()
    get_open_comments()
    get_satisfaction_ratings()
    get_course_data()

    for question, value_lst in lickert_scores.items():
        break_list_values(question, value_lst, "lickert")

    for question, value_lst in satisfaction_ratings.items():
        break_list_values(question, value_lst, "satisfaction")

    print('Completed processing of {file}'.format(file=excel_file))
    return survey_values


def write_csv(list_of_dicts):
    fieldnames = list(list_of_dicts[0].keys())
    f = open(CONFIG['output_file'], 'wt')
    try:
        writer = csv.DictWriter(f, lineterminator='\n', fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(list_of_dicts)
    finally:
        f.close()

# Get in place
os.chdir(os.path.join(os.path.expanduser('~'), CONFIG['working_dir']))

# Grab all excel files parse data inside them pass data to CSV
collected_data = [
    collect_data(x) for x in os.listdir(CONFIG["excel_dir"])]
write_csv(collected_data)
print("Files processed; data avaliable at {}".format(CONFIG['output_file']))
