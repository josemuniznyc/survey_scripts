#!/usr/bin/env python
# -- coding: utf-8 --
"""
sm2csv.py

Convert a directory of SM excel files into a CSV of student course evals
"""

import os
import xlrd
import csv
from future.builtins import round

CONFIG = {"data_dir":"/Users/jmuniz/Dropbox/SPS/data", 
        "excel":"/Users/jmuniz/Dropbox/SPS/m14/online/raw_data",
        "output":"/Users/jmuniz/Dropbox/SPS/m14/online_evals.csv"}

def collect_data(excel_file):
    
    workbook = xlrd.open_workbook(os.path.join(CONFIG["excel"], excel_file))
    worksheet = workbook.sheet_by_name('Questions')

    #create dicts to hold values
    survey_values = {}
    lickert_scores = {}
    satisfaction_ratings = {}

    def get_lickert_scores():
        
        start_col = 2
        end_col = 9

        #grab the question values
        #"1. In this course."
        lickert_scores["q1"] = worksheet.row_values(3, start_col, end_col)
        lickert_scores["q2"] = worksheet.row_values(4, start_col, end_col)
        lickert_scores["q3"] = worksheet.row_values(5, start_col, end_col)
        lickert_scores["q4"] = worksheet.row_values(6, start_col, end_col)
        lickert_scores["q5"] = worksheet.row_values(7, start_col, end_col)
        lickert_scores["q6"] = worksheet.row_values(8, start_col, end_col)
        lickert_scores["q7"] = worksheet.row_values(9, start_col, end_col)

        #"2. The instructor in the course"
        lickert_scores["q8"] = worksheet.row_values(16, start_col, end_col)
        lickert_scores["q9"] = worksheet.row_values(17, start_col, end_col)
        lickert_scores["q10"] = worksheet.row_values(18, start_col, end_col)
        lickert_scores["q11"] = worksheet.row_values(19, start_col, end_col)
        lickert_scores["q12"] = worksheet.row_values(20, start_col, end_col)
        
        #"3. The instructor in the course:"
        lickert_scores["q13"] = worksheet.row_values(27, start_col, end_col)
        lickert_scores["q14"] = worksheet.row_values(28, start_col, end_col)
        lickert_scores["q15"] = worksheet.row_values(29, start_col, end_col)

    def get_satisfaction_ratings():
        start_col = 2
        end_col = 8

        satisfaction_ratings["q16"] = worksheet.row_values(36, start_col, end_col)
        satisfaction_ratings["q17"] = worksheet.row_values(43, start_col, end_col)

    def get_open_comments():
        pass
    
    def get_course_data():
        survey_title = worksheet.row_values(0, 0, 1)
        survey_title = survey_title[0].split() #ST is in a list, have to break
        survey_values["Term"] = survey_title[0]
        survey_values["Year"] = survey_title[1]
        survey_values["Course"] = "{} {}".format(survey_title[2], 
            survey_title[3]) #Get Ins. name from external source. Too fragile.
        response_col = worksheet.col_values(9)
        response_nums = [x for x in response_col if isinstance(x, float)]
        survey_values["Response count"] = int(max(response_nums))

    def break_list_values(key, value, q_type):
        if q_type == "lickert":
            survey_values[key+"_SA"] = value[0]
            survey_values[key+"_A"] = value[1]
            survey_values[key+"_N"] = value[2]
            survey_values[key+"_D"] = value[3]
            survey_values[key+"_SD"] = value[4]
            survey_values[key+"_NA"] = value[5]
            survey_values[key+"_RA"] = round(value[6], 1)
        elif q_type == "satisfaction":
            survey_values[key+"_E"] = value[0]
            survey_values[key+"_A"] = value[1]
            survey_values[key+"_AA"] = value[2]
            survey_values[key+"_BA"] = value[3]
            survey_values[key+"_NA"] = value[4]
            survey_values[key+"_RA"] = round(value[5], 1)
        else:
            pass

    #Release the funcs!

    get_lickert_scores()
    get_open_comments()
    get_satisfaction_ratings()
    get_course_data()

    for question, value_lst in lickert_scores.iteritems():
        break_list_values(question, value_lst, "lickert")

    for question, value_lst in satisfaction_ratings.iteritems():
        break_list_values(question, value_lst, "satisfaction")        

    return survey_values

def write_csv(list_of_dicts):
    fieldnames = list(list_of_dicts[0].keys())
    f = open(CONFIG["output"], 'wt')
    try:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(list_of_dicts)
    finally:
        f.close()

raw_files = [x for x in os.listdir(CONFIG["excel"]) if x[-4:] == ".xls"]
collected_data = [collect_data(x) for x in raw_files]
write_csv(collected_data)

print("Files processed; data avaliable at {}".format(CONFIG["output"]))