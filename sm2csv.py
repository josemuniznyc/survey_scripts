#!/usr/bin/env python

"""
sm2csv.py

Convert a directory of SM excel files into a CSV of student course evals
"""

import os
import xlrd
import csv
import collections

CONFIG = { "excel" : "/Users/jmuniz/Dropbox/SPS/f14/online/raw_data",
        "output" : "/Users/jmuniz/Dropbox/SPS/f14/online_evals.csv",
        "roster" : "/Users/jmuniz/Dropbox/SPS/data/course_roster.xlsx"}

def collect_data(excel_file):
    
    workbook = xlrd.open_workbook(os.path.join(CONFIG["excel"], excel_file))
    worksheet = workbook.sheet_by_name('Questions')

    #create dicts to hold values
    survey_values = collections.OrderedDict()
    lickert_scores = collections.OrderedDict()
    satisfaction_ratings = collections.OrderedDict()

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

        comments_col = worksheet.col_values(2)
        comments_indicies = [i for i, x in enumerate(comments_col) \
            if x == u'Response Count'] #Find all start points for comments

        def extract_comment(comment_num):
            
            position = comment_num - 1
            count_cell = comments_indicies[position] + 1
            num_responses = int(comments_col[count_cell])
            comment = "No responses to this question."

            if num_responses == 0:
                comment = "No responses to this question."
            else:
                start_cell = count_cell + 5
                end_cell = start_cell + num_responses
                comments = comments_col[start_cell : end_cell]
                comments = [x for x in comments]
                collected_strings = []
                for i, x in enumerate(comments, start=1):
                    string = "{}: {}".format(i, x)
                    collected_strings.append(string)
                    comment = "\n".join(collected_strings)

            return comment

        survey_values["comment_1"] = extract_comment(1)
        survey_values["comment_2"] = extract_comment(2)
        survey_values["comment_3"] = extract_comment(3)
        survey_values["comment_4"] = extract_comment(4)
        survey_values["comment_5"] = extract_comment(5)

    def get_course_data():
        survey_title = worksheet.row_values(0, 0, 1)
        survey_title = survey_title[0].split() #ST is in a list, have to break
        survey_values["term"] = survey_title[0]
        survey_values["year"] = survey_title[1]
        survey_values["course"] = "{} {}".format(survey_title[2], 
            survey_title[3]) #Get Ins. name from external source. Too fragile.
        response_col = worksheet.col_values(9)
        response_nums = [x for x in response_col if isinstance(x, float)]
        survey_values["response_count"] = int(max(response_nums))
        
        # course = ROSTER[(ROSTER["course"] == survey_values["course"]) & 
        #     (ROSTER["term_year"] == int(survey_values["year"])) & 
        #     (ROSTER["term_semester"] == survey_values["term"])]
        # survey_values["course_id"] = int(course.index)

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
    get_satisfaction_ratings()
    get_course_data()
    get_open_comments()

    for question, value_lst in lickert_scores.items():
        break_list_values(question, value_lst, "lickert")

    for question, value_lst in satisfaction_ratings.items():
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