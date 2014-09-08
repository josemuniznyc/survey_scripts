#!/usr/bin/env python
# -- coding: utf-8 --

import pandas as pd
import os

EXCEL_DIRECTORY = "/Users/jmuniz/Dropbox/SPS/S14/F2F/jsm/"
ROSTER = "/Users/jmuniz/Desktop/Roster_to_m14.xlsx"

def cal_sat_rates(excel_file):
    course = pd.read_excel(excel_file)
    course_ID = excel_file.split("-")
    course_ID = course_ID[-2]
    response_count = course.Q3.count()
    ins_sat = round(course.Q13.mean(),1)
    cou_sat = round(course.Q14.mean(),1)
    return (course_ID, response_count, ins_sat, cou_sat)

excel_files = [x for x in os.listdir(EXCEL_DIRECTORY) if x[-5:] == ".xlsx"]
rates_list = [cal_sat_rates(os.path.join(EXCEL_DIRECTORY,x)) for x in excel_files]

eval_ratings = pd.DataFrame(rates_list, \
    columns=["ID", "Response_count", "Ins_Sat_Rating", "Cou_Sat_Rating"])

eval_ratings.to_csv("/Users/jmuniz/Desktop/eval_ratings.csv", index=False)