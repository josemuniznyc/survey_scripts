#!/usr/bin/env python
# -- coding: utf-8 --

import pandas as pd
import os

EXCEL_DIRECTORY = "/Users/jmuniz/Dropbox/Transfer/s14-f2f-excel"
ROSTER = "/Users/jmuniz/Desktop/Roster_to_m14.xlsx"

def cal_sat_rates(excel_file):
    course = pd.read_excel(excel_file)
    course_ID = excel_file.split("-")
    course_ID = course_ID[2]
    response_count = course.Q3.count()
    ins_sat = round(course.Q13.mean(),1)
    cou_sat = round(course.Q14.mean(),1)
    return (course_ID, response_count, ins_sat, cou_sat)

excel_files = [x for x in os.listdir(EXCEL_DIRECTORY) if x[-5:] == ".xlsx"]
rates_list = [cal_sat_rates(x) for x in excel_files]
print rates_list