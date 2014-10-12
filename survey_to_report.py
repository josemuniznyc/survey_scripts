#!/usr/bin/env python
# -- coding: utf-8 --
"""
Create CSV report from surveymonkey excel spreadsheets 
"""

import os, csv, xlrd

CONFIG = ['/Users/jmuniz/Dropbox/Statistics/Evals/',
        '/Users/jmuniz/Desktop/evaluation_report.csv',
        '/Users/jmuniz/Desktop/control_sheet.xls']

RESPONSE_COLLECTOR = []
SEAT_COUNT = []

class Survey:
    """Represents an individual survey"""
    def __init__(self, filename):
        self.responses = []
        workbook = xlrd.open_workbook(filename)
        self.worksheet = workbook.sheet_by_index(0)
        self.sid = []
        self.num_of_responses = []

    def get_course_information(self):
        title = self.worksheet.row_values(0, end_colx=1)
        title = title[0].rsplit() 
        title = title[3:] #Remove survey title
        del title[2] #Remove first dash
        del title[4] #...and now the second
        course_num_sec = title[3]
        course_sec = course_num_sec[-2:]
        course_num = course_num_sec[:-3]
        del title[3] #Remove the number/section
        title.insert(3, course_num) 
        title.insert(4, course_sec)
        instructor = title[5:]
        instructor_first = instructor[-1]
        instructor_last = instructor[:-1]
        title = title[:5] #We need to do something for non-standard last names
        instructor_last = ' '.join(instructor_last)
        instructor_first = instructor_first.capitalize()
        instructor_last = instructor_last.title()
        instructor_last = instructor_last[:-1]
        title.append(instructor_last) 
        title.append(instructor_first)
        self.responses.extend(title) 
        self.sid.extend(title[2:5]) #Use course info to ID the survey
        self.sid.append(instructor_last[:-1]) #Remove the comma

    def get_lickert_scores(self):
        score_holder = [
            self.worksheet.row_values(3, 2, 9),
            self.worksheet.row_values(4, 2, 9),
            self.worksheet.row_values(5, 2, 9),
            self.worksheet.row_values(6, 2, 9),
            self.worksheet.row_values(7, 2, 9),
            self.worksheet.row_values(8, 2, 9),
            self.worksheet.row_values(9, 2, 9),
            self.worksheet.row_values(16, 2, 9),
            self.worksheet.row_values(17, 2, 9),
            self.worksheet.row_values(18, 2, 9),
            self.worksheet.row_values(19, 2, 9),
            self.worksheet.row_values(20, 2, 9),
            self.worksheet.row_values(27, 2, 9),
            self.worksheet.row_values(28, 2, 9),
            self.worksheet.row_values(29, 2, 9),
            self.worksheet.row_values(36, 2, 8),
            self.worksheet.row_values(43, 2, 8),
            ]

        sample_row = score_holder[0] #Grab a sample row and cal responses
        self.num_of_responses = int(sum(sample_row[:-1]))#Drop RA col
        
        for s in score_holder:
            self.responses.extend(s)

    def get_comments(self):
        comments = []
        counter = []
        
        col = self.worksheet.col_values(2) #Grab the whole col
        for c, x in enumerate(col):
            if x == "Response Count":
                count = int(self.worksheet.cell_value(c+1, 2)) # how many
                first_row = c+6 # Row offset to the first response
                counter.append([count, first_row]) 
        for count, first_row in counter:
            if count == 0:
                comments.append('No responses.')#SM skips rows with 0 responses 
            else:
                replies = []
                reply_holder = self.worksheet.col_values(2, first_row,
                    count + first_row)
                for j, m in enumerate(reply_holder):
                    statement = str(j+1) + ': ' + m
                    replies.append(statement)
                replies = '\n'.join(replies) 
                replies = replies.encode('utf8')
                comments.append(replies)
        for r in comments: #Loop through the comments and send to responses
            self.responses.append(r)
        return

    def get_response_nums(self):
        response_nums = [self.num_of_responses,]
        course = self.sid[0] + ' ' + self.sid[1] + '.' + self.sid[2]
        for item in SEAT_COUNT:
            if item[0] == course.encode('utf8'):
                seats = item[1]
                response_rate = float(self.num_of_responses) / float(seats)
                response_rate = round(response_rate, 3)
                dep_code = item[2]
                course_title = item[3]
                response_nums.append(seats)
                response_nums.append(response_rate)
                response_nums.append(dep_code)
                response_nums.append(course_title)
        self.responses.extend(response_nums)
        return
    
    def publish_responses(self):
        """Called to output response list"""
        RESPONSE_COLLECTOR.append(self.responses)
        return

def write_header():
    header_row = [
        'Term', 'Year', 'C_prefix', 'C_num', 'C_sec', 'I_ln', 'I_fn', 'q1_sa',
        'q1_a', 'q1_neu', 'q1_d', 'q1_sd', 'q1_na', 'q1_ra', 'q2_sa', 'q2_a',
        'q2_neu', 'q2_d', 'q2_sd', 'q2_na', 'q2_ra', 'q3_sa', 'q3_a','q3_neu',
        'q3_d', 'q3_sd', 'q3_na', 'q3_ra', 'q4_sa', 'q4_a', 'q4_neu', 'q4_d',
        'q4_sd', 'q4_na', 'q4_ra', 'q5_sa', 'q5_a', 'q5_neu', 'q5_d', 'q5_sd',
        'q5_na', 'q5_ra', 'q6_sa', 'q6_a', 'q6_neu', 'q6_d', 'q6_sd', 'q6_na',
        'q6_ra', 'q7_sa', 'q7_a', 'q7_neu', 'q7_d', 'q7_sd', 'q7_na', 'q7_ra',
        'q8_sa', 'q8_a', 'q8_neu', 'q8_d', 'q8_sd', 'q8_na', 'q8_ra', 'q9_sa',
        'q9_a', 'q9_neu', 'q9_d', 'q9_sd', 'q9_na', 'q9_ra', 'q10_sa', 'q10_a',
        'q10_neu', 'q10_d', 'q10_sd', 'q10_na', 'q10_ra', 'q11_sa', 'q11_a',
        'q11_neu', 'q11_d', 'q11_sd', 'q11_na', 'q11_ra', 'q12_sa', 'q12_a',
        'q12_neu', 'q12_d', 'q12_sd', 'q12_na', 'q12_ra', 'q13_sa', 'q13_a',
        'q13_neu', 'q13_d', 'q13_sd', 'q13_na', 'q13_ra', 'q14_sa', 'q14_a',
        'q14_neu', 'q14_d', 'q14_sd', 'q14_na', 'q14_ra', 'q15_sa', 'q15_a',
        'q15_neu', 'q15_d', 'q15_sd', 'q15_na', 'q15_ra', 'I_sat_e',
        'I_sat_aa', 'I_sat_a', 'I_sat_ba', 'I_sat_na', 'I_sat_ra', 'C_sat_e',
        'C_sat_aa', 'C_sat_a', 'C_sat_ba', 'C_sat_na', 'C_sat_ra', 'R_1',
        'R_2', 'R_3', 'R_4', 'R_5', 'R_count', 'Seats', 'R_rate', 'Dep_code',
        'C_title',
        ] #OUCH: Should do something with this. 
    fake_rows = [header_row,] #We need this in order to use write_rows
    write_rows(fake_rows)
    return

def write_rows(rows):
    for r in rows:
        with open(CONFIG[1], 'a') as csvfile:
            w = csv.writer(csvfile)
            w.writerow(r)
    return

def rename_file():
    return

def get_seat_count():
    """load the registrar's seat report"""
    control_wk = xlrd.open_workbook(CONFIG[2])
    control_sheet = control_wk.sheet_by_index(0)
    count = control_sheet.nrows
    count = range(count)
    for x in count:
        if x == 0:
            pass
        else:
            values = control_sheet.row_values(x, 0, 13)
            # course, seats, dep_code, title
            values = [values[2], values[6], values[0], values[3]]
            values[1] = int(values[1])
            values[2] = int(values[2])
            SEAT_COUNT.append(values)

#Loop through all files in directory
os.chdir(CONFIG[0])
files = os.listdir(CONFIG[0])

get_seat_count()

for f in files:
    s = Survey(f)
    s.get_course_information()
    s.get_lickert_scores()
    s.get_comments()
    s.get_response_nums()
    s.publish_responses()

write_header()
write_rows(RESPONSE_COLLECTOR)
