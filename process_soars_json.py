#! /usr/bin/python

"""
process_soars_json
"""
import os
import json
import pandas as pd

# Storage
CONFIG = {}
CONFIG['working_dir'] = 'Dropbox/Documents/SPS/evals/s15/online_courses/soars/'
CONFIG['sid_id_map'] = 'soars_sid_id.xlsx'
CONFIG['json_file'] = 'raw_data.json'

# Get into working directory
os.chdir(os.path.join(os.path.expanduser('~'), CONFIG['working_dir']))

# Functions


# def get_response_counts(survey):
#     """
#     * pull response count
#     * match with eval_id
#     * send to REPSONSE_COUNTS
#     """
#     responses = surveys[survey]['metadata']['total responses']
#     eval_id = sid_id_map.get(int(survey), 'eval_id')
#     course = sid_id_map.get(int(survey), 'course')
#     print('{} has {} responses'.format(
#         course, responses))
#     REPSONSE_COUNTS[str(eval_id)] = responses


def create_csv(survey):
    """
    * Create a pandas dataframe using ['responses'] as data, and ['questions']
      for columns
    * Lookup eval_id and use as the filename
    """
    survey_df = pd.DataFrame(surveys[survey]['responses'],
                             columns=surveys[survey]['questions'])
    col_names = ['id',
                 'submit_date',
                 'marker',
                 'language',
                 'q1',
                 'q2',
                 'q3',
                 'q4',
                 'q5',
                 'q6',
                 'q7',
                 'q8',
                 'q9',
                 'q10',
                 'q11',
                 'q12',
                 'q13',
                 'q14',
                 'q15',
                 'q16',
                 'q17',
                 'c1',
                 'c2',
                 'c3',
                 'c4',
                 'c5'
                 ]
    survey_df.columns = col_names
    x_course = sid_id_map.get_value(int(survey), 'course')
    survey_df.to_csv('{course}.csv'.format(course=x_course), index=False,
                     encoding='utf-8')
    print('Created CSV file for {course}'.format(
        course=x_course))

# Load sid_id_map

print('Loading {file}'.format(file=CONFIG['sid_id_map']))
sid_id_map = pd.read_excel(CONFIG['sid_id_map'], index_col=0)

print('Loading JSON file {file}'.format(file=CONFIG['json_file']))
surveys = json.load(open(CONFIG['json_file']))
print('JSON file {file} successfully loaded'.format(file=CONFIG['json_file']))

# [get_response_counts(x) for x in surveys]
[create_csv(x) for x in surveys]
