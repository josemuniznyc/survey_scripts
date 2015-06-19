#! /usr/bin/python

"""
process_soars_csv
"""

import os
import pandas as pd

CONFIG = {}
CONFIG['working_dir'] = 'Dropbox/Documents/SPS/evals/s15/online_courses/soars/'
CONFIG['csv_dir'] = 'data/'
CONFIG['output_file'] = 'soars_data.csv'


def process(x):
    df = pd.read_csv(x, encoding='utf-8')
    meta_cols = df.iloc[:,0:4]
    cate_cols = df.iloc[:,4:21]
    comm_cols = df.iloc[:,21:]

    for col in cate_cols:
        df[x] = df[x].astype('category')

    


# Get into working directory
os.chdir(os.path.join(os.path.expanduser('~'), CONFIG['working_dir'],
                     CONFIG['csv_dir']))

[process(x) for x in os.listdir(".") if [-4:] == ".csv"]

