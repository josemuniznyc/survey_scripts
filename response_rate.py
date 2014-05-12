#!/usr/bin/env python
# -- coding: utf-8 --
"""
Convert downloaded SM html to produce a response rate report
"""

import pandas as pd
import os
from datetime import date

CONFIG = ['/Users/jmuniz/Dropbox/SPS/S14/Midpoint Response Rate Report/']
RESPONSE_DATA = []

def convert(html_file):
    dfs = pd.read_html(html_file)
    data = dfs[1]
    data = data[[0,6]]
    data = data[1:10]
    data_aslist = list(data.values.tolist())
    RESPONSE_DATA.extend(data_aslist)
    return

def load_html():
    """Return a list with surveys in data directory"""
    surveys = os.listdir(CONFIG[0])
    return surveys

def create_csv(response_list):
    headers = ['Course','Responses']
    data = pd.DataFrame(response_list, headers)
    today = date.today()
    csv_file1 = "/Users/jmuniz/Desktop/" + "ResponseRates-"
    csv_file2 = today.strftime("%y%m%d")
    csv_file3 = ".csv"
    csv_filename = csv_file1 + csv_file2 + csv_file3
    data.to_csv(csv_filename, index = False)

FILES = load_html()

for i in FILES:
    if i[-5:] != ".html":
        pass
    else:
        file_toread = CONFIG[0] + i
        convert(file_toread)

create_csv(RESPONSE_DATA)