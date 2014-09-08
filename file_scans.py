#!/usr/bin/env python

import shutil
import os

BASE_DIRECTORY = "/Users/jmuniz/Dropbox/SPS/scans_sce/"

scans = [x for x in os.listdir(os.path.join(BASE_DIRECTORY,"current/"))if x[-4:] == ".pdf"]
for scan in scans:
    term = scan[-7:-4]
    term = term.lower()
    source = os.path.join(BASE_DIRECTORY,"current/",scan)
    destination = os.path.join(BASE_DIRECTORY,term)
    shutil.copy2(source, destination)
