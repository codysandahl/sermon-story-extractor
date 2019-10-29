# coding: utf-8

import docx
import os
from dateutil import parser

# find all sermon docs
sermon_files = [os.path.join(root, name)
             for root, dirs, files in os.walk(".")
             for name in files
             if name == "Sermon.docx" or name == "Sermon - Pulpit.docx"]

# extract date from path
dt_start = parser.parse("2019-10-13") # date to start collecting sermons
sermons = []
for path in sermon_files:
    path_parts = path.split(" - ")
    path_date = (path_parts[0])[2:] # NOTE: "2:" gets rid of "./" on the path
    # get rid of "(Ash Wednesday)" and "(Easter)" descriptors
    if len(path_date) > 10:
        path_date = path_date[0:10]
    #print(path_date)
    # convert string to date
    dt = parser.parse(path_date)
    #print(dt)
    # add to list ONLY IF it's at least as new as the dt_start variable
    if (dt >= dt_start):
        sermons.append({'date': dt, 'path': path })
print(sermons)

# create the output file
output = open("sermon_stories.txt", "w")

# process all the sermons
for sermon in reversed(sermons): # NOTE: reversed makes it newest to oldest
    dt = sermon['date']
    path = sermon['path']
    output.write(path[2:]+"\n")
    # open the file
    doc = docx.Document(path)
    # print the bolded statements
    for par in doc.paragraphs:
        for r in par.runs:
            if r.bold:
                output.write(r.text+"\n")
    # print a spacer
    output.write("\n")
    output.write("###################################################################\n")
    output.write("\n")

# save and close
output.close()


