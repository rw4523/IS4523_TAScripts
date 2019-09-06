# Script to read csv or xls file from Blackboard "Weekly Report" tool and output a list of things to complete for the week.
# Should output everything that was submitted for that week, including items that may have already been completed
# Will add removing entries already graded by TA code later..

import os
import pandas as pd
from datetime import date, datetime, timedelta

directory = os.getcwd()
file = input('\nEnter file path to .csv or .xls file:  ')

# file type check
if '.xls' in file:
    df = pd.read_excel(file)
elif '.csv' in file:
    df = pd.read_csv(file)
else:
    print('Error: Input file must be CSV (.csv) or XLS (.xls)')

# Only want gradable entries made by Students of Module Quiz and Lab
df = df[df['Column'].str.contains('Module Quiz') | df['Column'].str.contains('Lab')]
df = df[df['Last Edited by: Role'].str.contains('S')]
df = df.drop(columns='Last Edited by: IP Address')
#df.head(10)

### FORMAT
#
# Student Name: \t [User]
# ABC123: \t [Last Edited by: Username]
# Status: \t [Event]
# Submit Date: \t [Attempt Submitted]
# Assignment: \t [Column]
# 
# \n\n
###

# Output printing:
print('To-do list for week of: '+ datetime.today().strftime('%B %d, %Y'))
output_entry = ''

name_list    = df['Last Edited by: Name'].tolist()
abc123_list  = df['Last Edited by: Username'].tolist()
status_list  = df['Event'].tolist()
submit_list  = df['Attempt Submitted'].tolist()
task_list    = df['Column'].tolist()

for entry in range(len(name_list)):   # issue: does not iterate more than once. (Will fix it later)
#for entry in range(0, multiple+1):
    output_entry = ('\nStudent Name: \t ' + name_list[entry] 
                    + '\nABC123: \t ' + abc123_list[entry]
                    + '\nStatus: \t ' + status_list[entry]
                    + '\nSubmit Date: \t ' + submit_list[entry]
                    + '\nAssignment: \t ' + task_list[entry] + '\n\n')
print(output_entry)
