# Script made for use with Python 3 on Windows 
# Please consult the README.md files before running.
# OSX might need to tweak directory path, or put script in same directory as the Module_Quiz excel files.
#     If going with this method, set "directory" equal to ''
# Best case scenario: directory path tweaking makes it work ; worst case scenario: you buy a PC or use the lab PCs
# Works only with Excel files (xlsx yes, xls maybe, csv no)

# imports, don't touch these if you want the script to work properly.
import pandas as pd
from pathlib import Path
import sys

# specifying which xlsx file to use
directory = 'C:\\Path\\to\\MQs\\'              # Directory path to module quiz file directory. Important to add the \\ at end.
file_name = input('Enter name of Quiz .xlsx file (example: Module 1 Quiz.xlsx): ')

# if the Excel extension is not already written, add it into the file name. 
if not '.xlsx' in file_name:
    file_name = file_name + '.xlsx'

# Will re-prompt until a valid file is entered, or exit if EXIT is typed.
while not Path(directory + file_name).is_file():
    print('\n' + file_name + ' does not exist')
    file_name = input('Enter name of Quiz .xlsx file (example: Module 1 Quiz.xlsx) or EXIT to quit: ')
    if 'EXIT' in file_name:
        sys.exit(0)
    elif not '.xlsx' in file_name:
        file_name = file_name + '.xlsx'
    
# prompt for what row "Quiz Questions" is on. (i.e.) if Quiz Questions is on row 20, enter 20
rowNum = input('What row is \'Quiz Questions\' located on: ')
while not rowNum.isdigit():
    print('\nIntegers only')
    rowNum = input('What row is \'Quiz Questions\' located on: ')

# =========================================================================================
# ==  If using Jupyter Notebook/ipynb or similar, put below code in separate code block  ==
# =========================================================================================

# leave this block alone.
rowNum = int(rowNum)
# df_obj = Learning objectives ; df_q = Quiz Questions
df_obj = pd.read_excel(directory + file_name, skiprows=1, header=None)
df_q   = pd.read_excel(directory + file_name, skiprows=rowNum)

# === Code below for assembling the feedback ===

# selecting first 2 columns and between rows 2 and whever the NaN starts-1
df_obj = df_obj.iloc[:, 0:2]
df_obj.columns = ['Symbol', 'Content']     # 'Symbol' = A1, A2, etc. ; 'Content' = the learning objective

# rowNum = min(df_obj[df_obj['Symbol'].isnull()].index)
df_obj = df_obj.iloc[0:rowNum, :] 
df_obj = df_obj.dropna(subset=['Content'])

# df_qq = Quiz Questions, matches to Learning Objectives
df_qq = df_q[['#', 'LO', 'Ans. Loc.']]
df_qq.columns = ['Question Number', 'LO', 'Answer Location']

# prompt for questions missed, removes empty strings (''), store as an integer list
missed = list(filter(None, input('Enter questions missed: ').split(',')))
missed = [int(i) for i in missed]

# instruction message to students
print('\nThe questions missed pertain to the following learning objectives. \n' +
      'In parentheses after each learning objective is specific course material that would \nbe good to study ' +
      'to better understand the respective learning objective, pertaining to questions missed.\n') # new line

# iterates through all of the questions entered in 'missed' for assembling feedback
for i in range(len(missed)):
    found = df_qq[df_qq['Question Number'] == missed[i]]
    quiz_lo = str(found['LO'].tolist()).replace('\'', '').replace('[', '').replace(']', '') # ['A1'] --> A1
    ans_loc = str(found['Answer Location'].tolist()).replace('\'', '').replace('[', '').replace(']', '')
    

#   found       : search value in df_qq
#   quiz_lo     : lists found(df_qq) for LO column as String
#   ans_loc     : lists found(df_qq) for Answer Location column as String
#                   (location of each answer section)
#   learn_obj   : search Symbol column for quiz_lo
#   c_obj       : lists learn_obj(df_obj) for Content column as string
#                   (content title)


    # in case of multiple answer locations, print them all
    ans_loc = ans_loc.split(' // ')
    ans = ''
    for x in ans_loc:
        ans += ' -- (' + str(x) + ')  \n'
    
    # these two variables used for creating the formatted output, leave them alone
    learn_obj = ''
    c_obj = ''
    
    # in case of multiple LO's, print them all
    if ',' in quiz_lo:
        multi_lo = quiz_lo.split(',')
        for x in multi_lo:
            x = x.strip()
            learn_obj = df_obj[df_obj['Symbol'] == x]
            c_obj += str(learn_obj['Content'].tolist()).replace('\'', '').replace('[', '').replace(']', '') + '\n' # ['A1'] --> A1
        c_obj = c_obj.strip()
    else:
        learn_obj = df_obj[df_obj['Symbol'] == quiz_lo]
        c_obj = str(learn_obj['Content'].tolist()).replace('\'', '').replace('[', '').replace(']', '') # ['A1'] --> A1
    print(str(c_obj) + '\n' + ans)

# ==  FORMAT OF OUTPUT  ==
#     
#     c_obj [learning objective]
#      -- ( ans [contents, Ans. Loc])
#     
#     c_obj (learning objective)
#      -- ( ans (contents, Ans. Loc))
#     
