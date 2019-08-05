import pandas as pd
from pathlib import Path

# specifying which xlsx file to use
directory = 'Put directory to where the .xlsx is located'              # Should probably add tkinter to this..
file_name = input('Enter name of Quiz .xlsx file (example: Module 1 Quiz.xlsx): ')

if not '.xlsx' in file_name:
    file_name = file_name + '.xlsx'

# Will re-prompt until a valid file is entered
while not Path(directory + file_name).is_file():
    print('\n' + file_name + ' does not exist')
    file_name = input('Enter name of Quiz .xlsx file (example: Module 1 Quiz.xlsx) or EXIT to quit: ')
    if not '.xlsx' in file_name:
        file_name = file_name + '.xlsx'
    if 'EXIT' in file_name:
        break

rowNum = input('What row is \'Quiz Questions\' located on: ')
while not rowNum.isdigit():
    print('\nIntegers only')
    rowNum = input('What row is \'Quiz Questions\' located on: ')

rowNum = int(rowNum)
# df_obj = Learning objectives | df_q = Quiz Questions
df_obj = pd.read_excel(directory + file_name, skiprows=1, header=None)
df_q   = pd.read_excel(directory + file_name, skiprows=rowNum)

# --- Code below for assembling the feedback ---

# selecting first 2 columns and between rows 2 and whever the NaN starts-1
df_obj = df_obj.iloc[:, 0:2]
df_obj.columns = ['Symbol', 'Content']
# rowNum = min(df_obj[df_obj['Symbol'].isnull()].index)
df_obj = df_obj.iloc[0:rowNum, :] 
df_obj = df_obj.dropna(subset=['Content'])

# df_qq = Quiz Questions, matches to learning Objectives
df_qq = df_q[['#', 'LO', 'Ans. Loc.']]
df_qq.columns = ['Question Number', 'LO', 'Answer Location']

# prompt for questions missed, stored as an integer list
missed = input('Enter questions missed: ').split(',')
missed = [int(i) for i in missed]

# instruction message to students
print('\nThe questions missed pertain to the following learning objectives. \n' +
      'In parentheses after each learning objective is specific course material that would \nbe good to study ' +
      'to better understand the respective learning objective, pertaining to questions missed.\n') # new line

'''
    found       : search value in df_qq
    quiz_lo     : lists found(df_qq) for LO column as String
    ans_loc     : lists found(df_qq) for Answer Location column as String
                    (location of each answer section)
    learn_obj   : search Symbol column for quiz_lo
    c_obj       : lists learn_obj(df_obj) for Content column as string
                    (content title)
'''

for i in range(len(missed)):
    found = df_qq[df_qq['Question Number'] == missed[i]]
    quiz_lo = str(found['LO'].tolist()).replace('\'', '').replace('[', '').replace(']', '') # ['A1'] --> A1
    ans_loc = str(found['Answer Location'].tolist()).replace('\'', '').replace('[', '').replace(']', '')
    
    # in case of multiple answer locations, print them all
    ans_loc = ans_loc.split(' // ')
    ans = ''
    for x in range(len(ans_loc)):
        ans += ' -- (' + str(ans_loc[x]) + ')  \n'

    learn_obj = df_obj[df_obj['Symbol'] == quiz_lo]
    c_obj = str(learn_obj['Content'].tolist()).replace('\'', '').replace('[', '').replace(']', '') # ['A1'] --> A1
    print(str(c_obj) + '\n' + ans)

# FORMAT
#     c_obj [learning objective]
#      -- ( ans [contents, Ans. Loc])
    
#     c_obj (learning objective)
#      -- ( ans (contents, Ans. Loc))
