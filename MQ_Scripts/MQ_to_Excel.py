import os
import pandas as pd
import re
from tkinter import *
from tkinter.filedialog import askopenfilename

f_path = Tk()
f_path.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("text files","*.txt"),("all files","*.*")))
file_path = str(f_path.filename)
fname = file_path[file_path.rfind('/')+1:]

lst = []      # append to lst when adding another row of questions
# lst.append(['Quiz Questions', ' '])
arr = ['#', 'LO', 'Ans. Loc.', 'Format', 'Question', 'a', ' ', 'b', ' ', 'c', ' ', 'd', ' ', 'e', ' ']
lst.append(arr)
    
def question_to_row(data):
    count = 0
    choices = ['a.', 'b.', 'c.', 'd.', 'e.']
    for i in choices:
        if i in data:
            count+=1
    # if 2 answer choices and only 1 * , it is True/False question
    if count == 2 and len(re.findall('\*', data)) == 1:
        q_format = 'TF'
    # if 2-5 choices and has multiple *, it is Multiple Answer
    elif (count > 2 and count <= 5) and len(re.findall('\*', data)) > 1: 
        q_format = 'MA'
    # if 2-5 choices and only 1 *, it is Multiple Choice
    elif (count > 2 and count <= 5) and len(re.findall('\*', data)) == 1: 
        q_format = 'MC'
                                                            # q_format question type (MA, TF, MC, etc.)
    q_number = data[0:data.index('.')]                      # question number
    
    if not '[' in data:
        LO = data[data.index('.')+3: re.match(re.compile(r'[A-Za-z]'))] ### ISSUE
    LO = data[data.index('.')+3: data.index(']')]           # learning objective of the question
    
    if not ' {' in data:
        question = data[data.index('] '):]
    # Ans.Loc of the question. If empty, ans_loc = --
        ans_loc = '--'
    else:
        question = data[data.index('] ')+2:data.index(' {')]    # The Question as a separate String
        ans_loc = data[data.index('{')+1:data.index('}')]

    # if '{' or '}' does not exist
    if (not ' {' in data) or (not '}' in data):
        answer_list = str(data.split('\n')).replace('a.\t', '').replace('b.\t', '').replace('c.\t', '').replace('d.\t', '').replace('e.\t', '').replace('\\t', '').replace('[', '').replace(']', '').replace('\'', '').split(',')
    else:
        answer_list = str(data[data.index('}')+2:].split('\n')).replace('a.\t', '').replace('b.\t', '').replace('c.\t', '').replace('d.\t', '').replace('e.\t', '').replace('\\t', '').replace('[', '').replace(']', '').replace('\'', '').split(',')
    tmp = []
    tmp.append(q_number.strip())
    tmp.append(LO.strip())
    tmp.append(ans_loc.strip())
    tmp.append(q_format.strip())
    tmp.append(question.strip())
    
    for i in answer_list:
        if '*' in i:                  # if/else block to get rid of the asterisk (*) 
            #i = i.replace('\t*', '')
            tmp.append(i.replace('\t*', '')[3:].strip())
            tmp.append('Correct')
        else:
            tmp.append(i[3:].strip())
            tmp.append('Incorrect')
    lst.append(tmp)

with open(file_path, encoding='utf-8', mode='r') as file:
    data = file.read()                                     # file read into a string (data)
    # print(data +'\n\n --- ')
    data_list = data.split('\n\n')
    for i in data_list:
        question_to_row(i)
    lst = [[ls.replace('*', '') for ls in l] for l in lst] # last check to remove asterisks

    # convert list to dataframe and export to Excel
    df = pd.DataFrame(lst, index=None)
    #df.columns = arr
    quiz_num = input('Enter Module Quiz Number: ')
    df.to_excel(fname + '.xlsx', index=False, header=None)
print('\nREAD:  ' + fname + '\nDONE: \'Module ' + quiz_num + ' Quiz.xlsx\' saved to ' + os.getcwd())

# ISSUES to fix (eventually)
# Issue: For whatever reason (not yet found), if an answer choice includes a comma (,) or apostrophe (')
#   it gets split into two different cells, resulting in more columns per row than there should be.
#   Example: "a. blah blah, bleh bleh bleh" --> [blah blah] | [bleh bleh bleh] ([ ] = cell)
##   Answer found @ line 53, 55 ; replace .split(',') with another character

# Issue: if text contains a special character (i.e. \ ), it will be printed as '\\'
# Need to append as raw string. 
