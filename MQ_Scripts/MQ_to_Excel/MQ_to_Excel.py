import argparse
import os
import pandas as pd
import re
import time

start_time = time.time()

# argparse
parser = argparse.ArgumentParser(description='Converts module quiz txt file into an excel spreadsheet following the BlackBoard Module Quiz format.')
parser.add_argument('-i', '--input', metavar='', required=True, help='file path to the .txt input file')
parser.add_argument('-o', '--output', metavar='', required=True, help='.xlsx file name that the output should be saved as')
args = parser.parse_args()
file_path = args.input
fname = args.output

# file extensions check
if not '.txt' in args.input:     # xlsx
    file_path = args.input + '.txt'
else:
    file_path = args.input
if not '.xlsx' in args.output:   # txt
    fname = args.output + '.xlsx'
else:
    fname = args.output

# append to lst when adding another row of questions
lst = []
arr = ['#', 'LO', 'Ans. Loc.', 'Format', 'Question', 'a', ' ', 'b', ' ', 'c', ' ', 'd', ' ', 'e', ' ']
lst.append(arr)

# row-by-row assembly of questions
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
        LO = data[data.index('.')+3: re.match(re.compile(r'[A-Za-z]'))] ### ISSUE to be fixed someday..
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
        answer_list = data.split('\n')
        for x in answer_list:
            if ('a.\t' in x) or ('b.\t' in x) or ('b.\t' in x) or ('b.\t' in x) or ('b.\t' in x):
                x = x[5:]
    else:
        answer_list = data[data.index('}')+2:].split('\n')
        for x in answer_list:
            if ('a.\t' in x) or ('b.\t' in x) or ('b.\t' in x) or ('b.\t' in x) or ('b.\t' in x):
                x = x[5:]

    # appending data to cells
    tmp = []
    tmp.append(q_number.strip())
    tmp.append(LO.strip())
    tmp.append(ans_loc.strip())
    tmp.append(q_format.strip())
    tmp.append(question.strip())
    
    # code block to get rid of the asterisk (*) 
    for i in answer_list:
        if '*' in i:
            tmp.append(i.replace('\t*', '')[i.index('.')+1:].strip())
            tmp.append('Correct')
        else:
            tmp.append(i[i.index('.')+1:].strip())
            tmp.append('Incorrect')
    lst.append(tmp)
    
if __name__ == '__main__':
    with open(file_path, encoding='utf-8', mode='r') as file:
        data = file.read()                                     # file read into a string (data)
        data_list = data.split('\n\n')
        for i in data_list:
            question_to_row(i)
        lst = [[ls.replace('*', '') for ls in l] for l in lst] # last check to remove asterisks

        # convert list to dataframe and export to Excel
        df = pd.DataFrame(lst, index=None)
        df.to_excel(fname[0:-4] + 'xlsx', index=False, header=None)
        time_taken = time.time() - start_time
        print(('\nREAD:  ' + file_path + '\nDONE: \'' + fname[0:-4] + 'xlsx\' saved to ' + os.getcwd() 
               + '\nTime taken: {:.2f} seconds').format(time_taken))
