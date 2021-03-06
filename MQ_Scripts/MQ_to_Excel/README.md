**For the how-to information, click [HERE](https://github.com/rw4523/IS4523_TAScripts/blob/master/MQ_Scripts/README.md#mq_to_excelpy).**

### Note: `MQ_to_Excel.py` is not 100% error-proof. 
   * The human using this script should still check the output, then copy in the `LO`'s and information about the LO's at the top rows
   * This is not made to get things done in a few seconds. You (the TA) will still need to read the slides/textbook/etc. to make the quiz questions.
        * This script will only speed up the time it takes to convert the questions you made 
   in a `.txt` file (assuming using UTF-8, not ANSI encoding) to the `.xlsx` format that Blackboard uses for the MQ's. 
   
##### Reserved symbols/characters (please use these only in the format specified [*here*](https://github.com/rw4523/IS4523_TAScripts/blob/master/MQ_Scripts/README.md#mq_to_excelpy))
   * `*` = Correct answer(s). Answer choices without an Asterisk (`*`) will be marked as "`Incorrect`"
   * `{` and `}` = Source of the answer (i.e. from Textbook, Slides, Notes, etc.)
   * `[` and `]` = Learning Objective(s). Multiple Learning Objectives can be added with a comma "`,`" (i.e. `A1, A2, B1`)

##### Requirements
   * A computer that can run `Python 3` or later. 
   * Having the `ArgParse` python library. (It should come pre-installed, if not, see below ↓)
   
##### Troubleshooting
   * having argparse installed is a must (if not installed, `pip install argparse`)
   * Usage: if script located in same directory as input file: 
        * `python3 mq_to_excel.py -i input.txt -o output.xlsx`
   * Usage if script located in different directory than input file:
        * `python3 mq_to_excel.py -i "C:\Path\To\File\input.txt" -o "C:\Preferred\Path\To\File\output.txt"`
   * Usage if you don't have "python" or "python3" as a command, but script and txt file are in same directory as current working directory:
        * `"C:\Path\To\Python3.exe" mq_to_excel.py -i input input.txt -o output.xlsx"`
   * Usage if everything looks right but it is still giving errors:
        * Make sure that you're running the python.exe executable out of the file directory it is installed in
