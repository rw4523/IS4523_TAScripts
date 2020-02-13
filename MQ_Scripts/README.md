# Module Quiz Scripts
Used for making/using `.xlsx` files in the format that BlackBoard's quiz machine accepts

Written in Python 3 (Does not work in 2.7), on a Windows 10 PC. 
Works with Mac/OSX if using Anaconda-Navigator/JupyterNotebook (Python 3) version. (I hope).
If you don't want to install new Anaconda/JupyterNotebook onto your computer, you can alternatively use [Google Colab](https://colab.research.google.com/). 
* Note: Use at your own risk because Google *might or might not* charge you. I just heard it's an option but I've never ran the `MQ_Scripts` with it.

### [MQ_Feedback.py](https://github.com/rw4523/IS4523_TAScripts/blob/master/MQ_Scripts/MQ_Feedback/MQ_Feedback.py):
   * Reads in the `.xlsx` file that stores the `LO`'s and `Ans.Loc.`'s of each question. Outputs it in the format below for questions missed 
     * Please note that this script **ONLY** works with `.xlsx` and maybe `.xls` files. `.csv` will likely not work.
   * for the `directory` value, please enter the file path to where the Module Quizz excel files are located.
     * If running on Windows, write the backslashes as doubles. 
       * Example, if the directory is `C:\Users\user123\Desktop\`, write it as `C:\\Users\\user123\\Desktop\\`
       * The ending backslash is important to note the directory.
       * If you're on Linux or OSX or another OS where directories use `/`, ignore this.
   * Will prompt for file name (i.e.: `quiz 1.xlsx`) , and which row that `Quiz Questions` starts on (i.e. if `Quiz Questions` is on row 20, enter `20`)
   * Will prompt for `Questions Missed`. If a student missed questions 1, 2, and 5, type in: `1,2,5`. 
     * **NOTE: This script will only work if the questions follow the same order. If the questions/order is randomized for each student, this script won't work properly.** 
   * Because multiple questions might have the same Learning Objective, make sure to copy the output into a Notepad first and combine the Answer Locations of the questions under the same Learning Objective.
 ```    
      This is where the LO looks like
      -- (Ans.Loc written to here)

      This is another LO
      -- (This is another Ans.Loc.)
      -- (More Ans.Loc. if there is more)
      
      This is the first LO
      This is the second LO
      -- (This is an Ans.Loc.)
```

---

### [MQ_to_Excel.py](https://github.com/rw4523/IS4523_TAScripts/blob/master/MQ_Scripts/MQ_to_Excel/MQ_to_Excel.py)
   * The actual questions should be done using numbered list in `.docx` and copy pasted into a .txt file. 
   * Save as a `.txt` file on the PC. Then space out the questions (1 empty row between questions).
   * Script will prompt for where the `.txt file` is, and will format it into the BlackBoard format as a `.xlsx` file
   * Questions should be done in this format:
```
      1.  [A1] True/False: This is a question. {Book pg. 2}
      a.  *True
      b.  False
      
      2.  [A9] This is MC question? {--}
      a.  *Yes
      b.  No
      c.  Also No but not the same answer
      d.  No again but different answer
      e.  Still not the right answer
      
      3.  [--] This is a MA question? {1.A Powerpoint Name slide 2}
      a.  *Correct Answer
      b.  *Another Correct Answer
      c.  *More Correct Answer
      d.  Incorrect Answer
      e.  Also Incorrect Answer
```
  * Notes:
     * `.txt` file **MUST** be saved using **UTF-8 encoding**. (ANSI will cause errors)
     * Looks like 2 spaces "`  `" on this `README.md` file, but that should be a tab (`\t`). 
        * If copy/pasting from `.docx`, it should do this automatically. 
        * You should still add 1 empty line between questions
     * Do not add apostrophes (`'`) in answer choices because it can break the script (Not yet sure why).
     * Backslashes can sometimes cause errors. (Need to figure out how to convert data stored in variables as raw string)
        * Example answer choice: `\tmp` gets outputted as `mp`
     * If the question does not have an `LO` or `Ans.Loc.`, use double dash (`--`)
     * After the `.xlsx` is created, you need to copy paste the `LO`'s into the `.xlsx` file, before the block that is the quiz questions.
        * Might still require a human to check it, as it is not 100% perfect but will reduce time taken to make `.xlsx` file.
     * May require installing `tkinter`. For installation details, consult your package manager (pip, conda, etc.)
