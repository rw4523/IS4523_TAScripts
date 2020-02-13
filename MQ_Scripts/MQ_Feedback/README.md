**For the how-to info, click [HERE](https://github.com/rw4523/IS4523_TAScripts/blob/master/MQ_Scripts/README.md#mq_feedbackpy).**

### NOTES
* The first time you copy this script over to wherever you're running it, be sure to change the `directory` variable (on line 13) to match where you're storing the `.xlsx` files for the Module Quizzes. If they are in another directory, set it to the folder that stores it. 
   * Example: if on Windows, it is stored in `C:\Users\user123\Desktop\4523\`, the `directory` variable should look like: `directory = 'C:\\Users\\user123\\Desktop\\4523\\'`.
   * Example: if on OSX, it is stored in `~/Desktop`, the `directory` variable should look like: `directory = '~/Desktop'`. Consult OSX's Finder window for the correct file paths.
   * Example: if this script is located in the same directory as the `.xlsx` files, then you can leave the `directory` variable set as `directory = ''`
* Please note that while this script can help speed up feedback-making time, it will **NOT** do all of the work for you.
* Please make sure to check the output to remove duplicates, and if 2+ questions have the same `LO` but different answer locations, try to put them under the same `LO`.
* This script only works if the questions are asked in the same order that the `Module Quiz.xlsx` files lists them. If they are randomized for every student, this script is kind of useless unless you can edit the code to accept randomized questions. If it is randomized and you aren't sure how to change the code, email me. (Email is in the `READ ME FIRST.pptx` file, slide 68-ish.)
