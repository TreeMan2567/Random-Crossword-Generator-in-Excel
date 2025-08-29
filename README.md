# Random Crossword Generator in Excel
This system is created for revising just like with flashcards and as a proof that there is a lot that can be done in Excel and VBA.

It is possible to enter your own entries and clues as a base for generating crosswords.

## Unblocking Scripts 
All excel scripts not written by current user are blocked by default, so first you need to unblock scripts in excel spreadsheet. To do so right-click excel file    
Properties -> Security -> *Unblock*
Now you're ready to open the file.

## How to use
Click **Generate Random Crossword** to generate random crossword. To show clue for entry, select cell inside the crossword you want to see a clue for, and click **Show Clue** button. Write letters in cells and when you are ready click **Check puzzle** to validate your work. If it's not correct, system will show you what letters are wrong, and you will lose your one of three lives. Loose all lives to lose the game. At last click **Give up** to show the correct answers.

## Customize your experience
In ***Dictionary*** tab, you can freely change and add the entries and clues for them. Changes will take effect in next crossword generation.

You can also change how many lives you have as well as how many words do you wish for your crossword to have in ***Settings*** sheet at the bottom of your screen. All changes will be applied in next crossword generation.

## How does it works
At the start, system choose first crossword entry randomly from ***Dictionary***, then it remembers in ***Possible Extensions*** sheet every cell where is a place to insert another entry. Then it recursevely chooses one of those letters and matches one of the words from dictonary until words count in ***Variables*** reaches max words setted in ***Settings***. During the whole process answers to crossword are built in ***Solution*** sheet and the clues for each entry are saved inside ***HintsTable***. At the end whole crossword is drawn onto the ***Crossword*** Sheet, based on ***Visuals***.