# SudokuSolver

This is a very simple algorithm for a sudokusolver which I made when I was in 11th grade when I got introduced to python in school. This algorithm is a baisc bruteforce method to get to one solution of any given valid sudoku problem.

<h2>How to use this program</h2>
First download both the data.xlsx and main.py file and put those in the same foler.

Now open the data.xlsx file and fill in your sudoku problem in the problem section and save the file and close it.

Now run the main.py file and enter yes for the asked input. Now, the solution of your sudoku problem will be filled in the solution section of the data.xlsx file.

<h2>How the code works?</h2>
This is a basic bruteforce method. Definitely not even close to efficiency and inconsideration of edge cases such as invalid sudoku problem (where the code might run into infinite loop or just throw up some error). But it works if the input is correct!

So it takes the input from excel sheet by using the openpyxl library first. All the data extracted is stored in 27 different variables. Now it loops through each Empty cells.

To fill an empty cell, we check the validity of each number 1 - 9 by first checking its agreement with column and row wise filling and then its block wise filling(i.e in which whole grid is divided into 9 blocks). If a number is matched, then it is filled, if any number is not matching, then it goes back to the previously filled box and tries the left out numbers which were not tried in that box. Then goes back to the current block after a number matches with the previous block. But if the previous block was also not agreeing with any number, then it goes to the previous block before that one too. This is sort of a recursion method. And this is the only rule that is followed throughout the code.

And yes this method provides only one solution to the puzzle where there might be multiple solutions and we don't consider that. But to a 11th grader this is already too much of an achievement :D
