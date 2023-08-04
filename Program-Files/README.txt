This project was designed as an inspection efficiency aid, automating the repetitive parts of inspection data recording

This program was written for S&C Electric Canada. 
The program takes data (Sales order #'s, drawing #'s, lot size, etc) and appropriately formats it into existing S&C excel files.
It also tells the inspector how much of the lot to inspect based on S&C quality standards.
Some known info is automatically filled (date, name, qty inspected) to decrease time wasted typing.
Uses an editable .txt file for managment to place important inspection reminders which will display during use of the program.
This program also circumvents the annoying Excel feature which doesnt allow multiple editors at once.
When another user has the sheet open, a temp sheet is used to store data ad then is uplaoded laterwhen possible.

IMPORTANT - For this program to work properly the following must be true:
- Must have the inspection file (in the same directory as the .exe) by the name "Small Packaging Inspection (CURRENT YEAR).xlsx"
- NONE of the files in Program-Files may be renamed or deleted, but may be edited
- TMs who opt to enter data manually use standard DD/MM/YYYY date format
- There cannot be any empty rows in the middle of the sheet

Code Written by: Jonathan DiGiorgio
ver 1.0 - Aug 4, 2023


Contact author at jddigior@uwaterloo.ca