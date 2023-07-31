"""
S&C Electric Canada
Developed by Jonathan DiGiorgio
July 2023 - August 2023
Python 3.11

Dependancies: 
- Openpyxl (version 3.1.2)
- Keyboard (version 0.13.5)
- Excel (Developed on Microsoft® Excel® for Microsoft 365 MSO Version 2305 Build 16.0.16501.20074 64-bit )

IT IS UNKNOWN IF THIS PROGRAM WILL FUNCTION ON UPDATED VERSIONS OF DEPENDANCIES
"""

# Modules and Libraries
import webbrowser
import time
import keyboard
from openpyxl import load_workbook
from datetime import date


# Variables
FILE_NAME = 'TestSP.xlsx' # ADD WAY TO CHANGE THIS BASED ON YEAR
LINK_TEMPLATE = 'https://sourceone.sandc.ws/apps/drawingsearch?query='

# Tries to load spreadsheet with exception handling for when another team member is using the file
def canLoadXl():
    try:
        wb = load_workbook(FILE_NAME, data_only = True)
    except PermissionError:
        print('Someone already has the document open. Try again later.')
        return False, 'Error'
    else:
        return True, wb
    
# Calls above function to attempt file load until sucessful or exited
def loadXl():
    isLoaded, wb = canLoadXl()
    while (not isLoaded):
        input('Press Enter to try again.')
        isLoaded, wb = canLoadXl()

    return wb

#Finds the next empty row in the catalog number column
def nextFreeRow():
    count = 0
    rowFound = False
    while not rowFound:
        count += 1
        if ws.cell(row = count, column = 3).value == None: # Catalog num are column 3
            rowFound = True
    return count

#checks if a date was already listed in the spreadsheet
def dateIsListed(count):
    found = False
    while (count > 0):
        #Checks the two most common date formats from the sheet, (00:00:00 is added by excel but not shown in the cell)
        if str(ws.cell(row = count, column = 1).value) in {date.today().strftime("%d/%m/%Y"), date.today().strftime("%Y-%m-%d 00:00:00")}:
            found = True
        count -= 1
    return found

# Uses sample size chart standard for the following numbers
def numToInspect(qty):
    qty = int(qty)

    if qty <= 2:
        return qty
    elif qty <= 25:
        return 2
    elif qty <= 50:
        return 3
    elif qty <= 90:
        return 5
    elif qty <= 150:
        return 8
    elif qty <= 280:
        return 13
    elif qty <= 500:
        return 20
    elif qty <= 1200:
        return 32
    elif qty <= 3200:
        return 50
    elif qty <= 10000:
        return 80
    elif qty <= 35000:
        return 125  #Standard ends here
    else:
        return int((125/35000)*qty) # Linearization from previous 

# Opens drawing in browser
def openDrawing(catalogNum):
    link = LINK_TEMPLATE + catalogNum
    webbrowser.open(link)

# Initializes sheet based on month
def initSheet(wb):
    sheets = wb.sheetnames
    month = date.today().month # Saves month as a number 1-12
    ws = wb[sheets[month - 1]]
    return ws

## ------------------ MAIN ------------------ ## 

# Check to see if xlsx is open alreadyand determine next free row
wb = loadXl()
ws = initSheet(wb)
writeRow = nextFreeRow()
wb.save(FILE_NAME)


user = input('Please enter your name (LASTNAME, FIRSTNAME): ')

anotherInspection = True
while anotherInspection: #Continues unti user is done inspecting

    # SALES ORDER INPUT
    SO = input('Enter Sales Order (Check Oracle): ')
    while (len(SO) != 6) or (not SO.isdigit()): #Sales orders must be 6 digit numbers
        SO = input('Error. Please enter a 6 digit Sales Order: ')

    # CATALOG NUMBER INPUT
    catNum = input('Enter Catalog Number: ')

    # LOT QUANTITY INPUT
    lotQty = input('Enter lot quantity: ')
    while not lotQty.isdigit():
        lotQty = input('Error! Please enter a number quantity : ')

    # SAMPLE SIZE INSPECTION STANDARD
    inspQty = numToInspect(lotQty)
    print("Please inspect {} parts".format(inspQty))
    time.sleep(1.5)

     # OPEN RESPECTIVE DRAWING
    openDrawing(catNum)

    # INSPECTION RESULTS INPUT
    result = input('Pass or Fail? : ')
    while result not in {'Pass','pass','Fail','fail'}:
        result = input('Invalid input, try again: ')

    # NOTE INPUT
    note = input('Enter a note (Press enter to skip): ')

    # Open sheet again to quickly add data
    wb = loadXl()
    ws = initSheet(wb)

    # Add date if needed
    if not dateIsListed(writeRow - 1):
        ws['A' + str(writeRow)] = date.today().strftime("%d/%m/%Y")

    # Add everything else
    ws['B' + str(writeRow)] = SO
    ws['C' + str(writeRow)] = catNum
    ws['D' + str(writeRow)] = inspQty
    ws['E' + str(writeRow)] = lotQty
    ws['F' + str(writeRow)] = result
    ws['G' + str(writeRow)] = user
    if note not in {None, ' '}:
        ws['H' + str(writeRow)] = note


        

    wb.save('TestSP.xlsx')

    # Go to next row if reiterated
    writeRow += 1

    break


input('PRESS ENTER TO CLOSE...')
