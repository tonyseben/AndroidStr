import os
import os.path
import sys
import math
import time
from openpyxl import load_workbook


BOOK_FILE = "/Users/tonyseben/Projects/PythonProjects/NodleCash_Strings_20200603-Final.xlsx"
COL_SLNO = 1
COL_KEY = 2
COL_DEFAULT = 3
COL_START = 4
COL_MAX = -1

ROW_HEADING = 1
ROW_LANG_CODE = 2
# Empty row in row 3
ROW_START = 4
ROW_MAX = -1

BOOK = None
SHEET = None

def isValidFile(path):
    print("Validating file ...")
    isFile = os.path.isfile(path)
    if not isFile:
        print("Error! Path invalid : " + path)
    return isFile

def verifySheetFormat(path):
    print("Verifiying sheet format ...")
    global BOOK
    global SHEET
    global ROW_MAX
    global COL_MAX
    sheetName = "Translations"

    book = load_workbook(path, read_only=True)
    sheet = book[sheetName]
    if sheet is None:
        print("Error! Sheet does not exist : " + sheetName)
        return False
    else:
        BOOK = book
        SHEET = sheet
        ROW_MAX = sheet.max_row
        COL_MAX = sheet.max_column
        return True


def generateString(langColumn):
    global SHEET
    global ROW_START
    global ROW_MAX
    global COL_KEY

    lang = SHEET.cell(ROW_HEADING, langColumn).value
    langCode = SHEET.cell(ROW_LANG_CODE, langColumn).value
    langDisplay = "%s (%s)" %(lang, langCode)
    #print("Generating strings for language %s(%s) ..." % (lang, langCode))
    
    # Create directory for language.
    dirPath = "./res/values-" + langCode
    os.makedirs(dirPath, exist_ok=True)

    strFile = open(dirPath + "/strings.xml", 'w')

    strFile.write("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n")
    for row in range(ROW_START, ROW_MAX+1):
        key = SHEET.cell(row, COL_KEY).value
        value = SHEET.cell(row, langColumn).value.rstrip("\n")

        if(key is not None):
            strFile.write("\n<string name=\"%s\">%s</string>" %(key, value))
        showProgress(langDisplay , row)

    sys.stdout.write("\n")
    strFile.write("\n\n</resources>")
    strFile.close()


def showProgress(langDisplay, currentRow):
    global ROW_START
    global ROW_MAX

    total = ROW_MAX - ROW_START
    row = currentRow - ROW_START
    progress = math.ceil(row/total*100)

    sys.stdout.write('\r* {0} {1}%'.format(langDisplay, progress))
    sys.stdout.flush()

def convertMillis(milli):
    seconds=(milli/1000)%60
    minutes=(milli/(1000*60))%60
    return "%d minutes %d seconds" % (minutes, seconds)

#**************************************************************#

startTime = int(round(time.time() * 1000))

if isValidFile(BOOK_FILE) and verifySheetFormat(BOOK_FILE):
    print("Sheet to Strings: START")
    for col in range(COL_START, COL_MAX):
        if SHEET.cell(ROW_LANG_CODE, col).value is not None:
            generateString(col)
    print("Sheet to Strings: COMPLETE")

finishTime = int(round(time.time() * 1000))
time = finishTime - startTime
print("Completed in %s\n" % convertMillis(time))


