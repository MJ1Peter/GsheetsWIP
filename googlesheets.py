import sys
import gspread
import csv

# see https://github.com/burnash/gspread
# see https://docs.gspread.org/en/latest/oauth2.html

# Google Sheets has a spreadsheet, which can contain multiple worksheets
# The gspread API supports this hierarchy

BASEDIR   =  "/files/python/gsheets"
CSVFILE   =  BASEDIR + "commentary-data-sheet.csv"
WSNAME    =  "Inserted Data"


def opensheet (sheetname,  wsname):
    print("using sheetname {}".format(sheetname))
    worksheet =  None 
    
    try:
        gc = gspread.oauth()
        sh = gc.open(sheetname)
        worksheet = sh.worksheet(wsname)
        print("returning {}/{}".format(sheetname,  wsname))

    except gspread.SpreadsheetNotFound:
        print("spreadsheet '{}' not found".format(sheetname))
    except gspread.WorksheetNotFound:
        print("worksheet '{}' not found".format(wsname))        
    except Exception as e:
        print("oauth failed: {}".format(e))

    return worksheet

def _get_worksheet(worksheets,  wsname):
    for s in worksheets:
        if s.title == wsname:
            return s
    return None


def _get_worksheet_id(worksheets,  wsname):
    for s in worksheets:
        if s.title == wsname:
            return s.id
    return None

'''
accepts spreadsheet name (string), 
worksheet name (string) returns a gspread.models.Spreadsheet object if successful, 
otherwise none.
'''

def  createsheet (sheetname):
    try:
        # executing 3rd party authentication, oauth returns gspread.client if successful
        gspreadclient = gspread.oauth()
        spreadsheet = gspreadclient.open(sheetname)
        print("found sheet {}, cant create".format(sheetname))
        return spreadsheet

    except gspread.SpreadsheetNotFound:
        pass

    try:
        # requesting that the client create a spreasheet for us.
        spreadsheet = gspreadclient.create('A new spreadsheet')
        return spreadsheet
 
    except Exception as e:
        print(f'error creating sheet {e} ')
        return None

def  insertsheet (sheetname,  wsname):
    try:
        gspreadclient = gspread.oauth()
        spreadsheet = gspreadclient.insert('WSNAME')
    pass

def  showsheet(worksheet):
    # <fix me>
    pass

def  deletesheet(sheetname,  wsname):
    # <fix me>
    pass


# execute the code below only if we ran this script directly
if  __name__ ==  "__main__":
    print(len(sys.argv),  sys.argv)
    if  len(sys.argv) < 3:
        print("specify <sheet-name> <operation> where operation is 'create', 'update', 'insert', show' or 'delete'")
        sys.exit()
        
    print("argument list:", str(sys.argv))
    sheetname =  sys.argv[1]
    operation =  sys.argv[2]

    if  operation ==  "delete":
        deletesheet(sheetname,  WSNAME)

    elif operation ==  "update":
        sheet = updatesheet (sheetname, WSNAME)
        #update sheet with query user for row,comlumn cell value.

    elif operation == "show":
        sheet = opensheet(sheetname, WSNAME)
        showsheet(sheet)

    elif operation ==  "create":
        createsheet(sheetname)

    elif operation ==  "insert":
        insertsheet(sheetname,  WSNAME)

    else:
        print("unsupported operation: {}".format(operation))
        sys.exit(1)