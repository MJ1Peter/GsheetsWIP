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
    print("using sheetname {sheetname}")
    worksheet =  None 
    
    try:
        gc = gspread.oauth()
        sh = gc.open(sheetname)
        worksheet = sh.worksheet(wsname)
        print(f"returning {sheetname}, {WSNAME}")

    except gspread.SpreadsheetNotFound:
        print(f"spreadsheet '{sheetname}' not found")
    except gspread.WorksheetNotFound:
        print(f"worksheet '{WSNAME} not found")        
    except Exception as e:
        print(f"oauth failed: {e}")

    return worksheet

def _get_worksheet(worksheets,  wsname):
    for s in worksheets:
        if s.title == wsname:
            return s
    return None

# def _get_worksheet_id(worksheets,  wsname):
#     for s in worksheets:
#         if s.title == wsname:
#             return s.id
#     return None

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
        print(f"found sheet {sheetname}, cant create")
        return spreadsheet

    except gspread.SpreadsheetNotFound:
        pass

    try:
        # requesting that the client create a spreasheet for us.
        spreadsheet = gspreadclient.create("A new spreadsheet")
        return spreadsheet
 
    except Exception as e:
        print(f"error creating sheet {e}")
        return None

def  insertsheet (sheetname, wsname):
    print(f"using sheetname {WSNAME}")
    
    try:
        gspreadclient = gspread.oauth()
        spreadsheet = gspreadclient.open(sheetname)
        print(f"Found sheet {sheetname}. Adding {WSNAME} to spreadsheet")
        return spreadsheet
    
    except gspreadclient.SpreadsheetnNotFound:
        print(f"Cannot locate sheet {sheetname}")
        return None
    
    except Exception as e:
        print(f"oauth failed:{e}")
        return None
    
    try:
        with open(CSVFILE, "r") as f:
            csvdata = list(csv.reader(f, delimiter=","))

        spreadsheet.add_worksheet(title=wsname, rows=1000, cols=50)
        spreadsheet.values_update(
            wsname,
            params={'valueInputOption': 'USER_ENTERED'},
            body={'values': csvdata}
        )
        print(f"finished loading data into {spreadsheet}")
    
    except Exception as e:
        print(f"error inserting data into sheet {WSNAME}")

    return spreadsheet

def  showsheet(worksheet):
    if worksheet is not None:
        lists = worksheet.get(return_type=gspread.utils.GridRangeType.ListOfLists)
        if lists is not None:
            print(f"sheet columns:{lists[0]}")
            for ndx in range(1,6):
                print(f"batsman {player[0]}, {player[1:4]}")
    return

def  deletesheet(sheetname,  wsname):
    print(f"using sheetname {sheetname}.")

    gspreadclient = gspread.oauth()

    try:
        spreadsheet = gspreadclient.open(sheetname)
        worksheets = spreadsheet.worksheets()
        ws = _get_worksheet(worksheets, wsname)
        if ws is not None:
            spreadsheet.del_worksheet(ws)
            print(f"{sheetname} not found can't delete it...")
            
    except gspread.SpreadsheetNotFound:
        print(f"sheetname not found, can't delete...")
    except gspread.exceptions.APIError as e:
        print(f"delete failed: {e}")

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
