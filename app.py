from dotenv import load_dotenv
load_dotenv()
import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from flask import Flask, request

#--------------GLOBALS--------------------------------------------

today = "Day 2 Lunch"

app = Flask(__name__)

SERVICE_ACCOUNT_FILE = "icaa-qr-generation-dc58b73870cd.json"

_service = None

SHEET_IDs = {
    "Day 1 Lunch" : "1kgKZe_6ZIBxDgZQVaBdcu_MIXOM10ezRLlBX2jT4Y9s",
    "Day 2 Lunch" : "1kgKZe_6ZIBxDgZQVaBdcu_MIXOM10ezRLlBX2jT4Y9s",
    "Day 3 Lunch" : "1kgKZe_6ZIBxDgZQVaBdcu_MIXOM10ezRLlBX2jT4Y9s",
    "Day 3 Dinner" : "1kgKZe_6ZIBxDgZQVaBdcu_MIXOM10ezRLlBX2jT4Y9s"
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ID_to_row_map = {}

#-----------------------------------------------------------------------

#---------------------------FUNCTIONS----------------------------------------

def get_sheets_service():
    global _service

    if _service is not None:
        return _service

    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]),
        scopes=SCOPES
    )

    _service = build("sheets", "v4", credentials=creds)
    return _service

def writeValuestoRange(sheetID,_values,_range):
    try:
        service = get_sheets_service()

        # IMPORTANT FIX: use the passed values and make them 2D
        body = {
            "values": _values# wrap inside another list
        }

        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=sheetID,
                range=_range,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        return result
    except HttpError as error:
        raise error

def readValueFromCell(sheetID: str, cell: str):
  try:
    service = get_sheets_service()

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=sheetID, range=cell)
        .execute()
    )
    values = result.get("values", [])
    return values
  except HttpError as err:
    print(err)


def add_headers(sheetIdKey:str):
    try:
        writeValuestoRange(SHEET_IDs[sheetIdKey], [["ID","Name","Food Taken"]], f"'{sheetIdKey}'!A1:C")
        print(f"Added headers to {sheetIdKey}")
    except HttpError as error:
        print(f"An error occurred: {error}")

def add_IDs(sheetIdKey:str):
    IDs = []
    for i in range(1,36):
        IDs.append([f"ICAA26_ORG_I{str(i).zfill(3)}"])
    for i in range(1,11):
        IDs.append([f"ICAA26_SSC_I{str(i).zfill(3)}"])
    for i in range(1,16):
        IDs.append([f"ICAA26_CMC_I{str(i).zfill(3)}"])
    for i in range(1,21):
        IDs.append([f"ICAA26_GPC_I{str(i).zfill(3)}"])
    for i in range(1,26):
        IDs.append([f"ICAA26_VLT_I{str(i).zfill(3)}"])
    for i in range(1,41):
        IDs.append([f"ICAA26_AUT_I{str(i).zfill(3)}"])
    for i in range(1,21):
        IDs.append([f"ICAA26_CPC_I{str(i).zfill(3)}"])
    total_IDs = len(IDs)
    try:
        writeValuestoRange(SHEET_IDs[sheetIdKey],IDs,f"{sheetIdKey}!A2:A{total_IDs+1}")
        print(f"Added IDs to {sheetIdKey}")
        for ID in IDs:
            ID_to_row_map[ID[0]] = 2 + IDs.index(ID)
    except HttpError as error:
        print(f"An error occurred: {error}")

def prepareSheets():
    for key in SHEET_IDs.keys():
        add_headers(key)
        add_IDs(key)


def updateFoodForID(ID:str):
    try:
        writeValuestoRange(SHEET_IDs[today],[["Yes"]],f"{today}!C{ID_to_row_map[ID]}")
    except HttpError as error:
        raise error

def checkIfFoodTaken(ID: str):
    value = readValueFromCell(SHEET_IDs[today],f"{today}!C{ID_to_row_map[ID]}")
    return value    
#-----------------------------------------------------------------------------

#-----------------------ROUTES------------------------------------------------
@app.route("/food")
def food():
    id = request.args.get("id")
    try:
        status = checkIfFoodTaken(id)
        #print(status[0][0]=='Yes', type(status))
        #updateFoodForID(id)
        #return status[0][0] == 'Yes'
        if status == []:
            updateFoodForID(id)
            return f"Updated record for {id}"
        elif status[0][0] == 'Yes':
            return f"Food already taken for {id}"
    except HttpError as error:
        print(error)
        updateFoodForID(id)
        return f"Some error occured: {error}"
#-----------------------------------------------------------------------------

#---------------------STARTUP-------------------------------------------------
if __name__ == "__main__":
    prepareSheets()
    app.run()
#-----------------------------------------------------------------------------