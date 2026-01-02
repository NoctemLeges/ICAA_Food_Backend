from dotenv import load_dotenv
load_dotenv()
import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from flask import Flask, request

#--------------GLOBALS--------------------------------------------

today = "Day 1 Lunch"

app = Flask(__name__)

SERVICE_ACCOUNT_FILE = "icaa-qr-generation-dc58b73870cd.json"

_service = None

SHEET_IDs = {
    "Day 1 Lunch"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 2 Lunch"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 3 Lunch"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 3 Dinner" : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4"
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ID_to_row_map = {'ICAA26_ORG_I001': 2, 'ICAA26_ORG_I002': 3, 'ICAA26_ORG_I003': 4, 'ICAA26_ORG_I004': 5, 'ICAA26_ORG_I005': 6, 'ICAA26_ORG_I006': 7, 'ICAA26_ORG_I007': 8, 'ICAA26_ORG_I008': 9, 'ICAA26_ORG_I009': 10, 'ICAA26_ORG_I010': 11, 'ICAA26_ORG_I011': 12, 'ICAA26_ORG_I012': 13, 'ICAA26_ORG_I013': 14, 'ICAA26_ORG_I014': 15, 'ICAA26_ORG_I015': 16, 'ICAA26_ORG_I016': 17, 'ICAA26_ORG_I017': 18, 'ICAA26_ORG_I018': 19, 'ICAA26_ORG_I019': 20, 'ICAA26_ORG_I020': 21, 'ICAA26_ORG_I021': 22, 'ICAA26_ORG_I022': 23, 'ICAA26_ORG_I023': 24, 'ICAA26_ORG_I024': 25, 'ICAA26_ORG_I025': 26, 'ICAA26_ORG_I026': 27, 'ICAA26_ORG_I027': 28, 'ICAA26_ORG_I028': 29, 'ICAA26_ORG_I029': 30, 'ICAA26_ORG_I030': 31, 'ICAA26_ORG_I031': 32, 'ICAA26_ORG_I032': 33, 'ICAA26_ORG_I033': 34, 'ICAA26_ORG_I034': 35, 'ICAA26_ORG_I035': 36, 'ICAA26_SSC_I001': 37, 'ICAA26_SSC_I002': 38, 'ICAA26_SSC_I003': 39, 'ICAA26_SSC_I004': 40, 'ICAA26_SSC_I005': 41, 'ICAA26_SSC_I006': 42, 'ICAA26_SSC_I007': 43, 'ICAA26_SSC_I008': 44, 'ICAA26_SSC_I009': 45, 'ICAA26_SSC_I010': 46, 'ICAA26_CMC_I001': 47, 'ICAA26_CMC_I002': 48, 'ICAA26_CMC_I003': 49, 'ICAA26_CMC_I004': 50, 'ICAA26_CMC_I005': 51, 'ICAA26_CMC_I006': 52, 'ICAA26_CMC_I007': 53, 'ICAA26_CMC_I008': 54, 'ICAA26_CMC_I009': 55, 'ICAA26_CMC_I010': 56, 'ICAA26_CMC_I011': 57, 'ICAA26_CMC_I012': 58, 'ICAA26_CMC_I013': 59, 'ICAA26_CMC_I014': 60, 'ICAA26_CMC_I015': 61, 'ICAA26_GPC_I001': 62, 'ICAA26_GPC_I002': 63, 'ICAA26_GPC_I003': 64, 'ICAA26_GPC_I004': 65, 'ICAA26_GPC_I005': 66, 'ICAA26_GPC_I006': 67, 'ICAA26_GPC_I007': 68, 'ICAA26_GPC_I008': 69, 'ICAA26_GPC_I009': 70, 'ICAA26_GPC_I010': 71, 'ICAA26_GPC_I011': 72, 'ICAA26_GPC_I012': 73, 'ICAA26_GPC_I013': 74, 'ICAA26_GPC_I014': 75, 'ICAA26_GPC_I015': 76, 'ICAA26_GPC_I016': 77, 'ICAA26_GPC_I017': 78, 'ICAA26_GPC_I018': 79, 'ICAA26_GPC_I019': 80, 'ICAA26_GPC_I020': 81, 'ICAA26_VLT_I001': 82, 'ICAA26_VLT_I002': 83, 'ICAA26_VLT_I003': 84, 'ICAA26_VLT_I004': 85, 'ICAA26_VLT_I005': 86, 'ICAA26_VLT_I006': 87, 'ICAA26_VLT_I007': 88, 'ICAA26_VLT_I008': 89, 'ICAA26_VLT_I009': 90, 'ICAA26_VLT_I010': 91, 'ICAA26_VLT_I011': 92, 'ICAA26_VLT_I012': 93, 'ICAA26_VLT_I013': 94, 'ICAA26_VLT_I014': 95, 'ICAA26_VLT_I015': 96, 'ICAA26_VLT_I016': 97, 'ICAA26_VLT_I017': 98, 'ICAA26_VLT_I018': 99, 'ICAA26_VLT_I019': 100, 'ICAA26_VLT_I020': 101, 'ICAA26_VLT_I021': 102, 'ICAA26_VLT_I022': 103, 'ICAA26_VLT_I023': 104, 'ICAA26_VLT_I024': 105, 'ICAA26_VLT_I025': 106, 'ICAA26_AUT_I001': 107, 'ICAA26_AUT_I002': 108, 'ICAA26_AUT_I003': 109, 'ICAA26_AUT_I004': 110, 'ICAA26_AUT_I005': 111, 'ICAA26_AUT_I006': 112, 'ICAA26_AUT_I007': 113, 'ICAA26_AUT_I008': 114, 'ICAA26_AUT_I009': 115, 'ICAA26_AUT_I010': 116, 'ICAA26_AUT_I011': 117, 'ICAA26_AUT_I012': 118, 'ICAA26_AUT_I013': 119, 'ICAA26_AUT_I014': 120, 'ICAA26_AUT_I015': 121, 'ICAA26_AUT_I016': 122, 'ICAA26_AUT_I017': 123, 'ICAA26_AUT_I018': 124, 'ICAA26_AUT_I019': 125, 'ICAA26_AUT_I020': 126, 'ICAA26_AUT_I021': 127, 'ICAA26_AUT_I022': 128, 'ICAA26_AUT_I023': 129, 'ICAA26_AUT_I024': 130, 'ICAA26_AUT_I025': 131, 'ICAA26_AUT_I026': 132, 'ICAA26_AUT_I027': 133, 'ICAA26_AUT_I028': 134, 'ICAA26_AUT_I029': 135, 'ICAA26_AUT_I030': 136, 'ICAA26_AUT_I031': 137, 'ICAA26_AUT_I032': 138, 'ICAA26_AUT_I033': 139, 'ICAA26_AUT_I034': 140, 'ICAA26_AUT_I035': 141, 'ICAA26_AUT_I036': 142, 'ICAA26_AUT_I037': 143, 'ICAA26_AUT_I038': 144, 'ICAA26_AUT_I039': 145, 'ICAA26_AUT_I040': 146, 'ICAA26_CPC_I001': 147, 'ICAA26_CPC_I002': 148, 'ICAA26_CPC_I003': 149, 'ICAA26_CPC_I004': 150, 'ICAA26_CPC_I005': 151, 'ICAA26_CPC_I006': 152, 'ICAA26_CPC_I007': 153, 'ICAA26_CPC_I008': 154, 'ICAA26_CPC_I009': 155, 'ICAA26_CPC_I010': 156, 'ICAA26_CPC_I011': 157, 'ICAA26_CPC_I012': 158, 'ICAA26_CPC_I013': 159, 'ICAA26_CPC_I014': 160, 'ICAA26_CPC_I015': 161, 'ICAA26_CPC_I016': 162, 'ICAA26_CPC_I017': 163, 'ICAA26_CPC_I018': 164, 'ICAA26_CPC_I019': 165, 'ICAA26_CPC_I020': 166}

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
        if status == []:
            updateFoodForID(id)
            return f"""
            <html>
                <body>
                    <h1>Updated record for {id}</h1>
                </body>
            </html>
            """
        elif status[0][0] == 'Yes':
            return f"""
            <html>
                <body>
                    <h1>Food already taken for {id}</h1>
                </body>
            </html>
            """
    except HttpError as error:
        print(error)
        updateFoodForID(id)
        return f"Some error occured: {error}"

@app.route("/ping")
def pong():
    return "pong"
#-----------------------------------------------------------------------------

#---------------------STARTUP-------------------------------------------------
if __name__ == "__main__":
    prepareSheets()
    app.run()
#-----------------------------------------------------------------------------