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

ID_to_row_map_trial = {}

ID_to_row_map = {'ICAA26_ORG_I001': 2, 'ICAA26_ORG_I002': 3, 'ICAA26_ORG_I003': 4, 'ICAA26_ORG_I004': 5, 'ICAA26_ORG_I005': 6, 'ICAA26_ORG_I006': 7, 'ICAA26_ORG_I007': 8, 'ICAA26_ORG_I008': 9, 'ICAA26_ORG_I009': 10, 'ICAA26_ORG_I010': 11, 'ICAA26_ORG_I011': 12, 'ICAA26_ORG_I012': 13, 'ICAA26_ORG_I013': 14, 'ICAA26_ORG_I014': 15, 'ICAA26_ORG_I015': 16, 'ICAA26_ORG_I016': 17, 'ICAA26_ORG_I017': 18, 'ICAA26_ORG_I018': 19, 'ICAA26_ORG_I019': 20, 'ICAA26_ORG_I020': 21, 'ICAA26_ORG_I021': 22, 'ICAA26_ORG_I022': 23, 'ICAA26_ORG_I023': 24, 'ICAA26_ORG_I024': 25, 'ICAA26_ORG_I025': 26, 'ICAA26_ORG_I026': 27, 'ICAA26_ORG_I027': 28, 'ICAA26_ORG_I028': 29, 'ICAA26_ORG_I029': 30, 'ICAA26_ORG_I030': 31, 'ICAA26_ORG_I031': 32, 'ICAA26_ORG_I032': 33, 'ICAA26_ORG_I033': 34, 'ICAA26_ORG_I034': 35, 'ICAA26_ORG_I035': 36, 'ICAA26_SSC_I001': 37, 'ICAA26_SSC_I002': 38, 'ICAA26_SSC_I003': 39, 'ICAA26_SSC_I004': 40, 'ICAA26_SSC_I005': 41, 'ICAA26_SSC_I006': 42, 'ICAA26_SSC_I007': 43, 'ICAA26_SSC_I008': 44, 'ICAA26_SSC_I009': 45, 'ICAA26_SSC_I010': 46, 'ICAA26_SSC_I011': 47, 'ICAA26_SSC_I012': 48, 'ICAA26_SSC_I013': 49, 'ICAA26_SSC_I014': 50, 'ICAA26_CMC_I001': 51, 'ICAA26_CMC_I002': 52, 'ICAA26_CMC_I003': 53, 'ICAA26_CMC_I004': 54, 'ICAA26_CMC_I005': 55, 'ICAA26_CMC_I006': 56, 'ICAA26_CMC_I007': 57, 'ICAA26_CMC_I008': 58, 'ICAA26_CMC_I009': 59, 'ICAA26_CMC_I010': 60, 'ICAA26_CMC_I011': 61, 'ICAA26_CMC_I012': 62, 'ICAA26_CMC_I013': 63, 'ICAA26_CMC_I014': 64, 'ICAA26_CMC_I015': 65, 'ICAA26_GPC_I001': 66, 'ICAA26_GPC_I002': 67, 'ICAA26_GPC_I003': 68, 'ICAA26_GPC_I004': 69, 'ICAA26_GPC_I005': 70, 'ICAA26_GPC_I006': 71, 'ICAA26_GPC_I007': 72, 'ICAA26_GPC_I008': 73, 'ICAA26_GPC_I009': 74, 'ICAA26_GPC_I010': 75, 'ICAA26_GPC_I011': 76, 'ICAA26_GPC_I012': 77, 'ICAA26_GPC_I013': 78, 'ICAA26_GPC_I014': 79, 'ICAA26_GPC_I015': 80, 'ICAA26_GPC_I016': 81, 'ICAA26_GPC_I017': 82, 'ICAA26_GPC_I018': 83, 'ICAA26_GPC_I019': 84, 'ICAA26_GPC_I020': 85, 'ICAA26_VLT_I001': 86, 'ICAA26_VLT_I002': 87, 'ICAA26_VLT_I003': 88, 'ICAA26_VLT_I004': 89, 'ICAA26_VLT_I005': 90, 'ICAA26_VLT_I006': 91, 'ICAA26_VLT_I007': 92, 'ICAA26_VLT_I008': 93, 'ICAA26_VLT_I009': 94, 'ICAA26_VLT_I010': 95, 'ICAA26_VLT_I011': 96, 'ICAA26_VLT_I012': 97, 'ICAA26_VLT_I013': 98, 'ICAA26_VLT_I014': 99, 'ICAA26_VLT_I015': 100, 'ICAA26_VLT_I016': 101, 'ICAA26_VLT_I017': 102, 'ICAA26_VLT_I018': 103, 'ICAA26_VLT_I019': 104, 'ICAA26_VLT_I020': 105, 'ICAA26_VLT_I021': 106, 'ICAA26_VLT_I022': 107, 'ICAA26_VLT_I023': 108, 'ICAA26_VLT_I024': 109, 'ICAA26_VLT_I025': 110, 'ICAA26_AUT_I001': 111, 'ICAA26_AUT_I002': 112, 'ICAA26_AUT_I003': 113, 'ICAA26_AUT_I004': 114, 'ICAA26_AUT_I005': 115, 'ICAA26_AUT_I006': 116, 'ICAA26_AUT_I007': 117, 'ICAA26_AUT_I008': 118, 'ICAA26_AUT_I009': 119, 'ICAA26_AUT_I010': 120, 'ICAA26_AUT_I011': 121, 'ICAA26_AUT_I012': 122, 'ICAA26_AUT_I013': 123, 'ICAA26_AUT_I014': 124, 'ICAA26_AUT_I015': 125, 'ICAA26_AUT_I016': 126, 'ICAA26_AUT_I017': 127, 'ICAA26_AUT_I018': 128, 'ICAA26_AUT_I019': 129, 'ICAA26_AUT_I020': 130, 'ICAA26_AUT_I021': 131, 'ICAA26_AUT_I022': 132, 'ICAA26_AUT_I023': 133, 'ICAA26_AUT_I024': 134, 'ICAA26_AUT_I025': 135, 'ICAA26_AUT_I026': 136, 'ICAA26_AUT_I027': 137, 'ICAA26_AUT_I028': 138, 'ICAA26_AUT_I029': 139, 'ICAA26_AUT_I030': 140, 'ICAA26_AUT_I031': 141, 'ICAA26_AUT_I032': 142, 'ICAA26_AUT_I033': 143, 'ICAA26_AUT_I034': 144, 'ICAA26_AUT_I035': 145, 'ICAA26_AUT_I036': 146, 'ICAA26_AUT_I037': 147, 'ICAA26_AUT_I038': 148, 'ICAA26_AUT_I039': 149, 'ICAA26_AUT_I040': 150, 'ICAA26_CPC_I001': 151, 'ICAA26_CPC_I002': 152, 'ICAA26_CPC_I003': 153, 'ICAA26_CPC_I004': 154, 'ICAA26_CPC_I005': 155, 'ICAA26_CPC_I006': 156, 'ICAA26_CPC_I007': 157, 'ICAA26_CPC_I008': 158, 'ICAA26_CPC_I009': 159, 'ICAA26_CPC_I010': 160, 'ICAA26_CPC_I011': 161, 'ICAA26_CPC_I012': 162, 'ICAA26_CPC_I013': 163, 'ICAA26_CPC_I014': 164, 'ICAA26_CPC_I015': 165, 'ICAA26_CPC_I016': 166, 'ICAA26_CPC_I017': 167, 'ICAA26_CPC_I018': 168, 'ICAA26_CPC_I019': 169, 'ICAA26_CPC_I020': 170, 'ICAA26_REV_I001': 171, 'ICAA26_REV_I002': 172, 'ICAA26_REV_I003': 173, 'ICAA26_REV_I004': 174, 'ICAA26_REV_I005': 175, 'ICAA26_REV_I006': 176, 'ICAA26_REV_I007': 177, 'ICAA26_REV_I008': 178, 'ICAA26_REV_I009': 179, 'ICAA26_REV_I010': 180, 'ICAA26_MSC_I001': 181, 'ICAA26_MSC_I002': 182, 'ICAA26_MSC_I003': 183, 'ICAA26_MSC_I004': 184, 'ICAA26_MSC_I005': 185, 'ICAA26_MSC_I006': 186, 'ICAA26_MSC_I007': 187, 'ICAA26_MSC_I008': 188, 'ICAA26_MSC_I009': 189, 'ICAA26_MSC_I010': 190, 'ICAA26_MSC_I011': 191, 'ICAA26_MSC_I012': 192, 'ICAA26_MSC_I013': 193, 'ICAA26_MSC_I014': 194, 'ICAA26_MSC_I015': 195, 'ICAA26_MSC_I016': 196, 'ICAA26_MSC_I017': 197, 'ICAA26_MSC_I018': 198, 'ICAA26_MSC_I019': 199, 'ICAA26_MSC_I020': 200, 'ICAA26_MSC_I021': 201, 'ICAA26_MSC_I022': 202, 'ICAA26_MSC_I023': 203, 'ICAA26_MSC_I024': 204, 'ICAA26_MSC_I025': 205, 'ICAA26_MSC_I026': 206, 'ICAA26_MSC_I027': 207, 'ICAA26_MSC_I028': 208, 'ICAA26_MSC_I029': 209, 'ICAA26_MSC_I030': 210, 'ICAA26_MSC_I031': 211, 'ICAA26_MSC_I032': 212, 'ICAA26_MSC_I033': 213, 'ICAA26_MSC_I034': 214, 'ICAA26_MSC_I035': 215, 'ICAA26_MSC_I036': 216, 'ICAA26_MSC_I037': 217, 'ICAA26_MSC_I038': 218, 'ICAA26_MSC_I039': 219, 'ICAA26_MSC_I040': 220, 'ICAA26_MSC_I041': 221, 'ICAA26_MSC_I042': 222, 'ICAA26_MSC_I043': 223, 'ICAA26_MSC_I044': 224, 'ICAA26_MSC_I045': 225, 'ICAA26_MSC_I046': 226, 'ICAA26_MSC_I047': 227, 'ICAA26_MSC_I048': 228, 'ICAA26_MSC_I049': 229, 'ICAA26_MSC_I050': 230, 'ICAA26_MSC_I051': 231, 'ICAA26_MSC_I052': 232, 'ICAA26_MSC_I053': 233, 'ICAA26_MSC_I054': 234, 'ICAA26_MSC_I055': 235, 'ICAA26_MSC_I056': 236, 'ICAA26_MSC_I057': 237, 'ICAA26_MSC_I058': 238, 'ICAA26_MSC_I059': 239, 'ICAA26_MSC_I060': 240, 'ICAA26_MSC_I061': 241, 'ICAA26_MSC_I062': 242, 'ICAA26_MSC_I063': 243, 'ICAA26_MSC_I064': 244, 'ICAA26_MSC_I065': 245, 'ICAA26_MSC_I066': 246, 'ICAA26_MSC_I067': 247, 'ICAA26_MSC_I068': 248, 'ICAA26_MSC_I069': 249, 'ICAA26_MSC_I070': 250, 'ICAA26_MSC_I071': 251, 'ICAA26_MSC_I072': 252, 'ICAA26_MSC_I073': 253, 'ICAA26_MSC_I074': 254, 'ICAA26_MSC_I075': 255}

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
    row = 2
    for i in range(1,36):
        IDs.append([f"ICAA26_ORG_I{str(i).zfill(3)}"])
    for i in range(1,15):
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
    for i in range(1,11):
        IDs.append([f"ICAA26_REV_I{str(i).zfill(3)}"])
    for i in range(1,76):
        IDs.append([f"ICAA26_MSC_I{str(i).zfill(3)}"])
    total_IDs = len(IDs)
    #for ID in IDs:
    #    ID_to_row_map_trial[ID[0]] = row
    #    row+=1
    try:
        writeValuestoRange(SHEET_IDs[sheetIdKey],IDs,f"{sheetIdKey}!A2:A{total_IDs+1}")
        print(f"Added IDs to {sheetIdKey}")
    except HttpError as error:
        print(f"An error occurred: {error}")

def prepareSheets():
    for key in SHEET_IDs.keys():
        add_headers(key)
        add_IDs(key)
    #print(ID_to_row_map_trial)


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