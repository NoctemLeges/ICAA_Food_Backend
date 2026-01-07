from dotenv import load_dotenv
load_dotenv()
import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from flask import Flask, request, jsonify, redirect, make_response, render_template
import jwt
from datetime import datetime, timedelta
from functools import wraps

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

ID_to_row_map = {'ICAA26_ORG_I001': 2, 'ICAA26_ORG_I002': 3, 'ICAA26_ORG_I003': 4, 'ICAA26_ORG_I004': 5, 'ICAA26_ORG_I005': 6, 'ICAA26_ORG_I006': 7, 'ICAA26_ORG_I007': 8, 'ICAA26_ORG_I008': 9, 'ICAA26_ORG_I009': 10, 'ICAA26_ORG_I010': 11, 'ICAA26_ORG_I011': 12, 'ICAA26_ORG_I012': 13, 'ICAA26_ORG_I013': 14, 'ICAA26_ORG_I014': 15, 'ICAA26_ORG_I015': 16, 'ICAA26_ORG_I016': 17, 'ICAA26_ORG_I017': 18, 'ICAA26_ORG_I018': 19, 'ICAA26_ORG_I019': 20, 'ICAA26_ORG_I020': 21, 'ICAA26_ORG_I021': 22, 'ICAA26_ORG_I022': 23, 'ICAA26_ORG_I023': 24, 'ICAA26_ORG_I024': 25, 'ICAA26_ORG_I025': 26, 'ICAA26_ORG_I026': 27, 'ICAA26_ORG_I027': 28, 'ICAA26_ORG_I028': 29, 'ICAA26_ORG_I029': 30, 'ICAA26_ORG_I030': 31, 'ICAA26_ORG_I031': 32, 'ICAA26_ORG_I032': 33, 'ICAA26_ORG_I033': 34, 'ICAA26_ORG_I034': 35, 'ICAA26_ORG_I035': 36, 'ICAA26_SSC_I001': 37, 'ICAA26_SSC_I002': 38, 'ICAA26_SSC_I003': 39, 'ICAA26_SSC_I004': 40, 'ICAA26_SSC_I005': 41, 'ICAA26_SSC_I006': 42, 'ICAA26_SSC_I007': 43, 'ICAA26_SSC_I008': 44, 'ICAA26_SSC_I009': 45, 'ICAA26_SSC_I010': 46, 'ICAA26_SSC_I011': 47, 'ICAA26_SSC_I012': 48, 'ICAA26_SSC_I013': 49, 'ICAA26_SSC_I014': 50, 'ICAA26_SSC_I015': 51, 'ICAA26_SSC_I016': 52, 'ICAA26_CMC_I001': 53, 'ICAA26_CMC_I002': 54, 'ICAA26_CMC_I003': 55, 'ICAA26_CMC_I004': 56, 'ICAA26_CMC_I005': 57, 'ICAA26_CMC_I006': 58, 'ICAA26_CMC_I007': 59, 'ICAA26_CMC_I008': 60, 'ICAA26_CMC_I009': 61, 'ICAA26_CMC_I010': 62, 'ICAA26_CMC_I011': 63, 'ICAA26_CMC_I012': 64, 'ICAA26_CMC_I013': 65, 'ICAA26_CMC_I014': 66, 'ICAA26_GPC_I001': 67, 'ICAA26_GPC_I002': 68, 'ICAA26_GPC_I003': 69, 'ICAA26_GPC_I004': 70, 'ICAA26_GPC_I005': 71, 'ICAA26_GPC_I006': 72, 'ICAA26_GPC_I007': 73, 'ICAA26_GPC_I008': 74, 'ICAA26_GPC_I009': 75, 'ICAA26_GPC_I010': 76, 'ICAA26_GPC_I011': 77, 'ICAA26_GPC_I012': 78, 'ICAA26_GPC_I013': 79, 'ICAA26_GPC_I014': 80, 'ICAA26_GPC_I015': 81, 'ICAA26_GPC_I016': 82, 'ICAA26_GPC_I017': 83, 'ICAA26_GPC_I018': 84, 'ICAA26_GPC_I019': 85, 'ICAA26_GPC_I020': 86, 'ICAA26_SPC_I001': 87, 'ICAA26_SPC_I002': 88, 'ICAA26_SPC_I003': 89, 'ICAA26_SPC_I004': 90, 'ICAA26_SPC_I005': 91, 'ICAA26_SPC_I006': 92, 'ICAA26_SPC_I007': 93, 'ICAA26_SPC_I008': 94, 'ICAA26_SPC_I009': 95, 'ICAA26_SPC_I010': 96, 'ICAA26_SPC_I011': 97, 'ICAA26_SPC_I012': 98, 'ICAA26_SPC_I013': 99, 'ICAA26_SPC_I014': 100, 'ICAA26_SPC_I015': 101, 'ICAA26_SPC_I016': 102, 'ICAA26_SPC_I017': 103, 'ICAA26_SPC_I018': 104, 'ICAA26_SPC_I019': 105, 'ICAA26_SPC_I020': 106, 'ICAA26_SPC_I021': 107, 'ICAA26_SPC_I022': 108, 'ICAA26_SPC_I023': 109, 'ICAA26_SPC_I024': 110, 'ICAA26_SPC_I025': 111, 'ICAA26_SPC_I026': 112, 'ICAA26_SPC_I027': 113, 'ICAA26_SPC_I028': 114, 'ICAA26_SPC_I029': 115, 'ICAA26_SPC_I030': 116, 'ICAA26_SPC_I031': 117, 'ICAA26_SPC_I032': 118, 'ICAA26_SPC_I033': 119, 'ICAA26_SPC_I034': 120, 'ICAA26_SPC_I035': 121, 'ICAA26_SPC_I036': 122, 'ICAA26_SPC_I037': 123, 'ICAA26_SPC_I038': 124, 'ICAA26_SPC_I039': 125, 'ICAA26_SPC_I040': 126, 'ICAA26_SPC_I041': 127, 'ICAA26_SPC_I042': 128, 'ICAA26_SPC_I043': 129, 'ICAA26_SPC_I044': 130, 'ICAA26_SPC_I045': 131, 'ICAA26_SPC_I046': 132, 'ICAA26_SPC_I047': 133, 'ICAA26_SPC_I048': 134, 'ICAA26_SPC_I049': 135, 'ICAA26_SPC_I050': 136, 'ICAA26_SPC_I051': 137, 'ICAA26_SPC_I052': 138, 'ICAA26_SPC_I053': 139, 'ICAA26_SPC_I054': 140, 'ICAA26_SPC_I055': 141, 'ICAA26_SPC_I056': 142, 'ICAA26_SPC_I057': 143, 'ICAA26_SPC_I058': 144, 'ICAA26_SPC_I059': 145, 'ICAA26_SPC_I060': 146, 'ICAA26_SPC_I061': 147, 'ICAA26_SPC_I062': 148, 'ICAA26_SPC_I063': 149, 'ICAA26_SPC_I064': 150, 'ICAA26_SPC_I065': 151, 'ICAA26_SPC_I066': 152, 'ICAA26_SPC_I067': 153, 'ICAA26_SPC_I068': 154, 'ICAA26_SPC_I069': 155, 'ICAA26_SPC_I070': 156, 'ICAA26_SPC_I071': 157, 'ICAA26_SPC_I072': 158, 'ICAA26_SPC_I073': 159, 'ICAA26_SPC_I074': 160, 'ICAA26_SPC_I075': 161, 'ICAA26_MPC_I001': 162, 'ICAA26_MPC_I002': 163, 'ICAA26_MPC_I003': 164, 'ICAA26_MPC_I004': 165, 'ICAA26_MPC_I005': 166, 'ICAA26_MPC_I006': 167, 'ICAA26_MPC_I007': 168, 'ICAA26_MPC_I008': 169, 'ICAA26_MPC_I009': 170, 'ICAA26_MPC_I010': 171, 'ICAA26_MPC_I011': 172, 'ICAA26_MPC_I012': 173, 'ICAA26_MPC_I013': 174, 'ICAA26_MPC_I014': 175, 'ICAA26_MPC_I015': 176, 'ICAA26_VLT_I001': 177, 'ICAA26_VLT_I002': 178, 'ICAA26_VLT_I003': 179, 'ICAA26_VLT_I004': 180, 'ICAA26_VLT_I005': 181, 'ICAA26_VLT_I006': 182, 'ICAA26_VLT_I007': 183, 'ICAA26_VLT_I008': 184, 'ICAA26_VLT_I009': 185, 'ICAA26_VLT_I010': 186, 'ICAA26_VLT_I011': 187, 'ICAA26_VLT_I012': 188, 'ICAA26_VLT_I013': 189, 'ICAA26_VLT_I014': 190, 'ICAA26_VLT_I015': 191, 'ICAA26_VLT_I016': 192, 'ICAA26_VLT_I017': 193, 'ICAA26_VLT_I018': 194, 'ICAA26_VLT_I019': 195, 'ICAA26_VLT_I020': 196, 'ICAA26_VLT_I021': 197, 'ICAA26_VLT_I022': 198, 'ICAA26_VLT_I023': 199, 'ICAA26_VLT_I024': 200, 'ICAA26_VLT_I025': 201, 'ICAA26_AUT_I001': 202, 'ICAA26_AUT_I002': 203, 'ICAA26_AUT_I003': 204, 'ICAA26_AUT_I004': 205, 'ICAA26_AUT_I005': 206, 'ICAA26_AUT_I006': 207, 'ICAA26_AUT_I007': 208, 'ICAA26_AUT_I008': 209, 'ICAA26_AUT_I009': 210, 'ICAA26_AUT_I010': 211, 'ICAA26_AUT_I011': 212, 'ICAA26_AUT_I012': 213, 'ICAA26_AUT_I013': 214, 'ICAA26_AUT_I014': 215, 'ICAA26_AUT_I015': 216, 'ICAA26_AUT_I016': 217, 'ICAA26_AUT_I017': 218, 'ICAA26_AUT_I018': 219, 'ICAA26_AUT_I019': 220, 'ICAA26_AUT_I020': 221, 'ICAA26_AUT_I021': 222, 'ICAA26_AUT_I022': 223, 'ICAA26_AUT_I023': 224, 'ICAA26_AUT_I024': 225, 'ICAA26_AUT_I025': 226, 'ICAA26_AUT_I026': 227, 'ICAA26_AUT_I027': 228, 'ICAA26_AUT_I028': 229, 'ICAA26_AUT_I029': 230, 'ICAA26_AUT_I030': 231, 'ICAA26_AUT_I031': 232, 'ICAA26_AUT_I032': 233, 'ICAA26_AUT_I033': 234, 'ICAA26_AUT_I034': 235, 'ICAA26_AUT_I035': 236, 'ICAA26_AUT_I036': 237, 'ICAA26_AUT_I037': 238, 'ICAA26_AUT_I038': 239, 'ICAA26_AUT_I039': 240, 'ICAA26_AUT_I040': 241, 'ICAA26_CPC_I001': 242, 'ICAA26_CPC_I002': 243, 'ICAA26_CPC_I003': 244, 'ICAA26_CPC_I004': 245, 'ICAA26_CPC_I005': 246, 'ICAA26_CPC_I006': 247, 'ICAA26_CPC_I007': 248, 'ICAA26_CPC_I008': 249, 'ICAA26_CPC_I009': 250, 'ICAA26_CPC_I010': 251, 'ICAA26_CPC_I011': 252, 'ICAA26_CPC_I012': 253, 'ICAA26_CPC_I013': 254, 'ICAA26_CPC_I014': 255, 'ICAA26_CPC_I015': 256, 'ICAA26_CPC_I016': 257, 'ICAA26_CPC_I017': 258, 'ICAA26_CPC_I018': 259, 'ICAA26_CPC_I019': 260, 'ICAA26_CPC_I020': 261, 'ICAA26_REV_I001': 262, 'ICAA26_REV_I002': 263, 'ICAA26_REV_I003': 264, 'ICAA26_REV_I004': 265, 'ICAA26_REV_I005': 266, 'ICAA26_REV_I006': 267, 'ICAA26_REV_I007': 268, 'ICAA26_REV_I008': 269, 'ICAA26_REV_I009': 270, 'ICAA26_REV_I010': 271, 'ICAA26_REV_I011': 272, 'ICAA26_MSC_I001': 273, 'ICAA26_MSC_I002': 274, 'ICAA26_MSC_I003': 275, 'ICAA26_MSC_I004': 276, 'ICAA26_MSC_I005': 277, 'ICAA26_MSC_I006': 278, 'ICAA26_MSC_I007': 279, 'ICAA26_MSC_I008': 280, 'ICAA26_MSC_I009': 281, 'ICAA26_MSC_I010': 282, 'ICAA26_MSC_I011': 283, 'ICAA26_MSC_I012': 284, 'ICAA26_MSC_I013': 285, 'ICAA26_MSC_I014': 286, 'ICAA26_MSC_I015': 287, 'ICAA26_MSC_I016': 288, 'ICAA26_MSC_I017': 289, 'ICAA26_MSC_I018': 290, 'ICAA26_MSC_I019': 291, 'ICAA26_MSC_I020': 292, 'ICAA26_MSC_I021': 293, 'ICAA26_MSC_I022': 294, 'ICAA26_MSC_I023': 295, 'ICAA26_MSC_I024': 296, 'ICAA26_MSC_I025': 297, 'ICAA26_MSC_I026': 298, 'ICAA26_MSC_I027': 299, 'ICAA26_MSC_I028': 300, 'ICAA26_MSC_I029': 301, 'ICAA26_MSC_I030': 302, 'ICAA26_MSC_I031': 303, 'ICAA26_MSC_I032': 304, 'ICAA26_MSC_I033': 305, 'ICAA26_MSC_I034': 306, 'ICAA26_MSC_I035': 307, 'ICAA26_MSC_I036': 308, 'ICAA26_MSC_I037': 309, 'ICAA26_MSC_I038': 310, 'ICAA26_MSC_I039': 311, 'ICAA26_MSC_I040': 312, 'ICAA26_MSC_I041': 313, 'ICAA26_MSC_I042': 314, 'ICAA26_MSC_I043': 315, 'ICAA26_MSC_I044': 316, 'ICAA26_MSC_I045': 317, 'ICAA26_MSC_I046': 318, 'ICAA26_MSC_I047': 319, 'ICAA26_MSC_I048': 320, 'ICAA26_MSC_I049': 321, 'ICAA26_MSC_I050': 322, 'ICAA26_MSC_I051': 323, 'ICAA26_MSC_I052': 324, 'ICAA26_MSC_I053': 325, 'ICAA26_MSC_I054': 326, 'ICAA26_MSC_I055': 327, 'ICAA26_MSC_I056': 328, 'ICAA26_MSC_I057': 329, 'ICAA26_MSC_I058': 330, 'ICAA26_MSC_I059': 331, 'ICAA26_MSC_I060': 332, 'ICAA26_MSC_I061': 333, 'ICAA26_MSC_I062': 334, 'ICAA26_MSC_I063': 335, 'ICAA26_MSC_I064': 336, 'ICAA26_MSC_I065': 337, 'ICAA26_MSC_I066': 338, 'ICAA26_MSC_I067': 339, 'ICAA26_MSC_I068': 340, 'ICAA26_MSC_I069': 341, 'ICAA26_MSC_I070': 342, 'ICAA26_MSC_I071': 343, 'ICAA26_MSC_I072': 344, 'ICAA26_MSC_I073': 345, 'ICAA26_MSC_I074': 346, 'ICAA26_MSC_I075': 347, 'ICAA26_KNS_I001': 348,'ICAA26_KNS_I002': 349, 'ICAA26_KNS_I003': 350, 'ICAA26_KNS_I004': 351, 'ICAA26_KNS_I005': 352, 'ICAA26_KNS_I006': 353, 'ICAA26_KNS_I007': 354, 'ICAA26_KNS_I008': 355, 'ICAA26_KNS_I009': 356, 'ICAA26_KNS_I010': 357}

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
    for i in range(1,17):
        IDs.append([f"ICAA26_SSC_I{str(i).zfill(3)}"])
    for i in range(1,15):
        IDs.append([f"ICAA26_CMC_I{str(i).zfill(3)}"])
    for i in range(1,21):
        IDs.append([f"ICAA26_GPC_I{str(i).zfill(3)}"])
    for i in range(1,76):
        IDs.append([f"ICAA26_SPC_I{str(i).zfill(3)}"])
    for i in range(1,16):
        IDs.append([f"ICAA26_MPC_I{str(i).zfill(3)}"])
    for i in range(1,26):
        IDs.append([f"ICAA26_VLT_I{str(i).zfill(3)}"])
    for i in range(1,41):
        IDs.append([f"ICAA26_AUT_I{str(i).zfill(3)}"])
    for i in range(1,21):
        IDs.append([f"ICAA26_CPC_I{str(i).zfill(3)}"])
    for i in range(1,12):
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

# def get_food_stats():
#     service = get_sheets_service()

#     rows_result = (
#         service.spreadsheets()
#         .values()
#         .get(
#             spreadsheetId=SHEET_IDs[today],
#             range=f"{today}!A2:A"
#         )
#         .execute()
#     )

#     food_result = (
#         service.spreadsheets()
#         .values()
#         .get(
#             spreadsheetId=SHEET_IDs[today],
#             range=f"{today}!C2:C"
#         )
#         .execute()
#     )

#     rows = rows_result.get("values", [])
#     food = food_result.get("values", [])

#     taken = 0
#     not_taken = 0

#     for i in range(len(rows)):
#         value = food[i][0].strip().lower() if i < len(food) and food[i] else ""

#         if value == "yes":
#             taken += 1
#         elif value == "duplicate":
#             continue
#         else:
#             not_taken += 1

#     return taken, not_taken

def get_food_stats():
    service = get_sheets_service()

    rows_result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_IDs[today],
        range=f"{today}!A2:A"
    ).execute()

    names_result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_IDs[today],
        range=f"{today}!B2:B"
    ).execute()

    food_result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_IDs[today],
        range=f"{today}!C2:C"
    ).execute()

    rows = rows_result.get("values", [])
    names = names_result.get("values", [])
    food = food_result.get("values", [])

    taken = 0
    not_taken = 0
    seen_names = set()

    for i in range(len(rows)):
        # ---- NAME CHECK (ADDED) ----
        name = names[i][0].strip().lower() if i < len(names) and names[i] else ""

        if not name:
            continue

        if name in seen_names:
            continue

        seen_names.add(name)
        # ----------------------------

        value = food[i][0].strip().lower() if i < len(food) and food[i] else ""

        if value == "yes":
            taken += 1
        elif value == "duplicate":
            continue
        else:
            not_taken += 1

    return taken, not_taken

def get_all_food_data():
    service = get_sheets_service()

    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SHEET_IDs[today],
            range=f"{today}!A2:C"
        )
        .execute()
    )

    return result.get("values", [])

def mark_food_with_duplicate_check(scanned_id: str):
    data = get_all_food_data()

    scanned_row = None
    scanned_name = None
    scanned_status = None

    for idx, row in enumerate(data):
        if len(row) >= 2 and row[0] == scanned_id:
            scanned_row = idx + 2
            scanned_name = row[1].strip()
            scanned_status = row[2].strip().lower() if len(row) >= 3 else ""
            break

    if not scanned_row:
        return "invalid", None

    if scanned_status == "yes":
        return "already_taken", scanned_name

    for row in data:
        if len(row) >= 3:
            name = row[1].strip().lower()
            food_status = row[2].strip().lower()

            if name == scanned_name.lower() and food_status == "yes":
                writeValuestoRange(
                    SHEET_IDs[today],
                    [["Duplicate"]],
                    f"{today}!C{scanned_row}"
                )
                return "duplicate", scanned_name

    writeValuestoRange(
        SHEET_IDs[today],
        [["Yes"]],
        f"{today}!C{scanned_row}"
    )
    return "yes", scanned_name
#-----------------------------------------------------------------------------

#---------------------------JWT Helpers---------------------------------------
JWT_SECRET = os.environ["JWT_SECRET"]
JWT_EXP_MINUTES = int(os.environ.get("JWT_EXP_MINUTES", 1440))
FOOD_ADMIN_PASSWORD = os.environ["FOOD_ADMIN_PASSWORD"]

def generate_token():
    payload = {
        "role": "food_admin",
        "exp": datetime.utcnow() + timedelta(minutes=JWT_EXP_MINUTES),
        "iat": datetime.utcnow()
    }
    return jwt.encode(payload, JWT_SECRET, algorithm="HS256")

def verify_token(token):
    try:
        return jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None
#-----------------------------------------------------------------------------

#---------------------------MIDDLEWARE----------------------------------------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.cookies.get("token")

        if not token or not verify_token(token):
            return redirect("/login")

        return f(*args, **kwargs)
    return decorated
#-----------------------------------------------------------------------------

#-----------------------ROUTES------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    taken, not_taken = get_food_stats()
    total = taken + not_taken

    return render_template(
    "dashboard.html",
    day=today,
    taken=taken,
    not_taken=not_taken,
    total=total
)

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        data = request.get_json()
        password = data.get("password")

        if password != FOOD_ADMIN_PASSWORD:
            return jsonify({"error": "Invalid password"}), 401

        token = generate_token()

        res = make_response(jsonify({"success": True}))
        res.set_cookie(
            "token",
            token,
            httponly=True,
            secure=False,
            samesite="Lax",
            max_age=24 * 60 * 60
        )
        return res

    return render_template("login.html")

@app.route("/food")
@login_required
def food():
    scanned_id = request.args.get("id")

    status, scanned_name = mark_food_with_duplicate_check(scanned_id)

    if status == "invalid":
        return render_template(
            "food_status.html",
            image="/images/Invalid.png",
            heading="Invalid QR",
            message="This QR code is not recognized, as this ID has not been assigned to anyone.",
            css_class="error"
        )

    if status == "already_taken":
        return render_template(
            "food_status.html",
            image="/images/Served.png",
            heading="Food Already Served",
            message="Food has already been served.",
            name=scanned_name,
            css_class="warning"
        )

    if status == "duplicate":
        return render_template(
            "food_status.html",
            image="/images/Served.png",
            heading="Food Already Served",
            message="Food already served using another ID.",
            name=scanned_name,
            css_class="warning"
        )

    return render_template(
        "food_status.html",
        image="/images/Served.png",
        heading="Food Served Successfully",
        message="Food Served to Guest.",
        name=scanned_name,
        id=scanned_id,
        css_class="success"
    )

@app.route("/ping")
def pong():
    return "pong"
#-----------------------------------------------------------------------------

#---------------------STARTUP-------------------------------------------------
if __name__ == "__main__":
    #prepareSheets()
    app.run()
#-----------------------------------------------------------------------------