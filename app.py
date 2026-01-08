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
from constants import ID_TO_ROW_MAP

#--------------GLOBALS--------------------------------------------

today = "Day 3 Lunch"

app = Flask(__name__)

SERVICE_ACCOUNT_FILE = "icaa-qr-generation-dc58b73870cd.json"

_service = None

SHEET_IDs = {
    "Day 1 Lunch"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 2 Lunch"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 2 Dinner"  : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4",
    "Day 3 Lunch" : "18LOeMF04XYSGKSy5TYnvTcUjWj8zy0BT82bhWGtRdz4"
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ID_to_row_map_trial = {}

ID_to_row_map = ID_TO_ROW_MAP

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

def get_food_stats():
    service = get_sheets_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_IDs[today],
        range=f"{today}!B2:C"
    ).execute()

    data = result.get("values", [])

    name_status = {}

    for row in data:
        if len(row) < 1:
            continue

        name = row[0].strip().lower()
        status = row[1].strip().lower() if len(row) > 1 else ""

        if not name:
            continue

        # If any ID for this name is marked yes → taken
        if name not in name_status:
            name_status[name] = status
        elif status == "yes":
            name_status[name] = "yes"

    taken = sum(1 for s in name_status.values() if s == "yes")
    not_taken = sum(1 for s in name_status.values() if s != "yes")

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
