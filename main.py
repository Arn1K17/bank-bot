import os
import re
import json
import logging
from io import BytesIO
from datetime import datetime

import openpyxl
import pdfplumber
import gspread

from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters
from google.oauth2.service_account import Credentials

# ---------------- LOGGING ----------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------- ENV ----------------
BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Реестр26")

# ---------------- ACCOUNTS ----------------
IBAN_MAP = {
    "KZ81722S000020084562": "Kaspi Зухра",
    "KZ97722S000050068041": "Kaspi Серик",
    "KZ38722S000002562577": "Kaspi Орынбаева",
    "KZ94722S000041906224": "Kaspi КОКО",
    "KZ78722S000036046839": "Kaspi ТОО",
    "KZ50722S000022286244": "Kaspi Дильназ",
    "KZ508562203134868457": "BCC ТОО",
    "KZ97722C000015235365": "Kaspi Gold",
    "KZ448562204152575978": "BCC ИП Серик",
}

# ---------------- HELPERS ----------------
def get_account(iban):
    return IBAN_MAP.get(iban, iban)

def is_empty(v):
    return v is None or str(v).strip() in ("", "None", "nan")

def parse_date(date_str):
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", date_str)
    if not m:
        return None
    d, mo, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    return f"{int(d):02d}/{int(mo):02d}/{y}"

def get_week(date_str):
    try:
        d = datetime.strptime(date_str, "%d/%m/%Y")
        return d.isocalendar()[1]
    except:
        return ""

def format_amount(x):
    try:
        x = float(x)
        return str(int(x)) if x.is_integer() else str(round(x, 2))
    except:
        return str(x)

# ---------------- XLSX (KASPI) ----------------
def process_xlsx(file_bytes):
    rows = []
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active

    iban = str(ws.cell(3, 3).value or "").strip()
    account = get_account(iban)

    for i in range(14, ws.max_row + 1):
        date_val = ws.cell(i, 2).value
        debit = ws.cell(i, 3).value
        credit = ws.cell(i, 4).value
        desc = str(ws.cell(i, 9).value or "")

        if is_empty(date_val):
            continue

        date_str = parse_date(str(date_val))
        if not date_str:
            continue

        amount = None

        if not is_empty(debit):
            amount = -float(str(debit).replace(",", ".").replace(" ", ""))
        elif not is_empty(credit):
            amount = float(str(credit).replace(",", ".").replace(" ", ""))
        else:
            continue

        month = date_str.split("/")[1]
        week = get_week(date_str)

        rows.append([
            date_str[-4:], month, week,
            date_str, format_amount(amount),
            account, desc
        ])

    return rows

# ---------------- PDF DETECT ----------------
def detect_bank(file_bytes):
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ""
            if "Kaspi" in text or "CASPKZKA" in text:
                return "kaspi"
            if "ЦентрКредит" in text or "BCC" in text:
                return "bcc"
    except:
        pass
    return "unknown"

# ---------------- KASPI PDF ----------------
def process_kaspi_pdf(file_bytes):
    rows = []
    account = "Kaspi Gold"

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    if not row or len(row) < 2:
                        continue

                    date_raw = str(row[0])
                    date_str = parse_date(date_raw)
                    if not date_str:
                        continue

                    amount_raw = str(row[1]).replace("₸", "").replace(" ", "").replace(",", ".")
                    sign = -1 if "-" in amount_raw else 1

                    try:
                        amount = sign * float(re.sub(r"[^\d.]", "", amount_raw))
                    except:
                        continue

                    desc = str(row[3] if len(row) > 3 else "")

                    month = date_str.split("/")[1]
                    week = get_week(date_str)

                    rows.append([
                        date_str[-4:], month, week,
                        date_str, format_amount(amount),
                        account, desc
                    ])

    return rows

# ---------------- BCC PDF ----------------
def process_bcc_pdf(file_bytes):
    rows = []
    account = "BCC"

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    if not row or len(row) < 9:
                        continue

                    date_str = parse_date(str(row[1] or ""))
                    if not date_str:
                        continue

                    debit = row[7]
                    credit = row[8]

                    amount = None
                    if not is_empty(debit):
                        amount = -float(str(debit).replace(" ", "").replace(",", "."))
                    elif not is_empty(credit):
                        amount = float(str(credit).replace(" ", "").replace(",", "."))

                    if amount is None:
                        continue

                    desc = str(row[-1] or "")
                    month = date_str.split("/")[1]
                    week = get_week(date_str)

                    rows.append([
                        date_str[-4:], month, week,
                        date_str, format_amount(amount),
                        account, desc
                    ])

    return rows

# ---------------- GOOGLE SHEETS ----------------
def push(rows):
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDENTIALS),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )

    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    sheet.append_rows(rows)

# ---------------- TELEGRAM ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return await update.message.reply_text("Send file")

    name = doc.file_name.lower()

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = bytes(await file.download_as_bytearray())

    rows = []

    if name.endswith(".xlsx"):
        rows = process_xlsx(file_bytes)
    else:
        bank = detect_bank(file_bytes)

        if bank == "kaspi":
            rows = process_kaspi_pdf(file_bytes)
        elif bank == "bcc":
            rows = process_bcc_pdf(file_bytes)

    if not rows:
        return await update.message.reply_text("No data found")

    push(rows)

    await update.message.reply_text(f"Done: {len(rows)} rows")

# ---------------- MAIN ----------------
def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    print("Bot started")
    app.run_polling()

if __name__ == "__main__":
    main()
