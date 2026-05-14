import os
import re
import logging
from io import BytesIO
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes
import gspread
from google.oauth2.service_account import Credentials
import json
from datetime import datetime
import openpyxl
import pdfplumber
import asyncio

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.environ.get("GOOGLE_CREDENTIALS")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
SHEET_NAME = os.environ.get("SHEET_NAME", "Реестр26")

IBAN_MAP = {
    "KZ81722S000020084562": "Каспи Имангазиева Зухра",
    "KZ97722S000050068041": "Каспи СЕРИК",
    "KZ38722S000002562577": "Каспи Орынбаева",
    "KZ94722S000041906224": "Каспи КОКО (Сулейменов)",
    "KZ78722S000036046839": "Каспи ТОО",
    "KZ50722S000022286244": "Каспи Имангазиева Дильназ",
    "KZ508562203134868457": "БЦК ТОО",
    "KZ97722C000015235365": "Арман каспи голд",
    "KZ038562204137753855": "БЦК Имангазиева",
    "KZ448562204152575978": "БЦК ИП Серик",
}


def is_empty(v):
    return v is None or str(v).strip() in ("", "None", "nan")


def format_date(val):
    s = str(val).strip()
    m = re.match(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", s)
    if not m:
        return s
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


def detect_bank(file_bytes):
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            text = pdf.pages[0].extract_text() or ""
            if "Kaspi" in text or "CASPKZKA" in text:
                return "kaspi"
            if "ЦентрКредит" in text or "KCJBKZKX" in text:
                return "bcc"
    except:
        pass
    return "unknown"


# ---------------- XLSX ----------------
def process_kaspi_xlsx(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active

    rows = []

    for i in range(14, ws.max_row + 1):
        date_val = ws.cell(i, 2).value
        debit = ws.cell(i, 3).value
        credit = ws.cell(i, 4).value
        desc = ws.cell(i, 9).value or ""

        if is_empty(date_val):
            continue

        date_str = format_date(date_val)

        if debit:
            amount = -float(str(debit).replace(",", "."))
        elif credit:
            amount = float(str(credit).replace(",", "."))
        else:
            continue

        rows.append([
            date_str[-4:],
            date_str,
            str(get_week(date_str)),
            str(amount),
            desc
        ])

    return rows


# ---------------- BCC PDF ----------------
def process_bcc(file_bytes):
    rows = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for r in table:
                    if not r or len(r) < 9:
                        continue

                    date_raw = (r[1] or "").split()[0]
                    m = re.match(r"(\d{1,2})\.(\d{2})\.(\d{4})", date_raw)
                    if not m:
                        continue

                    d, mo, y = m.groups()
                    date_str = f"{d.zfill(2)}/{mo}/{y}"

                    def parse(x):
                        try:
                            return float(re.sub(r"[^\d\-,.]", "", str(x)).replace(",", "."))
                        except:
                            return None

                    amount = parse(r[7]) or -parse(r[8]) if r[7] else parse(r[8])
                    if amount is None:
                        continue

                    rows.append([
                        y,
                        date_str,
                        str(get_week(date_str)),
                        str(amount),
                        r[-1] or ""
                    ])

    return rows


# ---------------- KASPI PDF ----------------
def process_kaspi_pdf(file_bytes):
    rows = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for t in tables:
                for r in t:
                    if not r or len(r) < 3:
                        continue

                    if not re.match(r"\d{1,2}\.\d{2}\.\d{2,4}", str(r[0])):
                        continue

                    date = str(r[0]).split()[0]
                    d, mo, y = date.split(".")
                    if len(y) == 2:
                        y = "20" + y

                    amount = float(
                        str(r[1])
                        .replace("₸", "")
                        .replace(" ", "")
                        .replace(",", ".")
                    )

                    rows.append([
                        y,
                        f"{d}/{mo}/{y}",
                        str(get_week(f"{d}/{mo}/{y}")),
                        str(amount),
                        r[2] or ""
                    ])

    return rows


# ---------------- SHEETS ----------------
def push(rows):
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDENTIALS),
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    sheet.append_rows(rows)


# ---------------- BOT ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if not doc:
        return await update.message.reply_text("Отправь файл")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = bytes(await file.download_as_bytearray())

    name = doc.file_name.lower()

    if name.endswith(".xlsx"):
        rows = process_kaspi_xlsx(file_bytes)

    else:
        bank = detect_bank(file_bytes)

        if bank == "kaspi":
            rows = process_kaspi_pdf(file_bytes)
        elif bank == "bcc":
            rows = process_bcc(file_bytes)
        else:
            rows = process_kaspi_pdf(file_bytes) or process_bcc(file_bytes)

    if not rows:
        return await update.message.reply_text("ничего не найдено")

    push(rows)

    await update.message.reply_text(f"готово: {len(rows)} строк")


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))
    app.run_polling()


if __name__ == "__main__":
    main()
