import os
import json
import logging
from io import BytesIO

import pdfplumber
import gspread
from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters
from google.oauth2.service_account import Credentials
from pdfminer.pdfparser import PDFSyntaxError

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")


# ---------------- GOOGLE SHEETS ----------------
def get_sheet():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")

    if not creds_json:
        raise Exception("GOOGLE_CREDENTIALS missing")

    creds_dict = json.loads(creds_json)

    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    client = gspread.authorize(creds)

    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

# ---------------- PDF PARSER ----------------
def parse_pdf(file_bytes: bytes):
    rows = []

    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                for line in text.split("\n"):
                    line = line.strip()
                    if line:
                        rows.append([line])

    except PDFSyntaxError:
        logging.error("Invalid PDF")
        return []
    except Exception as e:
        logging.error(f"PDF error: {e}")
        return []

    return rows


# ---------------- SHEETS ----------------
def push(rows):
    try:
        sheet = get_sheet()
        sheet.append_rows(rows, value_input_option="RAW")
    except Exception as e:
        logging.error(f"Sheets error: {e}")


# ---------------- HANDLER ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    file_name = (doc.file_name or "").lower()

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = bytes(await file.download_as_bytearray())

    if not file_name.endswith(".pdf"):
        await update.message.reply_text("Only PDF allowed")
        return

    rows = parse_pdf(file_bytes)

    if not rows:
        await update.message.reply_text("No data in file")
        return

    push(rows)

    await update.message.reply_text(f"Done: {len(rows)} rows")


# ---------------- MAIN (FIXED) ----------------
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    print("Bot started")

    # ❗ ЭТО ГЛАВНЫЙ FIX
    app.run_polling()


if __name__ == "__main__":
    main()


async def main():
    print("BOT_TOKEN:", BOT_TOKEN is not None)
    print("SPREADSHEET_ID:", SPREADSHEET_ID)
    print("GOOGLE_CREDENTIALS:", GOOGLE_CREDENTIALS is not None)

    app = Application.builder().token(BOT_TOKEN).build()
