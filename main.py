import os
import json
import logging
from io import BytesIO

import pdfplumber
import gspread
from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")


# ---------------- GOOGLE SHEETS ----------------
def get_sheet():
    creds_dict = json.loads(GOOGLE_CREDENTIALS)

    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


# ---------------- PDF PARSER (простая версия) ----------------
def parse_pdf(file_bytes):
    rows = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")

            for line in lines:
                if len(line) < 5:
                    continue

                # супер простой парсер (чтобы не падало)
                rows.append([line])

    return rows


# ---------------- PUSH ----------------
def push(rows):
    sheet = get_sheet()
    sheet.append_rows(rows)


# ---------------- TELEGRAM ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Send file")
        return

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = await file.download_as_bytearray()

    rows = parse_pdf(bytes(file_bytes))

    if not rows:
        await update.message.reply_text("No data found")
        return

    try:
        push(rows)
        await update.message.reply_text(f"Done: {len(rows)} rows")
    except Exception as e:
        logger.exception(e)
        await update.message.reply_text("Google Sheets error")


# ---------------- MAIN ----------------
def main():
    if not BOT_TOKEN:
        raise Exception("BOT_TOKEN missing")

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    print("Bot started")
    app.run_polling()


if __name__ == "__main__":
    main()
