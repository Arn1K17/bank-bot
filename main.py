import os
import json
import logging
import asyncio
from io import BytesIO

import pdfplumber
import gspread
from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")


def get_sheet():
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDENTIALS),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def parse_pdf(file_bytes):
    rows = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                rows.append([line])

    return rows


def push(rows):
    sheet = get_sheet()
    sheet.append_rows(rows)


async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = await file.download_as_bytearray()

    rows = parse_pdf(bytes(file_bytes))

    if rows:
        push(rows)
        await update.message.reply_text(f"Done: {len(rows)} rows")
    else:
        await update.message.reply_text("No data found")


async def run():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    print("Bot started")

    await asyncio.Event().wait()


if __name__ == "__main__":
    asyncio.run(run())
