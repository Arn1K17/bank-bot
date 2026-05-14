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

    except Exception as e:
        logging.error(f"PDF parse error: {e}")
        return []

    return rows


# ---------------- SHEETS PUSH ----------------
def push(rows):
    try:
        sheet = get_sheet()
        sheet.append_rows(rows)
    except Exception as e:
        logging.error(f"Google Sheets error: {e}")


# ---------------- HANDLER ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    file_name = doc.file_name.lower()

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = await file.download_as_bytearray()
    file_bytes = bytes(file_bytes)

    # ---- ONLY PDF ----
    if not file_name.endswith(".pdf"):
        await update.message.reply_text("❌ Only PDF supported right now")
        return

    rows = parse_pdf(file_bytes)

    if not rows:
        await update.message.reply_text("❌ No data found in file")
        return

    push(rows)

    await update.message.reply_text(f"✅ Done: {len(rows)} rows added")


# ---------------- MAIN ----------------
async def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    print("Bot started")

    await asyncio.Event().wait()


if __name__ == "__main__":
    asyncio.run(main())
