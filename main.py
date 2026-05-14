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
from pdfminer.pdfparser import PDFSyntaxError

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

            # защита от пустых/битых pdf
            if not pdf.pages:
                return []

            for page in pdf.pages:
                try:
                    text = page.extract_text()
                except Exception:
                    continue

                if not text:
                    continue

                for line in text.split("\n"):
                    line = line.strip()
                    if line:
                        rows.append([line])

    except PDFSyntaxError:
        logging.error("Invalid PDF file (PDFSyntaxError)")
        return []
    except Exception as e:
        logging.error(f"PDF parse error: {e}")
        return []

    return rows


# ---------------- SHEETS PUSH (FAST VERSION) ----------------
def push(rows):
    try:
        sheet = get_sheet()

        # быстрее чем append_rows для больших данных
        sheet.append_rows(rows, value_input_option="RAW")

    except Exception as e:
        logging.error(f"Google Sheets error: {e}")


# ---------------- HANDLER ----------------
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    file_name = (doc.file_name or "").lower()

    await update.message.reply_text("Processing...")

    file = await context.bot.get_file(doc.file_id)
    file_bytes = await file.download_as_bytearray()
    file_bytes = bytes(file_bytes)

    # ---- FILTER ----
    if not file_name.endswith(".pdf"):
        await update.message.reply_text("❌ Only PDF supported right now")
        return

    rows = parse_pdf(file_bytes)

    if not rows:
        await update.message.reply_text("❌ No readable text in PDF (or file is invalid)")
        return

    push(rows)

    await update.message.reply_text(f"✅ Done: {len(rows)} rows added")


# ---------------- MAIN ----------------
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ALL, handle))

    print("Bot started")
    app.run_polling()


if __name__ == "__main__":
    main()
