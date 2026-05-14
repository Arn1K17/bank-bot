import os
import re
import logging
import requests
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
}

def get_account_name(iban):
    return IBAN_MAP.get(iban, iban)

def format_date(val):
    s = str(val).strip()
    m = re.match(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", s)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if len(y) == 2:
            y = "20" + y
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return s

def get_month_from_payment(description):
    if not description:
        return None
    desc = str(description)
    m = re.search(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}", desc)
    if m:
        parts = re.split(r"[./]", m.group())
        if len(parts) >= 2:
            try:
                return int(parts[1])
            except:
                pass
    m = re.search(r"\d{4}-(\d{2})-\d{2}", desc)
    if m:
        return int(m.group(1))
    months = {
        "январ": 1, "феврал": 2, "март": 3, "апрел": 4,
        "май": 5, "мая": 5, "июн": 6, "июл": 7,
        "август": 8, "сентябр": 9, "октябр": 10,
        "ноябр": 11, "декабр": 12,
    }
    desc_lower = desc.lower()
    for word, num in months.items():
        if word in desc_lower:
            return num
    return None

def get_article(description, amount):
    if not description:
        return ""
    desc = str(description)
    commission_keywords = [
        "Оплата за услуги операций по картам Kaspi Gold",
        "Оплата рекламных услуг",
        "Оплата за услуги процессинга Без НДС",
        "Оплата услуги по обработке данных",
        "Оплата за информационно-технологические услуги",
        "Оплата рекламных услуг, в том числе НДС",
        "Бонусы за отзыв клиенту",
        "Оплата услуг по обработке данных, связанных с доставкой",
    ]
    for kw in commission_keywords:
        if kw.lower() in desc.lower():
            return "Комиссия за эквайринг"
    if "Возврат продаж с Kaspi.kz" in desc:
        return "Возврат от покупателя"
    if "Возврат" in desc and amount < 0:
        return "Возврат от покупателя"
    if "Возврат" in desc and amount > 0:
        return "Оплата от покупателя, выручка"
    if "Продажи с Kaspi.kz" in desc:
        return "Оплата от покупателя, выручка"
    return ""

def format_amount(val):
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return str(round(f, 2))
    except:
        return str(val)

def get_week_number(date_str):
    try:
        d = datetime.strptime(date_str, "%d/%m/%Y")
        return d.isocalendar()[1]
    except:
        return ""

def is_empty(val):
    return val is None or str(val).strip() in ("", "None", "nan")

def process_kaspi_xlsx(file_bytes):
    rows = []
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active

    # Get IBAN from row 3, column C (index 2)
    iban = ""
    try:
        iban = str(ws.cell(row=3, column=3).value).strip()
    except:
        pass
    account = get_account_name(iban)

    for idx in range(14, ws.max_row + 1):
        date_val = ws.cell(row=idx, column=2).value
        debit = ws.cell(row=idx, column=3).value
        credit = ws.cell(row=idx, column=4).value
        desc_val = ws.cell(row=idx, column=9).value

        if is_empty(date_val):
            continue
        date_str_raw = str(date_val).strip()
        if not re.search(r"\d{2}[.\-/]\d{2}[.\-/]\d{4}", date_str_raw):
            # Try datetime object
            if hasattr(date_val, 'strftime'):
                date_str_raw = date_val.strftime("%d.%m.%Y")
            else:
                continue

        desc = str(desc_val).strip() if not is_empty(desc_val) else ""

        has_debit = not is_empty(debit)
        has_credit = not is_empty(credit)

        if has_debit:
            try:
                amount = -float(str(debit).replace(" ", "").replace(",", "."))
            except:
                continue
        elif has_credit:
            try:
                amount = float(str(credit).replace(" ", "").replace(",", "."))
            except:
                continue
        else:
            continue

        date_str = format_date(date_str_raw)
        month = get_month_from_payment(desc)
        if month is None:
            m = re.match(r"\d{2}[.\-/](\d{2})[.\-/]\d{4}", date_str_raw)
            month = int(m.group(1)) if m else ""

        year = date_str[-4:] if date_str else ""
        week = get_week_number(date_str)
        article = get_article(desc, amount)
        amount_str = format_amount(amount)

        rows.append([year, str(month), str(week), date_str, amount_str,
                     str(month), account, article, desc, "", ""])
    return rows

def process_bcc_pdf(file_bytes):
    rows = []
    account = "БЦК ТОО"
    
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 4:
                        continue
                    date_val = str(row[0]).strip() if row[0] else ""
                    if not re.match(r"\d{2}\.\d{2}\.\d{4}", date_val):
                        continue
                    
                    desc = str(row[-1]).strip() if row[-1] else ""
                    amount_str_raw = ""
                    
                    # Try to find amount columns
                    for cell in row[1:]:
                        cell_str = str(cell).strip() if cell else ""
                        if re.match(r"[\d\s]+[.,]\d{2}", cell_str):
                            amount_str_raw = cell_str
                            break
                    
                    if not amount_str_raw:
                        continue
                    
                    try:
                        amount = float(amount_str_raw.replace(" ", "").replace(",", "."))
                    except:
                        continue

                    date_str = format_date(date_val)
                    month = get_month_from_payment(desc)
                    if month is None:
                        m = re.match(r"\d{2}\.(\d{2})\.\d{4}", date_val)
                        month = int(m.group(1)) if m else ""

                    year = date_str[-4:] if date_str else ""
                    week = get_week_number(date_str)
                    article = get_article(desc, amount)
                    amount_fmt = format_amount(amount)

                    rows.append([year, str(month), str(week), date_str, amount_fmt,
                                 str(month), account, article, desc, "", ""])
    return rows

def push_to_sheets(rows):
    creds_dict = json.loads(GOOGLE_CREDENTIALS)
    scopes = ["https://spreadsheets.google.com/feeds",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    sheet.append_rows(rows, value_input_option="RAW")

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Пожалуйста, отправьте файл выписки (.xlsx или .pdf)")
        return

    fname = (doc.file_name or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".pdf")):
        await update.message.reply_text("Поддерживаются только файлы .xlsx и .pdf")
        return

    await update.message.reply_text("⏳ Обрабатываю выписку...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = await file.download_as_bytearray()

        if fname.endswith(".xlsx"):
            rows = process_kaspi_xlsx(bytes(file_bytes))
        else:
            rows = process_bcc_pdf(bytes(file_bytes))

        if not rows:
            await update.message.reply_text("❌ Не удалось найти операции в файле.")
            return

        push_to_sheets(rows)
        await update.message.reply_text(
            f"✅ Готово! Добавлено {len(rows)} строк в Google Sheets."
        )
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    logger.info("Bot started!")
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    app.run_polling()

if __name__ == "__main__":
    main()
