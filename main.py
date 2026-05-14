import os
import re
import json
import logging
from io import BytesIO
from datetime import datetime

import openpyxl
import gspread
from telegram import Update
from telegram.ext import Application, MessageHandler, ContextTypes, filters
from google.oauth2.service_account import Credentials

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Реестр26")

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

def get_sheet():
    creds_dict = json.loads(os.getenv("GOOGLE_CREDENTIALS"))
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

def cell_val(cell):
    return cell.value if cell and cell.value is not None else ""

def format_date(val):
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    m = re.match(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})", s)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if len(y) == 2:
            y = "20" + y
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return s

def get_month(desc, date_str):
    if desc:
        m = re.search(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}", str(desc))
        if m:
            parts = re.split(r"[./]", m.group())
            if len(parts) >= 2:
                try:
                    return int(parts[1])
                except:
                    pass
        months = {"январ":1,"феврал":2,"март":3,"апрел":4,"май":5,"мая":5,
                  "июн":6,"июл":7,"август":8,"сентябр":9,"октябр":10,"ноябр":11,"декабр":12}
        for w, n in months.items():
            if w in str(desc).lower():
                return n
    try:
        return int(date_str.split("/")[1])
    except:
        return ""

def get_week(date_str):
    try:
        return datetime.strptime(date_str, "%d/%m/%Y").isocalendar()[1]
    except:
        return ""

def get_article(desc, amount):
    if not desc:
        return ""
    d = str(desc)
    for kw in ["Оплата за услуги операций по картам Kaspi Gold",
               "Оплата рекламных услуг","Оплата за услуги процессинга Без НДС",
               "Оплата услуги по обработке данных","Оплата за информационно-технологические услуги",
               "Бонусы за отзыв клиенту","Оплата услуг по обработке данных, связанных с доставкой"]:
        if kw.lower() in d.lower():
            return "Комиссия за эквайринг"
    if "Возврат продаж с Kaspi.kz" in d:
        return "Возврат от покупателя"
    if "Возврат" in d and amount < 0:
        return "Возврат от покупателя"
    if "Возврат" in d and amount > 0:
        return "Оплата от покупателя, выручка"
    if "Продажи с Kaspi.kz" in d:
        return "Оплата от покупателя, выручка"
    return ""

def fmt_amount(val):
    try:
        f = float(val)
        return str(int(f)) if f == int(f) else str(round(f, 2))
    except:
        return str(val)

def process_xlsx(file_bytes):
    rows = []
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active

    # Get IBAN from row 3, column C (row=3, col=3)
    iban = str(cell_val(ws.cell(row=3, column=3))).strip()
    account = IBAN_MAP.get(iban, iban)

    # Data starts from row 14
    for row_idx in range(14, ws.max_row + 1):
        date_val = cell_val(ws.cell(row=row_idx, column=2))
        debit = cell_val(ws.cell(row=row_idx, column=3))
        credit = cell_val(ws.cell(row=row_idx, column=4))
        desc = str(cell_val(ws.cell(row=row_idx, column=9)))

        if not date_val:
            continue

        date_str = format_date(date_val)
        if not re.search(r"\d{2}/\d{2}/\d{4}", date_str):
            continue

        has_d = debit not in ("", None, "None")
        has_c = credit not in ("", None, "None")

        if has_d:
            try:
                amount = -float(str(debit).replace(" ", "").replace(",", "."))
            except:
                continue
        elif has_c:
            try:
                amount = float(str(credit).replace(" ", "").replace(",", "."))
            except:
                continue
        else:
            continue

        month = get_month(desc, date_str)
        year = date_str[-4:] if date_str else ""
        week = get_week(date_str)

        rows.append([year, str(month), str(week), date_str,
                     fmt_amount(amount), str(month), account,
                     get_article(desc, amount), desc, "", ""])
    return rows

async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Отправьте файл выписки (.xlsx)")
        return

    fname = (doc.file_name or "").lower()
    if not fname.endswith(".xlsx"):
        await update.message.reply_text("Поддерживаются только .xlsx файлы")
        return

    await update.message.reply_text("⏳ Обрабатываю выписку...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())
        rows = process_xlsx(file_bytes)

        if not rows:
            await update.message.reply_text("❌ Операции не найдены в файле")
            return

        sheet = get_sheet()
        sheet.append_rows(rows, value_input_option="RAW")
        await update.message.reply_text(f"✅ Готово! Добавлено {len(rows)} строк в таблицу.")
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))
    logger.info("Bot started!")
    app.run_polling()

if __name__ == "__main__":
    main()
