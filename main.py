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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Реестр26")
SPREADSHEET_URL = f"https://docs.google.com/spreadsheets/d/{os.getenv('SPREADSHEET_ID')}"

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
    "KZ448562204152575978": "БЦК Серик ИП",
}

def get_spreadsheet():
    creds_dict = json.loads(os.getenv("GOOGLE_CREDENTIALS"))
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def get_sheet():
    return get_spreadsheet().worksheet(SHEET_NAME)

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

def cell_val(cell):
    return cell.value if cell and cell.value is not None else ""

def parse_num(s):
    s = str(s or "").replace(" ", "").replace("\xa0", "").replace(",", ".").replace("\n", "")
    try:
        return float(s)
    except:
        return 0

# ============ СВЕРКА ОСТАТКОВ ============
def check_balance(account_name, bank_closing_balance):
    """Сверяет остаток из банка с ДДС таблицей"""
    try:
        spreadsheet = get_spreadsheet()

        # Берём начальный остаток из Счета2026(Справка)
        справка = spreadsheet.worksheet("Счета2026(Справка)")
        справка_data = справка.get_all_values()

        initial_balance = None
        for row in справка_data:
            if row and row[0].strip() == account_name.strip():
                try:
                    initial_balance = float(str(row[1]).replace(" ", "").replace(",", ".").replace("\xa0", ""))
                    break
                except:
                    pass

        if initial_balance is None:
            return None, None, f"Счет '{account_name}' не найден в Счета2026(Справка)"

        # Суммируем все операции в Реестре по этому счету
        реестр = spreadsheet.worksheet(SHEET_NAME)
        реестр_data = реестр.get_all_values()

        total_operations = 0.0
        for row in реестр_data[1:]:  # пропускаем заголовок
            if len(row) >= 7 and row[6].strip() == account_name.strip():
                try:
                    total_operations += float(str(row[4]).replace(" ", "").replace(",", "."))
                except:
                    pass

        dds_balance = round(initial_balance + total_operations, 2)
        bank_balance = round(bank_closing_balance, 2)

        return dds_balance, bank_balance, None
    except Exception as e:
        return None, None, str(e)

# ============ XLSX ============
def process_xlsx(file_bytes):
    rows = []
    closing_balance = None

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    iban = str(cell_val(ws.cell(row=3, column=3))).strip()
    account = IBAN_MAP.get(iban, iban)

    # Ищем исходящий остаток
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = str(cell_val(ws.cell(row=row_idx, column=col_idx))).lower()
            if "исходящ" in cell and "сальдо" in cell:
                # Берём следующую ячейку или ищем число в строке
                for c in range(col_idx + 1, min(col_idx + 5, ws.max_column + 1)):
                    v = cell_val(ws.cell(row=row_idx, column=c))
                    if v:
                        try:
                            closing_balance = float(str(v).replace(" ", "").replace(",", "."))
                            break
                        except:
                            pass
                if closing_balance:
                    break
        if closing_balance:
            break

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

    return rows, account, closing_balance

# ============ PDF Kaspi Gold ============
def process_kaspi_gold_pdf(file_bytes):
    rows = []
    account = "Арман каспи голд"
    closing_balance = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        # Ищем исходящий остаток в тексте первой страницы
        first_text = pdf.pages[0].extract_text() or ""

        # "Доступно на 12.05.26: + 173 416,82 ₸"
        m = re.search(r"Доступно на \d{2}\.\d{2}\.\d{2,4}[:\s]+\+?\s*([\d\s,]+)\s*₸", first_text)
        if m:
            closing_balance = parse_num(m.group(1))

        # Также ищем IBAN если есть
        m_iban = re.search(r"Номер счета[:\s]+(KZ\w+)", first_text)
        if m_iban:
            iban = m_iban.group(1).strip()
            account = IBAN_MAP.get(iban, account)

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue
                    date_cell = str(row[0] or "").strip()
                    amount_cell = str(row[1] or "").strip()
                    desc_cell = str(row[2] or "").strip() if len(row) > 2 else ""
                    if len(row) > 3 and row[3]:
                        desc_cell = desc_cell + " " + str(row[3]).strip()

                    if not re.match(r"\d{2}\.\d{2}\.\d{2,4}", date_cell):
                        continue

                    amount_clean = amount_cell.replace(" ", "").replace("₸", "").replace("\xa0", "")
                    amount_clean = amount_clean.replace(",", ".")
                    sign = 1
                    if amount_clean.startswith("-"):
                        sign = -1
                        amount_clean = amount_clean[1:]
                    elif amount_clean.startswith("+"):
                        amount_clean = amount_clean[1:]

                    try:
                        amount = sign * float(amount_clean)
                    except:
                        continue

                    date_str = format_date(date_cell)
                    month = get_month(desc_cell, date_str)
                    year = date_str[-4:] if date_str else ""
                    week = get_week(date_str)

                    rows.append([year, str(month), str(week), date_str,
                                 fmt_amount(amount), str(month), account,
                                 get_article(desc_cell, amount), desc_cell, "", ""])

    return rows, account, closing_balance

# ============ PDF BCC ============
def process_bcc_pdf(file_bytes):
    rows = []
    iban = ""
    account = "БЦК Серик ИП"
    closing_balance = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        first_text = pdf.pages[0].extract_text() or ""

        m = re.search(r"ЖСК\s*/\s*ИИК\s*:\s*(KZ\w+)", first_text)
        if m:
            iban = m.group(1).strip()
            account = IBAN_MAP.get(iban, iban)

        # Ищем исходящее сальдо во всём тексте
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

        m_bal = re.search(r"[Ии]сходящее сальдо[:\s]*([\d\s,]+)", full_text)
        if m_bal:
            closing_balance = parse_num(m_bal.group(1))

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 12:
                        continue

                    date_cell = str(row[1] or "").replace("\n", "").strip()
                    debit_cell = str(row[7] or "").strip()
                    credit_cell = str(row[8] or "").strip()
                    desc_cell = str(row[11] or "").replace("\n", " ").strip()

                    date_m = re.search(r"\d{2}\.\d{2}\.\d{4}", date_cell)
                    if not date_m:
                        continue

                    date_str = format_date(date_m.group())
                    debit = parse_num(debit_cell)
                    credit = parse_num(credit_cell)

                    if debit > 0:
                        amount = -debit
                    elif credit > 0:
                        amount = credit
                    else:
                        continue

                    month = get_month(desc_cell, date_str)
                    year = date_str[-4:] if date_str else ""
                    week = get_week(date_str)

                    rows.append([year, str(month), str(week), date_str,
                                 fmt_amount(amount), str(month), account,
                                 get_article(desc_cell, amount), desc_cell, "", ""])

    return rows, account, closing_balance

# ============ HANDLER ============
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        await update.message.reply_text("Отправьте файл выписки (.xlsx или .pdf)")
        return

    fname = (doc.file_name or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".pdf")):
        await update.message.reply_text("Поддерживаются только .xlsx и .pdf файлы")
        return

    await update.message.reply_text("⏳ Обрабатываю выписку...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())

        if fname.endswith(".xlsx"):
            rows, account, closing_balance = process_xlsx(file_bytes)
        else:
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                first_page_text = pdf.pages[0].extract_text() or ""

            if "Kaspi Gold" in first_page_text or "KZ97722C" in first_page_text:
                rows, account, closing_balance = process_kaspi_gold_pdf(file_bytes)
            elif "ЦентрКредит" in first_page_text or "KZ44856" in first_page_text:
                rows, account, closing_balance = process_bcc_pdf(file_bytes)
            else:
                rows, account, closing_balance = process_kaspi_gold_pdf(file_bytes)

        if not rows:
            await update.message.reply_text("❌ Операции не найдены в файле")
            return

        sheet = get_sheet()
        sheet.append_rows(rows, value_input_option="RAW")

        # Базовое сообщение
        msg = f"✅ Готово! Добавлено {len(rows)} строк\nСчет: {account}\n"

        # Сверка остатков
        if closing_balance is not None:
            dds_balance, bank_balance, error = check_balance(account, closing_balance)
            if error:
                msg += f"\n⚠️ Сверка: {error}"
            else:
                diff = round(bank_balance - dds_balance, 2)
                if abs(diff) < 1:
                    msg += f"\n✅ Остаток сходится: {bank_balance:,.2f} ₸"
                else:
                    msg += f"\n❌ Остаток НЕ сходится!\n"
                    msg += f"Банк: {bank_balance:,.2f} ₸\n"
                    msg += f"ДДС: {dds_balance:,.2f} ₸\n"
                    msg += f"Разница: {diff:,.2f} ₸"
        else:
            msg += "\n⚠️ Исходящий остаток не найден в файле"

        msg += f"\n\n🔗 {SPREADSHEET_URL}"

        await update.message.reply_text(msg)

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))
    logger.info("Bot started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
