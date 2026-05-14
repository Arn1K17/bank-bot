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
    "KZ448562204152575978": "БЦК ИП Серик",
}

def get_account_name(iban):
    return IBAN_MAP.get(iban, iban)

def format_date(val):
    s = str(val).strip()
    # dd.mm.yy -> dd/mm/20yy
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


def detect_pdf_bank(file_bytes):
    """Определяем банк по тексту первой страницы PDF"""
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            first_page_text = pdf.pages[0].extract_text() or ""
            if "Kaspi" in first_page_text or "kaspi" in first_page_text or "CASPKZKA" in first_page_text:
                return "kaspi_gold"
            if "ЦентрКредит" in first_page_text or "BCC" in first_page_text or "KCJBKZKX" in first_page_text:
                return "bcc"
    except:
        pass
    return "unknown"


def process_kaspi_gold_pdf(file_bytes):
    """
    Парсит PDF выписку Kaspi Gold.
    Формат таблицы: Дата | Сумма | Операция | Детали
    Пример строки: 12.05.26 | - 2 670,00 ₸ | Покупка | YANDEX.GO
    """
    rows = []
    account = "Арман каспи голд"

    # Пытаемся найти IBAN в тексте
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                iban_match = re.search(r"KZ\w{18}", text)
                if iban_match:
                    found_iban = iban_match.group(0)
                    if found_iban in IBAN_MAP:
                        account = IBAN_MAP[found_iban]
                    break
    except:
        pass

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue

                    # Ищем строки где первая колонка — дата вида dd.mm.yy
                    date_val = str(row[0]).strip() if row[0] else ""
                    date_match = re.match(r"(\d{1,2})\.(\d{2})\.(\d{2,4})$", date_val)
                    if not date_match:
                        continue

                    # Нормализуем дату
                    d, mo, y = date_match.group(1), date_match.group(2), date_match.group(3)
                    if len(y) == 2:
                        y = "20" + y
                    date_str = f"{int(d):02d}/{int(mo):02d}/{y}"

                    # Сумма во второй колонке: "- 2 670,00 ₸" или "+ 200 000,00 ₸"
                    amount_raw = str(row[1]).strip() if row[1] else ""
                    amount_raw = amount_raw.replace("₸", "").replace("\u20b8", "").strip()
                    amount_raw = re.sub(r"\s+", "", amount_raw)  # убираем пробелы внутри числа
                    # Знак: "-" или "+"
                    sign = 1
                    if amount_raw.startswith("-"):
                        sign = -1
                        amount_raw = amount_raw[1:]
                    elif amount_raw.startswith("+"):
                        amount_raw = amount_raw[1:]
                    amount_raw = amount_raw.replace(",", ".")
                    try:
                        amount = sign * float(amount_raw)
                    except:
                        continue

                    # Тип операции (3-я колонка) и детали (4-я колонка)
                    operation = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                    details = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    # Иногда детали разбиты на 2 строки внутри ячейки
                    details = details.replace("\n", " ").strip()

                    desc = f"{operation}: {details}" if details else operation

                    month = int(mo)
                    year_str = y
                    week = get_week_number(date_str)
                    article = get_article(desc, amount)
                    amount_fmt = format_amount(amount)

                    rows.append([year_str, str(month), str(week), date_str, amount_fmt,
                                 str(month), account, article, desc, "", ""])

    return rows


def process_bcc_pdf(file_bytes):
    """
    Парсит PDF выписку BCC Business (Банк ЦентрКредит).
    Сложная таблица: №, Дата, БИК, ИИК, ИИН отправителя, Корреспондент,
                     ИИН получателя, Дебет, Кредит, КНП, Банк, Назначение платежа
    """
    rows = []
    account = "БЦК ИП Серик"

    # Пытаемся найти ИИК клиента
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                iban_match = re.search(r"KZ\d{2}\w{14,16}", text)
                if iban_match:
                    found_iban = iban_match.group(0)
                    if found_iban in IBAN_MAP:
                        account = IBAN_MAP[found_iban]
                    break
    except:
        pass

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 8:
                        continue

                    # В BCC таблице дата во 2-й колонке (индекс 1)
                    date_val = str(row[1]).strip() if row[1] else ""
                    # Убираем время если есть: "10.05.2026 02:44" -> "10.05.2026"
                    date_val = date_val.split()[0] if date_val else ""
                    date_match = re.match(r"(\d{1,2})\.(\d{2})\.(\d{4})$", date_val)
                    if not date_match:
                        continue

                    d, mo, y = date_match.group(1), date_match.group(2), date_match.group(3)
                    date_str = f"{int(d):02d}/{int(mo):02d}/{y}"

                    # Дебет — индекс 7, Кредит — индекс 8
                    debit_raw = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                    credit_raw = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                    desc = str(row[-1]).strip() if row[-1] else ""
                    desc = desc.replace("\n", " ").strip()

                    amount = None
                    def parse_bcc_amount(s):
                        s = re.sub(r"\s+", "", s).replace(",", ".")
                        try:
                            return float(s)
                        except:
                            return None

                    if debit_raw and re.search(r"\d", debit_raw):
                        val = parse_bcc_amount(debit_raw)
                        if val:
                            amount = -val  # расход
                    elif credit_raw and re.search(r"\d", credit_raw):
                        val = parse_bcc_amount(credit_raw)
                        if val:
                            amount = val  # приход

                    if amount is None:
                        continue

                    month = int(mo)
                    week = get_week_number(date_str)
                    article = get_article(desc, amount)
                    amount_fmt = format_amount(amount)

                    rows.append([y, str(month), str(week), date_str, amount_fmt,
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
        await update.message.reply_text(
            "Пожалуйста, отправьте файл выписки (.xlsx или .pdf)"
        )
        return

    fname = (doc.file_name or "").lower()

    if not (fname.endswith(".xlsx") or fname.endswith(".pdf")):
        await update.message.reply_text(
            "Поддерживаются только файлы .xlsx и .pdf"
        )
        return

    await update.message.reply_text("⏳ Обрабатываю выписку...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())

        # XLSX
        if fname.endswith(".xlsx"):
            rows = process_kaspi_xlsx(file_bytes)

        # PDF
        else:
            bank = detect_pdf_bank(file_bytes)

            logger.info(
                f"Detected bank: {bank} for file: {fname}"
            )

            # KASPI GOLD
            if bank == "kaspi_gold":

                rows = process_kaspi_gold_pdf(file_bytes)

            # BCC ДИАГНОСТИКА
elif bank == "bcc":
    rows = process_bcc_pdf(file_bytes)

            # UNKNOWN PDF
            else:

                rows = process_kaspi_gold_pdf(file_bytes)

                if not rows:
                    rows = process_bcc_pdf(file_bytes)

        # Нет операций
        if not rows:

            await update.message.reply_text(
                "❌ Не удалось найти операции в файле."
            )

            return

        # Загружаем в Google Sheets
        push_to_sheets(rows)

        await update.message.reply_text(
            f"✅ Готово! Добавлено "
            f"{len(rows)} строк в Google Sheets."
        )

    except Exception as e:

        logger.error(
            f"Error: {e}",
            exc_info=True
        )

        await update.message.reply_text(
            f"❌ Ошибка: {str(e)}"
        )

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    logger.info("Bot started!")
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    app.run_polling()

if __name__ == "__main__":
    main()
