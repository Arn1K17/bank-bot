import os
import re
import json
import logging
import time
import requests
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
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
WEBHOOK_URL = "https://bank-bot-w89l.onrender.com"
PORT = int(os.getenv("PORT", "10000"))

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
    "KZ448562204152575978": "БЦК Ип Серик",
    "KZ03601A231012849031": "Народный банк Ип Серик",
    "KZ438562203234869084": "БЦК Доллары ТОО",
    "KZ19722RU00001041014": "Депозит Каспи Ип Серик",
    "KZ31722RD00029080091": "Депозит Каспи кз",
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
        return val.strftime("%m/%d/%Y")
    s = str(val).replace("\n", "").strip()
    s = re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?$', '', s).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', s)
    if m:
        part1, part2, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if part1 > 12:
            return f"{part2:02d}/{part1:02d}/{y}"
        if part2 > 12:
            return f"{part1:02d}/{part2:02d}/{y}"
        return f"{part1:02d}/{part2:02d}/{y}"
    m = re.search(r"(\d{1,2})[.\-](\d{1,2})[.\-](\d{2,4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if len(y) == 2:
            y = "20" + y
        return f"{mo:02d}/{d:02d}/{y}"
    return s

def get_year_from_date(date_str):
    try:
        return date_str.split("/")[2]
    except:
        return ""

def get_month_from_date(date_str):
    try:
        return int(date_str.split("/")[0])
    except:
        return ""

def get_month_nachislenia(desc, date_str):
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
        return int(date_str.split("/")[0])
    except:
        return ""

def get_week(date_str):
    try:
        return datetime.strptime(date_str, "%m/%d/%Y").isocalendar()[1]
    except:
        return ""

def get_article(desc, amount):
    if not desc:
        return ""
    d = str(desc)
    for kw in ["Оплата за услуги операций по картам Kaspi Gold",
               "Оплата рекламных услуг","Оплата за услуги процессинга Без НДС",
               "Оплата услуги по обработке данных","Оплата за информационно-технологические услуги",
               "Бонусы за отзыв клиенту","Оплата услуг по обработке данных, связанных с доставкой",
               "Погашение комиссии за ведение счета",
               "Бонусы клиенту от продаж за"]:
        if kw.lower() in d.lower():
            return "Комиссия за эквайринг"
    if "Расчеты РЕСПУБЛИКИ по карточкам за" in d:
        return "Оплата от покупателя, выручка"
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

def parse_справка_num(s):
    s = str(s or "").strip().replace("\xa0", "").replace(" ", "")
    comma_count = s.count(",")
    dot_count = s.count(".")
    if comma_count > 1:
        s = s.replace(",", "")
    elif comma_count == 1 and dot_count == 0:
        parts = s.split(",")
        if len(parts[1]) == 3 and len(parts[0]) <= 3:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif dot_count > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except:
        return None

def parse_amount_from_registry(s):
    s = str(s or "").strip().replace(" ", "").replace("\xa0", "")
    if not s:
        return None
    dot_count = s.count(".")
    comma_count = s.count(",")
    if dot_count >= 1 and comma_count >= 1:
        s = s.replace(",", "")
    elif comma_count > 1:
        s = s.replace(",", "")
    elif comma_count == 1 and dot_count == 0:
        parts = s.lstrip("-").split(",")
        if len(parts) == 2 and len(parts[1]) == 3:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif dot_count > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except:
        return None

def make_row(date_str, amount, account, desc, supplier=""):
    year = get_year_from_date(date_str)
    month_oplaty = get_month_from_date(date_str)
    week = get_week(date_str)
    month_nachislenia = get_month_nachislenia(desc, date_str)
    return [
        year,
        str(month_oplaty),
        str(week),
        date_str,
        fmt_amount(amount),
        str(month_nachislenia),
        account,
        get_article(desc, amount),
        desc,
        "",
        str(supplier) if supplier else "",
    ]

# ============ ДЕДУПЛИКАЦИЯ ============

def _normalize_date(date_str: str) -> str:
    s = str(date_str).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%m/%d/%Y")
        except:
            pass
    parts = s.split("/")
    if len(parts) == 3:
        try:
            return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"
        except:
            pass
    return s

def _strip_doc_num(desc: str) -> str:
    return re.sub(r'^\d{7,12}\s+', '', str(desc).strip())

def _normalize_desc(desc: str) -> str:
    d = _strip_doc_num(desc)
    d = re.sub(r'\s+', ' ', d.replace("\n", " ").replace("\r", " ")).strip().lower()
    return d

def _normalize_amount(amount) -> str:
    amt_str = str(amount).strip().replace("\xa0", "").replace(" ", "").replace(",", "")
    try:
        return str(int(round(float(amt_str))))
    except:
        return amt_str

def _normalize_amount_exact(amount) -> str:
    amt_str = str(amount).strip().replace("\xa0", "").replace(" ", "").replace(",", "")
    try:
        return f"{float(amt_str):.2f}"
    except:
        return amt_str

def _extract_doc_num(desc_clean: str) -> str:
    desc_lower = desc_clean.lower()
    m = re.search(r'00ubs(\d+)', desc_lower)
    if m:
        return "ubs-" + m.group(1)
    m = re.match(r'^(g2-\d+)', desc_lower)
    if m:
        return m.group(1)
    m = re.match(r'^(nt-?\d+)', desc_lower)
    if m:
        return m.group(1)
    m = re.search(r'референс\s+(\d{8,12})', desc_lower)
    if m:
        return "ref-" + m.group(1)
    m = re.match(r'^(\d{7,12})\b', desc_clean)
    if m:
        return m.group(1)
    m = re.search(r'(?<!\d)(\d{7,12})(?!\d)', desc_clean)
    if m:
        return m.group(1)
    return ""

def _build_all_keys(date, amount, account, desc):
    nd = _normalize_date(str(date).strip())
    acc = str(account).strip()
    amt = _normalize_amount(amount)
    amt_exact = _normalize_amount_exact(amount)
    desc_raw = re.sub(r'\s+', ' ', str(desc).replace("\n", " ").replace("\r", " ")).strip()
    desc_norm = _normalize_desc(desc_raw)
    desc_short = desc_norm[:50]
    doc_num = _extract_doc_num(desc_raw)
    keys = []
    if doc_num:
        keys.append(("doc", nd, acc, doc_num))
    keys.append(("exact", nd, acc, amt_exact))
    keys.append(("amt", nd, acc, amt))
    if desc_short:
        keys.append(("text", nd, acc, desc_short))
    if desc_short:
        keys.append(("amt+text", nd, acc, amt, desc_short))
    return keys

def is_duplicate(r, existing_keys: set) -> bool:
    keys = _build_all_keys(r[3], r[4], r[6], r[8])
    for k in keys:
        if k in existing_keys:
            return True
    return False

def build_existing_keys(existing_data) -> tuple:
    key_to_row = {}
    for i, row in enumerate(existing_data[1:], start=2):
        if len(row) < 9:
            continue
        if not any(str(c).strip() for c in row):
            continue
        date_val = str(row[3]).strip()
        amount_val = str(row[4]).strip().replace(",", "").replace(" ", "").replace("\xa0", "")
        account_val = str(row[6]).strip()
        desc_val = str(row[8]).strip()
        if not date_val or not amount_val or not account_val:
            continue
        keys = _build_all_keys(date_val, amount_val, account_val, desc_val)
        for k in keys:
            if k not in key_to_row:
                key_to_row[k] = i
    return key_to_row, set(key_to_row.keys())

def find_existing_row(r, key_to_row: dict):
    keys = _build_all_keys(r[3], r[4], r[6], r[8])
    for k in keys:
        if k in key_to_row:
            return key_to_row[k]
    return None

def append_rows_from_col_a(sheet, rows):
    all_vals = sheet.get_all_values()
    last_row = 1
    for i, row in enumerate(all_vals, start=1):
        col_a = str(row[0]).strip() if len(row) > 0 else ""
        col_d = str(row[3]).strip() if len(row) > 3 else ""
        has_year = bool(re.match(r'^\d{4}$', col_a))
        has_date = bool(re.match(r'^\d{2}/\d{2}/\d{4}$', col_d))
        if has_year or has_date:
            last_row = i
    start_row = last_row + 1
    if not rows:
        return start_row
    end_col = "K"
    range_name = f"A{start_row}:{end_col}{start_row + len(rows) - 1}"
    sheet.update(range_name=range_name, values=rows, value_input_option="USER_ENTERED")
    return start_row

# ============ GROQ AI ============
def get_sheets_data_for_ai():
    try:
        spreadsheet = get_spreadsheet()
        реестр = spreadsheet.worksheet(SHEET_NAME)
        data = реестр.get_all_values()
        initial_balances = {}
        try:
            справка = spreadsheet.worksheet("Счета2026(Справка)")
            справка_data = справка.get_all_values()
            for row in справка_data:
                if row and row[0].strip() and len(row) > 1:
                    name = row[0].strip()
                    bal = parse_справка_num(row[1])
                    if bal is not None:
                        initial_balances[name] = bal
        except:
            pass
        if len(data) < 2:
            return "Данных в таблице пока нет."
        rows = data[1:]
        month_totals = {}
        account_totals = {}
        article_totals = {}
        for row in rows:
            try:
                month = row[1] if len(row) > 1 else ""
                amount_str = row[4] if len(row) > 4 else "0"
                account = row[6] if len(row) > 6 else ""
                article = row[7] if len(row) > 7 else ""
                amount = parse_amount_from_registry(amount_str)
                if amount is None:
                    continue
                if month:
                    month_totals[month] = month_totals.get(month, 0) + amount
                if account:
                    account_totals[account] = account_totals.get(account, 0) + amount
                if article:
                    article_totals[article] = article_totals.get(article, 0) + amount
            except:
                continue
        month_names = {
            "1":"Январь","2":"Февраль","3":"Март","4":"Апрель",
            "5":"Май","6":"Июнь","7":"Июль","8":"Август",
            "9":"Сентябрь","10":"Октябрь","11":"Ноябрь","12":"Декабрь"
        }
        summary = f"Всего строк в реестре: {len(rows)}\n\n"
        summary += "ОБОРОТЫ ПО МЕСЯЦАМ:\n"
        for m in sorted(month_totals.keys(), key=lambda x: int(x) if x.isdigit() else 99):
            name = month_names.get(m, f"Месяц {m}")
            summary += f"  {name}: {month_totals[m]:,.0f} ₸\n"
        summary += "\nТЕКУЩИЕ ОСТАТКИ ПО КАЖДОМУ СЧЕТУ:\n"
        total_all = 0.0
        for acc, ops_total in sorted(account_totals.items(), key=lambda x: x[0]):
            initial = 0.0
            best_score = 0.0
            for name, bal in initial_balances.items():
                if name.lower() == acc.lower():
                    initial = bal
                    break
                score = _account_similarity(acc, name)
                if score > best_score and score >= 0.7:
                    best_score = score
                    initial = bal
            current = initial + ops_total
            total_all += current
            summary += f"  {acc}: {current:,.0f} ₸  (нач.остаток: {initial:,.0f} + обороты: {ops_total:,.0f})\n"
        summary += f"  ИТОГО НА ВСЕХ СЧЕТАХ: {total_all:,.0f} ₸\n"
        summary += "\nОБОРОТЫ ПО СЧЕТАМ (только операции без начального остатка):\n"
        for acc, ops_total in sorted(account_totals.items(), key=lambda x: -abs(x[1])):
            summary += f"  {acc}: {ops_total:,.0f} ₸\n"
        summary += "\nПО СТАТЬЯМ:\n"
        for art, total in sorted(article_totals.items(), key=lambda x: -abs(x[1])):
            if art:
                summary += f"  {art}: {total:,.0f} ₸\n"
        return summary
    except Exception as e:
        return f"Ошибка получения данных: {e}"

def ask_ai(question: str) -> str:
    try:
        sheets_data = get_sheets_data_for_ai()
        today = datetime.now().strftime("%d.%m.%Y")
        prompt = f"""Ты финансовый помощник компании. Сегодня {today}.

{sheets_data}

ВАЖНЫЕ ПРАВИЛА:
- Отвечай ТОЛЬКО на русском языке
- Отвечай КОРОТКО — только финальный ответ, без рассуждений
- НЕ показывай свои вычисления, мысли, списки промежуточных шагов
- НЕ повторяй данные реестра целиком
- Если спрашивают остаток на счетах — дай цифры по каждому счёту и итог
- Если данных недостаточно — скажи об этом коротко

Вопрос пользователя: {question}

Дай ТОЛЬКО финальный ответ на русском языке. Никаких рассуждений."""

        headers = {
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": "llama-3.3-70b-versatile",
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 1000,
            "temperature": 0.3
        }
        resp = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers=headers, json=payload, timeout=30
        )
        resp.raise_for_status()
        result = resp.json()
        text = result["choices"][0]["message"].get("content") or "Нет ответа"
        return text.strip()
    except Exception as e:
        logger.error(f"Groq error: {e}")
        return f"❌ Ошибка ИИ: {str(e)}"

# ============ СВЕРКА ОСТАТКОВ ============
ACCOUNT_SYNONYMS = {
    "халык":       ["халык", "народный", "halyk", "hsbk"],
    "каспи":       ["каспи", "kaspi", "каспий"],
    "бцк":         ["бцк", "центркредит", "bcc", "kcjb"],
    "серик":       ["серик", "серік"],
    "имангазиева": ["имангазиева"],
    "орынбаева":   ["орынбаева"],
    "дильназ":     ["дильназ"],
    "коко":        ["коко", "сулейменов"],
    "арман":       ["арман"],
    "голд":        ["голд", "gold"],
    "тоо":         ["тоо"],
    "ип":          ["ип", "ip"],
    "депозит":     ["депозит", "deposit"],
    "usd":         ["usd", "доллар"],
}
BANK_CONFLICT_GROUPS = [
    {"халык", "народный"},
    {"каспи"},
    {"бцк", "центркредит"},
    {"втб"},
]

def _normalize_tokens(name: str) -> set:
    name = name.lower().strip()
    name = re.sub(r'[«»"\'(),.\-]', " ", name)
    words = name.split()
    tokens = set()
    for word in words:
        tokens.add(word)
        for canon, synonyms in ACCOUNT_SYNONYMS.items():
            if word in synonyms:
                tokens.add(canon)
                break
    return tokens

def _account_similarity(name_a: str, name_b: str) -> float:
    ta = _normalize_tokens(name_a)
    tb = _normalize_tokens(name_b)
    if not ta or not tb:
        return 0.0
    score = len(ta & tb) / max(len(ta), len(tb))
    for group in BANK_CONFLICT_GROUPS:
        a_in = bool(ta & group)
        b_in = bool(tb & group)
        a_other = any(ta & g for g in BANK_CONFLICT_GROUPS if g != group)
        b_other = any(tb & g for g in BANK_CONFLICT_GROUPS if g != group)
        if a_other and b_in and not (ta & group):
            score *= 0.3
        if b_other and a_in and not (tb & group):
            score *= 0.3
    return score

def find_account_in_справка(account_name: str, справка_data: list):
    search = account_name.strip().lower()
    for row in справка_data:
        if not row or not row[0].strip():
            continue
        candidate = row[0].strip()
        if candidate.lower() == search:
            balance = parse_справка_num(row[1] if len(row) > 1 else "")
            if balance is not None:
                return candidate, balance
    best_name = None
    best_balance = None
    best_score = 0.0
    for row in справка_data:
        if not row or not row[0].strip():
            continue
        candidate = row[0].strip()
        score = _account_similarity(account_name, candidate)
        if score > best_score:
            best_score = score
            best_name = candidate
            best_balance = parse_справка_num(row[1] if len(row) > 1 else "")
    if best_score >= 0.4 and best_balance is not None:
        return best_name, best_balance
    return None, None

def _clean_name_for_match(s):
    import unicodedata
    s = str(s)
    s = s.replace("\xa0", " ").replace("\u200b", "").replace("\n", " ").replace("\r", " ")
    s = unicodedata.normalize("NFKC", s)
    s = s.strip("'\"")
    words = s.split()
    return " ".join(words).lower()

def _matches_account_strict(row_acc, target):
    r = _clean_name_for_match(row_acc)
    t = _clean_name_for_match(target)
    return r == t

def _matches_account(row_acc, target):
    r = _clean_name_for_match(row_acc)
    t = _clean_name_for_match(target)
    if r == t:
        return True
    if _account_similarity(r, t) >= 0.80:
        return True
    return False

def get_account_balance(account_name: str):
    spreadsheet = get_spreadsheet()
    initial = 0.0
    try:
        справка = spreadsheet.worksheet("Счета2026(Справка)")
        справка_data = справка.get_all_values()
        _, bal = find_account_in_справка(account_name, справка_data)
        if bal is not None:
            initial = bal
    except Exception as e:
        logger.warning(f"get_account_balance: справка error: {e}")

    реестр = spreadsheet.worksheet(SHEET_NAME)
    реестр_data = реестр.get_all_values()
    ops_total = 0.0
    ops_count = 0
    for row in реестр_data[1:]:
        acc_val = str(row[6]).strip() if len(row) > 6 else ""
        amt_val = str(row[4]).strip() if len(row) > 4 else ""
        if not acc_val or not amt_val:
            continue
        if _matches_account_strict(acc_val, account_name):
            parsed = parse_amount_from_registry(amt_val)
            if parsed is not None:
                ops_total += parsed
                ops_count += 1

    dds = round(initial + ops_total, 2)
    return initial, round(ops_total, 2), dds, ops_count

def build_balance_msg(account, bank_closing_balance):
    if bank_closing_balance is None:
        return "\n⚠️ Исходящий остаток не найден в файле"
    try:
        initial, ops_total, dds, ops_count = get_account_balance(account)
        bank_balance = round(bank_closing_balance, 2)
        diff = round(bank_balance - dds, 2)
        msg = ""
        if abs(diff) < 1:
            msg += f"\n✅ Остаток сходится: {bank_balance:,.2f} ₸"
        else:
            msg += f"\n❌ Остаток НЕ сходится!\n"
            msg += f"  Банк: {bank_balance:,.2f} ₸\n"
            msg += f"  ДДС:  {dds:,.2f} ₸\n"
            msg += f"  Разница: {diff:,.2f} ₸"
        msg += f"\n\n📊 Расчёт ДДС:\n"
        msg += f"  Нач. остаток (Справка): {initial:,.2f} ₸\n"
        msg += f"  + Операции ({ops_count} строк): {ops_total:,.2f} ₸\n"
        msg += f"  = Итого ДДС: {dds:,.2f} ₸\n"
        msg += f"\n🏦 Банк (исходящий): {bank_balance:,.2f} ₸"
        if abs(diff) >= 1:
            msg += f"\n\n🔎 Требуется проверка операций и дублей."
        return msg
    except Exception as e:
        return f"\n⚠️ Ошибка сверки: {e}"

def get_account_rows(account_name):
    spreadsheet = get_spreadsheet()
    реестр = spreadsheet.worksheet(SHEET_NAME)
    реестр_data = реестр.get_all_values()
    matched = []
    for i, row in enumerate(реестр_data[1:], start=2):
        acc_val = row[6].strip() if len(row) > 6 else ""
        if acc_val and _matches_account(acc_val, account_name):
            matched.append((i, row))
    return matched

# ============ XLSX ============
def process_xlsx(file_bytes):
    rows = []
    closing_balance = None
    opening_balance = None

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    iban = str(cell_val(ws.cell(row=3, column=3))).strip()
    account = IBAN_MAP.get(iban, iban)

    opening_val = cell_val(ws.cell(row=9, column=3))
    try:
        opening_balance = float(str(opening_val).replace(" ", "").replace(",", "."))
    except:
        opening_balance = None

    closing_val = cell_val(ws.cell(row=10, column=3))
    try:
        closing_balance = float(str(closing_val).replace(" ", "").replace(",", "."))
    except:
        pass

    if not closing_balance:
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = str(cell_val(ws.cell(row=row_idx, column=col_idx))).lower()
                if "исходящ" in cell and "сальдо" in cell:
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

    data_start = 14
    for row_idx in range(12, min(20, ws.max_row + 1)):
        val = cell_val(ws.cell(row=row_idx, column=2))
        if val and re.search(r"\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}", str(val)):
            data_start = row_idx
            break

    for row_idx in range(data_start, ws.max_row + 1):
        date_val = cell_val(ws.cell(row=row_idx, column=2))
        debit = cell_val(ws.cell(row=row_idx, column=3))
        credit = cell_val(ws.cell(row=row_idx, column=4))
        supplier = str(cell_val(ws.cell(row=row_idx, column=5)) or "")
        desc_raw = str(cell_val(ws.cell(row=row_idx, column=9)) or "")
        doc_num_val = str(cell_val(ws.cell(row=row_idx, column=1)) or "").strip()
        if doc_num_val and re.match(r'^\d{7,}$', doc_num_val):
            desc = f"{doc_num_val} {desc_raw}".strip()
        else:
            desc = desc_raw

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
        rows.append(make_row(date_str, amount, account, desc, supplier))

    return rows, account, closing_balance, opening_balance

# ============ PDF Kaspi Gold ============
def _parse_kaspi_text_lines(all_text, account):
    """
    Текстовый парсер для Каспи Gold/депозит выписок.
    Формат строки: DD.MM.YY [+/-] сумма ₸ Операция Детали
    Строки вида "(- 66,00 USD)" добавляются в описание предыдущей операции.
    """
    rows = []
    lines = all_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(
            r'^(\d{2}\.\d{2}\.\d{2,4})\s+([+\-])\s*([\d\s]+,\d{2})\s*₸\s+(.+)$',
            line
        )
        if m:
            date_raw = m.group(1)
            sign = 1 if m.group(2) == '+' else -1
            amt_str = m.group(3).replace(" ", "").replace(",", ".")
            desc = m.group(4).strip()

            # Смотрим следующие строки — могут быть продолжение описания или USD
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                if not nxt:
                    j += 1
                    continue
                # Если следующая строка — новая операция, останавливаемся
                if re.match(r'^\d{2}\.\d{2}\.\d{2,4}', nxt):
                    break
                # Строка с валютой вида "(- 66,00 USD)" — добавляем в описание
                if re.match(r'^\([-+]?\s*[\d\s]+[,.][\d]+\s*(USD|EUR|RUB|CNY)\)', nxt, re.IGNORECASE):
                    desc = f"{desc} {nxt}".strip()
                    j += 1
                    continue
                # Строка начинается с "(" но не валюта — пропускаем
                if nxt.startswith("("):
                    j += 1
                    continue
                # Обычное продолжение описания (напр. "своего счета")
                desc = f"{desc} {nxt}".strip()
                j += 1
                break  # берём только одну строку продолжения

            i = j  # прыгаем к следующей необработанной строке

            try:
                amount = sign * float(amt_str)
                date_str = format_date(date_raw)
                rows.append(make_row(date_str, amount, account, desc))
            except:
                pass
        else:
            i += 1
    return rows

def process_kaspi_gold_pdf(file_bytes):
    rows = []
    account = "Арман каспи голд"
    closing_balance = None
    opening_balance = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        all_text = ""
        for p in pdf.pages:
            all_text += (p.extract_text() or "") + "\n"

        is_deposit = "По Депозиту" in first_text or "На Депозите" in first_text or "KZ19722RU" in first_text

        # Определяем счёт по IBAN
        m_iban = re.search(r"Номер счета[:\s]+(KZ\w+)", all_text)
        if not m_iban:
            m_iban = re.search(r"счет[оом]*\s+(KZ\w+)", all_text, re.IGNORECASE)
        if m_iban:
            iban = m_iban.group(1).strip()
            account = IBAN_MAP.get(iban, account)
        elif "KZ97722C000015235365" in all_text:
            account = IBAN_MAP.get("KZ97722C000015235365", account)
        elif "KZ19722RU00001041014" in all_text:
            account = IBAN_MAP.get("KZ19722RU00001041014", "Депозит Каспи Ип Серик")

        if not is_deposit:
            # Остаток: "Доступно на DD.MM.YY: + сумма ₸"
            bal_matches = re.findall(
                r"Доступно на\s+(\d{2}\.\d{2}\.\d{2,4})[:\s]*[+\-]?\s*([\d\s]+,\d+)\s*₸",
                all_text
            )
            if len(bal_matches) >= 2:
                opening_balance = parse_num(bal_matches[0][1])
                closing_balance = parse_num(bal_matches[-1][1])
            elif len(bal_matches) == 1:
                closing_balance = parse_num(bal_matches[0][1])

            # Текстовый парсинг (основной для этого формата PDF)
            rows = _parse_kaspi_text_lines(all_text, account)

            # Fallback на таблицы если текстовый не нашёл строки
            if not rows:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if not row or len(row) < 3:
                                continue
                            date_cell = str(row[0] or "").strip()
                            amount_cell = str(row[1] or "").strip()
                            op_cell = str(row[2] or "").strip() if len(row) > 2 else ""
                            det_cell = str(row[3] or "").strip() if len(row) > 3 else ""
                            desc_cell = f"{op_cell} {det_cell}".strip()
                            if not re.match(r"\d{2}\.\d{2}\.\d{2,4}", date_cell):
                                continue
                            amount_clean = amount_cell.replace(" ", "").replace("₸", "").replace("\xa0", "").replace(",", ".")
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
                            rows.append(make_row(date_str, amount, account, desc_cell))

        else:
            # Депозит
            dep_matches = re.findall(
                r"На Депозите\s+\d{2}\.\d{2}\.\d{2,4}\s+([\d\s]+[,.][\d]+)\s*₸",
                all_text
            )
            if len(dep_matches) >= 2:
                opening_balance = parse_num(dep_matches[0])
                closing_balance = parse_num(dep_matches[-1])
            elif len(dep_matches) == 1:
                closing_balance = parse_num(dep_matches[0])

            rows = _parse_kaspi_text_lines(all_text, account)

    if opening_balance is None and closing_balance is not None and rows:
        total_ops = sum(float(r[4]) for r in rows)
        opening_balance = round(closing_balance - total_ops, 2)

    logger.info(f"Kaspi Gold: счет={account}, строк={len(rows)}, входящий={opening_balance}, исходящий={closing_balance}")
    return rows, account, closing_balance, opening_balance

# ============ PDF BCC ============
def process_bcc_pdf(file_bytes):
    rows = []
    iban = ""
    account = "БЦК Ип Серик"
    closing_balance = None
    opening_balance = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

        m = re.search(r"ЖСК\s*/\s*ИИК\s*:\s*(KZ\w+)", full_text)
        if not m:
            m = re.search(r"ЖСК\s*/\s*ИИК\s+(KZ\w+)", full_text)
        if not m:
            m = re.search(r"IBAN[:\s]+(KZ\w+)", full_text)
        if m:
            iban = m.group(1).strip()
            account = IBAN_MAP.get(iban, iban)

        m_open = re.search(r"[Кк]іріс [қк]алдық\s*/\s*[Вв]ходящий остаток[:\s]*([\d\s]+[,.][\d]+)", full_text)
        if not m_open:
            m_open = re.search(r"[Кк]іріс сальдо\s*/\s*[Вв]ходящее сальдо[:\s]*([\d\s]+[,.][\d]+)", full_text)
        if m_open:
            opening_balance = parse_num(m_open.group(1))

        m_bal = re.search(r"[Шш]ығыс сальдо\s*/\s*[Ии]сходящее сальдо[:\s]*([\d\s]+[,.][\d]+)", full_text)
        if not m_bal:
            m_bal = re.search(r"[Ии]сходящее сальдо[:\s]*([\d\s]+[,.][\d]+)", full_text)
        if not m_bal:
            m_bal = re.search(r"[Шш]ығыс сальдо[:\s]*([\d\s]+[,.][\d]+)", full_text)
        if m_bal:
            closing_balance = parse_num(m_bal.group(1))

        all_table_rows = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) >= 12:
                        all_table_rows.append(list(row))

        merged_rows = []
        i = 0
        while i < len(all_table_rows):
            row = all_table_rows[i]
            c0 = str(row[0] or "").replace("\n", "").strip()
            c1 = str(row[1] or "").replace("\n", "").strip()

            if "итого" in c0.lower() or "жиынтығы" in c0.lower():
                i += 1
                continue
            if "дата" in c1.lower() or "күні" in c1.lower() or "реттік" in c0.lower():
                i += 1
                continue

            if c0 == "NT-" and i + 1 < len(all_table_rows):
                next_row = all_table_rows[i + 1]
                nc0 = str(next_row[0] or "").replace("\n", "").strip()
                if re.match(r"^\d+$", nc0):
                    glued = []
                    for ci in range(12):
                        pv = str(row[ci] or "").replace("\n", " ").strip()
                        nv = str(next_row[ci] or "").replace("\n", " ").strip()
                        if ci == 0:
                            glued.append("NT-" + nc0)
                        elif ci == 1:
                            if re.search(r"\d{1,2}\.\d{2}\.\d{4}", pv):
                                glued.append(pv)
                            elif re.search(r"\d{1,2}\.\d{2}\.\d{4}", nv):
                                glued.append(nv)
                            else:
                                glued.append((pv + nv).strip())
                        elif ci in (7, 8):
                            combined = (pv + nv).replace(" ", "")
                            glued.append(combined)
                        else:
                            glued.append((pv + " " + nv).strip() if (pv and nv) else (pv or nv))
                    merged_rows.append(glued)
                    i += 2
                    continue

            merged_rows.append(row)
            i += 1

        for row in merged_rows:
            c0 = str(row[0] or "").replace("\n", "").strip()
            c1 = str(row[1] or "").replace("\n", "").strip()
            c7 = str(row[7] or "").replace("\n", " ").strip()
            c8 = str(row[8] or "").replace("\n", " ").strip()
            c11 = str(row[11] or "").replace("\n", " ").strip()

            # Убираем валютную часть "(- 66,00 USD)" — берём только тенге
            c7 = re.sub(r'\(.*?(?:USD|EUR|RUB|CNY).*?\)', '', c7, flags=re.IGNORECASE).strip()
            c8 = re.sub(r'\(.*?(?:USD|EUR|RUB|CNY).*?\)', '', c8, flags=re.IGNORECASE).strip()

            if "итого" in c1.lower() or "жиынтығы" in c1.lower():
                continue

            date_m = re.search(r"(\d{1,2})\.(\d{2})\.(\d{4})", c1)
            if not date_m:
                continue
            day = int(date_m.group(1))
            month = int(date_m.group(2))
            year = date_m.group(3)
            date_str = f"{month:02d}/{day:02d}/{year}"

            debit = parse_num(c7)
            credit = parse_num(c8)

            if debit > 0:
                amount = -debit
            elif credit > 0:
                amount = credit
            else:
                continue

            desc = f"{c0} {c11}".strip() if c0 else c11
            rows.append(make_row(date_str, amount, account, desc))

    logger.info(f"BCC: счет={account}, строк={len(rows)}, входящий={opening_balance}, исходящий={closing_balance}")
    return rows, account, closing_balance, opening_balance

# ============ PDF Halyk ============
def parse_kz_num(s):
    s = str(s or "").strip().replace(" ", "").replace("\xa0", "")
    s = re.sub(r",(\d{3})(?=[\d,.])", r"\1", s)
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0

def process_halyk_pdf(file_bytes):
    rows = []
    account = "Народный банк Ип Серик"
    closing_balance = None
    opening_balance = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

    m_iban = re.search(r"Счет\(Валюта\)[:\s]+(KZ[\w]+)", full_text)
    if m_iban:
        iban = m_iban.group(1).strip()
        account = IBAN_MAP.get(iban, account)

    m_open = re.search(r"[Вв]ходящий остаток[:\s]*([\d\s,]+\.\d{2})", full_text)
    if m_open:
        opening_balance = parse_kz_num(m_open.group(1))

    m_bal = re.search(r"[Ии]сходящий остаток[:\s]*([\d\s,]+\.\d{2})", full_text)
    if m_bal:
        closing_balance = parse_kz_num(m_bal.group(1))

    lines = full_text.split("\n")
    blocks = []
    current = []
    for line in lines:
        line = line.strip()
        if re.match(r"^\d{2}\.\d{2}\.\d{4}\s+\S+\s+[\d,]+\.\d{2}", line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        blocks.append(current)

    for block in blocks:
        first = block[0]
        full_desc = " ".join(block)
        m = re.match(r"^(\d{2}\.\d{2}\.\d{4})\s+\S+\s+([\d,]+\.\d{2})", first)
        if not m:
            continue
        date_raw = m.group(1)
        amount = parse_kz_num(m.group(2))
        desc_lower = full_desc.lower()
        if any(kw in desc_lower for kw in ["снятие", "комиссия", "перевод", "cmstake"]):
            amount = -amount
        d, mo, y = date_raw.split(".")
        date_str = f"{int(mo):02d}/{int(d):02d}/{y}"
        doc_num_m = re.match(r'^\d{2}\.\d{2}\.\d{4}\s+(\S+)\s+', first)
        doc_num = doc_num_m.group(1) if doc_num_m else ""
        ubs_m = re.search(r'00UBS(\d+)', full_desc)
        ubs_num = ("00UBS" + ubs_m.group(1)) if ubs_m else ""
        clean_desc = re.sub(r'^\d{2}\.\d{2}\.\d{4}\s+\S+\s+[\d,]+\.\d{2}\s*', '', full_desc).strip()
        if ubs_num:
            desc_with_docnum = f"{ubs_num} {doc_num} {clean_desc}".strip()
        elif doc_num:
            desc_with_docnum = f"{doc_num} {clean_desc}".strip()
        else:
            desc_with_docnum = clean_desc
        rows.append(make_row(date_str, amount, account, desc_with_docnum[:200]))

    return rows, account, closing_balance, opening_balance

# ============ ОПРЕДЕЛЕНИЕ ТИПА PDF ============
def detect_pdf_type(first_page_text):
    if "По Депозиту" in first_page_text or "На Депозите" in first_page_text or "KZ19722RU" in first_page_text:
        return "kaspi_deposit"
    if "СПРАВКА" in first_page_text and ("Kaspi Gold" in first_page_text or "KZ97722C" in first_page_text or "CASPKZKA" in first_page_text):
        return "kaspi_gold"
    if "Kaspi Gold" in first_page_text or "KZ97722C" in first_page_text:
        return "kaspi_gold"
    if "Kaspi Bank" in first_page_text or "CASPKZKA" in first_page_text:
        return "kaspi_gold"
    if "Народный Банк" in first_page_text or "Halyk" in first_page_text or "HSBKKZKX" in first_page_text:
        return "halyk"
    if ("ЦентрКредит" in first_page_text or "ЦентрКре" in first_page_text
            or "KCJBKZKX" in first_page_text
            or "KZ44856" in first_page_text or "KZ03856" in first_page_text
            or "KZ50856" in first_page_text or "KZ43856" in first_page_text):
        return "bcc"
    return "kaspi_gold"

# ============ HANDLER ============
async def handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text:
        question = update.message.text.strip()

        if question.startswith("/start"):
            await update.message.reply_text(
                "👋 Привет! Я бухгалтерский бот.\n\n"
                "📎 Отправьте файл выписки (.xlsx или .pdf) — загружу в таблицу.\n\n"
                "💬 Или задайте вопрос текстом, например:\n"
                "• Какая разница между апрелем и маем?\n"
                "• Сколько пришло за май?\n"
                "• Какой счёт имеет наибольший оборот?\n\n"
                "🔍 Команды:\n"
                "/rows <название счёта> — показать все строки по счёту\n"
                "Пример: /rows Каспи ТОО"
            )
            return

        if question.startswith("/rows"):
            acc = question[5:].strip()
            if not acc:
                await update.message.reply_text("Укажи название счёта после команды.\nПример: /rows Каспи ТОО")
                return

            await update.message.reply_text(f"🔍 Ищу строки для «{acc}»...")
            try:
                matched = get_account_rows(acc)
            except Exception as e:
                await update.message.reply_text(f"❌ Ошибка: {e}")
                return

            if not matched:
                await update.message.reply_text(f"❌ Строк по счёту «{acc}» не найдено.")
                return

            total = 0.0
            for _, row in matched:
                try:
                    val = parse_amount_from_registry(row[4])
                    if val is not None:
                        total += val
                except:
                    pass

            header = (
                f"📋 Счёт: {acc}\n"
                f"Найдено строк: {len(matched)}\n"
                f"Сумма операций: {total:,.2f} ₸\n"
                f"{'─'*35}\n"
            )

            chunk_size = 50
            chunks = [matched[i:i+chunk_size] for i in range(0, len(matched), chunk_size)]

            for part_idx, chunk in enumerate(chunks):
                lines = []
                for sheet_row, row in chunk:
                    date = row[3] if len(row) > 3 else ""
                    amount = row[4] if len(row) > 4 else ""
                    desc = row[8] if len(row) > 8 else ""
                    desc_short = (desc[:40] + "…") if len(desc) > 40 else desc
                    lines.append(f"#{sheet_row} | {date} | {amount} | {desc_short}")

                part_text = "\n".join(lines)
                if part_idx == 0:
                    msg = header + part_text
                else:
                    msg = f"(продолжение {part_idx+1}/{len(chunks)})\n" + part_text

                await update.message.reply_text(msg)
            return

        await update.message.reply_text("🤔 Думаю...")
        answer = ask_ai(question)
        await update.message.reply_text(answer)
        return

    doc = update.message.document
    if not doc:
        await update.message.reply_text("Отправьте файл выписки (.xlsx или .pdf) или задайте вопрос текстом.")
        return

    fname = (doc.file_name or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".pdf")):
        await update.message.reply_text("Поддерживаются только .xlsx и .pdf файлы")
        return

    await update.message.reply_text("⏳ Обрабатываю выписку...")

    try:
        file = await context.bot.get_file(doc.file_id)
        file_bytes = bytes(await file.download_as_bytearray())

        opening_balance = None

        if fname.endswith(".xlsx"):
            rows, account, closing_balance, opening_balance = process_xlsx(file_bytes)
        else:
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                first_page_text = pdf.pages[0].extract_text() or ""
            pdf_type = detect_pdf_type(first_page_text)
            logger.info(f"PDF тип: {pdf_type}")
            if pdf_type in ("kaspi_gold", "kaspi_deposit"):
                rows, account, closing_balance, opening_balance = process_kaspi_gold_pdf(file_bytes)
            elif pdf_type == "halyk":
                rows, account, closing_balance, opening_balance = process_halyk_pdf(file_bytes)
            elif pdf_type == "bcc":
                rows, account, closing_balance, opening_balance = process_bcc_pdf(file_bytes)
            else:
                rows, account, closing_balance, opening_balance = process_kaspi_gold_pdf(file_bytes)

        if not rows:
            await update.message.reply_text("❌ Операции не найдены в файле")
            return

        time.sleep(2)
        sheet = get_sheet()
        existing_data = sheet.get_all_values()

        key_to_row, existing_keys = build_existing_keys(existing_data)

        dupe_sheet_rows = []
        dupe_rows = []
        for r in rows:
            if is_duplicate(r, existing_keys):
                dupe_rows.append(r)
                sheet_row = find_existing_row(r, key_to_row)
                dupe_sheet_rows.append(sheet_row)
        dupes = len(dupe_rows)

        def format_row_ranges(row_nums):
            if not row_nums:
                return ""
            sorted_nums = sorted(n for n in row_nums if n is not None)
            if not sorted_nums:
                return ""
            ranges = []
            start = end = sorted_nums[0]
            for n in sorted_nums[1:]:
                if n == end + 1:
                    end = n
                else:
                    ranges.append(f"{start}-{end}" if start != end else str(start))
                    start = end = n
            ranges.append(f"{start}-{end}" if start != end else str(start))
            return ", ".join(ranges)

        all_vals_for_count = sheet.get_all_values()
        last_filled = 1
        for i, row in enumerate(all_vals_for_count, start=1):
            if any(str(c).strip() for c in row):
                last_filled = i
        next_row = last_filled + 1

        if dupes == len(rows):
            dupe_range = format_row_ranges(dupe_sheet_rows)
            msg = (
                f"⚠️ Этот файл уже был загружен ранее!\n"
                f"Все {len(rows)} строк уже есть в таблице.\n"
                f"Строки в таблице: {dupe_range}\n"
                f"Ничего не добавлено."
            )
            msg += build_balance_msg(account, closing_balance)
            msg += f"\n\n🔗 {SPREADSHEET_URL}"
            await update.message.reply_text(msg)
            return

        if dupes > 0:
            dupe_range = format_row_ranges(dupe_sheet_rows)
            new_rows = [r for r in rows if not is_duplicate(r, existing_keys)]
            new_range = format_row_ranges(list(range(next_row, next_row + len(new_rows))))
            rows = new_rows
            await update.message.reply_text(
                f"⚠️ {dupes} строк уже есть в таблице\n"
                f"Дубли в строках: {dupe_range}\n"
                f"Добавляю {len(rows)} новых → строки {new_range}"
            )

        rows.sort(key=lambda r: datetime.strptime(r[3], "%m/%d/%Y") if r[3] else datetime.min)

        actual_start = append_rows_from_col_a(sheet, rows)
        time.sleep(3)

        added_range = format_row_ranges(list(range(actual_start, actual_start + len(rows))))
        msg = f"✅ Готово! Добавлено {len(rows)} строк\nСчет: {account}\n📋 Строки: {added_range}\n"
        msg += build_balance_msg(account, closing_balance)
        msg += f"\n\n🔗 {SPREADSHEET_URL}"
        await update.message.reply_text(msg)

    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle))
    app.add_handler(MessageHandler(filters.TEXT, handle))
    logger.info("Bot started!")
    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        webhook_url=f"{WEBHOOK_URL}/webhook",
        url_path="webhook",
        drop_pending_updates=True,
    )

if __name__ == "__main__":
    main()
