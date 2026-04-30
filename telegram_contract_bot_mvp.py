import os
import re
import json
import base64
import shutil
import subprocess
from datetime import datetime
from collections import defaultdict

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    InputFile,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

from docxtpl import DocxTemplate
from openai import OpenAI

# =========================================================
# TOKENS
# =========================================================
import os

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL = "gpt-4.1-mini"

client = OpenAI(api_key=OPENAI_API_KEY)

# =========================================================
# PATHS
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TMP_DIR = os.path.join(BASE_DIR, "tmp")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TMP_DIR, exist_ok=True)

TEMPLATES = {
    "rent": os.path.join(TEMPLATES_DIR, "rent_template.docx"),
    "sale": os.path.join(TEMPLATES_DIR, "buy_template.docx"),
}

# =========================================================
# STATES
# =========================================================
STATE_WAIT_TYPE = "wait_type"

STATE_OWNER_MODE = "owner_mode"
STATE_OWNER_PHOTOS = "owner_photos"
STATE_OWNER_MANUAL = "owner_manual"
STATE_OWNER_REVIEW = "owner_review"

STATE_TENANT_MODE = "tenant_mode"
STATE_TENANT_PHOTOS = "tenant_photos"
STATE_TENANT_MANUAL = "tenant_manual"
STATE_TENANT_REVIEW = "tenant_review"

STATE_PROPERTY_MODE = "property_mode"
STATE_PROPERTY_PHOTOS = "property_photos"
STATE_PROPERTY_TEXT = "property_text"
STATE_PROPERTY_REVIEW = "property_review"

STATE_FINANCE = "finance"
STATE_DATES = "dates"

STATE_FINAL_REVIEW = "final_review"

# =========================================================
# FIELDS
# =========================================================
PERSON_FIELDS = [
    "full_name",
    "birth_date",
    "tax_id",
    "passport_number",
    "passport_issued_by",
    "passport_record",
    "passport_date",
    "address",
    "phone",
]

PROPERTY_FIELDS = [
    "address",
    "apartment_number",
    "total_area",
    "ownership_doc_number",
    "ownership_doc_date",
    "building_street",
    "cold_water_meter",
    "hot_water_meter",
    "heat_meter",
    "electricity_meter",
]

FIELD_LABELS = {
    "full_name": "ПІБ",
    "birth_date": "Дата народження",
    "tax_id": "ІПН",
    "passport_number": "Паспорт",
    "passport_issued_by": "Ким виданий",
    "passport_record": "Запис",
    "passport_date": "Дата видачі паспорта",
    "address": "Адреса",
    "phone": "Телефон",

    "apartment_number": "Квартира №",
    "total_area": "Площа",
    "ownership_doc_number": "Документ власності №",
    "ownership_doc_date": "Дата документа власності",
    "building_street": "Вулиця будинку",
    "cold_water_meter": "Холодна вода",
    "hot_water_meter": "Гаряча вода",
    "heat_meter": "Тепло",
    "electricity_meter": "Електрика",
}

# =========================================================
# SESSION STORAGE
# =========================================================
sessions = {}


def ensure_session(user_id: int):
    if user_id not in sessions:
        sessions[user_id] = {
            "contract_type": None,
            "state": STATE_WAIT_TYPE,

            "owner_images": [],
            "tenant_images": [],
            "property_images": [],

            "owner_data": {},
            "tenant_data": {},
            "property_data": {},
            "finance_data": {},
            "dates_data": {},

            "owner_ocr_results": [],
            "tenant_ocr_results": [],
            "property_ocr_results": [],

            "pending_conflicts": [],
            "current_conflict_idx": 0,
            "conflict_target": None,   # owner / tenant / property
            "custom_edit_target": None,  # owner / tenant / property / final
        }


def reset_session(user_id: int):
    if user_id in sessions:
        del sessions[user_id]
    ensure_session(user_id)


# =========================================================
# BASIC HELPERS
# =========================================================
def now_ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_json_loads(text: str):
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)


def normalize_spaces(text: str) -> str:
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(value):
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    if value.lower() in {"null", "none", "невідомо", "unknown", "не видно", "нечитабельно"}:
        return None
    return normalize_spaces(value)


# =========================================================
# NORMALIZATION
# =========================================================
MONTHS_UA = {
    "січня": "01",
    "лютого": "02",
    "березня": "03",
    "квітня": "04",
    "травня": "05",
    "червня": "06",
    "липня": "07",
    "серпня": "08",
    "вересня": "09",
    "жовтня": "10",
    "листопада": "11",
    "грудня": "12",
}


def normalize_date(value: str):
    value = normalize_text(value)
    if not value:
        return None

    value = value.replace(",", " ")
    value = normalize_spaces(value.lower().replace("року", "").replace("р.", ""))

    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", value)
    if m:
        d, mth, y = m.groups()
        return f"{int(d):02d}.{int(mth):02d}.{y}"

    m = re.match(r"^(\d{1,2})\s+([а-щьюяіїєґ]+)\s+(\d{4})$", value)
    if m:
        d, mon_word, y = m.groups()
        if mon_word in MONTHS_UA:
            return f"{int(d):02d}.{MONTHS_UA[mon_word]}.{y}"

    return value


def normalize_tax_id(value: str):
    value = normalize_text(value)
    if not value:
        return None
    digits = re.sub(r"\D", "", value)
    return digits if digits else value


def normalize_passport_number(value: str):
    value = normalize_text(value)
    if not value:
        return None
    value = value.upper().replace(" ", "")
    return value


def normalize_phone(value: str):
    value = normalize_text(value)
    if not value:
        return None
    digits = re.sub(r"\D", "", value)
    return digits if digits else value


def normalize_person_name(value: str):
    value = normalize_text(value)
    if not value:
        return None

    value = value.replace("’", "'")
    value = normalize_spaces(value)

    if value.isupper():
        value = value.title()

    return value


def canonical_compare_value(field: str, value: str):
    value = normalize_text(value)
    if not value:
        return None

    if field in {"birth_date", "passport_date", "ownership_doc_date"}:
        return normalize_date(value)

    if field == "tax_id":
        return normalize_tax_id(value)

    if field == "passport_number":
        return normalize_passport_number(value)

    if field == "phone":
        return normalize_phone(value)

    if field == "full_name":
        norm_name = normalize_person_name(value)
        if not norm_name:
            return None
        tokens = sorted(norm_name.lower().split())
        return " ".join(tokens)

    if field == "address":
        value = value.lower()
        value = value.replace("місто", "м.")
        value = value.replace("будинок", "буд.")
        value = value.replace("квартира", "кв.")
        value = value.replace("область", "обл.")
        value = value.replace("район", "р-н")
        value = re.sub(r"[,\.;:]+", " ", value)
        value = normalize_spaces(value)
        return value

    value = value.lower().replace("’", "'")
    value = normalize_spaces(value)
    return value


def display_value(field: str, value: str):
    value = normalize_text(value)
    if not value:
        return None

    if field in {"birth_date", "passport_date", "ownership_doc_date"}:
        return normalize_date(value)

    if field == "tax_id":
        return normalize_tax_id(value)

    if field == "passport_number":
        return normalize_passport_number(value)

    if field == "phone":
        return normalize_phone(value)

    if field == "full_name":
        return normalize_person_name(value)

    return value


# =========================================================
# SOFT MERGE
# =========================================================
def choose_best_value(field: str, candidates: list[dict]):
    filtered = []
    for c in candidates:
        raw = normalize_text(c.get("value"))
        if not raw:
            continue
        filtered.append({
            "raw": display_value(field, raw),
            "cmp": canonical_compare_value(field, raw),
            "confidence": c.get("confidence", 0),
        })

    if not filtered:
        return None, []

    grouped = defaultdict(list)
    for item in filtered:
        grouped[item["cmp"]].append(item)

    if len(grouped) == 1:
        items = next(iter(grouped.values()))
        best = sorted(items, key=lambda x: x["confidence"], reverse=True)[0]
        return best["raw"], []

    variants = []
    for _, items in grouped.items():
        best = sorted(items, key=lambda x: x["confidence"], reverse=True)[0]
        variants.append({
            "value": best["raw"],
            "confidence": best["confidence"],
            "count": len(items),
        })

    variants = sorted(variants, key=lambda x: (x["count"], x["confidence"]), reverse=True)
    return variants[0]["value"], variants


def merge_results(results: list[dict], fields: list[str]):
    merged = {}
    conflicts = []

    for field in fields:
        candidates = []

        for result in results:
            field_data = result.get(field)
            if isinstance(field_data, dict):
                value = field_data.get("value")
                confidence = field_data.get("confidence", 0)
            else:
                value = field_data
                confidence = 0.5 if value else 0

            if normalize_text(value):
                candidates.append({
                    "value": value,
                    "confidence": confidence
                })

        best_value, variants = choose_best_value(field, candidates)
        merged[field] = best_value

        if len(variants) > 1:
            conflicts.append({
                "field": field,
                "variants": variants[:3],
                "best_guess": best_value,
            })

    return merged, conflicts


# =========================================================
# UI HELPERS
# =========================================================
def progress_text(session: dict):
    current = session["state"]

    if current in [STATE_OWNER_MODE, STATE_OWNER_PHOTOS, STATE_OWNER_MANUAL, STATE_OWNER_REVIEW]:
        step = "1/5 — Наймодавець"
    elif current in [STATE_TENANT_MODE, STATE_TENANT_PHOTOS, STATE_TENANT_MANUAL, STATE_TENANT_REVIEW]:
        step = "2/5 — Винаймач"
    elif current in [STATE_PROPERTY_MODE, STATE_PROPERTY_PHOTOS, STATE_PROPERTY_TEXT, STATE_PROPERTY_REVIEW]:
        step = "3/5 — Об'єкт"
    elif current == STATE_FINANCE:
        step = "4/5 — Фінанси"
    elif current == STATE_DATES:
        step = "5/5 — Дати"
    else:
        step = "Старт"

    return f"📍 Крок {step}"


def conflict_keyboard(conflict: dict, target: str, index: int):
    buttons = []
    for i, variant in enumerate(conflict["variants"], start=1):
        label = f"{i}. {variant['value'][:40]}"
        buttons.append([InlineKeyboardButton(label, callback_data=f"pick|{target}|{index}|{i-1}")])

    buttons.append([InlineKeyboardButton("✍️ Ввести свій варіант", callback_data=f"custom|{target}|{index}")])
    return InlineKeyboardMarkup(buttons)


def review_keyboard(target: str):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Підтвердити", callback_data=f"confirm|{target}")],
        [InlineKeyboardButton("✍️ Редагувати", callback_data=f"edit|{target}")],
    ])


def property_mode_keyboard():
    return ReplyKeyboardMarkup([["Фото", "Текст"]], resize_keyboard=True, one_time_keyboard=True)


def input_mode_keyboard():
    return ReplyKeyboardMarkup([["Фото", "Вручну"]], resize_keyboard=True, one_time_keyboard=True)


def final_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Підтвердити", callback_data="final_confirm")],
        [InlineKeyboardButton("✍️ Редагувати", callback_data="final_edit")],
    ])


def format_person_summary(title: str, data: dict):
    lines = [title]
    for key in PERSON_FIELDS:
        if data.get(key):
            lines.append(f"{FIELD_LABELS[key]}: {data[key]}")
    if len(lines) == 1:
        lines.append("Немає даних")
    return "\n".join(lines)


def format_property_summary(data: dict):
    lines = ["Об'єкт:"]
    for key in PROPERTY_FIELDS:
        if data.get(key):
            lines.append(f"{FIELD_LABELS[key]}: {data[key]}")
    if len(lines) == 1:
        lines.append("Немає даних")
    return "\n".join(lines)


def build_final_review_text(session: dict):
    ct = "Оренда" if session["contract_type"] == "rent" else "Купівля-продаж"

    parts = [
        f"Тип договору: {ct}",
        "",
        format_person_summary("Наймодавець / Продавець:", session["owner_data"]),
        "",
        format_person_summary("Винаймач / Покупець:", session["tenant_data"]),
        "",
        format_property_summary(session["property_data"]),
        "",
        "Фінанси:",
        f"Ціна: {session['finance_data'].get('price', '')}",
        f"Ціна словами: {session['finance_data'].get('price_words', '')}",
        f"Залог: {session['finance_data'].get('deposit', '')}",
        f"Комунальні: {session['finance_data'].get('communals', '')}",
               f"День оплати: {session['finance_data'].get('payment_day', '')}",
        "",
        "Дати:",
        f"Дата договору: {session['dates_data'].get('contract_date', '')}",
        f"Дата передачі: {session['dates_data'].get('transfer_date', '')}",
        f"Дата початку: {session['dates_data'].get('start_date', '')}",
        f"Дата кінця: {session['dates_data'].get('end_date', '')}",
    ]
    return "\n".join(parts)


def ask_conflict_text(conflict: dict):
    label = FIELD_LABELS.get(conflict["field"], conflict["field"])
    return f"Є розбіжність у полі: {label}\nОберіть правильний варіант."


# =========================================================
# PARSERS
# =========================================================
def parse_manual_person_fixes(text: str, data: dict):
    mapping = {
        "піб": "full_name",
        "дата народження": "birth_date",
        "іпн": "tax_id",
        "паспорт": "passport_number",
        "ким виданий": "passport_issued_by",
        "запис": "passport_record",
        "дата видачі паспорта": "passport_date",
        "адреса": "address",
        "телефон": "phone",
    }

    for line in text.splitlines():
        if ":" not in line:
            continue
        left, right = line.split(":", 1)
        key = left.strip().lower()
        val = right.strip()

        if key in mapping:
            data[mapping[key]] = display_value(mapping[key], val)

    return data


def parse_property_block(text: str):
    data = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        left, right = line.split(":", 1)
        key = left.strip().lower()
        val = right.strip()

        if key == "адреса":
            data["address"] = val
        elif key == "квартира №":
            data["apartment_number"] = val
        elif key == "площа":
            data["total_area"] = val
        elif key == "документ власності №":
            data["ownership_doc_number"] = val
        elif key == "дата документа власності":
            data["ownership_doc_date"] = display_value("ownership_doc_date", val)
        elif key == "вулиця будинку":
            data["building_street"] = val
        elif key == "холодна вода":
            data["cold_water_meter"] = val
        elif key == "гаряча вода":
            data["hot_water_meter"] = val
        elif key == "тепло":
            data["heat_meter"] = val
        elif key == "електрика":
            data["electricity_meter"] = val

    return data


def parse_finance_block(text: str):
    data = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        left, right = line.split(":", 1)
        key = left.strip().lower()
        val = right.strip()

        if key == "ціна":
            data["price"] = val
        elif key == "ціна словами":
            data["price_words"] = val
        elif key == "залог":
            data["deposit"] = val
        elif key == "комунальні":
            data["communals"] = val
        elif key == "день оплати":
            data["payment_day"] = val

    return data


def parse_dates_block(text: str):
    data = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        left, right = line.split(":", 1)
        key = left.strip().lower()
        val = right.strip()

        if key == "дата договору":
            data["contract_date"] = display_value("ownership_doc_date", val)
        elif key == "дата передачі":
            data["transfer_date"] = display_value("ownership_doc_date", val)
        elif key == "дата початку":
            data["start_date"] = display_value("ownership_doc_date", val)
        elif key == "дата кінця":
            data["end_date"] = display_value("ownership_doc_date", val)

    return data


# =========================================================
# OCR
# =========================================================
def extract_person_json_from_images(images_b64: list[str], role_name: str):
    schema = {
        "type": "object",
        "properties": {
            "full_name": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "birth_date": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "tax_id": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "passport_number": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "passport_issued_by": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "passport_record": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "passport_date": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "address": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "phone": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
        },
        "required": PERSON_FIELDS,
        "additionalProperties": False
    }

    content = [{
        "type": "input_text",
        "text": (
            f"Ти аналізуєш фото документів для ролі '{role_name}'. "
            "Витягни тільки чітко видимі дані. "
            "Не вигадуй. Якщо не впевнений — value = null. "
            "Поверни тільки JSON за схемою."
        )
    }]

    for image_b64 in images_b64:
        content.append({
            "type": "input_image",
            "image_url": f"data:image/jpeg;base64,{image_b64}"
        })

    response = client.responses.create(
        model=OPENAI_MODEL,
        input=[{"role": "user", "content": content}],
        text={
            "format": {
                "type": "json_schema",
                "name": f"{role_name}_ocr_result",
                "schema": schema,
                "strict": True
            }
        }
    )
    return safe_json_loads(response.output_text)


def extract_property_json_from_images(images_b64: list[str]):
    schema = {
        "type": "object",
        "properties": {
            "address": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "apartment_number": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "total_area": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "ownership_doc_number": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
            "ownership_doc_date": {
                "type": "object",
                "properties": {"value": {"type": ["string", "null"]}, "confidence": {"type": "number"}},
                "required": ["value", "confidence"],
                "additionalProperties": False
            },
        },
        "required": ["address", "apartment_number", "total_area", "ownership_doc_number", "ownership_doc_date"],
        "additionalProperties": False
    }

    content = [{
        "type": "input_text",
        "text": (
            "Ти аналізуєш фото документа на нерухомість. "
            "Витягни тільки чітко видимі дані по об'єкту. "
            "Не вигадуй. Якщо не впевнений — value = null. "
            "Поверни тільки JSON за схемою."
        )
    }]

    for image_b64 in images_b64:
        content.append({
            "type": "input_image",
            "image_url": f"data:image/jpeg;base64,{image_b64}"
        })

    response = client.responses.create(
        model=OPENAI_MODEL,
        input=[{"role": "user", "content": content}],
        text={
            "format": {
                "type": "json_schema",
                "name": "property_ocr_result",
                "schema": schema,
                "strict": True
            }
        }
    )
    return safe_json_loads(response.output_text)


# =========================================================
# FILE OUTPUT
# =========================================================
def convert_docx_to_pdf(docx_path: str):
    soffice = shutil.which("soffice")
    if not soffice:
        return None

    out_dir = os.path.dirname(docx_path)
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception:
        return None


def build_template_context(session: dict):
    owner = session["owner_data"]
    tenant = session["tenant_data"]
    prop = session["property_data"]
    fin = session["finance_data"]
    dates = session["dates_data"]

    end_date = dates.get("end_date", "")
    end_day, end_month, end_year = "", "", ""

    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", end_date)
    month_map = {
        "01": "січня", "02": "лютого", "03": "березня", "04": "квітня",
        "05": "травня", "06": "червня", "07": "липня", "08": "серпня",
        "09": "вересня", "10": "жовтня", "11": "листопада", "12": "грудня",
    }
    if m:
        end_day, month_num, end_year = m.groups()
        end_month = month_map.get(month_num, "")

    return {
        "CITY": "",
        "DATE": dates.get("contract_date", ""),

        "LANDLORD_NAME": owner.get("full_name", ""),
        "LANDLORD_TAX_ID": owner.get("tax_id", ""),
        "LANDLORD_PASSPORT_NUMBER": owner.get("passport_number", ""),
        "LANDLORD_PASSPORT_ISSUED_BY": owner.get("passport_issued_by", ""),
        "LANDLORD_PASSPORT_RECORD": owner.get("passport_record", ""),
        "LANDLORD_PASSPORT_DATE": owner.get("passport_date", ""),

        "TENANT_NAME": tenant.get("full_name", ""),
        "TENANT_TAX_ID": tenant.get("tax_id", ""),
        "TENANT_PASSPORT_NUMBER": tenant.get("passport_number", ""),
        "TENANT_PASSPORT_ISSUED_BY": tenant.get("passport_issued_by", ""),
        "TENANT_PASSPORT_RECORD": tenant.get("passport_record", ""),
        "TENANT_PASSPORT_DATE": tenant.get("passport_date", ""),

        "OWNERSHIP_DOC_NUMBER": prop.get("ownership_doc_number", ""),
        "OWNERSHIP_DOC_DATE": prop.get("ownership_doc_date", ""),

        "APARTMENT_NUMBER": prop.get("apartment_number", ""),
        "TOTAL_AREA": prop.get("total_area", ""),
        "ADDRESS": prop.get("address", ""),
        "BUILDING_STREET": prop.get("building_street", ""),

        "RENT_PRICE": fin.get("price", ""),
        "RENT_PRICE_WORDS": fin.get("price_words", ""),
        "DEPOSIT": fin.get("deposit", ""),
        "PAYMENT_DAY": fin.get("payment_day", ""),

        "TRANSFER_DATE": dates.get("transfer_date", ""),

        "LANDLORD_PHONE": owner.get("phone", ""),
        "TENANT_PHONE": tenant.get("phone", ""),

        "END_DATE_DAY": end_day,
        "END_DATE_MONTH": end_month,
        "END_DATE_YEAR": end_year,

        "COLD_WATER_METER": prop.get("cold_water_meter", ""),
        "HOT_WATER_METER": prop.get("hot_water_meter", ""),
        "HEAT_METER": prop.get("heat_meter", ""),
        "ELECTRICITY_METER": prop.get("electricity_meter", ""),
    }


# =========================================================
# PROCESSORS
# =========================================================
async def process_person_images(update: Update, session: dict, target: str):
    images = session["owner_images"] if target == "owner" else session["tenant_images"]

    if not images:
        await update.message.reply_text("Немає фото. Надішліть хоча б одне.")
        return

    await update.message.reply_text("📸 Аналізую документи...")

    results = []
    for image_b64 in images:
        try:
            one = extract_person_json_from_images([image_b64], target)
            results.append(one)
        except Exception as e:
            await update.message.reply_text(f"Помилка при аналізі фото: {e}")

    if target == "owner":
        session["owner_ocr_results"] = results
        merged, conflicts = merge_results(results, PERSON_FIELDS)
        session["owner_data"] = merged
    else:
        session["tenant_ocr_results"] = results
        merged, conflicts = merge_results(results, PERSON_FIELDS)
        session["tenant_data"] = merged

    if conflicts:
        session["pending_conflicts"] = conflicts
        session["current_conflict_idx"] = 0
        session["conflict_target"] = target
        session["state"] = STATE_OWNER_REVIEW if target == "owner" else STATE_TENANT_REVIEW

        first_conflict = conflicts[0]
        await update.message.reply_text(
            ask_conflict_text(first_conflict),
            reply_markup=conflict_keyboard(first_conflict, target, 0)
        )
        return

    if target == "owner":
        session["state"] = STATE_OWNER_REVIEW
        await update.message.reply_text(
            f"{progress_text(session)}\n\n{format_person_summary('Дані Наймодавця / Продавця:', session['owner_data'])}",
            reply_markup=review_keyboard("owner")
        )
    else:
        session["state"] = STATE_TENANT_REVIEW
        await update.message.reply_text(
            f"{progress_text(session)}\n\n{format_person_summary('Дані Винаймача / Покупця:', session['tenant_data'])}",
            reply_markup=review_keyboard("tenant")
        )


async def process_property_images(update: Update, session: dict):
    if not session["property_images"]:
        await update.message.reply_text("Немає фото документа на об'єкт.")
        return

    await update.message.reply_text("📄 Аналізую документ на об'єкт...")

    results = []
    for image_b64 in session["property_images"]:
        try:
            one = extract_property_json_from_images([image_b64])
            results.append(one)
        except Exception as e:
            await update.message.reply_text(f"Помилка при аналізі фото об'єкта: {e}")

    session["property_ocr_results"] = results
    merged, conflicts = merge_results(results, [
        "address", "apartment_number", "total_area", "ownership_doc_number", "ownership_doc_date"
    ])
    session["property_data"].update(merged)

    if conflicts:
        session["pending_conflicts"] = conflicts
        session["current_conflict_idx"] = 0
        session["conflict_target"] = "property"
        session["state"] = STATE_PROPERTY_REVIEW

        first_conflict = conflicts[0]
        await update.message.reply_text(
            ask_conflict_text(first_conflict),
            reply_markup=conflict_keyboard(first_conflict, "property", 0)
        )
        return

    session["state"] = STATE_PROPERTY_REVIEW
    await update.message.reply_text(
        f"{progress_text(session)}\n\n{format_property_summary(session['property_data'])}",
        reply_markup=review_keyboard("property")
    )


# =========================================================
# COMMANDS
# =========================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    reset_session(user_id)

    kb = [["Оренда", "Купівля-продаж"], ["Брокер"]]
    await update.message.reply_text(
        "Оберіть тип договору:",
        reply_markup=ReplyKeyboardMarkup(kb, resize_keyboard=True)
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    reset_session(user_id)
    await update.message.reply_text("Скасовано. Напишіть /start", reply_markup=ReplyKeyboardRemove())


async def status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ensure_session(user_id)
    s = sessions[user_id]
    await update.message.reply_text(
        f"Стан: {s['state']}\n"
        f"Тип: {s['contract_type']}\n"
        f"Фото owner: {len(s['owner_images'])}\n"
        f"Фото tenant: {len(s['tenant_images'])}\n"
        f"Фото property: {len(s['property_images'])}"
    )


# =========================================================
# PHOTO HANDLER
# =========================================================
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ensure_session(user_id)
    s = sessions[user_id]

    if s["state"] not in [STATE_OWNER_PHOTOS, STATE_TENANT_PHOTOS, STATE_PROPERTY_PHOTOS]:
        await update.message.reply_text("Зараз я не чекаю фото. Завершіть поточний крок або напишіть /start.")
        return

    if not update.message.photo:
        await update.message.reply_text("Фото не знайдено.")
        return

    photo = update.message.photo[-1]
    tg_file = await photo.get_file()

    file_name = f"{user_id}_{now_ts()}.jpg"
    file_path = os.path.join(TMP_DIR, file_name)
    await tg_file.download_to_drive(file_path)

    with open(file_path, "rb") as f:
        image_b64 = base64.b64encode(f.read()).decode("utf-8")

    if s["state"] == STATE_OWNER_PHOTOS:
        s["owner_images"].append(image_b64)
        await update.message.reply_text("Фото додано. Надішліть ще або напишіть: ГОТОВО НАЙМОДАВЕЦЬ")
        return

    if s["state"] == STATE_TENANT_PHOTOS:
        s["tenant_images"].append(image_b64)
        await update.message.reply_text("Фото додано. Надішліть ще або напишіть: ГОТОВО ВИНАЙМАЧ")
        return

    if s["state"] == STATE_PROPERTY_PHOTOS:
        s["property_images"].append(image_b64)
        await update.message.reply_text("Фото об'єкта додано. Надішліть ще або напишіть: ГОТОВО ОБ'ЄКТ")
        return


# =========================================================
# TEXT HANDLER
# =========================================================
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    ensure_session(user_id)
    s = sessions[user_id]
    text = update.message.text.strip()

    if s["state"] == STATE_WAIT_TYPE:
        if text == "Оренда":
            s["contract_type"] = "rent"
            s["state"] = STATE_OWNER_MODE
            await update.message.reply_text(
                "📍 Крок 1/5 — Наймодавець\n\nОберіть спосіб введення: Фото або Вручну",
                reply_markup=input_mode_keyboard()
            )
            return

        if text == "Купівля-продаж":
            s["contract_type"] = "sale"
            s["state"] = STATE_OWNER_MODE
            await update.message.reply_text(
                "📍 Крок 1/5 — Продавець\n\nОберіть спосіб введення: Фото або Вручну",
                reply_markup=input_mode_keyboard()
            )
            return

        await update.message.reply_text("Оберіть тип договору кнопкою.")
        return

    if s["state"] == STATE_OWNER_MODE:
        low = text.lower()
        if low in {"фото", "документ", "документи"}:
            s["state"] = STATE_OWNER_PHOTOS
            await update.message.reply_text(
                "Надішліть фото паспорта / ID та ІПН Наймодавця.\nКоли завершите — напишіть: ГОТОВО НАЙМОДАВЕЦЬ",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        if low in {"вручну", "текст"}:
            s["state"] = STATE_OWNER_MANUAL
            await update.message.reply_text(
                "Надішліть дані Наймодавця одним повідомленням:\n\n"
                "ПІБ: ...\nДата народження: ...\nІПН: ...\nПаспорт: ...\n"
                "Ким виданий: ...\nЗапис: ...\nДата видачі паспорта: ...\nАдреса: ...\nТелефон: ...",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        await update.message.reply_text("Напишіть: Фото або Вручну")
        return

    if s["state"] == STATE_OWNER_PHOTOS:
        if text.upper() == "ГОТОВО НАЙМОДАВЕЦЬ":
            await process_person_images(update, s, "owner")
            return
        await update.message.reply_text("Я чекаю фото або команду: ГОТОВО НАЙМОДАВЕЦЬ")
        return

    if s["state"] == STATE_OWNER_MANUAL:
        s["owner_data"] = parse_manual_person_fixes(text, {})
        s["state"] = STATE_OWNER_REVIEW
        await update.message.reply_text(
            f"{progress_text(s)}\n\n{format_person_summary('Дані Наймодавця / Продавця:', s['owner_data'])}",
            reply_markup=review_keyboard("owner")
        )
        return

    if s["state"] == STATE_TENANT_MODE:
        low = text.lower()
        if low in {"фото", "документ", "документи"}:
            s["state"] = STATE_TENANT_PHOTOS
            await update.message.reply_text(
                "Надішліть фото паспорта / ID та ІПН Винаймача.\nКоли завершите — напишіть: ГОТОВО ВИНАЙМАЧ",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        if low in {"вручну", "текст"}:
            s["state"] = STATE_TENANT_MANUAL
            await update.message.reply_text(
                "Надішліть дані Винаймача одним повідомленням:\n\n"
                "ПІБ: ...\nДата народження: ...\nІПН: ...\nПаспорт: ...\n"
                "Ким виданий: ...\nЗапис: ...\nДата видачі паспорта: ...\nАдреса: ...\nТелефон: ...",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        await update.message.reply_text("Напишіть: Фото або Вручну")
        return

    if s["state"] == STATE_TENANT_PHOTOS:
        if text.upper() == "ГОТОВО ВИНАЙМАЧ":
            await process_person_images(update, s, "tenant")
            return
        await update.message.reply_text("Я чекаю фото або команду: ГОТОВО ВИНАЙМАЧ")
        return

    if s["state"] == STATE_TENANT_MANUAL:
        s["tenant_data"] = parse_manual_person_fixes(text, {})
        s["state"] = STATE_TENANT_REVIEW
        await update.message.reply_text(
            f"{progress_text(s)}\n\n{format_person_summary('Дані Винаймача / Покупця:', s['tenant_data'])}",
            reply_markup=review_keyboard("tenant")
        )
        return

    if s["state"] == STATE_PROPERTY_MODE:
        low = text.lower()
        if low in {"фото", "документ", "надішлю фото"}:
            s["state"] = STATE_PROPERTY_PHOTOS
            await update.message.reply_text(
                "📍 Крок 3/5 — Об'єкт\n\nНадішліть фото документа на квартиру / об'єкт.\nКоли завершите — напишіть: ГОТОВО ОБ'ЄКТ",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        if low in {"текст", "вручну"}:
            s["state"] = STATE_PROPERTY_TEXT
            await update.message.reply_text(
                "📍 Крок 3/5 — Об'єкт\n\nНадішліть дані одним повідомленням:\n\n"
                "Адреса: ...\nКвартира №: ...\nПлоща: ...\nДокумент власності №: ...\n"
                "Дата документа власності: ...\nВулиця будинку: ...\n"
                "Холодна вода: ...\nГаряча вода: ...\nТепло: ...\nЕлектрика: ...",
                reply_markup=ReplyKeyboardRemove()
            )
            return
        await update.message.reply_text("Напишіть: Фото або Текст")
        return

    if s["state"] == STATE_PROPERTY_PHOTOS:
        if text.upper() == "ГОТОВО ОБ'ЄКТ":
            await process_property_images(update, s)
            return
        await update.message.reply_text("Я чекаю фото або команду: ГОТОВО ОБ'ЄКТ")
        return

    if s["state"] == STATE_PROPERTY_TEXT:
        s["property_data"].update(parse_property_block(text))
        s["state"] = STATE_PROPERTY_REVIEW
        await update.message.reply_text(
            f"{progress_text(s)}\n\n{format_property_summary(s['property_data'])}",
            reply_markup=review_keyboard("property")
        )
        return

    if s["state"] == STATE_FINANCE:
        s["finance_data"] = parse_finance_block(text)
        s["state"] = STATE_DATES
        await update.message.reply_text(
            "📍 Крок 5/5 — Дати\n\n"
            "Надішліть одним повідомленням:\n\n"
            "Дата договору: ...\n"
            "Дата передачі: ...\n"
            "Дата початку: ...\n"
            "Дата кінця: ..."
        )
        return

    if s["state"] == STATE_DATES:
        s["dates_data"] = parse_dates_block(text)
        s["state"] = STATE_FINAL_REVIEW
        await update.message.reply_text(
            build_final_review_text(s),
            reply_markup=final_keyboard()
        )
        return

    if s["custom_edit_target"] == "owner":
        s["owner_data"] = parse_manual_person_fixes(text, s["owner_data"])
        s["custom_edit_target"] = None
        s["state"] = STATE_OWNER_REVIEW
        await update.message.reply_text(
            format_person_summary("Оновлені дані Наймодавця / Продавця:", s["owner_data"]),
            reply_markup=review_keyboard("owner")
        )
        return

    if s["custom_edit_target"] == "tenant":
        s["tenant_data"] = parse_manual_person_fixes(text, s["tenant_data"])
        s["custom_edit_target"] = None
        s["state"] = STATE_TENANT_REVIEW
        await update.message.reply_text(
            format_person_summary("Оновлені дані Винаймача / Покупця:", s["tenant_data"]),
            reply_markup=review_keyboard("tenant")
        )
        return

    if s["custom_edit_target"] == "property":
        s["property_data"].update(parse_property_block(text))
        s["custom_edit_target"] = None
        s["state"] = STATE_PROPERTY_REVIEW
        await update.message.reply_text(
            format_property_summary(s["property_data"]),
            reply_markup=review_keyboard("property")
        )
        return

    if s["custom_edit_target"] == "final":
        s["custom_edit_target"] = None
        await update.message.reply_text("Для простоти: /start і введіть дані заново, або я можу додати точне фінальне редагування окремо.")
        return

    await update.message.reply_text("Не зрозумів. Напишіть /start")


# =========================================================
# CALLBACK HANDLER
# =========================================================
async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    user_id = query.from_user.id
    ensure_session(user_id)
    s = sessions[user_id]
    data = query.data

    if data.startswith("pick|"):
        _, target, idx_str, variant_str = data.split("|")
        idx = int(idx_str)
        variant_idx = int(variant_str)

        conflict = s["pending_conflicts"][idx]
        field = conflict["field"]
        chosen = conflict["variants"][variant_idx]["value"]

        if target == "owner":
            s["owner_data"][field] = chosen
        elif target == "tenant":
            s["tenant_data"][field] = chosen
        elif target == "property":
            s["property_data"][field] = chosen

        s["current_conflict_idx"] += 1

        if s["current_conflict_idx"] < len(s["pending_conflicts"]):
            next_conflict = s["pending_conflicts"][s["current_conflict_idx"]]
            await query.edit_message_text(
                ask_conflict_text(next_conflict),
                reply_markup=conflict_keyboard(next_conflict, target, s["current_conflict_idx"])
            )
        else:
            s["pending_conflicts"] = []
            s["current_conflict_idx"] = 0

            if target == "owner":
                s["state"] = STATE_OWNER_REVIEW
                await query.edit_message_text(format_person_summary("Дані Наймодавця / Продавця:", s["owner_data"]))
                await query.message.reply_text("Перевірте дані:", reply_markup=review_keyboard("owner"))
            elif target == "tenant":
                s["state"] = STATE_TENANT_REVIEW
                await query.edit_message_text(format_person_summary("Дані Винаймача / Покупця:", s["tenant_data"]))
                await query.message.reply_text("Перевірте дані:", reply_markup=review_keyboard("tenant"))
            elif target == "property":
                s["state"] = STATE_PROPERTY_REVIEW
                await query.edit_message_text(format_property_summary(s["property_data"]))
                await query.message.reply_text("Перевірте дані:", reply_markup=review_keyboard("property"))
        return

    if data.startswith("custom|"):
        _, target, _ = data.split("|")
        s["custom_edit_target"] = target
        await query.edit_message_text("Введіть свій правильний варіант одним повідомленням.")
        return

    if data == "confirm|owner":
        s["state"] = STATE_TENANT_MODE
        await query.edit_message_text("Дані Наймодавця підтверджено ✅")
        await query.message.reply_text(
            "📍 Крок 2/5 — Винаймач / Покупець\n\nОберіть спосіб введення: Фото або Вручну",
            reply_markup=input_mode_keyboard()
        )
        return

    if data == "confirm|tenant":
        s["state"] = STATE_PROPERTY_MODE
        await query.edit_message_text("Дані Винаймача підтверджено ✅")
        await query.message.reply_text(
            "📍 Крок 3/5 — Об'єкт\n\nВведення даних: Фото або Текст",
            reply_markup=property_mode_keyboard()
        )
        return

    if data == "confirm|property":
        s["state"] = STATE_FINANCE
        await query.edit_message_text("Дані об'єкта підтверджено ✅")
        await query.message.reply_text(
            "📍 Крок 4/5 — Фінанси\n\nНадішліть одним повідомленням:\n\n"
            "Ціна: ...\n"
            "Ціна словами: ...\n"
            "Залог: ...\n"
            "Комунальні: ...\n"
            "День оплати: ..."
        )
        return

    if data == "edit|owner":
        s["custom_edit_target"] = "owner"
        await query.edit_message_text(
            "Введіть виправлення для Наймодавця одним повідомленням, наприклад:\n\n"
            "ПІБ: ...\nІПН: ...\nПаспорт: ..."
        )
        return

    if data == "edit|tenant":
        s["custom_edit_target"] = "tenant"
        await query.edit_message_text(
            "Введіть виправлення для Винаймача одним повідомленням, наприклад:\n\n"
            "ПІБ: ...\nІПН: ...\nПаспорт: ..."
        )
        return

    if data == "edit|property":
        s["custom_edit_target"] = "property"
        await query.edit_message_text(
            "Введіть виправлення для об'єкта одним повідомленням, наприклад:\n\n"
            "Адреса: ...\nКвартира №: ...\nПлоща: ..."
        )
        return

    if data == "final_confirm":
        try:
            template_path = TEMPLATES[s["contract_type"]]
            context_data = build_template_context(s)

            doc = DocxTemplate(template_path)
            doc.render(context_data)

            file_base = f"contract_{user_id}_{now_ts()}"
            docx_path = os.path.join(OUTPUT_DIR, f"{file_base}.docx")
            doc.save(docx_path)

            pdf_path = convert_docx_to_pdf(docx_path)

            await query.edit_message_text("Формую договір...")

            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    await query.message.reply_document(document=InputFile(f, filename=os.path.basename(pdf_path)))

            with open(docx_path, "rb") as f:
                await query.message.reply_document(document=InputFile(f, filename=os.path.basename(docx_path)))

            await query.message.reply_text("Готово ✅\nДля нового договору напишіть /start", reply_markup=ReplyKeyboardRemove())
            reset_session(user_id)
        except Exception as e:
            await query.message.reply_text(f"Помилка при генерації: {e}")
        return

    if data == "final_edit":
        s["custom_edit_target"] = "final"
        await query.edit_message_text("Напишіть, що саме змінити. Або простіше — /start і введіть заново.")
        return


# =========================================================
# MAIN
# =========================================================
def main():
    if TELEGRAM_TOKEN == "PASTE_TELEGRAM_TOKEN":
        raise ValueError("Вставте TELEGRAM_TOKEN у код")
    if OPENAI_API_KEY == "PASTE_OPENAI_API_KEY":
        raise ValueError("Вставте OPENAI_API_KEY у код")

    app = (
        ApplicationBuilder()
        .token(TELEGRAM_TOKEN)
        .http_version("1.1")
        .connect_timeout(30)
        .read_timeout(30)
        .write_timeout(30)
        .pool_timeout(30)
        .get_updates_http_version("1.1")
        .get_updates_connect_timeout(30)
        .get_updates_read_timeout(30)
        .get_updates_write_timeout(30)
        .get_updates_pool_timeout(30)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    print("Бот запущено...")
    app.run_polling(drop_pending_updates=True, bootstrap_retries=-1)


if __name__ == "__main__":
    main()
