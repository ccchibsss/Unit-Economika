# -*- coding: utf-8 -*-
"""
Юнит-экономика FBS Яндекс Маркет — монолит Streamlit.

Запуск:
    pip install streamlit pandas numpy plotly requests xlsxwriter openpyxl
    streamlit run streamlit_app.py

Что умеет:
    - загрузка каталога: Артикул, Бренд, Категория, Длина, Ширина, Высота, Цена;
    - себестоимость и вес необязательны;
    - расчёт до 300 000+ SKU;
    - ABC/XYZ, цена безубыточности, цена с наценкой, цена под целевую маржу;
    - новые категории вручную и через CSV;
    - сохранение и восстановление загруженных данных;
    - Excel с живыми формулами без ограничения 300k строк: при превышении лимита Excel
      строки автоматически разбиваются на несколько листов Расчет_1, Расчет_2, ...
"""

from __future__ import annotations

import io
import json
import logging
import math
import os
import pickle
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


# =============================================================================
# НАСТРОЙКИ И КОНСТАНТЫ
# =============================================================================
APP_NAME = "Юнит-экономика FBS — Яндекс Маркет"
APP_VERSION = "5.0.0"

BASE_DIR = Path(__file__).parent.resolve() if "__file__" in globals() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
SAVE_DIR = DATA_DIR / "saved_catalogs"
DATA_DIR.mkdir(exist_ok=True)
SAVE_DIR.mkdir(parents=True, exist_ok=True)

BIG_DATA_THRESHOLD = 20_000
EXCEL_MAX_ROWS = 1_048_576
EXCEL_DATA_ROWS_PER_SHEET = 1_000_000
DEFAULT_COST_FALLBACK = 0.65
DEFAULT_DENSITY_KG_PER_L = 0.30
DEFAULT_TARGET_MARGIN = 0.20
DEFAULT_MARKUP = 0.15

DEFAULT_BASE_TARIFF = {
    "commission_rate": 0.14,
    "min_commission": 45.0,
    "logistics_base": 45.0,
    "logistics_per_kg": 14.0,
    "storage_per_day_per_liter": 0.25,
    "acquiring_fee": 0.02,
    "return_fee": 0.02,
}

DEFAULT_SPECIAL_COSTS = {
    "packaging": 45.0,
    "chestny_znak": 1.5,
    "labeling": 3.0,
    "warranty_reserve": 0.02,
    "hazard_surcharge": 0.01,
    "fragile_surcharge": 0.005,
}

DEFAULT_SPECIAL_TARIFFS = {
    "шины": {
        "label": "Шины",
        "commission_rate": 0.12,
        "logistics_base": 90.0,
        "storage_per_day_per_liter": 0.50,
        "reason": "Крупногабаритный",
    },
    "аккумулятор": {
        "label": "Аккумуляторы",
        "commission_rate": 0.13,
        "logistics_base": 75.0,
        "storage_per_day_per_liter": 0.40,
        "reason": "Опасный груз",
    },
    "двигател": {
        "label": "Двигатели",
        "commission_rate": 0.11,
        "logistics_base": 120.0,
        "storage_per_day_per_liter": 0.60,
        "reason": "Крупногабаритный/тяжелый",
    },
    "кпп": {
        "label": "КПП",
        "commission_rate": 0.11,
        "logistics_base": 110.0,
        "storage_per_day_per_liter": 0.60,
        "reason": "Крупногабаритный/тяжелый",
    },
}

DEFAULT_CATEGORIES = {
    "фильтры": {"volume_l": 1.5, "weight_kg": 0.5, "hazardous": False, "fragile": False},
    "масла": {"volume_l": 5.0, "weight_kg": 4.0, "hazardous": True, "fragile": False},
    "колодки": {"volume_l": 0.8, "weight_kg": 1.2, "hazardous": False, "fragile": False},
    "диски": {"volume_l": 3.0, "weight_kg": 4.0, "hazardous": False, "fragile": True},
    "амортизаторы": {"volume_l": 4.0, "weight_kg": 3.5, "hazardous": False, "fragile": True},
    "аккумуляторы": {"volume_l": 12.0, "weight_kg": 15.0, "hazardous": True, "fragile": True},
    "шины": {"volume_l": 25.0, "weight_kg": 10.0, "hazardous": False, "fragile": False},
    "фары": {"volume_l": 6.0, "weight_kg": 2.5, "hazardous": False, "fragile": True},
    "двигател": {"volume_l": 50.0, "weight_kg": 80.0, "hazardous": True, "fragile": True},
    "кпп": {"volume_l": 40.0, "weight_kg": 50.0, "hazardous": True, "fragile": True},
}

FALLBACK_CATEGORY = {"volume_l": 2.0, "weight_kg": 1.0, "hazardous": False, "fragile": False}

SYNONYMS = {
    "Артикул": ["артикул", "sku", "код", "код товара", "offer id", "offerid", "shop-sku", "артикул товара"],
    "Бренд": ["бренд", "brand", "производитель", "марка", "vendor", "торговая марка", "изготовитель"],
    "Категория": ["категория", "category", "группа", "тип товара", "раздел", "категория товара"],
    "Цена": ["цена", "price", "розничная цена", "цена продажи", "цена на маркете", "цена, руб", "цена, ₽"],
    "Себестоимость": ["себестоимость", "cost", "закупка", "закупочная цена", "закуп", "себестоимость, руб"],
    "Вес_кг": ["вес_кг", "вес", "weight", "масса", "вес, кг", "вес упаковки"],
    "Длина": ["длина", "length", "длина упаковки", "длина, см", "длина, мм"],
    "Ширина": ["ширина", "width", "ширина упаковки", "ширина, см", "ширина, мм"],
    "Высота": ["высота", "height", "высота упаковки", "высота, см", "высота, мм"],
    "Объем_л": ["объем_л", "объем", "объём", "volume", "объем, л", "объём, л"],
    "Оборачиваемость_дней": ["оборачиваемость", "turnover", "срок хранения", "оборачиваемость, дней"],
    "Опасный": ["опасный", "hazardous", "опасный груз"],
    "Хрупкий": ["хрупкий", "fragile", "хрупкий груз"],
}

DEMO_BRANDS = [
    "Bosch", "Mann-Filter", "Sachs", "Brembo", "Mahle", "Denso", "Valeo", "TRW",
    "NGK", "Febi Bilstein", "Lemforder", "Hella", "Continental", "Michelin", "Varta",
]

DEMO_CATS = [
    ("Фильтры", 280, 900, 22, 14, 14, 0.6, 25),
    ("Масла", 700, 4200, 28, 16, 28, 4.4, 20),
    ("Колодки", 1500, 3800, 18, 12, 8, 1.4, 30),
    ("Диски", 2500, 9500, 62, 62, 22, 9.2, 45),
    ("Амортизаторы", 3200, 7800, 66, 16, 16, 3.6, 40),
    ("Аккумуляторы", 5200, 12500, 35, 26, 26, 17.0, 35),
    ("Шины", 4200, 14500, 70, 70, 26, 10.6, 15),
    ("Фары", 2800, 18500, 52, 26, 26, 2.6, 50),
    ("Двигатели", 65000, 180000, 100, 62, 72, 95.0, 90),
    ("КПП", 48000, 120000, 82, 58, 56, 50.0, 90),
]

CALC_EXPORT_COLUMNS = [
    "Артикул", "Бренд", "Категория", "Длина", "Ширина", "Высота", "Объем_л", "Вес_кг", "Оплач_вес",
    "Оборачиваемость_дней", "Цена", "Себестоимость", "Себестоимость_оценка", "Ставка_комиссии",
    "Логистика_база", "Логистика_за_кг", "Ставка_хранения", "Ставка_эквайринга", "Переменная_ставка",
    "Фикс_расходы_без_цены", "Комиссия_руб", "Логистика_руб", "Хранение_руб", "Эквайринг_руб",
    "Спец_расходы_FBS", "Итого_расходы", "Прибыль", "Маржа_%", "Безубыток_руб", "Цена_с_наценкой",
    "Цена_целевая_маржа", "Рекомендованная_цена", "Маржа_при_рекоменд_цене", "ABC", "XYZ", "ABC_XYZ",
    "Спецтариф_применен", "Причина_спецтарифа", "Опасный", "Хрупкий",
]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("unit_economics_fbs")


# =============================================================================
# CSS И ФОРМАТИРОВАНИЕ
# =============================================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Manrope:wght@700;800&display=swap');
html, body, [class*="css"] { font-family: Inter, sans-serif; }
.block-container { max-width: 1420px; padding-top: 1rem; }
.hero {
  background: linear-gradient(135deg,#020617 0%,#172554 36%,#4338ca 70%,#f97316 130%);
  color: white; border-radius: 24px; padding: 24px 26px; position: relative; overflow: hidden;
  box-shadow: 0 24px 60px rgba(30,41,59,.22);
}
.hero:before { content:""; position:absolute; inset:-80px -140px auto auto; width:420px; height:260px;
  background: radial-gradient(circle, rgba(251,191,36,.34), transparent 70%); }
.hero h1 { font-family: Manrope, sans-serif; font-size: 1.8rem; font-weight: 800; margin:0; position:relative; }
.hero p { color:#c7d2fe; margin:.35rem 0 0; position:relative; }
.badge { display:inline-flex; align-items:center; gap:.35rem; padding:.25rem .65rem; border-radius:999px;
  background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); color:#eef2ff; font-size:.75rem; font-weight:700; }
.card { background:#fff; border:1px solid #e2e8f0; border-radius:20px; padding:18px; box-shadow:0 8px 26px rgba(15,23,42,.05); }
.metric-card { background:white; border:1px solid #e2e8f0; border-radius:18px; padding:16px; position:relative; overflow:hidden; min-height:112px; }
.metric-card:before { content:""; position:absolute; left:0; right:0; top:0; height:4px; }
.mc-indigo:before{ background:linear-gradient(90deg,#4f46e5,#8b5cf6); }
.mc-emerald:before{ background:linear-gradient(90deg,#059669,#14b8a6); }
.mc-rose:before{ background:linear-gradient(90deg,#e11d48,#f43f5e); }
.mc-amber:before{ background:linear-gradient(90deg,#f59e0b,#f97316); }
.mc-sky:before{ background:linear-gradient(90deg,#0284c7,#06b6d4); }
.small { color:#64748b; font-size:.78rem; font-weight:600; }
.section-title { font-family:Manrope,sans-serif; font-weight:800; font-size:1.15rem; margin:.25rem 0 .65rem; color:#0f172a; }
.muted { color:#64748b; font-size:.84rem; }
.stButton button { border-radius: 12px !important; font-weight: 700 !important; }
div[data-testid="stMetricValue"] { font-size: 1.45rem; }
</style>
"""


def money(v: float, digits: int = 0) -> str:
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return "0 ₽"
    return f"{v:,.{digits}f} ₽".replace(",", " ").replace(".", ",")


def money_short(v: float) -> str:
    v = float(v or 0)
    av = abs(v)
    if av >= 1_000_000_000:
        return f"{v / 1_000_000_000:.2f} млрд ₽".replace(".", ",")
    if av >= 1_000_000:
        return f"{v / 1_000_000:.2f} млн ₽".replace(".", ",")
    if av >= 1_000:
        return f"{v / 1_000:.1f} тыс ₽".replace(".", ",")
    return money(v)


def pct(v: float, digits: int = 1) -> str:
    if v is None or not np.isfinite(float(v)):
        return "0%"
    return f"{float(v) * 100:.{digits}f}%".replace(".", ",")


def fmt_num(v: float, digits: int = 0) -> str:
    return f"{float(v or 0):,.{digits}f}".replace(",", " ").replace(".", ",")


def fmt_size(bytes_count: int) -> str:
    if bytes_count >= 1024 * 1024:
        return f"{bytes_count / 1024 / 1024:.1f} МБ".replace(".", ",")
    return f"{bytes_count / 1024:.0f} КБ"


# =============================================================================
# SESSION STATE
# =============================================================================
def default_tariff() -> Dict[str, Any]:
    return {
        "base": dict(DEFAULT_BASE_TARIFF),
        "special_enabled": True,
        "special_tariffs": {k: dict(v) for k, v in DEFAULT_SPECIAL_TARIFFS.items()},
        "special_costs": dict(DEFAULT_SPECIAL_COSTS),
        "cost_fallback": DEFAULT_COST_FALLBACK,
        "density_kg_per_l": DEFAULT_DENSITY_KG_PER_L,
        "use_category_rates": True,
        "category_rates": {},
    }


def ensure_state() -> None:
    defaults = {
        "step": 0,
        "tariff": default_tariff(),
        "custom_categories": {},
        "df_raw": None,
        "df_calc": None,
        "parse_info": None,
        "target_margin": DEFAULT_TARGET_MARGIN,
        "markup_rate": DEFAULT_MARKUP,
        "last_calc_seconds": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


# =============================================================================
# УТИЛИТЫ ДАННЫХ
# =============================================================================
def to_num(value: Any) -> float:
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value) if np.isfinite(value) else np.nan
    text = str(value).strip().replace("\u00a0", "").replace(" ", "").replace("₽", "").replace("%", "")
    if not text or text == "-":
        return np.nan
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return np.nan


def to_bool(value: Any) -> Optional[bool]:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"да", "yes", "true", "1", "есть", "+", "y"}:
        return True
    if text in {"нет", "no", "false", "0", "-", "n"}:
        return False
    return None


def normalize_header(text: str) -> str:
    return str(text or "").strip().lower().replace("ё", "е")


def resolve_columns(headers: Iterable[str]) -> Dict[str, Optional[str]]:
    headers = [str(h).strip() for h in headers]
    normalized = [normalize_header(h) for h in headers]
    result: Dict[str, Optional[str]] = {}
    for canon, syns in SYNONYMS.items():
        syns_norm = [normalize_header(s) for s in syns]
        found = None
        for idx, h_norm in enumerate(normalized):
            if h_norm in syns_norm:
                found = headers[idx]
                break
        if not found:
            for idx, h_norm in enumerate(normalized):
                if any(s in h_norm for s in syns_norm):
                    found = headers[idx]
                    break
        result[canon] = found
    return result


def dim_factor(header: Optional[str]) -> float:
    h = normalize_header(header or "")
    if "мм" in h or "mm" in h:
        return 0.1
    if "метр" in h:
        return 100.0
    return 1.0


def all_categories() -> Dict[str, Dict[str, Any]]:
    merged = {k.lower(): dict(v) for k, v in DEFAULT_CATEGORIES.items()}
    for k, v in st.session_state.get("custom_categories", {}).items():
        merged[str(k).lower()] = dict(v)
    return merged


def category_defaults(category: str, categories: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    cat = normalize_header(category)
    for key, vals in categories.items():
        if key in cat:
            return vals
    return FALLBACK_CATEGORY


def build_demo_rows(count: int = 24, with_cost: bool = True) -> pd.DataFrame:
    rng = np.random.default_rng(20260408 + count)
    rows: List[Dict[str, Any]] = []
    for i in range(count):
        name, p0, p1, l0, w0, h0, kg0, turnover = DEMO_CATS[i % len(DEMO_CATS)]
        brand = DEMO_BRANDS[(i * 7 + i % 3) % len(DEMO_BRANDS)]
        price = int(round(float(rng.uniform(p0, p1)) / 10) * 10)
        dim_j = float(rng.uniform(0.85, 1.22))
        cost = int(round(price * float(rng.uniform(0.55, 0.77)))) if with_cost else np.nan
        rows.append(
            {
                "Артикул": f"{brand[:3].upper()}-{name[:3].upper()}-{i + 1:06d}",
                "Бренд": brand,
                "Категория": name,
                "Длина": round(l0 * dim_j, 0),
                "Ширина": round(w0 * dim_j, 0),
                "Высота": round(h0 * dim_j, 0),
                "Цена": price,
                "Себестоимость": cost,
                "Вес_кг": round(kg0 * float(rng.uniform(0.8, 1.25)), 2),
                "Оборачиваемость_дней": int(round(turnover * float(rng.uniform(0.7, 1.45)))),
            }
        )
    return pd.DataFrame(rows)


def parse_dataframe(df_raw: pd.DataFrame, file_name: str, file_size: int) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    started = time.time()
    df_raw = df_raw.copy()
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    mapping = resolve_columns(df_raw.columns)
    missing = [c for c in ["Артикул", "Категория", "Цена"] if not mapping.get(c)]
    if missing:
        raise ValueError(f"Не найдены обязательные колонки: {', '.join(missing)}")

    rename = {src: canon for canon, src in mapping.items() if src}
    df = df_raw.rename(columns=rename)
    keep = [c for c in SYNONYMS.keys() if c in df.columns]
    df = df[keep].copy()

    for required in ["Артикул", "Бренд", "Категория", "Цена"]:
        if required not in df.columns:
            df[required] = "" if required != "Цена" else 0

    numeric_cols = ["Цена", "Себестоимость", "Вес_кг", "Длина", "Ширина", "Высота", "Объем_л", "Оборачиваемость_дней"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].map(to_num)

    for col in ["Длина", "Ширина", "Высота"]:
        if col in df.columns and mapping.get(col):
            df[col] = df[col] * dim_factor(mapping[col])

    for col in ["Опасный", "Хрупкий"]:
        if col in df.columns:
            df[col] = df[col].map(to_bool)

    df["Артикул"] = df["Артикул"].astype(str).str.strip()
    df["Бренд"] = df["Бренд"].astype(str).str.strip().replace("", "Без бренда")
    df["Категория"] = df["Категория"].astype(str).str.strip().replace("", "Без категории")
    df["Цена"] = pd.to_numeric(df["Цена"], errors="coerce").fillna(0.0)

    non_empty = (df["Артикул"] != "") | (df["Категория"] != "Без категории") | (df["Цена"] > 0)
    skipped = int((~non_empty).sum())
    df = df[non_empty].reset_index(drop=True)
    if df.empty:
        raise ValueError("После очистки не осталось строк с данными")

    blank_art = df["Артикул"] == ""
    if blank_art.any():
        df.loc[blank_art, "Артикул"] = [f"SKU-{i + 1}" for i in df.index[blank_art]]

    info = {
        "file_name": file_name,
        "file_size": file_size,
        "matched": {k: v for k, v in mapping.items() if v},
        "missing": missing,
        "has_cost": bool(mapping.get("Себестоимость")),
        "has_dims": bool(mapping.get("Длина") and mapping.get("Ширина") and mapping.get("Высота")),
        "skipped": skipped,
        "rows": int(len(df)),
        "parse_seconds": time.time() - started,
    }
    return df, info


def read_uploaded_file(uploaded_file) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    raw = uploaded_file.getvalue()
    name = uploaded_file.name.lower()
    if name.endswith((".csv", ".txt", ".tsv")):
        sample = raw[:4096]
        sep = ";" if sample.count(b";") >= sample.count(b",") else ","
        df_raw = pd.read_csv(io.BytesIO(raw), sep=sep, dtype=str, keep_default_na=False)
        if df_raw.shape[1] == 1:
            alt_sep = "," if sep == ";" else ";"
            df_raw = pd.read_csv(io.BytesIO(raw), sep=alt_sep, dtype=str, keep_default_na=False)
    else:
        if len(raw) > 80 * 1024 * 1024:
            raise ValueError("Excel-файл больше 80 МБ. Для 300k+ SKU сохраните файл как CSV — он читается быстрее.")
        df_raw = pd.read_excel(io.BytesIO(raw), dtype=str, keep_default_na=False)
    return parse_dataframe(df_raw, uploaded_file.name, len(raw))


# =============================================================================
# РАСЧЕТ
# =============================================================================
def effective_tariff_for_category(category_lower: str, tariff: Dict[str, Any]) -> Dict[str, Any]:
    base = dict(tariff["base"])
    base["special_applied"] = False
    base["reason"] = ""
    if tariff.get("special_enabled", True):
        for key, rule in tariff.get("special_tariffs", {}).items():
            if key in category_lower:
                base["commission_rate"] = float(rule["commission_rate"])
                base["logistics_base"] = float(rule["logistics_base"])
                base["storage_per_day_per_liter"] = float(rule["storage_per_day_per_liter"])
                base["special_applied"] = True
                base["reason"] = rule.get("reason", "Спецтариф")
                break
    if tariff.get("use_category_rates") and tariff.get("category_rates"):
        for key, rate in tariff["category_rates"].items():
            if normalize_header(key) in category_lower:
                base["commission_rate"] = float(rate)
                break
    return base


def apply_pricing_columns(df: pd.DataFrame, target_margin: float, markup_rate: float) -> pd.DataFrame:
    out = df.copy()
    fixed = out["Фикс_расходы_без_цены"].astype(float).values
    var_rate = out["Переменная_ставка"].astype(float).values
    price = out["Цена"].astype(float).values
    denom_breakeven = np.clip(1 - var_rate, 0.03, 0.97)
    denom_target = np.clip(1 - var_rate - target_margin, 0.03, 0.97)
    breakeven = fixed / denom_breakeven
    target_price = fixed / denom_target
    markup_price = price * (1 + markup_rate)
    recommended = np.maximum.reduce([breakeven, target_price, markup_price])
    recommended = np.where(np.isfinite(recommended), recommended, 0)
    out["Безубыток_руб"] = np.round(breakeven, 2)
    out["Цена_с_наценкой"] = np.round(markup_price, 2)
    out["Цена_целевая_маржа"] = np.round(target_price, 2)
    out["Рекомендованная_цена"] = np.round(recommended, 0)
    out["Маржа_при_рекоменд_цене"] = np.where(
        recommended > 0,
        (recommended - (fixed + recommended * var_rate)) / recommended,
        0,
    )
    return out


def calculate_df(
    df: pd.DataFrame,
    tariff: Dict[str, Any],
    categories: Dict[str, Dict[str, Any]],
    target_margin: float,
    markup_rate: float,
) -> pd.DataFrame:
    started = time.time()
    n = len(df)
    out = df.copy()
    for col in ["Себестоимость", "Вес_кг", "Длина", "Ширина", "Высота", "Объем_л", "Оборачиваемость_дней"]:
        if col not in out.columns:
            out[col] = np.nan

    cat = out["Категория"].fillna("Без категории").astype(str)
    cat_lower = cat.map(normalize_header)
    unique_cats = cat_lower.unique()
    tariff_cache = {c: effective_tariff_for_category(c, tariff) for c in unique_cats}
    default_cache = {c: category_defaults(c, categories) for c in unique_cats}

    commission_rate = cat_lower.map(lambda c: tariff_cache[c]["commission_rate"]).astype(float).values
    min_commission = cat_lower.map(lambda c: tariff_cache[c]["min_commission"]).astype(float).values
    logistics_base = cat_lower.map(lambda c: tariff_cache[c]["logistics_base"]).astype(float).values
    logistics_per_kg = cat_lower.map(lambda c: tariff_cache[c]["logistics_per_kg"]).astype(float).values
    storage_rate = cat_lower.map(lambda c: tariff_cache[c]["storage_per_day_per_liter"]).astype(float).values
    acquiring_rate = cat_lower.map(lambda c: tariff_cache[c]["acquiring_fee"]).astype(float).values

    def_volume = cat_lower.map(lambda c: default_cache[c]["volume_l"]).astype(float).values
    def_weight = cat_lower.map(lambda c: default_cache[c]["weight_kg"]).astype(float).values
    def_hazard = cat_lower.map(lambda c: default_cache[c]["hazardous"]).astype(bool).values
    def_fragile = cat_lower.map(lambda c: default_cache[c]["fragile"]).astype(bool).values

    price = pd.to_numeric(out["Цена"], errors="coerce").fillna(0).astype(float).values
    cost_raw = pd.to_numeric(out["Себестоимость"], errors="coerce")
    cost_estimated = cost_raw.isna().values | (cost_raw.fillna(0).values <= 0)
    cost = np.where(cost_estimated, price * float(tariff["cost_fallback"]), cost_raw.fillna(0).values)

    length = pd.to_numeric(out["Длина"], errors="coerce").fillna(0).astype(float).values
    width = pd.to_numeric(out["Ширина"], errors="coerce").fillna(0).astype(float).values
    height = pd.to_numeric(out["Высота"], errors="coerce").fillna(0).astype(float).values
    has_dims = (length > 0) & (width > 0) & (height > 0)
    volume_raw = pd.to_numeric(out["Объем_л"], errors="coerce").fillna(0).astype(float).values
    volume_from_dims = np.where(has_dims, length * width * height / 1000.0, 0)
    volume = np.where(volume_raw > 0, volume_raw, np.where(volume_from_dims > 0, volume_from_dims, def_volume))

    weight_raw = pd.to_numeric(out["Вес_кг"], errors="coerce").fillna(0).astype(float).values
    weight_raw = np.where(weight_raw > 100, weight_raw / 1000.0, weight_raw)
    weight_est = volume * float(tariff["density_kg_per_l"])
    weight = np.where(weight_raw > 0, weight_raw, np.where(has_dims, np.maximum(0.1, weight_est), def_weight))
    volumetric_weight = np.where(has_dims, length * width * height / 5000.0, 0)
    billable_weight = np.maximum.reduce([weight, volumetric_weight, np.full(n, 0.1)])

    turnover = pd.to_numeric(out["Оборачиваемость_дней"], errors="coerce").fillna(30).astype(float).values
    turnover = np.where(turnover > 0, turnover, 30)

    hazard_raw = out["Опасный"].map(to_bool).values if "Опасный" in out.columns else np.array([None] * n)
    fragile_raw = out["Хрупкий"].map(to_bool).values if "Хрупкий" in out.columns else np.array([None] * n)
    hazard = np.array([def_hazard[i] if hazard_raw[i] is None else bool(hazard_raw[i]) for i in range(n)], dtype=bool)
    fragile = np.array([def_fragile[i] if fragile_raw[i] is None else bool(fragile_raw[i]) for i in range(n)], dtype=bool)

    sc = tariff["special_costs"]
    fixed_special = float(sc["packaging"]) + float(sc["chestny_znak"]) + float(sc["labeling"])
    variable_special_rate = (
        float(sc["warranty_reserve"])
        + np.where(hazard, float(sc["hazard_surcharge"]), 0)
        + np.where(fragile, float(sc["fragile_surcharge"]), 0)
    )

    commission = np.maximum(price * commission_rate, min_commission)
    logistics = logistics_base + billable_weight * logistics_per_kg
    storage = volume * storage_rate * turnover
    acquiring = price * acquiring_rate
    special_costs = fixed_special + price * variable_special_rate
    total = cost + commission + logistics + storage + acquiring + special_costs
    profit = price - total
    margin = np.where(price > 0, profit / price, 0)

    special_applied = cat_lower.map(lambda c: tariff_cache[c]["special_applied"]).astype(bool).values
    special_reason = cat_lower.map(lambda c: tariff_cache[c]["reason"]).astype(str).values

    fixed_cost_no_price = cost + logistics + storage + fixed_special
    variable_rate = commission_rate + acquiring_rate + variable_special_rate

    # ABC: вклад SKU в общую выручку. A = первые 80%, B = 80-95%, C = хвост.
    revenue = np.maximum(price, 0)
    abc = np.full(n, "C", dtype=object)
    total_revenue = float(revenue.sum())
    if total_revenue > 0:
        order = np.argsort(-revenue)
        cum_share = np.cumsum(revenue[order]) / total_revenue
        abc[order] = np.where(cum_share <= 0.80, "A", np.where(cum_share <= 0.95, "B", "C"))

    # XYZ: устойчивость по оборачиваемости и марже.
    xyz = np.where((turnover <= 30) & (margin >= 0.15), "X", np.where((turnover <= 60) & (margin >= 0), "Y", "Z"))

    out["Артикул"] = out["Артикул"].astype(str)
    out["Бренд"] = out["Бренд"].astype(str).replace("", "Без бренда")
    out["Категория"] = cat
    out["Цена"] = price
    out["Себестоимость"] = cost
    out["Себестоимость_оценка"] = cost_estimated
    out["Длина"] = length
    out["Ширина"] = width
    out["Высота"] = height
    out["Объем_л"] = volume
    out["Вес_кг"] = weight
    out["Оплач_вес"] = billable_weight
    out["Оборачиваемость_дней"] = turnover
    out["Опасный"] = hazard
    out["Хрупкий"] = fragile
    out["Ставка_комиссии"] = commission_rate
    out["Логистика_база"] = logistics_base
    out["Логистика_за_кг"] = logistics_per_kg
    out["Ставка_хранения"] = storage_rate
    out["Ставка_эквайринга"] = acquiring_rate
    out["Переменная_ставка"] = variable_rate
    out["Фикс_расходы_без_цены"] = fixed_cost_no_price
    out["Комиссия_руб"] = commission
    out["Логистика_руб"] = logistics
    out["Хранение_руб"] = storage
    out["Эквайринг_руб"] = acquiring
    out["Спец_расходы_FBS"] = special_costs
    out["Итого_расходы"] = total
    out["Прибыль"] = profit
    out["Маржа_%"] = margin
    out["ABC"] = abc
    out["XYZ"] = xyz
    out["ABC_XYZ"] = [f"{a}{x}" for a, x in zip(abc, xyz)]
    out["Спецтариф_применен"] = special_applied
    out["Причина_спецтарифа"] = special_reason
    out = apply_pricing_columns(out, target_margin, markup_rate)
    st.session_state.last_calc_seconds = time.time() - started
    return out


def totals_row(df: pd.DataFrame) -> Dict[str, float]:
    revenue = float(df["Цена"].sum())
    expenses = float(df["Итого_расходы"].sum())
    profit = float(df["Прибыль"].sum())
    return {
        "count": int(len(df)),
        "revenue": revenue,
        "expenses": expenses,
        "profit": profit,
        "avg_margin": profit / revenue if revenue else 0,
        "loss": int((df["Прибыль"] < 0).sum()),
        "spec": int(df["Спецтариф_применен"].sum()) if "Спецтариф_применен" in df.columns else 0,
        "estimated_cost": int(df["Себестоимость_оценка"].sum()) if "Себестоимость_оценка" in df.columns else 0,
        "commission": float(df["Комиссия_руб"].sum()),
        "logistics": float(df["Логистика_руб"].sum()),
        "storage": float(df["Хранение_руб"].sum()),
        "cost": float(df["Себестоимость"].sum()),
        "recommended_revenue": float(df["Рекомендованная_цена"].sum()) if "Рекомендованная_цена" in df.columns else 0,
    }


def summarize(df: pd.DataFrame, by: str) -> pd.DataFrame:
    grouped = df.groupby(by, dropna=False).agg(
        SKU=("Артикул", "count"),
        Выручка=("Цена", "sum"),
        Расходы=("Итого_расходы", "sum"),
        Прибыль=("Прибыль", "sum"),
        Рекоменд_выручка=("Рекомендованная_цена", "sum"),
        Убыточных=("Прибыль", lambda s: int((s < 0).sum())),
        Спецтариф=("Спецтариф_применен", "sum"),
    ).reset_index().rename(columns={by: "Группа"})
    grouped["Маржа"] = np.where(grouped["Выручка"] > 0, grouped["Прибыль"] / grouped["Выручка"], 0)
    grouped["Потенциал_выручки"] = grouped["Рекоменд_выручка"] - grouped["Выручка"]
    return grouped.sort_values("Выручка", ascending=False)


def abc_xyz_matrix(df: pd.DataFrame) -> pd.DataFrame:
    pivot = pd.pivot_table(
        df,
        index="ABC",
        columns="XYZ",
        values="Цена",
        aggfunc=["count", "sum"],
        fill_value=0,
    )
    rows = []
    for abc in ["A", "B", "C"]:
        for xyz in ["X", "Y", "Z"]:
            subset = df[(df["ABC"] == abc) & (df["XYZ"] == xyz)]
            rows.append(
                {
                    "ABC_XYZ": abc + xyz,
                    "SKU": int(len(subset)),
                    "Выручка": float(subset["Цена"].sum()),
                    "Прибыль": float(subset["Прибыль"].sum()),
                    "Маржа": float(subset["Прибыль"].sum() / subset["Цена"].sum()) if subset["Цена"].sum() else 0,
                }
            )
    return pd.DataFrame(rows)


# =============================================================================
# EXCEL EXPORT
# =============================================================================
def add_excel_formats(workbook: xlsxwriter.Workbook) -> Dict[str, Any]:
    return {
        "title": workbook.add_format({"bold": True, "font_size": 18, "font_color": "#0F172A"}),
        "subtitle": workbook.add_format({"font_color": "#64748B", "font_size": 10}),
        "header": workbook.add_format({
            "bold": True, "bg_color": "#0F3460", "font_color": "#FFFFFF", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "header2": workbook.add_format({
            "bold": True, "bg_color": "#4338CA", "font_color": "#FFFFFF", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "money": workbook.add_format({"num_format": "#,##0.00", "border": 1}),
        "money_green": workbook.add_format({"num_format": "#,##0.00", "border": 1, "font_color": "#047857"}),
        "money_red": workbook.add_format({"num_format": "#,##0.00", "border": 1, "font_color": "#BE123C"}),
        "percent": workbook.add_format({"num_format": "0.00%", "border": 1}),
        "number": workbook.add_format({"num_format": "#,##0.00", "border": 1}),
        "integer": workbook.add_format({"num_format": "#,##0", "border": 1}),
        "text": workbook.add_format({"border": 1}),
        "input": workbook.add_format({"border": 1, "bg_color": "#EEF2FF"}),
        "formula": workbook.add_format({"border": 1, "bg_color": "#F8FAFC"}),
        "good": workbook.add_format({"bg_color": "#DCFCE7", "font_color": "#166534", "border": 1}),
        "warn": workbook.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E", "border": 1}),
        "bad": workbook.add_format({"bg_color": "#FFE4E6", "font_color": "#9F1239", "border": 1}),
        "kpi_blue": workbook.add_format({"bold": True, "font_color": "#1E3A8A", "bg_color": "#DBEAFE", "border": 1}),
        "kpi_green": workbook.add_format({"bold": True, "font_color": "#065F46", "bg_color": "#D1FAE5", "border": 1}),
        "kpi_orange": workbook.add_format({"bold": True, "font_color": "#92400E", "bg_color": "#FEF3C7", "border": 1}),
        "kpi_red": workbook.add_format({"bold": True, "font_color": "#9F1239", "bg_color": "#FFE4E6", "border": 1}),
    }


def write_parameters_sheet(workbook: xlsxwriter.Workbook, tariff: Dict[str, Any], target_margin: float, markup_rate: float) -> None:
    fmt = add_excel_formats(workbook)
    ws = workbook.add_worksheet("Параметры")
    ws.set_column("A:A", 42)
    ws.set_column("B:B", 18)
    ws.write("A1", "Параметры расчета FBS", fmt["title"])
    ws.write("A2", "Меняйте значения в колонке B — формулы на листах Расчет_* пересчитаются", fmt["subtitle"])
    ws.write_row("A4", ["Параметр", "Значение"], fmt["header"])
    rows = [
        ("Базовая комиссия", tariff["base"]["commission_rate"], "percent"),
        ("Минимальная комиссия, руб", tariff["base"]["min_commission"], "money"),
        ("Логистика база, руб", tariff["base"]["logistics_base"], "money"),
        ("Логистика за кг, руб", tariff["base"]["logistics_per_kg"], "money"),
        ("Хранение за литр/сутки, руб", tariff["base"]["storage_per_day_per_liter"], "money"),
        ("Эквайринг", tariff["base"]["acquiring_fee"], "percent"),
        ("Возвраты", tariff["base"]["return_fee"], "percent"),
        ("Упаковка FBS, руб", tariff["special_costs"]["packaging"], "money"),
        ("Честный знак, руб", tariff["special_costs"]["chestny_znak"], "money"),
        ("Маркировка, руб", tariff["special_costs"]["labeling"], "money"),
        ("Гарантийный резерв", tariff["special_costs"]["warranty_reserve"], "percent"),
        ("Надбавка опасный груз", tariff["special_costs"]["hazard_surcharge"], "percent"),
        ("Надбавка хрупкий груз", tariff["special_costs"]["fragile_surcharge"], "percent"),
        ("Себестоимость по умолчанию", tariff["cost_fallback"], "percent"),
        ("Плотность для оценки веса, кг/л", tariff["density_kg_per_l"], "number"),
        ("Целевая маржа", target_margin, "percent"),
        ("Наценка к текущей цене", markup_rate, "percent"),
    ]
    for i, (label, val, fmt_key) in enumerate(rows, start=4):
        ws.write(i, 0, label, fmt["text"])
        ws.write(i, 1, val, fmt[fmt_key])
    ws.freeze_panes(4, 0)


def excel_formula_calc_row(row_num: int, col: Dict[str, int], formula_col: str) -> str:
    r = row_num
    price = f"{xl_col_to_name(col['Цена'])}{r}"
    cost = f"{xl_col_to_name(col['Себестоимость'])}{r}"
    comm_rate = f"{xl_col_to_name(col['Ставка_комиссии'])}{r}"
    min_comm = "Параметры!$B$6"  # actual row in params: row 6 = B6 min commission
    billable = f"{xl_col_to_name(col['Оплач_вес'])}{r}"
    log_base = f"{xl_col_to_name(col['Логистика_база'])}{r}"
    log_kg = f"{xl_col_to_name(col['Логистика_за_кг'])}{r}"
    volume = f"{xl_col_to_name(col['Объем_л'])}{r}"
    storage_rate = f"{xl_col_to_name(col['Ставка_хранения'])}{r}"
    turnover = f"{xl_col_to_name(col['Оборачиваемость_дней'])}{r}"
    acq_rate = f"{xl_col_to_name(col['Ставка_эквайринга'])}{r}"
    var_rate = f"{xl_col_to_name(col['Переменная_ставка'])}{r}"
    fixed = f"{xl_col_to_name(col['Фикс_расходы_без_цены'])}{r}"
    commission = f"{xl_col_to_name(col['Комиссия_руб'])}{r}"
    logistics = f"{xl_col_to_name(col['Логистика_руб'])}{r}"
    storage = f"{xl_col_to_name(col['Хранение_руб'])}{r}"
    acquiring = f"{xl_col_to_name(col['Эквайринг_руб'])}{r}"
    special = f"{xl_col_to_name(col['Спец_расходы_FBS'])}{r}"
    total = f"{xl_col_to_name(col['Итого_расходы'])}{r}"
    profit = f"{xl_col_to_name(col['Прибыль'])}{r}"
    breakeven = f"{xl_col_to_name(col['Безубыток_руб'])}{r}"
    markup = f"{xl_col_to_name(col['Цена_с_наценкой'])}{r}"
    target = f"{xl_col_to_name(col['Цена_целевая_маржа'])}{r}"
    recommended = f"{xl_col_to_name(col['Рекомендованная_цена'])}{r}"
    target_margin = "Параметры!$B$20"
    markup_rate = "Параметры!$B$21"

    formulas = {
        "Комиссия_руб": f"=MAX({price}*{comm_rate},{min_comm})",
        "Логистика_руб": f"={log_base}+{billable}*{log_kg}",
        "Хранение_руб": f"={volume}*{storage_rate}*{turnover}",
        "Эквайринг_руб": f"={price}*{acq_rate}",
        "Спец_расходы_FBS": f"=({fixed}-{cost}-{logistics}-{storage})+{price}*({var_rate}-{comm_rate}-{acq_rate})",
        "Итого_расходы": f"={cost}+{commission}+{logistics}+{storage}+{acquiring}+{special}",
        "Прибыль": f"={price}-{total}",
        "Маржа_%": f"=IF({price}>0,{profit}/{price},0)",
        "Безубыток_руб": f"=IF(1-{var_rate}>0,{fixed}/(1-{var_rate}),0)",
        "Цена_с_наценкой": f"={price}*(1+{markup_rate})",
        "Цена_целевая_маржа": f"=IF(1-{var_rate}-{target_margin}>0,{fixed}/(1-{var_rate}-{target_margin}),0)",
        "Рекомендованная_цена": f"=ROUND(MAX({breakeven},{markup},{target}),0)",
        "Маржа_при_рекоменд_цене": f"=IF({recommended}>0,({recommended}-({fixed}+{recommended}*{var_rate}))/{recommended},0)",
    }
    return formulas[formula_col]


def write_calc_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    df_part: pd.DataFrame,
    fmt: Dict[str, Any],
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    ws.freeze_panes(1, 3)
    ws.set_zoom(85)
    ws.write_row(0, 0, CALC_EXPORT_COLUMNS, fmt["header"])
    col = {name: idx for idx, name in enumerate(CALC_EXPORT_COLUMNS)}

    input_cols = {
        "Артикул", "Бренд", "Категория", "Длина", "Ширина", "Высота", "Объем_л", "Вес_кг", "Оплач_вес",
        "Оборачиваемость_дней", "Цена", "Себестоимость", "Себестоимость_оценка", "Ставка_комиссии",
        "Логистика_база", "Логистика_за_кг", "Ставка_хранения", "Ставка_эквайринга", "Переменная_ставка",
        "Фикс_расходы_без_цены", "ABC", "XYZ", "ABC_XYZ", "Спецтариф_применен", "Причина_спецтарифа", "Опасный", "Хрупкий",
    }
    formula_cols = {
        "Комиссия_руб", "Логистика_руб", "Хранение_руб", "Эквайринг_руб", "Спец_расходы_FBS", "Итого_расходы",
        "Прибыль", "Маржа_%", "Безубыток_руб", "Цена_с_наценкой", "Цена_целевая_маржа", "Рекомендованная_цена",
        "Маржа_при_рекоменд_цене",
    }

    for r_idx, (_, row) in enumerate(df_part.iterrows(), start=1):
        excel_row = r_idx + 1
        for c_idx, name in enumerate(CALC_EXPORT_COLUMNS):
            val = row.get(name, "")
            if name in formula_cols:
                ws.write_formula(r_idx, c_idx, excel_formula_calc_row(excel_row, col, name), fmt["formula"], val)
            elif isinstance(val, (bool, np.bool_)):
                ws.write_boolean(r_idx, c_idx, bool(val), fmt["input"])
            elif isinstance(val, (int, float, np.integer, np.floating)) and np.isfinite(val):
                number_fmt = fmt["percent"] if name.startswith("Ставка") or name in {"Переменная_ставка", "Маржа_%", "Маржа_при_рекоменд_цене"} else fmt["input"]
                ws.write_number(r_idx, c_idx, float(val), number_fmt)
            else:
                ws.write(r_idx, c_idx, str(val), fmt["input"])

    widths = {
        "Артикул": 18, "Бренд": 16, "Категория": 20, "Длина": 10, "Ширина": 10, "Высота": 10,
        "Объем_л": 10, "Вес_кг": 10, "Оплач_вес": 11, "Оборачиваемость_дней": 16, "Цена": 12,
        "Себестоимость": 14, "Себестоимость_оценка": 16, "Причина_спецтарифа": 26,
    }
    for name, idx in col.items():
        ws.set_column(idx, idx, widths.get(name, 13))
    last_row = len(df_part) + 1
    last_col = len(CALC_EXPORT_COLUMNS) - 1
    ws.autofilter(0, 0, last_row, last_col)
    ws.conditional_format(1, col["Маржа_%"], last_row, col["Маржа_%"], {"type": "cell", "criteria": "<", "value": 0, "format": fmt["bad"]})
    ws.conditional_format(1, col["Маржа_%"], last_row, col["Маржа_%"], {"type": "cell", "criteria": ">=", "value": 0.15, "format": fmt["good"]})
    ws.conditional_format(1, col["Прибыль"], last_row, col["Прибыль"], {"type": "cell", "criteria": "<", "value": 0, "format": fmt["bad"]})
    ws.conditional_format(1, col["Рекомендованная_цена"], last_row, col["Рекомендованная_цена"], {"type": "cell", "criteria": ">", "value": 0, "format": fmt["good"]})


def write_summary_sheet(workbook: xlsxwriter.Workbook, sheet_name: str, summary: pd.DataFrame, group_label: str, fmt: Dict[str, Any]) -> None:
    ws = workbook.add_worksheet(sheet_name)
    ws.write("A1", f"Сводка: {group_label}", fmt["title"])
    ws.write("A2", "Отсортировано по выручке. Цвета помогают быстро найти проблемные группы.", fmt["subtitle"])
    headers = list(summary.columns)
    ws.write_row(3, 0, headers, fmt["header"])
    for r_idx, (_, row) in enumerate(summary.iterrows(), start=4):
        for c_idx, name in enumerate(headers):
            val = row[name]
            if isinstance(val, (int, float, np.integer, np.floating)) and np.isfinite(val):
                cell_fmt = fmt["percent"] if name == "Маржа" else fmt["money"] if name in {"Выручка", "Расходы", "Прибыль", "Рекоменд_выручка", "Потенциал_выручки"} else fmt["integer"]
                ws.write_number(r_idx, c_idx, float(val), cell_fmt)
            else:
                ws.write(r_idx, c_idx, str(val), fmt["text"])
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 12)
    ws.set_column("C:F", 16)
    ws.set_column("G:H", 14)
    ws.freeze_panes(4, 0)
    ws.autofilter(3, 0, len(summary) + 3, len(headers) - 1)
    if "Маржа" in headers:
        c = headers.index("Маржа")
        ws.conditional_format(4, c, len(summary) + 3, c, {"type": "cell", "criteria": "<", "value": 0, "format": fmt["bad"]})
        ws.conditional_format(4, c, len(summary) + 3, c, {"type": "cell", "criteria": ">=", "value": 0.15, "format": fmt["good"]})


def write_dashboard_sheet(workbook: xlsxwriter.Workbook, df: pd.DataFrame, fmt: Dict[str, Any]) -> None:
    ws = workbook.add_worksheet("Дашборд")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 4)
    ws.set_column("B:I", 16)
    total = totals_row(df)
    ws.write("B2", APP_NAME, fmt["title"])
    ws.write("B3", f"Дата формирования: {datetime.now():%d.%m.%Y %H:%M}", fmt["subtitle"])
    kpis = [
        ("SKU", total["count"], fmt["kpi_blue"]),
        ("Выручка", total["revenue"], fmt["kpi_blue"]),
        ("Расходы", total["expenses"], fmt["kpi_orange"]),
        ("Прибыль", total["profit"], fmt["kpi_green"] if total["profit"] >= 0 else fmt["kpi_red"]),
        ("Маржа", total["avg_margin"], fmt["kpi_green"] if total["avg_margin"] >= 0.15 else fmt["kpi_orange"]),
        ("Убыточных SKU", total["loss"], fmt["kpi_red"] if total["loss"] else fmt["kpi_green"]),
        ("Реком. выручка", total["recommended_revenue"], fmt["kpi_green"]),
    ]
    for i, (name, val, kfmt) in enumerate(kpis):
        row = 5 + (i // 4) * 3
        col = 1 + (i % 4) * 2
        ws.merge_range(row, col, row, col + 1, name, kfmt)
        value_fmt = fmt["percent"] if name == "Маржа" else fmt["money"] if "выруч" in name.lower() or name in {"Расходы", "Прибыль"} else fmt["integer"]
        ws.merge_range(row + 1, col, row + 1, col + 1, val, value_fmt)

    costs = pd.DataFrame(
        [
            ("Себестоимость", total["cost"]),
            ("Комиссия", total["commission"]),
            ("Логистика", total["logistics"]),
            ("Хранение", total["storage"]),
            ("Эквайринг", float(df["Эквайринг_руб"].sum())),
            ("Спец. расходы", float(df["Спец_расходы_FBS"].sum())),
        ],
        columns=["Расход", "Сумма"],
    )
    start = 13
    ws.write_row(start, 1, costs.columns.tolist(), fmt["header"])
    for i, row in costs.iterrows():
        ws.write(start + 1 + i, 1, row["Расход"], fmt["text"])
        ws.write_number(start + 1 + i, 2, row["Сумма"], fmt["money"])
    chart = workbook.add_chart({"type": "doughnut"})
    chart.add_series({"name": "Структура расходов", "categories": ["Дашборд", start + 1, 1, start + len(costs), 1], "values": ["Дашборд", start + 1, 2, start + len(costs), 2]})
    chart.set_title({"name": "Структура расходов"})
    chart.set_style(10)
    ws.insert_chart("E13", chart, {"x_scale": 1.2, "y_scale": 1.2})


def write_legend_sheet(workbook: xlsxwriter.Workbook, fmt: Dict[str, Any]) -> None:
    ws = workbook.add_worksheet("Легенда")
    ws.set_column("A:A", 30)
    ws.set_column("B:B", 92)
    ws.write_row("A1", ["Колонка", "Описание"], fmt["header"])
    legend = [
        ("ABC", "A — SKU формируют первые 80% выручки; B — следующие 15%; C — хвост."),
        ("XYZ", "X — быстрая/стабильная маржа; Y — средний риск; Z — медленно или убыточно."),
        ("Безубыток_руб", "Цена, при которой прибыль около нуля с учетом комиссии, эквайринга и переменных резервов."),
        ("Цена_целевая_маржа", "Цена для достижения целевой маржи из листа Параметры."),
        ("Цена_с_наценкой", "Текущая цена, увеличенная на процент наценки из листа Параметры."),
        ("Рекомендованная_цена", "MAX(безубыток, цена целевой маржи, цена с наценкой), округлено до рублей."),
        ("Маржа_при_рекоменд_цене", "Ожидаемая маржа, если поставить рекомендованную цену."),
        ("Себестоимость_оценка", "TRUE, если себестоимость была рассчитана как % от цены, потому что в файле ее не было."),
    ]
    for i, row in enumerate(legend, start=1):
        ws.write_row(i, 0, row, fmt["text"])


def export_excel_formulas(df: pd.DataFrame, tariff: Dict[str, Any], target_margin: float, markup_rate: float) -> bytes:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    try:
        workbook = xlsxwriter.Workbook(tmp.name, {"constant_memory": True, "nan_inf_to_errors": True})
        fmt = add_excel_formats(workbook)
        write_parameters_sheet(workbook, tariff, target_margin, markup_rate)
        write_dashboard_sheet(workbook, df, fmt)
        write_summary_sheet(workbook, "Сводка_Категории", summarize(df, "Категория"), "Категории", fmt)
        write_summary_sheet(workbook, "Сводка_Бренды", summarize(df, "Бренд"), "Бренды", fmt)
        write_summary_sheet(workbook, "ABC_XYZ", abc_xyz_matrix(df), "ABC/XYZ", fmt)
        write_legend_sheet(workbook, fmt)
        chunks = math.ceil(len(df) / EXCEL_DATA_ROWS_PER_SHEET)
        for i in range(chunks):
            start = i * EXCEL_DATA_ROWS_PER_SHEET
            end = min((i + 1) * EXCEL_DATA_ROWS_PER_SHEET, len(df))
            write_calc_sheet(workbook, f"Расчет_{i + 1}", df.iloc[start:end], fmt)
        workbook.close()
        return Path(tmp.name).read_bytes()
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


def export_excel_values(df: pd.DataFrame, tariff: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        fmt = add_excel_formats(workbook)
        write_parameters_sheet(workbook, tariff, st.session_state.target_margin, st.session_state.markup_rate)
        write_dashboard_sheet(workbook, df, fmt)
        values_cols = [c for c in CALC_EXPORT_COLUMNS if c in df.columns]
        df[values_cols].to_excel(writer, sheet_name="Расчет_значения", index=False)
        ws = writer.sheets["Расчет_значения"]
        ws.freeze_panes(1, 3)
        ws.autofilter(0, 0, len(df), len(values_cols) - 1)
        for idx, name in enumerate(values_cols):
            ws.write(0, idx, name, fmt["header"])
            ws.set_column(idx, idx, 16)
        write_summary_sheet(workbook, "Сводка_Категории", summarize(df, "Категория"), "Категории", fmt)
        write_summary_sheet(workbook, "Сводка_Бренды", summarize(df, "Бренд"), "Бренды", fmt)
        write_summary_sheet(workbook, "ABC_XYZ", abc_xyz_matrix(df), "ABC/XYZ", fmt)
        write_legend_sheet(workbook, fmt)
    return output.getvalue()


def export_csv(df: pd.DataFrame) -> bytes:
    output = io.StringIO()
    df.to_csv(output, sep=";", index=False, encoding="utf-8-sig")
    return output.getvalue().encode("utf-8-sig")


# =============================================================================
# СОХРАНЕНИЕ / ЗАГРУЗКА
# =============================================================================
def save_current_project() -> Path:
    payload = {
        "df_raw": st.session_state.df_raw,
        "df_calc": st.session_state.df_calc,
        "tariff": st.session_state.tariff,
        "custom_categories": st.session_state.custom_categories,
        "parse_info": st.session_state.parse_info,
        "target_margin": st.session_state.target_margin,
        "markup_rate": st.session_state.markup_rate,
        "saved_at": datetime.now().isoformat(),
    }
    path = SAVE_DIR / f"fbs_project_{datetime.now():%Y%m%d_%H%M%S}.pkl"
    with open(path, "wb") as f:
        pickle.dump(payload, f, protocol=pickle.HIGHEST_PROTOCOL)
    return path


def load_project(path: Path) -> None:
    with open(path, "rb") as f:
        payload = pickle.load(f)
    for key in ["df_raw", "df_calc", "tariff", "custom_categories", "parse_info", "target_margin", "markup_rate"]:
        if key in payload:
            st.session_state[key] = payload[key]


def recalculate_current() -> None:
    if st.session_state.df_raw is None:
        return
    with st.spinner("Пересчитываем юнит-экономику..."):
        st.session_state.df_calc = calculate_df(
            st.session_state.df_raw,
            st.session_state.tariff,
            all_categories(),
            st.session_state.target_margin,
            st.session_state.markup_rate,
        )


# =============================================================================
# UI HELPERS
# =============================================================================
def hero() -> None:
    st.markdown(
        f"""
<div class="hero">
  <div style="display:flex;justify-content:space-between;gap:18px;align-items:center;flex-wrap:wrap;position:relative">
    <div>
      <h1>🚗 {APP_NAME} <span style="color:#fbbf24">FBS</span></h1>
      <p>До 300 000+ SKU · Excel с живыми формулами · ABC/XYZ · рекомендованные цены · v{APP_VERSION}</p>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <span class="badge">⚡ 300k+ SKU</span>
      <span class="badge">📊 Красочный Excel</span>
      <span class="badge">💾 Сохранение проектов</span>
      <span class="badge">🧩 Свои категории</span>
    </div>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


def stepper() -> None:
    steps = [("⚙️", "Тарифы"), ("📦", "Данные"), ("📊", "Дашборд"), ("📥", "Экспорт")]
    cols = st.columns(4)
    max_step = 0 if st.session_state.df_raw is None else 3 if st.session_state.df_calc is not None else 1
    for idx, (icon, label) in enumerate(steps):
        with cols[idx]:
            if st.button(f"{icon} {label}", key=f"step_{idx}", disabled=idx > max_step, type="primary" if st.session_state.step == idx else "secondary", use_container_width=True):
                st.session_state.step = idx
                st.rerun()


def metric_card(title: str, value: str, subtitle: str, tone: str) -> None:
    st.markdown(
        f"""
<div class="metric-card {tone}">
  <div class="small">{title}</div>
  <div style="font-family:Manrope,sans-serif;font-size:1.55rem;font-weight:800;color:#0f172a;margin-top:6px">{value}</div>
  <div class="small" style="margin-top:6px">{subtitle}</div>
</div>
""",
        unsafe_allow_html=True,
    )


# =============================================================================
# APP
# =============================================================================
st.set_page_config(page_title=APP_NAME, layout="wide", page_icon="🚗")
st.markdown(CSS, unsafe_allow_html=True)
ensure_state()
hero()
st.write("")
stepper()
st.divider()

tariff = st.session_state.tariff


# -----------------------------------------------------------------------------
# ШАГ 0. ТАРИФЫ И КАТЕГОРИИ
# -----------------------------------------------------------------------------
if st.session_state.step == 0:
    st.info("Новичок? Оставьте настройки по умолчанию, перейдите к шагу «Данные» и загрузите пример на 24 SKU.")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚙️ Базовый тариф Яндекс Маркета</div>', unsafe_allow_html=True)
        base = tariff["base"]
        base["commission_rate"] = st.number_input("Комиссия маркетплейса, %", 0.0, 60.0, base["commission_rate"] * 100, 0.5) / 100
        base["min_commission"] = st.number_input("Минимальная комиссия, руб", 0.0, 3000.0, base["min_commission"], 5.0)
        base["logistics_base"] = st.number_input("Логистика: база, руб", 0.0, 3000.0, base["logistics_base"], 5.0)
        base["logistics_per_kg"] = st.number_input("Логистика: за 1 кг, руб", 0.0, 500.0, base["logistics_per_kg"], 0.5)
        base["storage_per_day_per_liter"] = st.number_input("Хранение за 1 л/сутки, руб", 0.0, 10.0, base["storage_per_day_per_liter"], 0.05, format="%.2f")
        base["acquiring_fee"] = st.number_input("Эквайринг, %", 0.0, 20.0, base["acquiring_fee"] * 100, 0.1) / 100
        base["return_fee"] = st.number_input("Возвраты, %", 0.0, 20.0, base["return_fee"] * 100, 0.1) / 100
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">🔧 Спецтарифы автозапчастей</div>', unsafe_allow_html=True)
        tariff["special_enabled"] = st.toggle("Применять спецтарифы", value=tariff["special_enabled"])
        for key, rule in tariff["special_tariffs"].items():
            with st.expander(f"{rule['label']} · ключ: {key}", expanded=False):
                c1, c2, c3 = st.columns(3)
                rule["commission_rate"] = c1.number_input(f"Комиссия, % ({key})", 0.0, 60.0, float(rule["commission_rate"]) * 100, 0.5, key=f"sp_comm_{key}") / 100
                rule["logistics_base"] = c2.number_input(f"Логистика, руб ({key})", 0.0, 3000.0, float(rule["logistics_base"]), 5.0, key=f"sp_log_{key}")
                rule["storage_per_day_per_liter"] = c3.number_input(f"Хранение, руб/л ({key})", 0.0, 10.0, float(rule["storage_per_day_per_liter"]), 0.05, key=f"sp_st_{key}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🧩 Себестоимость, вес и доп. расходы</div>', unsafe_allow_html=True)
    sc = tariff["special_costs"]
    c1, c2, c3 = st.columns(3)
    sc["packaging"] = c1.number_input("Упаковка FBS, руб", 0.0, 1000.0, sc["packaging"], 5.0)
    sc["chestny_znak"] = c2.number_input("Честный знак, руб", 0.0, 100.0, sc["chestny_znak"], 0.5)
    sc["labeling"] = c3.number_input("Маркировка/стикеровка, руб", 0.0, 100.0, sc["labeling"], 0.5)
    c1, c2, c3 = st.columns(3)
    sc["warranty_reserve"] = c1.number_input("Гарантийный резерв, %", 0.0, 30.0, sc["warranty_reserve"] * 100, 0.5) / 100
    sc["hazard_surcharge"] = c2.number_input("Надбавка опасный груз, %", 0.0, 30.0, sc["hazard_surcharge"] * 100, 0.5) / 100
    sc["fragile_surcharge"] = c3.number_input("Надбавка хрупкий груз, %", 0.0, 30.0, sc["fragile_surcharge"] * 100, 0.5) / 100
    c1, c2, c3, c4 = st.columns(4)
    tariff["cost_fallback"] = c1.slider("Себестоимость по умолчанию, % от цены", 1, 99, int(tariff["cost_fallback"] * 100), 1) / 100
    tariff["density_kg_per_l"] = c2.number_input("Плотность для оценки веса, кг/л", 0.01, 5.0, float(tariff["density_kg_per_l"]), 0.05)
    st.session_state.target_margin = c3.slider("Целевая маржа для рекомендации, %", 0, 80, int(st.session_state.target_margin * 100), 1) / 100
    st.session_state.markup_rate = c4.slider("Наценка к текущей цене, %", 0, 200, int(st.session_state.markup_rate * 100), 5) / 100
    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")
    with st.expander("📚 Свои категории и характеристики", expanded=True):
        st.caption("CSV-формат: Категория;Объем_л;Вес_кг;Опасный;Хрупкий. Значения Да/Нет допустимы.")
        uploaded_cat = st.file_uploader("Загрузить CSV справочника категорий", type=["csv"], key="categories_csv")
        if uploaded_cat is not None:
            try:
                raw = uploaded_cat.getvalue()
                sep = ";" if raw[:4096].count(b";") >= raw[:4096].count(b",") else ","
                cdf = pd.read_csv(io.BytesIO(raw), sep=sep)
                added = 0
                for _, row in cdf.iterrows():
                    name = str(row.get("Категория", "")).strip().lower()
                    if not name:
                        continue
                    st.session_state.custom_categories[name] = {
                        "volume_l": float(to_num(row.get("Объем_л", 2)) or 2),
                        "weight_kg": float(to_num(row.get("Вес_кг", 1)) or 1),
                        "hazardous": bool(to_bool(row.get("Опасный", False)) or False),
                        "fragile": bool(to_bool(row.get("Хрупкий", False)) or False),
                    }
                    added += 1
                st.success(f"Загружено категорий: {added}")
            except Exception as exc:
                st.error(f"Ошибка загрузки категорий: {exc}")
        c1, c2, c3, c4, c5 = st.columns([2, 1, 1, 1, 1])
        new_name = c1.text_input("Новая категория")
        new_vol = c2.number_input("Объем, л", 0.01, 10000.0, 2.0, 0.1)
        new_w = c3.number_input("Вес, кг", 0.01, 10000.0, 1.0, 0.1)
        new_h = c4.toggle("Опасный")
        new_f = c5.toggle("Хрупкий")
        if st.button("➕ Добавить категорию") and new_name.strip():
            st.session_state.custom_categories[new_name.strip().lower()] = {"volume_l": new_vol, "weight_kg": new_w, "hazardous": new_h, "fragile": new_f}
            st.success(f"Категория добавлена: {new_name}")
        all_cat_df = pd.DataFrame([
            {"Категория": k, "Объем_л": v["volume_l"], "Вес_кг": v["weight_kg"], "Опасный": "Да" if v["hazardous"] else "Нет", "Хрупкий": "Да" if v["fragile"] else "Нет"}
            for k, v in all_categories().items()
        ])
        st.download_button("⬇️ Скачать текущий справочник категорий", data=all_cat_df.to_csv(index=False, sep=";").encode("utf-8-sig"), file_name="categories_fbs.csv", mime="text/csv")
        st.dataframe(all_cat_df.head(80), use_container_width=True, hide_index=True)

    with st.expander("🌐 Индивидуальные ставки категорий из API Яндекс Маркета", expanded=False):
        c1, c2, c3 = st.columns([2, 1, 1])
        token = c1.text_input("OAuth-токен", type="password")
        campaign_id = c2.number_input("Campaign ID", min_value=1, value=123456, step=1)
        if c3.button("Обновить из API", use_container_width=True):
            if not token:
                st.warning("Введите OAuth-токен")
            else:
                try:
                    resp = requests.get(
                        f"https://api.partner.market.yandex.ru/v2/campaigns/{int(campaign_id)}/categories/commissions",
                        headers={"Authorization": f"OAuth {token}", "Accept": "application/json"},
                        timeout=15,
                    )
                    if resp.status_code != 200:
                        st.error(f"API вернул {resp.status_code}: {resp.text[:300]}")
                    else:
                        cats = resp.json().get("result", {}).get("categories", [])
                        rates = {}
                        for item in cats:
                            name = normalize_header(item.get("categoryName", ""))
                            if name:
                                rates[name] = float(item.get("commissionPercent", 14)) / 100
                        tariff["category_rates"] = rates
                        tariff["use_category_rates"] = True
                        st.success(f"Получены ставки для {len(rates)} категорий")
                except Exception as exc:
                    st.error(f"Ошибка запроса API: {exc}")
        if st.button("🎲 Загрузить демо-ставки"):
            tariff["category_rates"] = {"шины": 0.12, "аккумуляторы": 0.13, "фильтры": 0.14, "масла": 0.15, "двигатели": 0.11, "кпп": 0.11}
            tariff["use_category_rates"] = True
            st.success("Демо-ставки загружены")
        tariff["use_category_rates"] = st.toggle("Использовать индивидуальные ставки", value=tariff.get("use_category_rates", True))
        if tariff.get("category_rates"):
            st.json(tariff["category_rates"])

    if st.session_state.df_raw is not None:
        if st.button("🔁 Пересчитать загруженные данные с новыми тарифами", type="primary"):
            recalculate_current()
            st.success("Расчет обновлен")


# -----------------------------------------------------------------------------
# ШАГ 1. ДАННЫЕ
# -----------------------------------------------------------------------------
elif st.session_state.step == 1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📦 Загрузка каталога</div>', unsafe_allow_html=True)
    st.caption("Минимально: Артикул, Категория, Цена. Рекомендуется: Бренд, Длина, Ширина, Высота. Себестоимость и вес необязательны.")
    uploaded = st.file_uploader("CSV / XLSX / XLS / TXT", type=["csv", "xlsx", "xls", "txt", "tsv"])
    st.markdown("</div>", unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    template = "Артикул;Бренд;Категория;Длина;Ширина;Высота;Цена;Себестоимость;Вес_кг;Оборачиваемость_дней\nMAN-FLT-000001;Mann-Filter;Фильтры;22;14;14;450;220;0,5;25\nMIC-TR-000002;Michelin;Шины;70;70;26;5400;3650;10,5;15\nVAR-BAT-000003;Varta;Аккумуляторы;35;26;26;6500;;16,5;30\n"
    c1.download_button("📋 Шаблон CSV", data=("\ufeff" + template).encode("utf-8-sig"), file_name="template_fbs.csv", mime="text/csv", use_container_width=True)
    if c2.button("🎲 24 SKU", use_container_width=True):
        df_demo = build_demo_rows(24)
        st.session_state.df_raw, st.session_state.parse_info = parse_dataframe(df_demo, "demo_24.csv", 0)
        recalculate_current()
        st.success("Демо 24 SKU загружено")
        st.rerun()
    if c3.button("⚡ 50 000 SKU", use_container_width=True):
        with st.spinner("Генерируем и считаем 50 000 SKU..."):
            df_demo = build_demo_rows(50_000)
            st.session_state.df_raw, st.session_state.parse_info = parse_dataframe(df_demo, "demo_50000.csv", 0)
            recalculate_current()
        st.rerun()
    if c4.button("🚀 300 000 SKU", use_container_width=True):
        with st.spinner("Генерируем и считаем 300 000 SKU..."):
            df_demo = build_demo_rows(300_000)
            st.session_state.df_raw, st.session_state.parse_info = parse_dataframe(df_demo, "demo_300000.csv", 0)
            recalculate_current()
        st.rerun()
    if c5.button("💾 Сохранить", use_container_width=True, disabled=st.session_state.df_calc is None):
        path = save_current_project()
        st.success(f"Сохранено: {path.name}")

    saved = sorted(SAVE_DIR.glob("fbs_project_*.pkl"), key=lambda p: p.stat().st_mtime, reverse=True)
    if saved:
        with st.expander("📂 Подгрузить сохраненные данные", expanded=False):
            selected = st.selectbox("Сохраненный проект", [p.name for p in saved[:20]])
            if st.button("Подгрузить выбранный проект"):
                load_project(SAVE_DIR / selected)
                st.success("Проект восстановлен")
                st.rerun()

    if uploaded is not None:
        try:
            started = time.time()
            with st.spinner("Читаем файл и распознаем колонки..."):
                st.session_state.df_raw, st.session_state.parse_info = read_uploaded_file(uploaded)
            with st.spinner("Считаем юнит-экономику..."):
                recalculate_current()
            st.success(f"Загружено {len(st.session_state.df_raw):,} SKU за {time.time() - started:.1f} сек".replace(",", " "))
            st.rerun()
        except Exception as exc:
            logger.exception(exc)
            st.error(f"Ошибка обработки файла: {exc}")

    if st.session_state.parse_info:
        info = st.session_state.parse_info
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">🔎 Предпросмотр: {info["file_name"]}</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Строк", f"{info['rows']:,}".replace(",", " "))
        c2.metric("Размер", fmt_size(info["file_size"]))
        c3.metric("Пропущено пустых", info["skipped"])
        c4.metric("Чтение, сек", f"{info['parse_seconds']:.1f}")
        if not info["has_cost"]:
            st.warning("Колонка Себестоимость не найдена: значения будут рассчитаны как % от цены.")
        if not info["has_dims"]:
            st.info("Габариты не найдены полностью: объем и вес будут взяты из справочника категорий.")
        if info.get("matched"):
            st.caption("Колонки: " + " · ".join(f"{k} ← {v}" for k, v in info["matched"].items()))
        if st.session_state.df_raw is not None:
            st.dataframe(st.session_state.df_raw.head(10), use_container_width=True, hide_index=True)
        if st.button("📊 Перейти к дашборду", type="primary", disabled=st.session_state.df_calc is None):
            st.session_state.step = 2
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# ШАГ 2. ДАШБОРД
# -----------------------------------------------------------------------------
elif st.session_state.step == 2:
    df = st.session_state.df_calc
    if df is None or df.empty:
        st.warning("Сначала загрузите данные")
        st.stop()

    # Обновляем ценовые рекомендации при изменении слайдеров без полного пересчета.
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📈 Управление маржой и рекомендованной ценой</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1, 2])
    new_target = c1.slider("Целевая маржа, %", 0, 80, int(st.session_state.target_margin * 100), 1) / 100
    new_markup = c2.slider("Наценка к текущей цене, %", 0, 200, int(st.session_state.markup_rate * 100), 5) / 100
    if new_target != st.session_state.target_margin or new_markup != st.session_state.markup_rate:
        st.session_state.target_margin = new_target
        st.session_state.markup_rate = new_markup
        st.session_state.df_calc = apply_pricing_columns(st.session_state.df_calc, new_target, new_markup)
        df = st.session_state.df_calc
    c3.info("Рекомендованная цена = максимум из безубытка, цены под целевую маржу и цены с наценкой. Эти же формулы будут в Excel.")
    st.markdown("</div>", unsafe_allow_html=True)

    totals = totals_row(df)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Выручка", money_short(totals["revenue"]), f"{totals['count']:,} SKU".replace(",", " "), "mc-indigo")
    with c2:
        metric_card("Прибыль", money_short(totals["profit"]), "после всех расходов", "mc-emerald" if totals["profit"] >= 0 else "mc-rose")
    with c3:
        metric_card("Маржа", pct(totals["avg_margin"]), f"Убыточных: {totals['loss']:,}".replace(",", " "), "mc-emerald" if totals["avg_margin"] >= 0.15 else "mc-amber")
    with c4:
        uplift = totals["recommended_revenue"] - totals["revenue"]
        metric_card("Потенциал цены", money_short(uplift), "если поставить рекомендованные цены", "mc-sky")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Комиссия", money_short(totals["commission"]))
    c2.metric("Логистика", money_short(totals["logistics"]))
    c3.metric("Хранение", money_short(totals["storage"]))
    c4.metric("Себест. оценена", f"{totals['estimated_cost']:,}".replace(",", " "))
    c5.metric("Спецтариф", f"{totals['spec']:,}".replace(",", " "))

    mode = st.radio("Разрез аналитики", ["Категория", "Бренд", "ABC/XYZ"], horizontal=True)
    if mode in {"Категория", "Бренд"}:
        summary = summarize(df, mode)
        top = summary.head(15)
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(top, x="Группа", y="Маржа", color="Маржа", color_continuous_scale="RdYlGn", title=f"Маржа: {mode}", range_color=[-0.2, 0.35])
            fig.update_layout(height=360, xaxis_tickangle=-25, coloraxis_showscale=False)
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(top, x="Группа", y="Прибыль", color="Прибыль", color_continuous_scale="RdYlGn", title=f"Прибыль: {mode}")
            fig.update_layout(height=360, xaxis_tickangle=-25, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            summary.head(50).style.format({
                "Выручка": "{:,.0f} ₽", "Расходы": "{:,.0f} ₽", "Прибыль": "{:,.0f} ₽",
                "Рекоменд_выручка": "{:,.0f} ₽", "Потенциал_выручки": "{:,.0f} ₽", "Маржа": "{:.1%}",
            }),
            use_container_width=True,
            hide_index=True,
        )
    else:
        matrix = abc_xyz_matrix(df)
        fig = px.treemap(matrix, path=["ABC_XYZ"], values="Выручка", color="Маржа", color_continuous_scale="RdYlGn", title="ABC/XYZ матрица по выручке")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(matrix.style.format({"Выручка": "{:,.0f} ₽", "Прибыль": "{:,.0f} ₽", "Маржа": "{:.1%}"}), use_container_width=True, hide_index=True)

    costs = pd.DataFrame(
        [
            ("Себестоимость", totals["cost"]),
            ("Комиссия", totals["commission"]),
            ("Логистика", totals["logistics"]),
            ("Хранение", totals["storage"]),
            ("Эквайринг", float(df["Эквайринг_руб"].sum())),
            ("Спец. расходы", float(df["Спец_расходы_FBS"].sum())),
        ],
        columns=["Расход", "Сумма"],
    )
    fig = px.pie(costs, names="Расход", values="Сумма", hole=0.58, title="Структура расходов")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">🧮 Расчет по SKU</div>', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
    query = f1.text_input("Поиск: артикул / бренд / категория")
    brand_filter = f2.selectbox("Бренд", ["Все"] + df["Бренд"].value_counts().head(200).index.tolist())
    cat_filter = f3.selectbox("Категория", ["Все"] + df["Категория"].value_counts().head(200).index.tolist())
    only_loss = f4.toggle("Только убыточные")
    view = df
    if query:
        q = query.lower()
        view = view[view["Артикул"].str.lower().str.contains(q, na=False) | view["Бренд"].str.lower().str.contains(q, na=False) | view["Категория"].str.lower().str.contains(q, na=False)]
    if brand_filter != "Все":
        view = view[view["Бренд"] == brand_filter]
    if cat_filter != "Все":
        view = view[view["Категория"] == cat_filter]
    if only_loss:
        view = view[view["Прибыль"] < 0]
    sort_col = st.selectbox("Сортировка", ["Прибыль", "Маржа_%", "Рекомендованная_цена", "Цена", "ABC", "XYZ", "Бренд", "Категория"])
    ascending = st.toggle("По возрастанию", value=False)
    view = view.sort_values(sort_col, ascending=ascending)
    show_cols = [c for c in ["Артикул", "Бренд", "Категория", "Цена", "Себестоимость", "Итого_расходы", "Прибыль", "Маржа_%", "Безубыток_руб", "Рекомендованная_цена", "Маржа_при_рекоменд_цене", "ABC", "XYZ", "ABC_XYZ"] if c in view.columns]
    st.caption(f"Показано {min(len(view), 5000):,} из {len(view):,} строк. Полный объем доступен в Excel/CSV.".replace(",", " "))
    st.dataframe(view[show_cols].head(5000), use_container_width=True, hide_index=True, height=560)

    col1, col2 = st.columns(2)
    if col1.button("💾 Сохранить проект", type="secondary", use_container_width=True):
        path = save_current_project()
        st.success(f"Сохранено: {path.name}")
    if col2.button("📥 Перейти к экспорту", type="primary", use_container_width=True):
        st.session_state.step = 3
        st.rerun()


# -----------------------------------------------------------------------------
# ШАГ 3. ЭКСПОРТ
# -----------------------------------------------------------------------------
else:
    df = st.session_state.df_calc
    if df is None or df.empty:
        st.warning("Сначала загрузите данные")
        st.stop()
    totals = totals_row(df)
    st.markdown(
        f"""
<div class="card" style="background:linear-gradient(135deg,#059669,#0d9488);color:white;border:none">
  <h3 style="margin:0;color:white">📥 Отчет готов</h3>
  <p style="margin:6px 0 0;color:#d1fae5">{len(df):,} SKU · Выручка {money_short(totals['revenue'])} · Прибыль {money_short(totals['profit'])} · Маржа {pct(totals['avg_margin'])}</p>
</div>
""".replace(",", " "),
        unsafe_allow_html=True,
    )
    st.write("")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📊 Excel с живыми формулами</div>', unsafe_allow_html=True)
        sheets = math.ceil(len(df) / EXCEL_DATA_ROWS_PER_SHEET)
        st.caption(f"Формулы в каждой строке. Если строк больше {EXCEL_DATA_ROWS_PER_SHEET:,}, файл разбивается на листы Расчет_1..N.".replace(",", " "))
        if st.button("Скачать Excel с формулами", type="primary", use_container_width=True):
            with st.spinner("Формируем большой Excel с живыми формулами. Для 300k SKU это может занять несколько минут..."):
                data = export_excel_formulas(df, st.session_state.tariff, st.session_state.target_margin, st.session_state.markup_rate)
            st.download_button(
                "⬇️ Скачать .xlsx",
                data=data,
                file_name=f"unit_economy_fbs_formulas_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚡ Excel со значениями</div>', unsafe_allow_html=True)
        st.caption("Самый быстрый вариант: все расчеты уже посчитаны, плюс дашборд, сводки, ABC/XYZ и легенда.")
        if st.button("Скачать Excel со значениями", use_container_width=True):
            with st.spinner("Формируем Excel со значениями..."):
                data = export_excel_values(df, st.session_state.tariff)
            st.download_button(
                "⬇️ Скачать .xlsx",
                data=data,
                file_name=f"unit_economy_fbs_values_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📄 CSV со всеми SKU</div>', unsafe_allow_html=True)
        st.caption("Для BI, Power Query и архивного хранения. Разделитель ';', UTF-8 BOM.")
        st.download_button(
            "Скачать CSV",
            data=export_csv(df),
            file_name=f"unit_economy_fbs_{datetime.now():%Y%m%d_%H%M%S}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.info("Excel с формулами больше не ограничен 300 000 строками: ограничение только техническое — максимум строк листа Excel. При превышении файл автоматически делится на несколько листов.")


st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:.78rem;margin-top:24px'>{APP_NAME} · монолит Streamlit v{APP_VERSION}</p>", unsafe_allow_html=True)
