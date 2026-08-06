# -*- coding: utf-8 -*-
"""
Unit Economics FBS for Yandex Market
Monolithic Streamlit application, version 4.1.0.

Run:
    pip install streamlit pandas numpy plotly requests xlsxwriter openpyxl
    streamlit run streamlit_app.py

The application is designed for catalogs up to 300,000 SKU. CSV is the
recommended format for large files. All business logic, UI, API integration,
analytics and exports are intentionally kept in this single file.
"""

from __future__ import annotations

import csv
import io
import json
import logging
import math
import os
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


# =============================================================================
# APPLICATION CONFIGURATION
# =============================================================================

APP_NAME = "Юнит-экономика FBS"
APP_VERSION = "5.0.0"
MAX_SKU = 1_048_500
FORMULA_EXCEL_LIMIT = 1_048_500
TABLE_PREVIEW_LIMIT = 5_000
API_URL = "https://api.partner.market.yandex.ru"
CATEGORY_CSV_TEMPLATE = (
    "key;label;volume_l;weight_kg;is_hazardous;is_fragile;commission_rate;logistics_base;storage_per_day_per_liter\n"
    "фильтры;Фильтры;1.5;0.5;false;false;0.14;45;0.25\n"
    "шины;Шины;25;10;false;false;0.12;90;0.5\n"
    "аккумулятор;Аккумуляторы;12;15;true;true;0.13;75;0.4\n"
)

BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
SETTINGS_FILE = DATA_DIR / "unit_economics_settings.json"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("UnitEconomicsFBS")

st.set_page_config(
    page_title=f"{APP_NAME} | Яндекс Маркет",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# CONSTANTS AND DEFAULTS
# =============================================================================

PAYOUT_SCHEDULE_RATES = {
    "weekly_4w": 0.016,
    "weekly_2w": 0.023,
    "weekly_1w": 0.028,
    "daily": 0.033,
}

MAX_DENSITY = 12.0   # кг/л — плотнее свинца физически невозможно
MIN_DENSITY = 0.01   # кг/л — легче пенопласта

# XYZ по абсолютным порогам оборачиваемости (экономический смысл, а не терцили)
XYZ_X_DAYS = 30      # быстрый оборот
XYZ_Y_DAYS = 60      # средний; свыше — Z (замороженные деньги)

ABC_XYZ_MEANING = {
    "AX": "★ Ядро — держать сток, не снижать цену",
    "AY": "Хиты со средним оборотом — следить за остатками",
    "AZ": "Дорогие «зависающие» — риск заморозки денег",
    "BX": "Стабильный середняк — можно поднять цену",
    "BY": "Обычные товары — плановый контроль",
    "BZ": "Медленные середняки — сократить закупку",
    "CX": "Дешёвые ходовые — кандидаты на подъём цены",
    "CY": "Малозначимые — минимум внимания",
    "CZ": "⚠ Балласт — вывести или распродать",
}

def srednya_mila_py(volume_liters: float) -> float:
    v = math.ceil(max(0, volume_liters))
    if v <= 0:
        return 0.0
    if v <= 1:
        return 80.0
    if v <= 30:
        return 80 + 9 * (v - 1)
    if v <= 200:
        return 80 + 9 * 29 + 7 * (v - 30)
    return min(80 + 9 * 29 + 7 * 170 + 5 * (v - 200), 5500)

DEFAULT_SETTINGS: Dict[str, Any] = {
    "commission_rate": 0.14,
    "min_commission": 45.0,
    "logistics_base": 45.0,
    "logistics_per_kg": 14.0,
    "storage_per_day_per_liter": 0.25,
    "acquiring_fee": 0.02,
    "return_fee": 0.02,
    "packaging": 45.0,
    "chestny_znak": 1.5,
    "labeling": 3.0,
    "warranty_reserve": 0.02,
    "hazard_surcharge": 0.01,
    "fragile_surcharge": 0.005,
    "cost_fallback_rate": 0.65,
    "density_kg_per_liter": 0.30,
    "special_enabled": True,
    "use_category_rates": True,
    "category_rates": {},
    "custom_categories": [],
    "pricing": {"mode": "none", "markupPercent": 15, "targetMargin": 0.20},
    "model": "FBS",
    "use_fixed_tariff": True,
    "fbs_speed_discount": 0.0,
    "payout_schedule": "weekly_2w",
    "storage_free_days": 90,
    "special_tariffs": {
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
            "reason": "Крупногабаритный/тяжёлый",
        },
        "кпп": {
            "label": "КПП",
            "commission_rate": 0.11,
            "logistics_base": 110.0,
            "storage_per_day_per_liter": 0.60,
            "reason": "Крупногабаритный/тяжёлый",
        },
    },
}

CATEGORY_DEFAULTS: Dict[str, Tuple[float, float, bool, bool]] = {
    "фильтры": (1.5, 0.5, False, False),
    "масла": (5.0, 4.0, True, False),
    "колодки": (0.8, 1.2, False, False),
    "диски": (3.0, 4.0, False, True),
    "амортизаторы": (4.0, 3.5, False, True),
    "аккумуляторы": (12.0, 15.0, True, True),
    "шины": (25.0, 10.0, False, False),
    "фары": (6.0, 2.5, False, True),
    "двигател": (50.0, 80.0, True, True),
    "кпп": (40.0, 50.0, True, True),
}

DEMO_CATEGORY_RATES: Dict[str, float] = {
    "фильтры": 0.14,
    "масла": 0.15,
    "колодки": 0.13,
    "диски": 0.16,
    "амортизаторы": 0.15,
    "аккумуляторы": 0.13,
    "шины": 0.12,
    "фары": 0.17,
    "двигатели": 0.11,
    "кпп": 0.11,
}

COLUMN_SYNONYMS: Dict[str, List[str]] = {
    "Артикул": [
        "артикул", "sku", "артикул товара", "артикул поставщика", "код",
        "код товара", "offer id", "offerid", "shop-sku", "shop sku",
    ],
    "Бренд": [
        "бренд", "brand", "производитель", "марка", "торговая марка",
        "vendor", "изготовитель",
    ],
    "Категория": [
        "категория", "категория товара", "category", "группа", "тип товара",
        "раздел",
    ],
    "ID_категории": [
        "id категории", "category id", "categoryid", "market category id",
        "marketcategoryid", "market_category_id",
    ],
    "Цена": [
        "цена", "цена, ₽", "цена руб", "цена, руб", "розничная цена",
        "price", "цена продажи", "цена на маркете",
    ],
    "Себестоимость": [
        "себестоимость", "себестоимость, ₽", "cost", "закупка",
        "закупочная цена", "закупочная", "закуп",
    ],
    "Вес_кг": [
        "вес_кг", "вес, кг", "вес", "вес (кг)", "weight", "масса",
        "вес брутто",
    ],
    "Длина": [
        "длина", "длина, см", "длина, мм", "длина (см)", "length",
        "длина упаковки",
    ],
    "Ширина": [
        "ширина", "ширина, см", "ширина, мм", "ширина (см)", "width",
        "ширина упаковки",
    ],
    "Высота": [
        "высота", "высота, см", "высота, мм", "высота (см)", "height",
        "высота упаковки",
    ],
    "Объем_л": [
        "объем_л", "объем, л", "объем", "объём, л", "объём", "volume",
    ],
    "Оборачиваемость_дней": [
        "оборачиваемость_дней", "оборачиваемость", "оборачиваемость, дней",
        "turnover", "оборачиваемость, дн", "срок хранения",
    ],
    "Продажи_шт": [
        "продажи_шт", "продажи", "продано", "количество", "кол-во", "qty",
        "quantity", "sales", "заказы", "заказов", "штук",
    ],
    "Опасный": [
        "опасный", "опасный груз", "опасность", "hazardous", "опасный_груз",
    ],
    "Хрупкий": [
        "хрупкий", "хрупкий груз", "хрупкость", "fragile", "хрупкий_груз",
    ],
}

REQUIRED_COLUMNS = ["Артикул", "Категория", "Цена"]

RESULT_COLUMNS = [
    "Артикул", "Бренд", "Категория", "ID_категории", "Длина", "Ширина",
    "Высота", "Объем_л", "Вес_кг", "Оплач_вес", "Цена", "Себестоимость",
    "Себестоимость_оценка", "Ставка_комиссии", "Комиссия_руб",
    "Логистика_руб", "Хранение_руб", "Эквайринг_руб", "Возвраты_руб", "Спец_расходы_FBS",
    "Итого_расходы", "Прибыль", "Маржа_%", "Продажи_шт", "Выручка", "Прибыль_период",
    "Рекомендованная_цена", "Цена_с_наценкой",
    "Прибыль_с_наценкой", "Маржа_с_наценкой_%", "ABC", "XYZ", "ABC_XYZ", "ABC_XYZ_смысл",
    "Выручка_доля", "Оборачиваемость_дней", "Спецтариф_применён", "Причина_спецтарифа",
    "Рекомендация", "Валидация", "Сумма_сторон", "Плотность",
]

MONEY_COLUMNS = [
    "Цена", "Себестоимость", "Комиссия_руб", "Логистика_руб",
    "Хранение_руб", "Эквайринг_руб", "Возвраты_руб", "Спец_расходы_FBS",
    "Итого_расходы", "Прибыль",
]


# =============================================================================
# VISUAL STYLE
# =============================================================================

st.markdown(
    """
<style>
    :root {
        --ink: #0f172a;
        --muted: #64748b;
        --line: #e2e8f0;
        --indigo: #4f46e5;
        --violet: #7c3aed;
        --green: #059669;
        --amber: #d97706;
        --red: #e11d48;
        --card-shadow: 0 4px 16px rgba(15,23,42,.06), 0 1px 3px rgba(15,23,42,.04);
        --card-shadow-hover: 0 12px 32px rgba(15,23,42,.10), 0 4px 12px rgba(15,23,42,.06);
    }
    .stApp {
        background:
          radial-gradient(1200px 600px at 20% -10%, #e0e7ff 0%, transparent 50%),
          radial-gradient(1000px 500px at 90% 0%, #fce7f3 0%, transparent 50%),
          radial-gradient(800px 600px at 50% 110%, #dbeafe 0%, transparent 50%),
          #f1f5f9;
        color: var(--ink);
    }
    .block-container { max-width: 1480px; padding-top: 1.25rem; padding-bottom: 4rem; }
    /* ── SIDEBAR — контраст исправлен ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e1b4b 50%, #312e81 100%);
        border-right: 1px solid rgba(255,255,255,.06);
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stText"] {
        color: #e2e8f0;
    }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.10); }
    /* Метрики в сайдбаре — белые карточки с тёмным текстом (исправлено) */
    [data-testid="stSidebar"] div[data-testid="stMetric"] {
        background: white !important;
        border: 1px solid rgba(255,255,255,.9) !important;
        border-radius: 16px !important;
        padding: 16px 16px !important;
        box-shadow: 0 8px 24px rgba(0,0,0,.14), 0 2px 8px rgba(0,0,0,.08) !important;
    }
    [data-testid="stSidebar"] div[data-testid="stMetricLabel"] label,
    [data-testid="stSidebar"] div[data-testid="stMetricLabel"] p,
    [data-testid="stSidebar"] div[data-testid="stMetricLabel"] div {
        color: #64748b !important;
        font-weight: 700 !important;
        font-size: 11px !important;
        text-transform: uppercase;
        letter-spacing: .06em;
    }
    [data-testid="stSidebar"] div[data-testid="stMetricValue"] div,
    [data-testid="stSidebar"] div[data-testid="stMetricValue"] p {
        color: #0f172a !important;
        font-weight: 850 !important;
        font-size: 22px !important;
        letter-spacing: -.02em;
    }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] [data-baseweb="select"] * { color: #0f172a !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background: white !important; border-radius: 10px;
    }
    /* Кнопки в сайдбаре — белый фон, тёмный текст (исправлено) */
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] .stDownloadButton > button {
        background: white !important;
        color: #0f172a !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        min-height: 42px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }
    [data-testid="stSidebar"] .stButton > button:hover,
    [data-testid="stSidebar"] .stDownloadButton > button:hover {
        background: #f8fafc !important;
        color: #0f172a !important;
        transform: translateY(-1px);
        box-shadow: 0 8px 20px rgba(0,0,0,.12);
        border-color: #cbd5e1 !important;
    }
    [data-testid="stSidebar"] .stButton > button:active { transform: translateY(0); }
    [data-testid="stSidebar"] .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #6366f1, #7c3aed) !important;
        color: white !important;
        border: 0 !important;
        box-shadow: 0 4px 16px rgba(99,102,241,.35);
    }
    [data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg, #4f46e5, #6d28d9) !important;
        color: white !important;
        box-shadow: 0 8px 24px rgba(99,102,241,.45);
    }
    [data-testid="stSidebar"] .stButton > button p,
    [data-testid="stSidebar"] .stDownloadButton > button p {
        color: inherit !important; font-weight: 700;
    }
    [data-testid="stSidebar"] .stButton > button[kind="primary"] p { color: white !important; }
    /* ── HERO — супер-современный ── */
    .hero {
        position: relative; overflow: hidden; border-radius: 24px;
        padding: 32px 36px; color: white; margin-bottom: 20px;
        background:
          radial-gradient(600px 400px at 15% -20%, rgba(99,102,241,.35), transparent 60%),
          radial-gradient(700px 500px at 85% 0%, rgba(168,85,247,.30), transparent 60%),
          linear-gradient(115deg, #0f172a 0%, #1e1b4b 30%, #312e81 60%, #4c1d95 100%);
        box-shadow: 0 20px 60px rgba(30,41,59,.22), 0 4px 16px rgba(0,0,0,.08);
        border: 1px solid rgba(255,255,255,.08);
    }
    .hero:after {
        content: ""; position: absolute; width: 420px; height: 420px;
        right: -120px; top: -200px; border-radius: 999px;
        border: 60px solid rgba(255,255,255,.05);
        pointer-events: none;
    }
    .hero:before {
        content: ""; position: absolute; inset: 0;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,.04), transparent);
        transform: translateX(-100%);
        animation: shimmer 3.5s infinite;
    }
    @keyframes shimmer { 100% { transform: translateX(100%); } }
    .hero-kicker { font-size: 11px; font-weight: 800; letter-spacing: .16em; color: #fbbf24; text-transform: uppercase; display:flex; align-items:center; gap:8px; }
    .hero-kicker:before { content:""; width:28px; height:2px; background:#fbbf24; border-radius:999px; display:inline-block; }
    .hero h1 { margin: 8px 0 6px; font-size: clamp(26px, 4vw, 40px); line-height: 1.05; font-weight: 850; letter-spacing:-.03em; }
    .hero p { margin: 0; color: #e0e7ff; font-size: 14.5px; line-height:1.5; max-width: 720px; }
    .hero-badge {
        display: inline-flex; align-items:center; gap:6px; margin-top: 16px; padding: 7px 13px;
        border: 1px solid rgba(255,255,255,.15); border-radius: 999px;
        background: rgba(255,255,255,.10); backdrop-filter: blur(8px);
        font-size: 11px; font-weight: 700; letter-spacing:.02em;
        box-shadow: 0 2px 12px rgba(0,0,0,.12);
    }
    .section-title { margin: 6px 0 4px; font-size: 24px; font-weight: 850; letter-spacing:-.025em; color: #0f172a; }
    .section-sub { margin: 0 0 18px; color: var(--muted); font-size: 13.5px; line-height:1.5; }
    .info-box, .warn-box, .success-box, .danger-box {
        border-radius: 16px; padding: 14px 16px; margin: 10px 0 16px;
        font-size: 13px; line-height: 1.55; border: 1px solid;
        display:flex; gap:10px; align-items:flex-start;
        box-shadow: var(--card-shadow);
    }
    .info-box:before, .warn-box:before, .success-box:before, .danger-box:before {
        font-size:16px; line-height:1; margin-top:1px;
    }
    .info-box { background: linear-gradient(135deg, #eff6ff, #dbeafe); color: #1e40af; border-color: #bfdbfe; }
    .info-box:before { content:"💡"; }
    .warn-box { background: linear-gradient(135deg, #fffbeb, #fef3c7); color: #92400e; border-color: #fde68a; }
    .warn-box:before { content:"⚠️"; }
    .success-box { background: linear-gradient(135deg, #ecfdf5, #d1fae5); color: #065f46; border-color: #a7f3d0; }
    .success-box:before { content:"✅"; }
    .danger-box { background: linear-gradient(135deg, #fff1f2, #ffe4e6); color: #9f1239; border-color: #fecdd3; }
    .danger-box:before { content:"🚨"; }
    .metric-card {
        min-height: 138px; background: white; border: 1px solid var(--line);
        border-radius: 20px; padding: 18px 20px; position: relative; overflow: hidden;
        box-shadow: var(--card-shadow);
        transition: all .25s cubic-bezier(.2,.8,.2,1);
    }
    .metric-card:hover { transform: translateY(-3px); box-shadow: var(--card-shadow-hover); border-color: #cbd5e1; }
    .metric-card:before { content: ""; position: absolute; inset: 0 0 auto 0; height: 4px; background: var(--accent); }
    .metric-card:after {
        content:""; position:absolute; right:-20px; top:-20px; width:80px; height:80px;
        background: var(--accent); opacity:.06; border-radius:50%; pointer-events:none;
    }
    .metric-label { font-size: 10.5px; color: var(--muted); font-weight: 800; letter-spacing: .09em; text-transform: uppercase; }
    .metric-value { margin-top: 10px; font-size: clamp(22px, 2.6vw, 32px); font-weight: 850; letter-spacing: -.04em; color: #0f172a; line-height:1; }
    .metric-note { margin-top: 6px; color: var(--muted); font-size: 11.5px; font-weight:500; }
    .chip {
        display: inline-flex; align-items:center; gap:5px; padding: 5px 10px; border-radius: 999px;
        margin: 3px; font-size: 11px; font-weight: 700; letter-spacing:.01em;
        background: linear-gradient(135deg, #eef2ff, #e0e7ff); color: #4338ca; border: 1px solid #c7d2fe;
        box-shadow: 0 1px 4px rgba(99,102,241,.12);
        transition:.15s;
    }
    .chip:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(99,102,241,.18); }
    .soft-panel {
        background: white; border: 1px solid var(--line); border-radius: 20px;
        padding: 20px; box-shadow: var(--card-shadow);
        transition:.2s;
    }
    .soft-panel:hover { box-shadow: var(--card-shadow-hover); }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, white, #f8fafc) !important;
        border: 1px solid var(--line) !important; border-radius: 16px !important;
        padding: 16px 18px !important; box-shadow: var(--card-shadow) !important;
        transition:.2s;
    }
    div[data-testid="stMetric"]:hover { transform: translateY(-2px); box-shadow: var(--card-shadow-hover) !important; }
    div[data-testid="stMetricLabel"] label { color: var(--muted) !important; font-weight:700 !important; font-size:11px !important; text-transform:uppercase; letter-spacing:.06em; }
    div[data-testid="stMetricValue"] div { color: #0f172a !important; font-weight:800 !important; }
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px !important; font-weight: 700 !important; min-height: 42px;
        border: 1px solid #e2e8f0 !important; transition: all .2s cubic-bezier(.2,.8,.2,1) !important;
        box-shadow: 0 2px 8px rgba(15,23,42,.04);
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-2px); box-shadow: 0 8px 20px rgba(15,23,42,.10) !important; border-color: #cbd5e1 !important;
    }
    .stButton > button:active { transform: translateY(0); }
    .stButton > button[kind="primary"] {
        border: 0 !important; color: white !important;
        background: linear-gradient(90deg, var(--indigo), var(--violet)) !important;
        box-shadow: 0 4px 16px rgba(99,102,241,.30) !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg, #4338ca, #6d28d9) !important;
        box-shadow: 0 8px 24px rgba(99,102,241,.40) !important;
    }
    div[data-baseweb="tab-list"] { gap: 6px; background: #f1f5f9; padding:6px; border-radius:14px; }
    div[data-baseweb="tab"] { border-radius: 10px !important; padding: 8px 16px !important; font-weight:600; border:0 !important; }
    div[data-baseweb="tab"][aria-selected="true"] { background: white !important; color: var(--indigo) !important; box-shadow: 0 2px 8px rgba(15,23,42,.08) !important; }
    [data-testid="stDataFrame"] { border-radius: 16px; overflow: hidden; border: 1px solid var(--line); box-shadow: var(--card-shadow); }
    .small-muted { font-size: 11px; color: var(--muted); }
    /* Улучшенные заголовки таблиц */
    [data-testid="stDataFrame"] thead tr th { background:#f8fafc !important; font-weight:700 !important; color:#334155 !important; }
    @media (max-width: 700px) {
        .block-container { padding-left: .8rem; padding-right: .8rem; }
        .hero { padding: 23px 20px; border-radius: 18px; }
    }
</style>
""",
    unsafe_allow_html=True,
)


# =============================================================================
# GENERIC HELPERS
# =============================================================================

def deep_copy_json(value: Any) -> Any:
    return json.loads(json.dumps(value, ensure_ascii=False))


def deep_merge(base: Dict[str, Any], patch: Dict[str, Any]) -> Dict[str, Any]:
    result = deep_copy_json(base)
    for key, value in patch.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_settings() -> Dict[str, Any]:
    if SETTINGS_FILE.exists():
        try:
            with SETTINGS_FILE.open("r", encoding="utf-8") as fh:
                return deep_merge(DEFAULT_SETTINGS, json.load(fh))
        except Exception as exc:
            logger.warning("Cannot load settings: %s", exc)
    return deep_copy_json(DEFAULT_SETTINGS)


def save_settings(settings: Dict[str, Any]) -> None:
    try:
        with SETTINGS_FILE.open("w", encoding="utf-8") as fh:
            json.dump(settings, fh, ensure_ascii=False, indent=2)
    except Exception as exc:
        logger.warning("Cannot save settings: %s", exc)


def money(value: float, digits: int = 0) -> str:
    if value is None or not np.isfinite(value):
        return "0 ₽"
    return f"{value:,.{digits}f}".replace(",", " ").replace(".", ",") + " ₽"


def money_short(value: float) -> str:
    sign = "-" if value < 0 else ""
    value = abs(float(value))
    if value >= 1_000_000_000:
        return f"{sign}{value / 1_000_000_000:.2f}".replace(".", ",") + " млрд ₽"
    if value >= 1_000_000:
        return f"{sign}{value / 1_000_000:.2f}".replace(".", ",") + " млн ₽"
    if value >= 1_000:
        return f"{sign}{value / 1_000:.1f}".replace(".", ",") + " тыс ₽"
    return money((-value if sign else value), 0)


def percent(value: float, digits: int = 1) -> str:
    if value is None or not np.isfinite(value):
        return "0%"
    return f"{value * 100:.{digits}f}%".replace(".", ",")


def file_size_label(size: int) -> str:
    if size >= 1024 ** 3:
        return f"{size / 1024 ** 3:.2f} ГБ"
    if size >= 1024 ** 2:
        return f"{size / 1024 ** 2:.1f} МБ"
    return f"{size / 1024:.0f} КБ"


def stable_settings_hash(settings: Dict[str, Any]) -> str:
    return json.dumps(settings, sort_keys=True, ensure_ascii=False, separators=(",", ":"))


def temporary_path(suffix: str) -> str:
    fh = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    path = fh.name
    fh.close()
    return path


def read_and_remove(path: str) -> bytes:
    try:
        with open(path, "rb") as fh:
            return fh.read()
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def metric_card(label: str, value: str, note: str, accent: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card" style="--accent:{accent}">
          <div class="metric-label">{label}</div>
          <div class="metric-value">{value}</div>
          <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# SESSION STATE
# =============================================================================

SESSION_DEFAULTS = {
    "settings": load_settings(),
    "raw_df": None,
    "result_df": None,
    "source_name": "",
    "source_size": 0,
    "parse_meta": {},
    "calculated_settings_hash": "",
    "calculation_seconds": 0.0,
    "export_bytes": None,
    "export_name": "",
    "export_mime": "",
    "last_api_result": None,
}

for state_key, state_value in SESSION_DEFAULTS.items():
    if state_key not in st.session_state:
        st.session_state[state_key] = state_value


def clear_export() -> None:
    st.session_state.export_bytes = None
    st.session_state.export_name = ""
    st.session_state.export_mime = ""


def mark_dirty() -> None:
    clear_export()


def calculation_is_current() -> bool:
    return (
        st.session_state.result_df is not None
        and st.session_state.calculated_settings_hash
        == stable_settings_hash(st.session_state.settings)
    )


# =============================================================================
# INPUT PARSING AND NORMALIZATION
# =============================================================================

def normalize_header(value: Any) -> str:
    return " ".join(str(value).strip().lower().replace("ё", "е").split())


def resolve_columns(columns: Iterable[Any]) -> Dict[str, str]:
    original = [str(c).strip() for c in columns]
    normalized = [normalize_header(c) for c in original]
    result: Dict[str, str] = {}

    for canonical, synonyms in COLUMN_SYNONYMS.items():
        normalized_synonyms = [normalize_header(s) for s in synonyms]
        found: Optional[str] = None

        for idx, header in enumerate(normalized):
            if header in normalized_synonyms:
                found = original[idx]
                break

        if found is None:
            for idx, header in enumerate(normalized):
                if any(synonym in header for synonym in normalized_synonyms):
                    found = original[idx]
                    break

        if found is not None:
            result[canonical] = found

    return result


def detect_csv_format(data: bytes) -> Tuple[str, str]:
    encodings = ["utf-8-sig", "utf-8", "cp1251"]
    decoded = None
    encoding = "utf-8-sig"

    sample_bytes = data[:128_000]
    for candidate in encodings:
        try:
            decoded = sample_bytes.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue

    if decoded is None:
        encoding = "latin-1"
        decoded = sample_bytes.decode(encoding, errors="replace")

    first_lines = "\n".join(decoded.splitlines()[:20])
    try:
        dialect = csv.Sniffer().sniff(first_lines, delimiters=";,\t|")
        separator = dialect.delimiter
    except csv.Error:
        counts = {sep: first_lines.count(sep) for sep in [";", ",", "\t", "|"]}
        separator = max(counts, key=counts.get)

    return encoding, separator


def read_uploaded_file(uploaded: Any) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    started = time.perf_counter()
    name = uploaded.name.lower()
    data = uploaded.getvalue()

    if name.endswith((".csv", ".txt", ".tsv")):
        encoding, separator = detect_csv_format(data)
        frame = pd.read_csv(
            io.BytesIO(data),
            sep=separator,
            encoding=encoding,
            low_memory=False,
            on_bad_lines="warn",
        )
        input_type = f"CSV · разделитель {repr(separator)} · {encoding}"
    elif name.endswith((".xlsx", ".xls")):
        if len(data) > 80 * 1024 ** 2:
            raise ValueError(
                "Excel-файл больше 80 МБ. Сохраните его в CSV: для 300 000 SKU "
                "CSV работает быстрее и требует меньше памяти."
            )
        frame = pd.read_excel(io.BytesIO(data), sheet_name=0)
        input_type = "Excel · первый лист"
    else:
        raise ValueError("Поддерживаются CSV, TXT, TSV, XLSX и XLS.")

    frame.columns = [str(c).strip() for c in frame.columns]
    frame = frame.dropna(how="all")
    return frame, {
        "input_type": input_type,
        "raw_rows": len(frame),
        "raw_columns": len(frame.columns),
        "parse_seconds": time.perf_counter() - started,
    }


def numeric_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype("string")
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace("₽", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def bool_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(series):
        return series.astype("boolean")
    values = series.astype("string").str.strip().str.lower()
    yes = values.isin(["да", "yes", "true", "1", "есть", "+"])
    no = values.isin(["нет", "no", "false", "0", "-"])
    result = pd.Series(pd.NA, index=series.index, dtype="boolean")
    result.loc[yes] = True
    result.loc[no] = False
    return result


def dimension_factor(source_header: str) -> float:
    header = normalize_header(source_header)
    if "мм" in header or "mm" in header:
        return 0.1
    if "метр" in header or header.endswith(" m"):
        return 100.0
    return 1.0


def prepare_input_frame(frame: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    mapping = resolve_columns(frame.columns)
    missing = [column for column in REQUIRED_COLUMNS if column not in mapping]
    if missing:
        raise ValueError(
            "Не найдены обязательные колонки: " + ", ".join(missing) + ". "
            "Минимальный набор: Артикул, Категория, Цена."
        )

    count = len(frame)
    result = pd.DataFrame(index=np.arange(count))

    result["Артикул"] = (
        frame[mapping["Артикул"]].astype("string").fillna("").str.strip().reset_index(drop=True)
    )
    empty_articles = result["Артикул"].eq("")
    if empty_articles.any():
        generated = "SKU-" + (np.arange(count) + 1).astype(str)
        result.loc[empty_articles, "Артикул"] = generated[empty_articles.to_numpy()]

    if "Бренд" in mapping:
        result["Бренд"] = (
            frame[mapping["Бренд"]].astype("string").fillna("").str.strip().reset_index(drop=True)
        )
        result["Бренд"] = result["Бренд"].mask(result["Бренд"].eq(""), "Без бренда")
    else:
        result["Бренд"] = "Без бренда"

    result["Категория"] = (
        frame[mapping["Категория"]].astype("string").fillna("").str.strip().reset_index(drop=True)
    )
    result["Категория"] = result["Категория"].mask(
        result["Категория"].eq(""), "Без категории"
    )

    if "ID_категории" in mapping:
        result["ID_категории"] = (
            numeric_series(frame[mapping["ID_категории"]]).reset_index(drop=True).fillna(0).astype("int64")
        )
    else:
        result["ID_категории"] = 0

    result["Цена"] = numeric_series(frame[mapping["Цена"]]).reset_index(drop=True).fillna(0.0)

    if "Себестоимость" in mapping:
        result["Себестоимость"] = numeric_series(
            frame[mapping["Себестоимость"]]
        ).reset_index(drop=True)
    else:
        result["Себестоимость"] = np.nan

    numeric_optional = [
        "Вес_кг", "Длина", "Ширина", "Высота", "Объем_л",
        "Оборачиваемость_дней", "Продажи_шт",
    ]
    for column in numeric_optional:
        if column in mapping:
            values = numeric_series(frame[mapping[column]]).reset_index(drop=True)
            if column in ["Длина", "Ширина", "Высота"]:
                values = values * dimension_factor(mapping[column])
            result[column] = values
        else:
            result[column] = np.nan

    for column in ["Опасный", "Хрупкий"]:
        if column in mapping:
            result[column] = bool_series(frame[mapping[column]]).reset_index(drop=True)
        else:
            result[column] = pd.Series(pd.NA, index=result.index, dtype="boolean")

    # Remove rows that contain neither a useful identifier nor a price/category.
    valid_mask = ~(
        result["Артикул"].eq("")
        & result["Категория"].eq("Без категории")
        & result["Цена"].eq(0)
    )
    result = result.loc[valid_mask].reset_index(drop=True)

    meta = {
        "mapping": mapping,
        "missing_optional": [c for c in COLUMN_SYNONYMS if c not in mapping],
        "has_cost": "Себестоимость" in mapping,
        "has_brand": "Бренд" in mapping,
        "has_dimensions": all(c in mapping for c in ["Длина", "Ширина", "Высота"]),
        "prepared_rows": len(result),
        "zero_prices": int(result["Цена"].le(0).sum()),
    }
    return result, meta


def build_template_csv() -> bytes:
    rows = [
        [
            "Артикул", "Бренд", "Категория", "ID_категории", "Длина",
            "Ширина", "Высота", "Цена", "Себестоимость", "Вес_кг",
            "Оборачиваемость_дней", "Продажи_шт",
        ],
        ["MAN-FLT-000001", "Mann-Filter", "Фильтры", "", 22, 14, 14, 450, 220, "0,5", 25, 1200],
        ["MIC-TR-000002", "Michelin", "Шины", "", 70, 70, 26, 5400, 3650, "10,5", 15, 180],
        ["VAR-BAT-000003", "Varta", "Аккумуляторы", "", 35, 26, 26, 6500, "", "16,5", 30, 95],
    ]
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=";", lineterminator="\n")
    writer.writerows(rows)
    return ("\ufeff" + buffer.getvalue()).encode("utf-8")


def generate_demo_catalog(count: int, include_cost: bool = True) -> pd.DataFrame:
    rng = np.random.default_rng(20260408 + count)
    brands = np.array(
        [
            "Bosch", "Mann-Filter", "Sachs", "Brembo", "Mahle", "Denso",
            "Valeo", "TRW", "NGK", "Febi Bilstein", "Lemforder", "Hella",
            "Continental", "Michelin", "Varta",
        ],
        dtype=object,
    )
    categories = np.array(
        [
            "Фильтры", "Масла", "Колодки", "Диски", "Амортизаторы",
            "Аккумуляторы", "Шины", "Фары", "Двигатели", "КПП",
        ],
        dtype=object,
    )
    base_price = np.array([450, 1800, 2400, 6800, 5200, 7600, 6900, 7200, 105000, 74000])
    lengths = np.array([22, 28, 18, 62, 66, 35, 70, 52, 100, 82])
    widths = np.array([14, 16, 12, 62, 16, 26, 70, 26, 62, 58])
    heights = np.array([14, 28, 8, 22, 16, 26, 26, 26, 72, 56])
    weights = np.array([0.6, 4.4, 1.4, 9.2, 3.6, 17, 10.6, 2.6, 95, 50])
    turnovers = np.array([25, 20, 30, 45, 40, 35, 15, 50, 90, 90])

    idx = np.arange(count)
    cat_idx = idx % len(categories)
    price = np.round(base_price[cat_idx] * rng.uniform(0.72, 1.48, count) / 10) * 10
    dim_factor = rng.uniform(0.88, 1.15, count)
    brand = brands[(idx * 7 + idx % 3) % len(brands)]

    frame = pd.DataFrame(
        {
            "Артикул": [f"SKU-{i + 1:06d}" for i in range(count)],
            "Бренд": brand,
            "Категория": categories[cat_idx],
            "ID_категории": np.zeros(count, dtype=np.int64),
            "Цена": price.astype(float),
            "Себестоимость": (
                price * rng.uniform(0.54, 0.76, count) if include_cost else np.full(count, np.nan)
            ),
            "Вес_кг": weights[cat_idx] * rng.uniform(0.87, 1.14, count),
            "Длина": np.round(lengths[cat_idx] * dim_factor),
            "Ширина": np.round(widths[cat_idx] * dim_factor),
            "Высота": np.round(heights[cat_idx] * dim_factor),
            "Объем_л": np.full(count, np.nan),
            "Оборачиваемость_дней": np.round(
                turnovers[cat_idx] * rng.uniform(0.7, 1.35, count)
            ),
            # Продажи: дешёвые ходовые продаются чаще (закон Парето) — для корректного ABC
            "Продажи_шт": np.maximum(1, np.round(
                (900_000 / np.maximum(price, 100)) * rng.uniform(0.3, 2.0, count)
            )),
            "Опасный": pd.Series(pd.NA, index=np.arange(count), dtype="boolean"),
            "Хрупкий": pd.Series(pd.NA, index=np.arange(count), dtype="boolean"),
        }
    )
    return frame


# =============================================================================
# VECTORIZED CALCULATION ENGINE
# =============================================================================

def calculate_unit_economics(
    source: pd.DataFrame,
    settings: Dict[str, Any],
    progress: Optional[Any] = None,
) -> pd.DataFrame:
    """Calculate the full model without per-row Python loops."""
    started = time.perf_counter()
    df = source.copy()
    n = len(df)

    def update(value: int, text: str) -> None:
        if progress is not None:
            progress.progress(value, text=text)

    update(5, "Подготовка данных")

    cat = df["Категория"].astype("string").fillna("Без категории")
    cat_lower = cat.str.lower().str.replace("ё", "е", regex=False)

    default_volume = np.full(n, 2.0, dtype=np.float64)
    default_weight = np.full(n, 1.0, dtype=np.float64)
    default_hazard = np.zeros(n, dtype=bool)
    default_fragile = np.zeros(n, dtype=bool)

    for key, values in CATEGORY_DEFAULTS.items():
        mask = cat_lower.str.contains(key, regex=False, na=False).to_numpy()
        default_volume[mask] = values[0]
        default_weight[mask] = values[1]
        default_hazard[mask] = values[2]
        default_fragile[mask] = values[3]

    for custom in settings.get("custom_categories", []):
        key = normalize_header(str(custom.get("key", "")))
        if not key:
            continue
        mask = cat_lower.str.contains(key, regex=False, na=False).to_numpy()
        default_volume[mask] = float(custom.get("volume_l", 2.0))
        default_weight[mask] = float(custom.get("weight_kg", 1.0))
        default_hazard[mask] = bool(custom.get("is_hazardous", False))
        default_fragile[mask] = bool(custom.get("is_fragile", False))

    update(18, "Габариты, объём и оплачиваемый вес")

    length = pd.to_numeric(df["Длина"], errors="coerce").fillna(0).clip(lower=0).to_numpy()
    width = pd.to_numeric(df["Ширина"], errors="coerce").fillna(0).clip(lower=0).to_numpy()
    height = pd.to_numeric(df["Высота"], errors="coerce").fillna(0).clip(lower=0).to_numpy()
    has_dimensions = (length > 0) & (width > 0) & (height > 0)
    dimension_volume = length * width * height / 1000.0

    supplied_volume = pd.to_numeric(df["Объем_л"], errors="coerce").to_numpy()
    volume = np.where(
        np.isfinite(supplied_volume) & (supplied_volume > 0),
        supplied_volume,
        np.where(has_dimensions, dimension_volume, default_volume),
    )

    supplied_weight = pd.to_numeric(df["Вес_кг"], errors="coerce").to_numpy()
    estimated_weight = np.maximum(0.1, volume * float(settings["density_kg_per_liter"]))
    has_weight = np.isfinite(supplied_weight) & (supplied_weight > 0)
    # ВАЖНО: не делим слепо при >100 кг (двигатель 95 кг легитимен!).
    # Граммы определяем только если плотность физически невозможна (>12 кг/л)
    raw_density = np.divide(supplied_weight, np.maximum(volume, 0.001),
                            out=np.zeros_like(volume), where=has_weight)
    looks_like_grams = has_weight & (supplied_weight > 100) & (raw_density > MAX_DENSITY)
    grams_fixed_density = np.divide(supplied_weight / 1000.0, np.maximum(volume, 0.001),
                                    out=np.zeros_like(volume), where=has_weight)
    convert_grams = looks_like_grams & (grams_fixed_density >= MIN_DENSITY) & (grams_fixed_density <= MAX_DENSITY)
    corrected_weight = np.where(convert_grams, supplied_weight / 1000.0, supplied_weight)
    weight = np.where(
        has_weight,
        corrected_weight,
        np.where(has_dimensions, estimated_weight, default_weight),
    )
    volumetric_weight = np.where(has_dimensions, length * width * height / 5000.0, 0.0)
    billable_weight = np.maximum(np.maximum(weight, volumetric_weight), 0.1)

    # Валидация габаритов и веса по лимитам FBS
    sum_sides = length + width + height
    density = np.divide(weight, np.maximum(volume, 0.001), out=np.zeros_like(volume), where=volume > 0)
    max_side = np.maximum(np.maximum(length, width), height)
    validation = np.full(n, "", dtype=object)
    def add_flag(mask: np.ndarray, text: str) -> None:
        idxs = np.where(mask)[0]
        for i in idxs:
            validation[i] = (validation[i] + " · " + text) if validation[i] else text
    add_flag(~has_dimensions, "нет габаритов")
    add_flag(~has_weight & has_dimensions, "вес оценён по объёму")
    add_flag(~has_weight & ~has_dimensions, "вес из справочника")
    add_flag(convert_grams, "вес переведён из граммов")
    add_flag(looks_like_grams & ~convert_grams, "плотность неправдоподобна")
    add_flag(has_dimensions & (max_side > 220), "превышены габариты СЦ")
    add_flag(has_dimensions & (sum_sides > 150), "сумма сторон > 150 см (только СЦ)")
    add_flag(weight > 100, "вес > 100 кг — не примет СЦ")
    add_flag((weight > 20) & (weight <= 100), "вес > 20 кг — только СЦ, не ПВЗ")

    update(32, "Себестоимость и специальные расходы")

    price = pd.to_numeric(df["Цена"], errors="coerce").fillna(0).clip(lower=0).to_numpy()
    supplied_cost = pd.to_numeric(df["Себестоимость"], errors="coerce").to_numpy()
    cost_estimated = ~np.isfinite(supplied_cost) | (supplied_cost <= 0)
    cost = np.where(cost_estimated, price * float(settings["cost_fallback_rate"]), supplied_cost)

    supplied_hazard = df["Опасный"].astype("boolean")
    supplied_fragile = df["Хрупкий"].astype("boolean")
    hazard = supplied_hazard.fillna(pd.Series(default_hazard, index=df.index)).to_numpy(dtype=bool)
    fragile = supplied_fragile.fillna(pd.Series(default_fragile, index=df.index)).to_numpy(dtype=bool)

    turnover = (
        pd.to_numeric(df["Оборачиваемость_дней"], errors="coerce")
        .fillna(30.0)
        .clip(lower=1.0)
        .to_numpy()
    )
    # Продажи за период — критично для экономически верного ABC
    if "Продажи_шт" in df.columns:
        qty = pd.to_numeric(df["Продажи_шт"], errors="coerce").fillna(1.0).clip(lower=0.0).to_numpy()
        qty = np.where(qty <= 0, 1.0, qty)
    else:
        qty = np.ones(n)

    special_costs = (
        float(settings["packaging"])
        + float(settings["chestny_znak"])
        + float(settings["labeling"])
        + price * float(settings["warranty_reserve"])
        + np.where(hazard, price * float(settings["hazard_surcharge"]), 0.0)
        + np.where(fragile, price * float(settings["fragile_surcharge"]), 0.0)
    )

    update(48, "Комиссии и специальные тарифы")

    commission_rate = np.full(n, float(settings["commission_rate"]), dtype=np.float64)
    logistics_base = np.full(n, float(settings["logistics_base"]), dtype=np.float64)
    storage_rate = np.full(
        n, float(settings["storage_per_day_per_liter"]), dtype=np.float64
    )
    special_applied = np.zeros(n, dtype=bool)
    special_reason = np.full(n, "", dtype=object)

    # Custom category overrides (before special tariffs, so special still wins if enabled)
    for custom in settings.get("custom_categories", []):
        key = normalize_header(str(custom.get("key", "")))
        if not key:
            continue
        mask = cat_lower.str.contains(key, regex=False, na=False).to_numpy()
        if custom.get("commission_rate") is not None:
            try:
                commission_rate[mask] = float(custom["commission_rate"])
            except Exception:
                pass
        if custom.get("logistics_base") is not None:
            try:
                logistics_base[mask] = float(custom["logistics_base"])
            except Exception:
                pass
        if custom.get("storage_per_day_per_liter") is not None:
            try:
                storage_rate[mask] = float(custom["storage_per_day_per_liter"])
            except Exception:
                pass

    if settings.get("special_enabled", True):
        for key, rule in settings.get("special_tariffs", {}).items():
            mask = cat_lower.str.contains(key, regex=False, na=False).to_numpy()
            commission_rate[mask] = float(rule["commission_rate"])
            logistics_base[mask] = float(rule["logistics_base"])
            storage_rate[mask] = float(rule["storage_per_day_per_liter"])
            special_applied[mask] = True
            special_reason[mask] = str(rule.get("reason", "Спецтариф"))

    # FBS speed discount: subtract from commission (only for FBS)
    fbs_discount = float(settings.get("fbs_speed_discount", 0.0)) if settings.get("model", "FBS") == "FBS" else 0.0
    if fbs_discount:
        commission_rate = np.maximum(0.01, commission_rate - fbs_discount)

    # Individual rates from API/CSV have the highest priority (already above, but FBS discount already applied — API rates are final)
    if settings.get("use_category_rates", True) and settings.get("category_rates"):
        rates = {
            normalize_header(key): float(value)
            for key, value in settings["category_rates"].items()
        }
        exact = cat_lower.map(rates)
        exact_mask = exact.notna().to_numpy()
        commission_rate[exact_mask] = np.maximum(0.01, exact[exact_mask].to_numpy(dtype=float) - fbs_discount) if fbs_discount else exact[exact_mask].to_numpy(dtype=float)

        # Substring fallback for a small rate dictionary.
        missing_rate = ~exact_mask
        for key, value in rates.items():
            if not missing_rate.any():
                break
            mask = (
                cat_lower.str.contains(key, regex=False, na=False).to_numpy()
                & missing_rate
            )
            commission_rate[mask] = max(0.01, value - fbs_discount) if fbs_discount else value
            missing_rate[mask] = False

    update(65, "Логистика (ср. миля 2026), хранение и эквайринг")

    # Acquiring rate from payout schedule
    acquiring_rate = float(PAYOUT_SCHEDULE_RATES.get(settings.get("payout_schedule", "weekly_2w"), float(settings.get("acquiring_fee", 0.023))))

    # Fixed tariff detection: price ≤300 and volume ≤5 and enabled (new 2026)
    use_fixed = bool(settings.get("use_fixed_tariff", True))
    model = settings.get("model", "FBS")
    fixed_rate_fbs = 0.42 - (fbs_discount if model == "FBS" else 0)
    fixed_rate_fby = 0.35
    fixed_rate_val = fixed_rate_fbs if model == "FBS" else fixed_rate_fby
    fixed_rate_arr = np.full(n, fixed_rate_val, dtype=float)
    is_fixed = use_fixed & (price > 0) & (price <= 300) & (volume <= 5)

    # Srednya mila (new grid 2026) — объёмная, не зависит от веса
    vector_srednya = np.vectorize(srednya_mila_py)
    srednya = vector_srednya(volume)

    # Commission / logistics / storage with fixed branch
    commission = np.where(is_fixed, 0.0, np.maximum(price * commission_rate, float(settings["min_commission"])))
    logistics = np.where(is_fixed, price * fixed_rate_arr, srednya)
    # Storage: FBS = 0 (your warehouse), FBY = volume * rate * max(0, turnover - freeDays)
    storage_free = int(settings.get("storage_free_days", 90))
    if model == "FBS":
        storage = np.zeros(n)
    else:
        paid_days = np.maximum(0, turnover - storage_free)
        storage = volume * storage_rate * paid_days

    acquiring = price * acquiring_rate
    # Returns: for fixed tariff returns are included, so 0; else reserve
    returns = np.where(is_fixed, 0.0, price * float(settings["return_fee"]))

    total_expenses = (
        cost + commission + logistics + storage + acquiring + returns + special_costs
    )
    profit = price - total_expenses
    margin = np.divide(profit, price, out=np.zeros_like(profit), where=price > 0)

    # Recommended breakeven price (vectorized) — учитывает фиксированный тариф
    spec_fixed = float(settings["packaging"]) + float(settings["chestny_znak"]) + float(settings["labeling"])
    # variable rate for non-fixed: commission + acquiring + warranty + surcharges + return
    # for fixed: fixedRate + acquiring + warranty + surcharges (returns already in fixed)
    variable_rate_std = (
        commission_rate
        + acquiring_rate
        + float(settings["warranty_reserve"])
        + np.where(hazard, float(settings["hazard_surcharge"]), 0.0)
        + np.where(fragile, float(settings["fragile_surcharge"]), 0.0)
        + float(settings["return_fee"])
    )
    variable_rate_fixed = (
        fixed_rate_arr
        + acquiring_rate
        + float(settings["warranty_reserve"])
        + np.where(hazard, float(settings["hazard_surcharge"]), 0.0)
        + np.where(fragile, float(settings["fragile_surcharge"]), 0.0)
    )
    variable_rate = np.where(is_fixed, variable_rate_fixed, variable_rate_std)
    denom = 1.0 - variable_rate
    denom = np.where(denom < 0.05, 0.05, denom)
    # Для обычного тарифа logistics = srednya (константа), для фиксированного logistics = price*fixed (уже в variable)
    # Поэтому в формуле: для fixed используем cost+storage+specFixed, для std — cost+logistics+storage+specFixed
    numer_std = cost + logistics + storage + spec_fixed
    numer_fixed = cost + storage + spec_fixed
    numer = np.where(is_fixed, numer_fixed, numer_std)
    # Для fixed: denom already includes fixedRate, для std: denom includes commission
    recommended = numer / denom * 1.01
    # Корректировка минимума комиссии только для не-фиксированных
    low_comm_mask = (~is_fixed) & (recommended * commission_rate < float(settings["min_commission"]))
    if np.any(low_comm_mask):
        denom2 = 1.0 - (
            acquiring_rate
            + float(settings["warranty_reserve"])
            + np.where(hazard, float(settings["hazard_surcharge"]), 0.0)
            + np.where(fragile, float(settings["fragile_surcharge"]), 0.0)
            + float(settings["return_fee"])
        )
        denom2 = np.where(denom2 < 0.05, 0.05, denom2)
        recommended_alt = (cost + logistics + storage + spec_fixed + float(settings["min_commission"])) / denom2 * 1.01
        recommended = np.where(low_comm_mask, recommended_alt, recommended)
    recommended = np.maximum(recommended, cost + 10)
    recommended = np.where(price > 0, recommended, 0)

    # Price with markup / target margin — пересчёт по тем же правилам
    pricing = settings.get("pricing", {"mode": "none", "markupPercent": 15, "targetMargin": 0.20})
    mode = pricing.get("mode", "none")
    markup_pct = float(pricing.get("markupPercent", 15))
    target_margin = float(pricing.get("targetMargin", 0.20))
    price_with_markup = price.copy()
    if mode == "markup" and markup_pct != 0:
        price_with_markup = price * (1 + markup_pct / 100)
    elif mode == "targetMargin":
        # Для фиксированных: logistics фиксированной доли, storage=0
        if np.any(is_fixed):
            # Пересчитаем target цену отдельно для fixed и std
            # Для fixed: numer = cost+storage+specFixed, denom = 1 - fixedRate - acquiring - warranty - surcharges - targetMargin
            fixed_var = fixed_rate_arr + acquiring_rate + float(settings["warranty_reserve"]) + np.where(hazard, float(settings["hazard_surcharge"]), 0.0) + np.where(fragile, float(settings["fragile_surcharge"]), 0.0)
            target_fixed = (cost + storage + spec_fixed) / np.maximum(0.05, 1 - fixed_var - target_margin)
            price_with_markup = np.where(is_fixed, target_fixed, price_with_markup)
            # min commission не применяется к fixed
            low_target_std = (~is_fixed) & (price_with_markup * commission_rate < float(settings["min_commission"]))
        else:
            low_target_std = price_with_markup * commission_rate < float(settings["min_commission"])
        if np.any(low_target_std):
            denom_t2 = 1.0 - (
                acquiring_rate
                + float(settings["warranty_reserve"])
                + np.where(hazard, float(settings["hazard_surcharge"]), 0.0)
                + np.where(fragile, float(settings["fragile_surcharge"]), 0.0)
                + float(settings["return_fee"])
                + target_margin
            )
            denom_t2 = np.where(denom_t2 < 0.05, 0.05, denom_t2)
            price_alt = (cost + srednya + storage + spec_fixed + float(settings["min_commission"])) / denom_t2
            price_with_markup = np.where(low_target_std, price_alt, price_with_markup)
        price_with_markup = np.maximum(price_with_markup, cost + 10)

    # Пересчёт расходов для новой цены с учётом фиксированного тарифа
    commission_markup = np.where(is_fixed, 0.0, np.maximum(price_with_markup * commission_rate, float(settings["min_commission"])))
    # Для fixed logistics = price * fixedRate, иначе srednya (пересчитываем на случай если объём изменился? объём константа)
    logistics_markup = np.where(is_fixed, price_with_markup * fixed_rate_arr, srednya)
    storage_markup = storage  # storage не зависит от новой цены (для FBS 0, для FBY от объёма)
    acquiring_markup = price_with_markup * acquiring_rate
    returns_markup = np.where(is_fixed, 0.0, price_with_markup * float(settings["return_fee"]))
    special_markup = (
        spec_fixed
        + price_with_markup * float(settings["warranty_reserve"])
        + np.where(hazard, price_with_markup * float(settings["hazard_surcharge"]), 0.0)
        + np.where(fragile, price_with_markup * float(settings["fragile_surcharge"]), 0.0)
    )
    total_markup = cost + commission_markup + logistics_markup + storage_markup + acquiring_markup + returns_markup + special_markup
    profit_markup = price_with_markup - total_markup
    margin_markup = np.divide(profit_markup, price_with_markup, out=np.zeros_like(profit_markup), where=price_with_markup > 0)

    total_rev = float(np.sum(price))
    revenue_share = np.divide(price, total_rev, out=np.zeros_like(price), where=total_rev > 0)

    update(82, "Формирование итоговой таблицы")

    result = pd.DataFrame(
        {
            "Артикул": df["Артикул"].astype("string"),
            "Бренд": df["Бренд"].astype("string"),
            "Категория": cat,
            "ID_категории": df["ID_категории"].astype("int64"),
            "Длина": length,
            "Ширина": width,
            "Высота": height,
            "Объем_л": volume,
            "Вес_кг": weight,
            "Оплач_вес": billable_weight,
            "Цена": price,
            "Себестоимость": cost,
            "Себестоимость_оценка": cost_estimated,
            "is_hazardous": hazard,
            "is_fragile": fragile,
            "Ставка_комиссии": commission_rate,
            "Логистика_база": logistics_base,
            "Ставка_за_кг": np.full(n, float(settings["logistics_per_kg"])),
            "Ставка_хранения": storage_rate,
            "Комиссия_руб": commission,
            "Логистика_руб": logistics,
            "Хранение_руб": storage,
            "Эквайринг_руб": acquiring,
            "Возвраты_руб": returns,
            "Спец_расходы_FBS": special_costs,
            "Итого_расходы": total_expenses,
            "Прибыль": profit,
            "Маржа_%": margin,
            "Продажи_шт": qty,
            "Выручка": price * qty,
            "Прибыль_период": profit * qty,
            "Валидация": np.where(validation == "", "OK", validation),
            "Сумма_сторон": sum_sides,
            "Плотность": density,
            "Фиксированный_тариф": is_fixed,
            "Рекомендованная_цена": recommended,
            "Цена_с_наценкой": price_with_markup,
            "Прибыль_с_наценкой": profit_markup,
            "Маржа_с_наценкой_%": margin_markup,
            "Выручка_доля": revenue_share,
            "Оборачиваемость_дней": turnover,
            "Спецтариф_применён": special_applied,
            "Причина_спецтарифа": special_reason,
        }
    )
    # Рекомендация по цене
    result["Рекомендация"] = np.where(
        result["Прибыль"] < 0,
        "↑ Поднять до " + result["Рекомендованная_цена"].round(0).astype(str) + " ₽",
        np.where(result["Маржа_%"] < 0.05, "⚠ Критично: <5%", np.where(result["Маржа_%"] < 0.15, "→ Можно +10%", "✓ ОК")),
    )
    # Добавляем пометку про фикс
    result.loc[result["Фиксированный_тариф"], "Рекомендация"] = result.loc[result["Фиксированный_тариф"], "Рекомендация"] + " · фикс 42/35%"
    # ABC / XYZ — экономически корректно: ABC по ВЫРУЧКЕ ЗА ПЕРИОД (Цена × Продажи_шт)
    if n > 0:
        revenue_period = price * qty
        order = np.argsort(-revenue_period, kind="stable")
        cumsum = np.cumsum(revenue_period[order])
        total = cumsum[-1] if cumsum.size else 1
        abc = np.full(n, "C", dtype=object)
        cum_share = cumsum / (total if total > 0 else 1)
        abc[cum_share <= 0.80] = "A"
        mask_b = (cum_share > 0.80) & (cum_share <= 0.95)
        abc[mask_b] = "B"
        # Товары без выручки — всегда C
        abc[revenue_period[order] <= 0] = "C"
        abc_final = np.empty(n, dtype=object)
        abc_final[order] = abc
        cum_final = np.empty(n, dtype=float)
        cum_final[order] = cum_share
        # XYZ по АБСОЛЮТНЫМ порогам оборачиваемости (терцили дают ложные 33/33/33)
        xyz = np.where(turnover <= XYZ_X_DAYS, "X", np.where(turnover <= XYZ_Y_DAYS, "Y", "Z"))
        result["ABC"] = abc_final
        result["ABC_накопл"] = cum_final
        result["XYZ"] = xyz
        result["ABC_XYZ"] = result["ABC"].astype(str) + result["XYZ"].astype(str)
        result["ABC_XYZ_смысл"] = result["ABC_XYZ"].map(ABC_XYZ_MEANING).fillna("")
        # Текстовая рекомендация по цене
        result["Рекомендация"] = np.where(
            result["Прибыль"] < 0,
            "↑ Поднять до " + result["Рекомендованная_цена"].round(0).astype(str) + " ₽",
            np.where(result["Маржа_%"] < 0.05, "⚠ Критично: <5%", np.where(result["Маржа_%"] < 0.15, "→ Можно +10%", "✓ ОК"))
        )
    else:
        result["ABC"] = pd.Series(dtype="string")
        result["ABC_накопл"] = pd.Series(dtype="float")
        result["XYZ"] = pd.Series(dtype="string")
        result["ABC_XYZ"] = pd.Series(dtype="string")
        result["ABC_XYZ_смысл"] = pd.Series(dtype="string")
        result["Рекомендация"] = pd.Series(dtype="string")

    result.attrs["calculation_seconds"] = time.perf_counter() - started
    update(100, "Готово")
    return result


def run_calculation() -> None:
    raw = st.session_state.raw_df
    if raw is None or raw.empty:
        st.warning("Сначала загрузите каталог.")
        return

    progress = st.progress(0, text="Подготовка расчёта")
    started = time.perf_counter()
    try:
        result = calculate_unit_economics(raw, st.session_state.settings, progress)
        st.session_state.result_df = result
        st.session_state.calculation_seconds = time.perf_counter() - started
        st.session_state.calculated_settings_hash = stable_settings_hash(
            st.session_state.settings
        )
        clear_export()
        time.sleep(0.08)
        progress.empty()
        st.success(
            f"Рассчитано {len(result):,} SKU за "
            f"{st.session_state.calculation_seconds:.2f} с".replace(",", " ")
        )
    except Exception as exc:
        progress.empty()
        logger.exception("Calculation failed")
        st.error(f"Ошибка расчёта: {exc}")


# =============================================================================
# AGGREGATIONS AND ANALYTICS
# =============================================================================

def calculate_totals(df: pd.DataFrame) -> Dict[str, float]:
    # Экономически верно: выручка и прибыль ЗА ПЕРИОД (× продажи)
    qty = df["Продажи_шт"] if "Продажи_шт" in df.columns else 1
    revenue = float((df["Цена"] * qty).sum())
    expenses = float((df["Итого_расходы"] * qty).sum())
    profit = float((df["Прибыль"] * qty).sum())
    return {
        "count": len(df),
        "revenue": revenue,
        "expenses": expenses,
        "profit": profit,
        "margin": profit / revenue if revenue > 0 else 0.0,
        "loss_count": int(df["Прибыль"].lt(0).sum()),
        "special_count": int(df["Спецтариф_применён"].sum()),
        "estimated_cost_count": int(df["Себестоимость_оценка"].sum()),
        "commission": float(df["Комиссия_руб"].sum()),
        "logistics": float(df["Логистика_руб"].sum()),
        "storage": float(df["Хранение_руб"].sum()),
        "acquiring": float(df["Эквайринг_руб"].sum()),
        "returns": float(df["Возвраты_руб"].sum()),
        "special_costs": float(df["Спец_расходы_FBS"].sum()),
        "cost": float(df["Себестоимость"].sum()),
    }


def aggregate_by(df: pd.DataFrame, column: str) -> pd.DataFrame:
    source = df[
        [column, "Артикул", "Цена", "Итого_расходы", "Прибыль", "Спецтариф_применён"]
    ].copy()
    qty = df["Продажи_шт"] if "Продажи_шт" in df.columns else 1
    # Агрегируем за период: выручка = цена × продажи
    source["Выручка_п"] = df["Цена"] * qty
    source["Расходы_п"] = df["Итого_расходы"] * qty
    source["Прибыль_п"] = df["Прибыль"] * qty
    source["_loss"] = source["Прибыль"].lt(0).astype("int8")
    grouped = (
        source.groupby(column, observed=True, sort=False)
        .agg(
            SKU=("Артикул", "size"),
            Выручка=("Выручка_п", "sum"),
            Расходы=("Расходы_п", "sum"),
            Прибыль=("Прибыль_п", "sum"),
            Убыточных_SKU=("_loss", "sum"),
            Спецтариф_SKU=("Спецтариф_применён", "sum"),
        )
        .reset_index()
    )
    grouped["Маржа_%"] = np.divide(
        grouped["Прибыль"],
        grouped["Выручка"],
        out=np.zeros(len(grouped), dtype=float),
        where=grouped["Выручка"].to_numpy() > 0,
    )
    return grouped.sort_values("Выручка", ascending=False, kind="stable")


def cost_structure(totals: Dict[str, float]) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Статья": [
                "Себестоимость", "Комиссия", "Логистика", "Хранение",
                "Эквайринг", "Возвраты", "Спец. расходы",
            ],
            "Сумма": [
                totals["cost"], totals["commission"], totals["logistics"],
                totals["storage"], totals["acquiring"], totals["returns"],
                totals["special_costs"],
            ],
        }
    )


# =============================================================================
# YANDEX MARKET API
# =============================================================================

def api_headers(token: str, token_type: str) -> Dict[str, str]:
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    if token_type == "Api-Key":
        headers["Api-Key"] = token.strip()
    else:
        headers["Authorization"] = f"Bearer {token.strip()}"
    return headers


def calculate_tariff_via_api(
    token: str,
    token_type: str,
    campaign_id: int,
    category_id: int,
    price: float,
    length: float,
    width: float,
    height: float,
    weight: float,
) -> Tuple[bool, str, Optional[pd.DataFrame]]:
    """Current official tariff calculator: POST /v2/tariffs/calculate."""
    url = f"{API_URL}/v2/tariffs/calculate"
    payload = {
        "parameters": {
            "campaignId": int(campaign_id),
            "frequency": "DAILY",
            "currency": "RUR",
        },
        "offers": [
            {
                "categoryId": int(category_id),
                "price": float(price),
                "length": float(length),
                "width": float(width),
                "height": float(height),
                "weight": float(weight),
                "quantity": 1,
            }
        ],
    }
    try:
        response = requests.post(
            url,
            headers=api_headers(token, token_type),
            json=payload,
            timeout=25,
        )
        if response.status_code != 200:
            return False, f"API вернул {response.status_code}: {response.text[:500]}", None

        body = response.json()
        offers = body.get("offers", body.get("result", {}).get("offers", []))
        if not offers:
            return False, "API не вернул расчёт услуг.", None

        tariffs = offers[0].get("tariffs", [])
        rows = []
        for item in tariffs:
            rows.append(
                {
                    "Тип услуги": item.get("type", ""),
                    "Сумма": float(item.get("amount", 0) or 0),
                    "Валюта": item.get("currency", "RUR"),
                    "Параметры": json.dumps(
                        item.get("parameters", []), ensure_ascii=False
                    ),
                }
            )
        return True, f"Получено услуг: {len(rows)}", pd.DataFrame(rows)
    except requests.RequestException as exc:
        return False, f"Ошибка соединения с API: {exc}", None
    except Exception as exc:
        return False, f"Ошибка разбора ответа API: {exc}", None


def legacy_category_commissions(
    token: str, token_type: str, campaign_id: int
) -> Tuple[bool, str, Dict[str, float]]:
    """Compatibility with the legacy category commission endpoint."""
    url = f"{API_URL}/v2/campaigns/{int(campaign_id)}/categories/commissions"
    try:
        response = requests.get(
            url,
            headers=api_headers(token, token_type),
            timeout=20,
        )
        if response.status_code != 200:
            return (
                False,
                "Старый метод комиссий недоступен. Используйте актуальный калькулятор "
                "POST /v2/tariffs/calculate ниже.",
                {},
            )
        categories = response.json().get("result", {}).get("categories", [])
        rates: Dict[str, float] = {}
        for category in categories:
            name = str(
                category.get("categoryName") or category.get("categoryId") or ""
            ).strip().lower()
            if name:
                rates[name] = float(category.get("commissionPercent", 14)) / 100.0
        return True, f"Загружено ставок: {len(rates)}", rates
    except Exception as exc:
        return False, f"Ошибка запроса: {exc}", {}


# =============================================================================
# EXCEL AND CSV EXPORT
# =============================================================================

def tariff_table(settings: Dict[str, Any]) -> pd.DataFrame:
    pricing = settings.get("pricing", {"mode":"none","markupPercent":15,"targetMargin":0.2})
    rows = [
        ("Комиссия", settings["commission_rate"], "доля"),
        ("Минимальная комиссия", settings["min_commission"], "₽"),
        ("Логистика: база", settings["logistics_base"], "₽"),
        ("Логистика: за кг", settings["logistics_per_kg"], "₽/кг"),
        ("Хранение", settings["storage_per_day_per_liter"], "₽/л/сутки"),
        ("Эквайринг", settings["acquiring_fee"], "доля"),
        ("Возвраты", settings["return_fee"], "доля"),
        ("Упаковка FBS", settings["packaging"], "₽"),
        ("Честный знак", settings["chestny_znak"], "₽"),
        ("Маркировка", settings["labeling"], "₽"),
        ("Гарантийный резерв", settings["warranty_reserve"], "доля"),
        ("Надбавка: опасный", settings["hazard_surcharge"], "доля"),
        ("Надбавка: хрупкий", settings["fragile_surcharge"], "доля"),
        ("Себестоимость по умолчанию", settings["cost_fallback_rate"], "доля от цены"),
        ("Плотность для оценки веса", settings["density_kg_per_liter"], "кг/л"),
        ("Режим ценообразования", pricing.get("mode","none"), ""),
        ("Наценка, %", pricing.get("markupPercent",15), "%"),
        ("Целевая маржа", pricing.get("targetMargin",0.2), "доля"),
        ("Фиксированный тариф ≤300₽/5л", "Да" if settings.get("use_fixed_tariff", True) else "Нет", "B20"),
        ("Скидка FBS за скорость", settings.get("fbs_speed_discount", 0.0), "доля"),
        ("Льготных дней хранения FBY", settings.get("storage_free_days", 90), "дней"),
        ("Модель работы", settings.get("model", "FBS"), "FBS/FBY"),
    ]
    if settings.get("custom_categories"):
        for cat in settings["custom_categories"]:
            rows.append((f"Категория: {cat.get('label', cat.get('key'))}", f"{cat.get('volume_l')}л/{cat.get('weight_kg')}кг", f"{'опасн' if cat.get('is_hazardous') else ''}"))
    return pd.DataFrame(rows, columns=["Параметр", "Значение", "Единица"])


def update_settings_from_tariff_csv(
    data: bytes, current: Dict[str, Any]
) -> Tuple[bool, str, Dict[str, Any]]:
    """Support both the old one-row tariff CSV and the new parameter table."""
    try:
        encoding, separator = detect_csv_format(data)
        table = pd.read_csv(io.BytesIO(data), sep=separator, encoding=encoding)
        if table.empty:
            return False, "CSV тарифа пуст.", current

        updated = deep_copy_json(current)
        technical_keys = {
            "commission_rate", "min_commission", "logistics_base",
            "logistics_per_kg", "storage_per_day_per_liter", "acquiring_fee",
            "return_fee", "packaging", "chestny_znak", "labeling",
            "warranty_reserve", "hazard_surcharge", "fragile_surcharge",
            "cost_fallback_rate", "density_kg_per_liter",
        }

        # Original format: one row with technical names as columns.
        if technical_keys.intersection(table.columns):
            row = table.iloc[0]
            changed = 0
            for key in technical_keys:
                if key in table.columns and pd.notna(row[key]):
                    updated[key] = float(str(row[key]).replace(",", "."))
                    changed += 1
            return True, f"Обновлено параметров: {changed}", updated

        # New format produced by the application itself.
        parameter_column = next(
            (c for c in table.columns if normalize_header(c) == "параметр"), None
        )
        value_column = next(
            (c for c in table.columns if normalize_header(c) == "значение"), None
        )
        if parameter_column is None or value_column is None:
            return False, "Нужны колонки «Параметр» и «Значение».", current

        label_to_key = {
            "комиссия": "commission_rate",
            "минимальная комиссия": "min_commission",
            "логистика: база": "logistics_base",
            "логистика: за кг": "logistics_per_kg",
            "хранение": "storage_per_day_per_liter",
            "эквайринг": "acquiring_fee",
            "возвраты": "return_fee",
            "упаковка fbs": "packaging",
            "честный знак": "chestny_znak",
            "маркировка": "labeling",
            "гарантийный резерв": "warranty_reserve",
            "надбавка: опасный": "hazard_surcharge",
            "надбавка: хрупкий": "fragile_surcharge",
            "себестоимость по умолчанию": "cost_fallback_rate",
            "плотность для оценки веса": "density_kg_per_liter",
        }
        changed = 0
        for _, row in table.iterrows():
            label = normalize_header(row[parameter_column])
            key = label_to_key.get(label)
            if key and pd.notna(row[value_column]):
                updated[key] = float(str(row[value_column]).replace(",", "."))
                changed += 1
        if not changed:
            return False, "В CSV не найдены известные параметры тарифа.", current
        return True, f"Обновлено параметров: {changed}", updated
    except Exception as exc:
        return False, f"Ошибка CSV тарифа: {exc}", current


def parse_category_csv(data: bytes, existing: List[Dict[str, Any]]) -> Tuple[bool, str, List[Dict[str, Any]]]:
    try:
        encoding, separator = detect_csv_format(data)
        table = pd.read_csv(io.BytesIO(data), sep=separator, encoding=encoding)
        if table.empty:
            return False, "CSV категорий пуст.", existing
        table.columns = [normalize_header(c) for c in table.columns]
        required = ["key", "label"]
        for req in required:
            if req not in table.columns:
                return False, f"Нет колонки {req} в CSV категорий.", existing
        updated = deep_copy_json(existing)
        existing_keys = {normalize_header(str(c.get("key",""))) : i for i,c in enumerate(updated)}
        added = 0
        for _, row in table.iterrows():
            key = str(row.get("key","")).strip()
            if not key:
                continue
            nkey = normalize_header(key)
            cfg = {
                "key": key.lower().strip(),
                "label": str(row.get("label","")).strip() or key,
                "volume_l": float(str(row.get("volume_l", 2)).replace(",",".")) if pd.notna(row.get("volume_l")) else 2.0,
                "weight_kg": float(str(row.get("weight_kg", 1)).replace(",",".")) if pd.notna(row.get("weight_kg")) else 1.0,
                "is_hazardous": str(row.get("is_hazardous","")).lower() in ["true","1","да","yes"],
                "is_fragile": str(row.get("is_fragile","")).lower() in ["true","1","да","yes"],
                "commission_rate": None,
                "logistics_base": None,
                "storage_per_day_per_liter": None,
            }
            if "commission_rate" in table.columns and pd.notna(row.get("commission_rate")):
                v = float(str(row.get("commission_rate")).replace(",","."))
                cfg["commission_rate"] = v/100 if v>1 else v
            if "logistics_base" in table.columns and pd.notna(row.get("logistics_base")):
                cfg["logistics_base"] = float(str(row.get("logistics_base")).replace(",","."))
            if "storage_per_day_per_liter" in table.columns and pd.notna(row.get("storage_per_day_per_liter")):
                cfg["storage_per_day_per_liter"] = float(str(row.get("storage_per_day_per_liter")).replace(",","."))
            if nkey in existing_keys:
                updated[existing_keys[nkey]] = cfg
            else:
                updated.append(cfg)
                added += 1
        return True, f"Категорий добавлено/обновлено: {len(table)} (новых {added})", updated
    except Exception as exc:
        return False, f"Ошибка CSV категорий: {exc}", existing


def get_saved_datasets() -> List[Dict[str, Any]]:
    try:
        path = DATA_DIR / "saved_datasets.json"
        if not path.exists():
            return []
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_dataset(name: str, raw_df: pd.DataFrame, parse_meta: Dict[str, Any]) -> str:
    try:
        datasets = get_saved_datasets()
        # Store raw data as parquet-like via pickle base64? For simplicity store as csv string
        # But to avoid huge JSON, save to separate pickle file
        ds_id = datetime.now().strftime("%Y%m%d_%H%M%S_") + str(len(datasets))
        pickle_path = DATA_DIR / f"dataset_{ds_id}.pkl"
        raw_df.to_pickle(pickle_path)
        entry = {
            "id": ds_id,
            "name": name,
            "createdAt": datetime.now().isoformat(),
            "rows": len(raw_df),
            "fileName": parse_meta.get("fileName", name),
            "pickle": str(pickle_path),
            "parse_meta": parse_meta,
        }
        datasets.append(entry)
        # keep last 20
        datasets = datasets[-20:]
        with (DATA_DIR / "saved_datasets.json").open("w", encoding="utf-8") as f:
            json.dump(datasets, f, ensure_ascii=False, indent=2)
        return ds_id
    except Exception as exc:
        logger.exception("save dataset failed")
        return ""


def load_dataset(ds_id: str) -> Tuple[Optional[pd.DataFrame], Dict[str, Any], str]:
    try:
        datasets = get_saved_datasets()
        for ds in datasets:
            if ds["id"] == ds_id:
                p = Path(ds["pickle"])
                if p.exists():
                    df = pd.read_pickle(p)
                    return df, ds.get("parse_meta", {}), ds.get("name","")
        return None, {}, ""
    except Exception:
        return None, {}, ""


def delete_dataset(ds_id: str) -> None:
    try:
        datasets = get_saved_datasets()
        new_list = []
        for ds in datasets:
            if ds["id"] == ds_id:
                try:
                    Path(ds["pickle"]).unlink(missing_ok=True)
                except Exception:
                    pass
            else:
                new_list.append(ds)
        with (DATA_DIR / "saved_datasets.json").open("w", encoding="utf-8") as f:
            json.dump(new_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def legend_table() -> pd.DataFrame:
    rows = [
        ("Артикул", "Уникальный идентификатор товара", "Файл пользователя"),
        ("Бренд", "Производитель / торговая марка", "Файл или 'Без бренда'"),
        ("Категория", "Группа товара для тарифов и аналитики", "Файл пользователя"),
        ("Объем_л", "Д×Ш×В / 1000", "Габариты / файл / справочник"),
        ("Оплач_вес", "MAX(вес; Д×Ш×В / 5000; 0,1)", "Формула"),
        ("Себестоимость", "Закупочная цена или % от цены", "Файл / настройки"),
        ("Комиссия_руб", "MAX(Цена × ставка; мин. комиссия)", "Тариф"),
        ("Логистика_руб", "База + оплач. вес × ставка за кг", "Тариф + габариты"),
        ("Хранение_руб", "Объём × ставка × дни", "Тариф + оборачиваемость"),
        ("Эквайринг_руб", "Цена × ставка эквайринга", "Тариф"),
        ("Возвраты_руб", "Цена × резерв возвратов", "Тариф"),
        ("Спец_расходы_FBS", "Упаковка, маркировка, резервы", "Настройки"),
        ("Рекомендованная_цена", "Безубыточная +1% (покрывает все сборы)", "Формула"),
        ("Цена_с_наценкой", "Цена после наценки / целевой маржи", "Формула (B18/B19 тарифа)"),
        ("Прибыль_с_наценкой", "Прибыль по новой цене", "Формула"),
        ("Маржа_с_наценкой_%", "Маржа по новой цене", "Формула"),
        ("ABC", "A 80%, B 15%, C 5% выручки", "Расчёт"),
        ("XYZ", "X стабильные, Y средние, Z нестабильные", "По оборачиваемости"),
        ("ABC_XYZ", "Матрица 9 ячеек", "Расчёт"),
        ("Итого_расходы", "Себестоимость + все расходы", "Формула"),
        ("Прибыль", "Цена - итого расходы", "Формула"),
        ("Маржа_%", "Прибыль / Цена", "Формула"),
    ]
    return pd.DataFrame(rows, columns=["Колонка", "Описание", "Источник / логика"])


def configure_excel_formats(workbook: Any) -> Dict[str, Any]:
    return {
        "title": workbook.add_format(
            {"bold": True, "font_size": 16, "font_color": "#0F3460"}
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#0F172A",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "header_accent": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#4F46E5",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "header_warm": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#EA580C",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "money": workbook.add_format({"num_format": "#,##0.00", "border": 1, "border_color": "#E2E8F0"}),
        "money_bold": workbook.add_format({"num_format": "#,##0.00", "border": 1, "bold": True, "border_color": "#E2E8F0"}),
        "percent": workbook.add_format({"num_format": "0.00%", "border": 1, "border_color": "#E2E8F0"}),
        "percent_bold": workbook.add_format({"num_format": "0.00%", "border": 1, "bold": True, "border_color": "#E2E8F0"}),
        "number": workbook.add_format({"num_format": "#,##0.00", "border": 1, "border_color": "#E2E8F0"}),
        "red": workbook.add_format({"bg_color": "#FECACA", "font_color": "#991B1B", "bold": True, "border": 1, "border_color": "#FCA5A5"}),
        "red_light": workbook.add_format({"bg_color": "#FFF1F2", "font_color": "#9F1239", "border": 1}),
        "green": workbook.add_format({"bg_color": "#A7F3D0", "font_color": "#065F46", "bold": True, "border": 1, "border_color": "#6EE7B7"}),
        "green_light": workbook.add_format({"bg_color": "#ECFDF5", "font_color": "#065F46", "border": 1}),
        "amber": workbook.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E", "bold": True, "border": 1, "border_color": "#FDE68A"}),
        "amber_light": workbook.add_format({"bg_color": "#FFFBEB", "font_color": "#92400E", "border": 1}),
        "blue": workbook.add_format({"bg_color": "#DBEAFE", "font_color": "#1E40AF", "bold": True, "border": 1}),
        "band": workbook.add_format({"bg_color": "#F8FAFC", "border": 1, "border_color": "#E2E8F0"}),
    }


def write_dataframe_header(
    worksheet: Any, dataframe: pd.DataFrame, header_format: Any
) -> None:
    for col, value in enumerate(dataframe.columns):
        worksheet.write(0, col, value, header_format)


def write_summary_sheet(
    writer: pd.ExcelWriter,
    sheet_name: str,
    group_name: str,
    summary: pd.DataFrame,
    formats: Dict[str, Any],
) -> None:
    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
    ws = writer.sheets[sheet_name]
    ws.write(0, 0, f"СВОДКА: {group_name.upper()}", formats["title"])
    for col, value in enumerate(summary.columns):
        ws.write(2, col, value, formats["header"])
    ws.freeze_panes(3, 0)
    ws.autofilter(2, 0, len(summary) + 2, len(summary.columns) - 1)
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 4, 16, formats["money"])
    ws.set_column(5, 5, 12, formats["percent"])
    ws.set_column(6, 7, 15)

    if not summary.empty:
        # ── Максимальная подсветка сводки ──
        # Маржа: 4 уровня (красный <0, светло-красный 0-5%, янтарь 5-15%, зелёный ≥15%)
        margin_col = list(summary.columns).index("Маржа_%") if "Маржа_%" in summary.columns else (5 if summary.shape[1] > 5 else -1)
        profit_col = list(summary.columns).index("Прибыль") if "Прибыль" in summary.columns else 4
        revenue_col = list(summary.columns).index("Выручка") if "Выручка" in summary.columns else 2
        loss_col = list(summary.columns).index("Убыточных_SKU") if "Убыточных_SKU" in summary.columns else -1
        # Градации маржи + светофор
        if margin_col >= 0:
            ws.conditional_format(3, margin_col, len(summary)+2, margin_col, {"type":"cell","criteria":"<","value":0,"format":formats["red"]})
            ws.conditional_format(3, margin_col, len(summary)+2, margin_col, {"type":"cell","criteria":"between","minimum":0,"maximum":0.05,"format":formats["red_light"]})
            ws.conditional_format(3, margin_col, len(summary)+2, margin_col, {"type":"cell","criteria":"between","minimum":0.05,"maximum":0.15,"format":formats["amber"]})
            ws.conditional_format(3, margin_col, len(summary)+2, margin_col, {"type":"cell","criteria":">=","value":0.15,"format":formats["green"]})
            ws.conditional_format(3, margin_col, len(summary)+2, margin_col, {"type":"icon_set","icon_style":"3_traffic_lights"})
        if profit_col >= 0:
            ws.conditional_format(3, profit_col, len(summary)+2, profit_col, {"type":"cell","criteria":"<","value":0,"format":formats["red"]})
            ws.conditional_format(3, profit_col, len(summary)+2, profit_col, {"type":"cell","criteria":">=","value":0,"format":formats["green_light"]})
            ws.conditional_format(3, profit_col, len(summary)+2, profit_col, {"type":"3_color_scale","min_color":"#FECACA","mid_color":"#FEF3C7","max_color":"#A7F3D0"})
        if revenue_col >= 0:
            ws.conditional_format(3, revenue_col, len(summary)+2, revenue_col, {"type":"data_bar","bar_color":"#6366F1"})
        if loss_col >= 0:
            ws.conditional_format(3, loss_col, len(summary)+2, loss_col, {"type":"cell","criteria":">","value":0,"format":formats["amber_light"]})
            ws.conditional_format(3, loss_col, len(summary)+2, loss_col, {"type":"cell","criteria":">","value":5,"format":formats["red_light"]})

        chart_rows = min(15, len(summary))
        chart = writer.book.add_chart({"type": "column"})
        chart.add_series(
            {
                "name": "Прибыль",
                "categories": [sheet_name, 3, 0, chart_rows + 2, 0],
                "values": [sheet_name, 3, 4, chart_rows + 2, 4],
                "fill": {"color": "#10B981"},
                "border": {"none": True},
            }
        )
        chart.set_title({"name": f"Прибыль: топ-{chart_rows}"})
        chart.set_legend({"none": True})
        chart.set_style(10)
        ws.insert_chart("J3", chart, {"x_scale": 1.45, "y_scale": 1.25})


def export_summary_excel(df: pd.DataFrame, settings: Dict[str, Any]) -> bytes:
    path = temporary_path(".xlsx")
    try:
        totals = calculate_totals(df)
        category = aggregate_by(df, "Категория")
        brand = aggregate_by(df, "Бренд")
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            workbook = writer.book
            formats = configure_excel_formats(workbook)

            total_df = pd.DataFrame(
                [
                    ["Количество SKU", totals["count"]],
                    ["Выручка", totals["revenue"]],
                    ["Расходы", totals["expenses"]],
                    ["Прибыль", totals["profit"]],
                    ["Маржа", totals["margin"]],
                    ["Убыточных SKU", totals["loss_count"]],
                    ["Себестоимость оценена", totals["estimated_cost_count"]],
                ],
                columns=["Метрика", "Значение"],
            )
            total_df.to_excel(writer, sheet_name="Итоги", index=False, startrow=2)
            ws = writer.sheets["Итоги"]
            ws.write(0, 0, "ИТОГИ ПО КАТАЛОГУ", formats["title"])
            for col, value in enumerate(total_df.columns):
                ws.write(2, col, value, formats["header"])
            ws.set_column("A:A", 32)
            ws.set_column("B:B", 20, formats["money"])

            write_summary_sheet(writer, "Сводка_Категории", "Категории", category, formats)
            write_summary_sheet(writer, "Сводка_Бренды", "Бренды", brand, formats)

            tariffs = tariff_table(settings)
            tariffs.to_excel(writer, sheet_name="Тариф", index=False)
            ws_t = writer.sheets["Тариф"]
            write_dataframe_header(ws_t, tariffs, formats["header"])
            ws_t.set_column("A:A", 40)
            ws_t.set_column("B:C", 18)

            legend = legend_table()
            legend.to_excel(writer, sheet_name="Легенда", index=False)
            ws_l = writer.sheets["Легенда"]
            write_dataframe_header(ws_l, legend, formats["header"])
            ws_l.set_column("A:A", 25)
            ws_l.set_column("B:C", 55)
        return read_and_remove(path)
    except Exception:
        try:
            os.remove(path)
        except OSError:
            pass
        raise


def export_values_excel(df: pd.DataFrame, settings: Dict[str, Any]) -> bytes:
    """Full values workbook with bounded memory use for up to 300,000 rows."""
    path = temporary_path(".xlsx")
    try:
        category = aggregate_by(df, "Категория")
        brand = aggregate_by(df, "Бренд")
        totals = calculate_totals(df)
        export_df = df[RESULT_COLUMNS]

        workbook = xlsxwriter.Workbook(
            path,
            {
                "constant_memory": True,
                "strings_to_urls": False,
                "nan_inf_to_errors": True,
            },
        )
        formats = configure_excel_formats(workbook)

        # Write the largest sheet in strict row order. Only one row is kept in RAM.
        ws = workbook.add_worksheet("Расчет_FBS")
        ws.write_row(0, 0, RESULT_COLUMNS, formats["header"])
        for row_number, row in enumerate(export_df.itertuples(index=False, name=None), start=1):
            clean_row = []
            for value in row:
                if value is pd.NA or value is None:
                    clean_row.append("")
                elif isinstance(value, np.generic):
                    clean_row.append(value.item())
                else:
                    clean_row.append(value)
            ws.write_row(row_number, 0, clean_row)

        ws.freeze_panes(1, 3)
        ws.autofilter(0, 0, len(export_df), len(RESULT_COLUMNS) - 1)
        ws.set_column(0, 0, 18)
        ws.set_column(1, 2, 20)
        ws.set_column(3, 9, 12, formats["number"])
        ws.set_column(10, 10, 14, formats["money"])
        ws.set_column(11, 11, 14, formats["money"])
        ws.set_column(12, 12, 18)
        ws.set_column(13, 13, 14, formats["percent"])
        ws.set_column(14, 21, 15, formats["money"])
        ws.set_column(22, 22, 12, formats["percent"])
        ws.set_column(23, 25, 20)

        margin_col = RESULT_COLUMNS.index("Маржа_%")
        profit_col = RESULT_COLUMNS.index("Прибыль")
        rec_col = RESULT_COLUMNS.index("Рекомендованная_цена") if "Рекомендованная_цена" in RESULT_COLUMNS else -1
        price_col = RESULT_COLUMNS.index("Цена") if "Цена" in RESULT_COLUMNS else -1
        markup_profit_col = RESULT_COLUMNS.index("Прибыль_с_наценкой") if "Прибыль_с_наценкой" in RESULT_COLUMNS else -1
        markup_margin_col = RESULT_COLUMNS.index("Маржа_с_наценкой_%") if "Маржа_с_наценкой_%" in RESULT_COLUMNS else -1
        abc_col = RESULT_COLUMNS.index("ABC") if "ABC" in RESULT_COLUMNS else -1
        # ── Максимальная палитра: маржа, прибыль, рекомендованная, ABC/XYZ, валидация ──
        # Маржа — светофор + 4 градации
        ws.conditional_format(1, margin_col, len(export_df), margin_col, {"type": "icon_set", "icon_style": "3_traffic_lights"})
        ws.conditional_format(1, margin_col, len(export_df), margin_col, {"type": "cell", "criteria": "<", "value": 0, "format": formats["red"]})
        ws.conditional_format(1, margin_col, len(export_df), margin_col, {"type": "cell", "criteria": "between", "minimum": 0, "maximum": 0.05, "format": formats["red_light"]})
        ws.conditional_format(1, margin_col, len(export_df), margin_col, {"type": "cell", "criteria": "between", "minimum": 0.05, "maximum": 0.15, "format": formats["amber"]})
        ws.conditional_format(1, margin_col, len(export_df), margin_col, {"type": "cell", "criteria": ">=", "value": 0.15, "format": formats["green"]})
        # Прибыль — тепловые карты + светофор
        ws.conditional_format(1, profit_col, len(export_df), profit_col, {"type": "cell", "criteria": "<", "value": 0, "format": formats["red"]})
        ws.conditional_format(1, profit_col, len(export_df), profit_col, {"type": "cell", "criteria": ">=", "value": 0, "format": formats["green_light"]})
        ws.conditional_format(1, profit_col, len(export_df), profit_col, {"type": "3_color_scale", "min_color": "#FECACA", "mid_color": "#FEF3C7", "max_color": "#A7F3D0"})
        # Рекомендованная > цены — оранжевый, >1.2× — красный
        if rec_col >= 0 and price_col >= 0:
            ws.conditional_format(1, rec_col, len(export_df), rec_col, {"type": "formula", "criteria": f"=${chr(65+rec_col)}2>${chr(65+price_col)}2", "format": formats["amber"]})
            ws.conditional_format(1, rec_col, len(export_df), rec_col, {"type": "formula", "criteria": f"=${chr(65+rec_col)}2>${chr(65+price_col)}2*1.2", "format": formats["red"]})
            ws.conditional_format(1, rec_col, len(export_df), rec_col, {"type": "3_color_scale", "min_color": "#ECFDF5", "mid_color": "#FEF3C7", "max_color": "#FECACA"})
        if markup_profit_col >= 0:
            ws.conditional_format(1, markup_profit_col, len(export_df), markup_profit_col, {"type": "cell", "criteria": "<", "value": 0, "format": formats["red"]})
            ws.conditional_format(1, markup_profit_col, len(export_df), markup_profit_col, {"type": "3_color_scale", "min_color": "#FECACA", "mid_color": "#FEF3C7", "max_color": "#A7F3D0"})
        if markup_margin_col >= 0:
            ws.conditional_format(1, markup_margin_col, len(export_df), markup_margin_col, {"type": "cell", "criteria": "<", "value": 0, "format": formats["red"]})
            ws.conditional_format(1, markup_margin_col, len(export_df), markup_margin_col, {"type": "cell", "criteria": "between", "minimum": 0, "maximum": 0.15, "format": formats["amber"]})
            ws.conditional_format(1, markup_margin_col, len(export_df), markup_margin_col, {"type": "cell", "criteria": ">=", "value": 0.15, "format": formats["green"]})
            ws.conditional_format(1, markup_margin_col, len(export_df), markup_margin_col, {"type": "icon_set", "icon_style": "3_traffic_lights"})
        if abc_col >= 0:
            ws.conditional_format(1, abc_col, len(export_df), abc_col, {"type": "cell", "criteria": "equal to", "value": '"A"', "format": formats["green"]})
            ws.conditional_format(1, abc_col, len(export_df), abc_col, {"type": "cell", "criteria": "equal to", "value": '"B"', "format": formats["amber"]})
            ws.conditional_format(1, abc_col, len(export_df), abc_col, {"type": "cell", "criteria": "equal to", "value": '"C"', "format": formats["blue"]})
        # XYZ — X зелёный, Y янтарь, Z красный
        xyz_col = RESULT_COLUMNS.index("XYZ") if "XYZ" in RESULT_COLUMNS else -1
        if xyz_col >= 0:
            ws.conditional_format(1, xyz_col, len(export_df), xyz_col, {"type": "cell", "criteria": "equal to", "value": '"X"', "format": formats["green"]})
            ws.conditional_format(1, xyz_col, len(export_df), xyz_col, {"type": "cell", "criteria": "equal to", "value": '"Y"', "format": formats["amber_light"]})
            ws.conditional_format(1, xyz_col, len(export_df), xyz_col, {"type": "cell", "criteria": "equal to", "value": '"Z"', "format": formats["red_light"]})
        abcx_col = RESULT_COLUMNS.index("ABC_XYZ") if "ABC_XYZ" in RESULT_COLUMNS else -1
        if abcx_col >= 0:
            ws.conditional_format(1, abcx_col, len(export_df), abcx_col, {"type": "cell", "criteria": "equal to", "value": '"AX"', "format": formats["green"]})
            ws.conditional_format(1, abcx_col, len(export_df), abcx_col, {"type": "cell", "criteria": "equal to", "value": '"CZ"', "format": formats["red"]})
            ws.conditional_format(1, abcx_col, len(export_df), abcx_col, {"type": "cell", "criteria": "equal to", "value": '"AZ"', "format": formats["amber"]})
        # Выручка и продажи — гистограммы в ячейках + тепловая карта
        for col_name, col_color in [("Выручка","#6366F1"),("Продажи_шт","#0D9488"),("Цена","#0EA5E9"),("Прибыль_период","#10B981"),("Выручка_доля","#7C3AED")]:
            if col_name in RESULT_COLUMNS:
                cidx = RESULT_COLUMNS.index(col_name)
                ws.conditional_format(1, cidx, len(export_df), cidx, {"type": "data_bar", "bar_color": col_color})
                if col_name in ["Выручка","Прибыль_период"]:
                    ws.conditional_format(1, cidx, len(export_df), cidx, {"type": "3_color_scale", "min_color": "#FECACA", "mid_color": "#FEF3C7", "max_color": "#A7F3D0"})
        # Комиссия/Логистика/Хранение — чем больше, тем краснее
        for col_name in ["Комиссия_руб","Логистика_руб","Хранение_руб","Эквайринг_руб","Спец_расходы_FBS","Итого_расходы"]:
            if col_name in RESULT_COLUMNS:
                cidx = RESULT_COLUMNS.index(col_name)
                ws.conditional_format(1, cidx, len(export_df), cidx, {"type": "3_color_scale", "min_color": "#FFFFFF", "mid_color": "#FEF3C7", "max_color": "#FECACA"})
        # Рекомендация — текстовая
        if "Рекомендация" in RESULT_COLUMNS:
            rec_text_col = RESULT_COLUMNS.index("Рекомендация")
            ws.conditional_format(1, rec_text_col, len(export_df), rec_text_col, {"type": "text", "criteria": "containing", "value": "Поднять", "format": formats["red"]})
            ws.conditional_format(1, rec_text_col, len(export_df), rec_text_col, {"type": "text", "criteria": "containing", "value": "Критично", "format": formats["amber"]})
            ws.conditional_format(1, rec_text_col, len(export_df), rec_text_col, {"type": "text", "criteria": "containing", "value": "ОК", "format": formats["green"]})
        # Валидация — всё, что не OK
        if "Валидация" in RESULT_COLUMNS:
            val_col = RESULT_COLUMNS.index("Валидация")
            ws.conditional_format(1, val_col, len(export_df), val_col, {"type": "text", "criteria": "containing", "value": "нет габаритов", "format": formats["amber_light"]})
            ws.conditional_format(1, val_col, len(export_df), val_col, {"type": "text", "criteria": "containing", "value": "не примет", "format": formats["red"]})
            ws.conditional_format(1, val_col, len(export_df), val_col, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["green_light"]})

        # Compact totals sheet.
        ws_total = workbook.add_worksheet("Итоги")
        ws_total.write(0, 0, "ИТОГИ ПО КАТАЛОГУ", formats["title"])
        ws_total.write_row(2, 0, ["Метрика", "Значение"], formats["header"])
        total_rows = [
            ("Количество SKU", totals["count"]),
            ("Выручка", totals["revenue"]),
            ("Расходы", totals["expenses"]),
            ("Прибыль", totals["profit"]),
            ("Маржа", totals["margin"]),
            ("Убыточных SKU", totals["loss_count"]),
            ("Себестоимость оценена", totals["estimated_cost_count"]),
        ]
        for row_number, values in enumerate(total_rows, start=3):
            ws_total.write(row_number, 0, values[0])
            if values[0] == "Маржа":
                ws_total.write_number(row_number, 1, float(values[1]), formats["percent"])
            elif values[0] in ["Выручка", "Расходы", "Прибыль"]:
                ws_total.write_number(row_number, 1, float(values[1]), formats["money"])
            else:
                ws_total.write_number(row_number, 1, float(values[1]), formats["number"])
        ws_total.set_column("A:A", 34)
        ws_total.set_column("B:B", 20)

        def write_direct_summary(sheet_name: str, title: str, summary: pd.DataFrame) -> None:
            sheet = workbook.add_worksheet(sheet_name)
            sheet.write(0, 0, title, formats["title"])
            sheet.write_row(2, 0, summary.columns.tolist(), formats["header"])
            for row_number, values in enumerate(
                summary.itertuples(index=False, name=None), start=3
            ):
                sheet.write_row(row_number, 0, list(values))
            sheet.freeze_panes(3, 0)
            sheet.autofilter(2, 0, len(summary) + 2, len(summary.columns) - 1)
            sheet.set_column(0, 0, 28)
            sheet.set_column(1, 1, 12)
            sheet.set_column(2, 4, 16, formats["money"])
            sheet.set_column(5, 5, 12, formats["percent"])
            sheet.set_column(6, 7, 15)
            if not summary.empty:
                chart_rows = min(15, len(summary))
                chart = workbook.add_chart({"type": "column"})
                chart.add_series(
                    {
                        "name": "Прибыль",
                        "categories": [sheet_name, 3, 0, chart_rows + 2, 0],
                        "values": [sheet_name, 3, 4, chart_rows + 2, 4],
                        "fill": {"color": "#10B981"},
                        "border": {"none": True},
                    }
                )
                chart.set_title({"name": f"Прибыль: топ-{chart_rows}"})
                chart.set_legend({"none": True})
                sheet.insert_chart("J3", chart, {"x_scale": 1.4, "y_scale": 1.2})

        write_direct_summary("Сводка_Категории", "СВОДКА: КАТЕГОРИИ", category)
        write_direct_summary("Сводка_Бренды", "СВОДКА: БРЕНДЫ", brand)

        tariffs = tariff_table(settings)
        ws_t = workbook.add_worksheet("Тариф")
        ws_t.write_row(0, 0, tariffs.columns.tolist(), formats["header"])
        for row_number, values in enumerate(tariffs.itertuples(index=False, name=None), start=1):
            ws_t.write_row(row_number, 0, list(values))
        ws_t.set_column("A:A", 40)
        ws_t.set_column("B:C", 20)

        legend = legend_table()
        ws_l = workbook.add_worksheet("Легенда")
        ws_l.write_row(0, 0, legend.columns.tolist(), formats["header"])
        for row_number, values in enumerate(legend.itertuples(index=False, name=None), start=1):
            ws_l.write_row(row_number, 0, list(values))
        ws_l.set_column("A:A", 25)
        ws_l.set_column("B:C", 55)

        # ── ABC / XYZ матрица ──
        try:
            ws_abc = workbook.add_worksheet("ABC_XYZ")
            ws_abc.write_row(0, 0, ["ABC / XYZ — МАТРИЦА 9 ЯЧЕЕК"], formats["header_warm"])
            ws_abc.write_row(1, 0, ["", "X ≤30дн", "Y ≤60дн", "Z >60дн"], formats["amber_light"])
            matrix = df.groupby(["ABC","XYZ"]).size().unstack(fill_value=0)
            for r_idx, abc in enumerate(["A","B","C"], start=2):
                label = {"A":"A — ядро 80%","B":"B — середняки 15%","C":"C — хвост 5%"}[abc]
                vals = [int(matrix.loc[abc, xyz] if abc in matrix.index and xyz in matrix.columns else 0) for xyz in ["X","Y","Z"]]
                ws_abc.write_row(r_idx, 0, [label] + vals)
                for c in range(1,4):
                    ws_abc.conditional_format(r_idx, c, r_idx, c, {"type":"data_bar","bar_color":"#6366F1"})
            ws_abc.set_column("A:A", 18); ws_abc.set_column("B:D", 14)
            # Смысл
            ws_abc.write_row(6, 0, ["AX: ★ Ядро","AZ: Риск заморозки","CX: Поднять цену","CZ: ⚠ Балласт"], formats["band"])
        except Exception:
            pass
        # ── Валидация ──
        try:
            bad = df[df["Валидация"] != "OK"]
            if not bad.empty:
                ws_val = workbook.add_worksheet("Валидация")
                ws_val.write_row(0, 0, ["Валидация — проблемные SKU (топ 5000)"], formats["header"])
                cols = ["Артикул","Бренд","Категория","Цена","Вес_кг","Объем_л","Плотность","Сумма_сторон","Валидация"]
                ws_val.write_row(1, 0, cols, formats["header"])
                for rn, (_, r) in enumerate(bad.head(5000).iterrows(), start=2):
                    ws_val.write_row(rn, 0, [r["Артикул"], r["Бренд"], r["Категория"], r["Цена"], r["Вес_кг"], r["Объем_л"], r["Плотность"], r["Сумма_сторон"], r["Валидация"]])
                ws_val.set_column("A:A", 18); ws_val.set_column("B:C", 16); ws_val.set_column("D:H", 12)
                ws_val.autofilter(1, 0, min(len(bad),5000)+1, len(cols)-1)
                ws_val.freeze_panes(2, 0)
                ws_val.conditional_format(2, 8, min(len(bad),5000)+1, 8, {"type":"text","criteria":"containing","value":"не примет","format":formats["red"]})
        except Exception:
            pass

        workbook.close()
        return read_and_remove(path)
    except Exception:
        try:
            os.remove(path)
        except OSError:
            pass
        raise


def export_formula_excel(df: pd.DataFrame, settings: Dict[str, Any]) -> bytes:
    """Красочный Excel с живыми формулами. Формулы синхронизированы с движком."""
    path = temporary_path(".xlsx")
    try:
        input_columns = [
            "Артикул", "Бренд", "Категория", "Цена", "Себестоимость", "Объем_л",
            "Оплач_вес", "Оборачиваемость_дней", "Продажи_шт", "is_hazardous",
            "is_fragile", "Ставка_комиссии", "Ставка_хранения",
        ]
        input_df = df[input_columns].copy()
        category = aggregate_by(df, "Категория")
        brand = aggregate_by(df, "Бренд")

        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            workbook = writer.book
            formats = configure_excel_formats(workbook)

            tariffs = tariff_table(settings)
            tariffs.to_excel(writer, sheet_name="Тариф", index=False)
            ws_t = writer.sheets["Тариф"]
            write_dataframe_header(ws_t, tariffs, formats["header"])
            ws_t.set_column("A:A", 40)
            ws_t.set_column("B:C", 18)

            input_df.to_excel(writer, sheet_name="Входные_Данные", index=False)
            ws_in = writer.sheets["Входные_Данные"]
            write_dataframe_header(ws_in, input_df, formats["header_accent"])
            ws_in.freeze_panes(1, 1)
            ws_in.autofilter(0, 0, len(input_df), len(input_df.columns) - 1)
            ws_in.set_column("A:A", 18)
            ws_in.set_column("B:C", 20)
            ws_in.set_column("D:M", 14)
            # Цветовые индикаторы входных данных
            ws_in.conditional_format(1, 3, len(input_df), 3, {"type":"data_bar", "bar_color":"#6366F1"})
            ws_in.conditional_format(1, 8, len(input_df), 8, {"type":"data_bar", "bar_color":"#10B981"})

            ws_calc = workbook.add_worksheet("Расчет_FBS")
            headers = [
                "Артикул", "Бренд", "Категория", "Цена", "Себестоимость", "Объем_л",
                "Рекоменд_цена", "Цена_с_наценкой", "Комиссия_руб", "Логистика_руб",
                "Хранение_руб", "Эквайринг_руб", "Возвраты_руб", "Спец_расходы_FBS",
                "Итого_расходы", "Прибыль", "Маржа_%", "Продажи_шт", "Выручка",
                "Прибыль_период", "ABC", "XYZ", "ABC/XYZ", "Рекомендация", "Валидация",
            ]
            ws_calc.write_row(0, 0, headers, formats["header_warm"])
            ws_calc.write_row(1, 0, [
                "","","","","","","★ ЖИВАЯ","★ ЖИВАЯ","★ ЖИВАЯ","★ СР.МИЛЯ",
                "★ ЖИВАЯ","★ ЖИВАЯ","★ ЖИВАЯ","★ ЖИВАЯ","★ E+I:N","★ H−O","★ P/H",
                "Входные","★ H×R","★ P×R","стат.","стат.","стат.","★ совет","флаги",
            ], formats["amber_light"])
            for idx, row in enumerate(df.itertuples(index=False), start=2):
                excel_row = idx + 1   # Расчет: две строки шапки → первая строка данных 3
                source_row = idx      # Входные: одна строка шапки → первая строка данных 2
                # Входные: D цена, E с/с, F объём, H оборот, I продажи, J опасный, K хрупкий, L effRate, M storageRate
                is_fixed = f"AND(H{excel_row}<=300,F{excel_row}<=5,'Тариф'!$B$20=\"Да\")"
                fix_rate = f"IF('Тариф'!$B$23=\"FBS\",0.42-'Тариф'!$B$21,0.35)"
                srednya = f"IF(F{excel_row}<=1,80,IF(F{excel_row}<=30,80+9*(CEILING(F{excel_row},1)-1),IF(F{excel_row}<=200,341+7*(CEILING(F{excel_row},1)-30),MIN(5500,1531+5*(CEILING(F{excel_row},1)-200)))))"
                storage = f"IF('Тариф'!$B$23=\"FBS\",0,F{excel_row}*'Входные_Данные'!M{source_row}*MAX(0,'Входные_Данные'!H{source_row}-'Тариф'!$B$22))"
                rec_formula = (
                    f"=LET(eff,'Входные_Данные'!L{source_row},fixR,{fix_rate},sred,{srednya},stor,{storage},"
                    f"specF,'Тариф'!$B$9+'Тариф'!$B$10+'Тариф'!$B$11,"
                    f"varS,eff+'Тариф'!$B$7+'Тариф'!$B$12+IF('Входные_Данные'!J{source_row},'Тариф'!$B$13,0)+IF('Входные_Данные'!K{source_row},'Тариф'!$B$14,0)+'Тариф'!$B$8,"
                    f"varF,fixR+'Тариф'!$B$7+'Тариф'!$B$12+IF('Входные_Данные'!J{source_row},'Тариф'!$B$13,0)+IF('Входные_Данные'!K{source_row},'Тариф'!$B$14,0),"
                    f"pctP,(E{excel_row}+sred+stor+specF)/(1-varS),minP,(E{excel_row}+sred+stor+specF+'Тариф'!$B$3)/(1-(varS-eff)),fixP,(E{excel_row}+stor+specF)/(1-varF),"
                    f"IF(AND('Тариф'!$B$20=\"Да\",F{excel_row}<=5,fixP<=300),fixP*1.01,IF(pctP*eff>='Тариф'!$B$3,pctP,minP)*1.01))"
                )
                price_markup_formula = (
                    f"=LET(eff,'Входные_Данные'!L{source_row},fixR,{fix_rate},sred,{srednya},stor,{storage},"
                    f"specF,'Тариф'!$B$9+'Тариф'!$B$10+'Тариф'!$B$11,"
                    f"varS,eff+'Тариф'!$B$7+'Тариф'!$B$12+IF('Входные_Данные'!J{source_row},'Тариф'!$B$13,0)+IF('Входные_Данные'!K{source_row},'Тариф'!$B$14,0)+'Тариф'!$B$8,"
                    f"varF,fixR+'Тариф'!$B$7+'Тариф'!$B$12+IF('Входные_Данные'!J{source_row},'Тариф'!$B$13,0)+IF('Входные_Данные'!K{source_row},'Тариф'!$B$14,0),"
                    f"pctP,(E{excel_row}+sred+stor+specF)/(1-varS-'Тариф'!$B$19),"
                    f"minP,(E{excel_row}+sred+stor+specF+'Тариф'!$B$3)/(1-(varS-eff)-'Тариф'!$B$19),"
                    f"fixP,(E{excel_row}+stor+specF)/(1-varF-'Тариф'!$B$19),"
                    f"IF('Тариф'!$B$17=\"markup\",D{excel_row}*(1+'Тариф'!$B$18/100),"
                    f"IF('Тариф'!$B$17=\"targetMargin\",IF(AND('Тариф'!$B$20=\"Да\",F{excel_row}<=5,fixP<=300),fixP,IF(pctP*eff>='Тариф'!$B$3,pctP,minP)),D{excel_row})))"
                )
                ws_calc.write_formula(idx, 0, f"='Входные_Данные'!A{source_row}", None, row.Артикул)
                ws_calc.write_formula(idx, 1, f"='Входные_Данные'!B{source_row}", None, row.Бренд)
                ws_calc.write_formula(idx, 2, f"='Входные_Данные'!C{source_row}", None, row.Категория)
                ws_calc.write_formula(idx, 3, f"='Входные_Данные'!D{source_row}", formats["money"], row.Цена)
                ws_calc.write_formula(idx, 4, f"='Входные_Данные'!E{source_row}", formats["money"], row.Себестоимость)
                ws_calc.write_formula(idx, 5, f"='Входные_Данные'!F{source_row}", formats["number"], row.Объем_л)
                ws_calc.write_formula(idx, 6, rec_formula, formats["money"], float(row.Рекомендованная_цена))
                ws_calc.write_formula(idx, 7, price_markup_formula, formats["money"], float(row.Цена_с_наценкой))
                ws_calc.write_formula(idx, 8, f"=IF({is_fixed},0,MAX(H{excel_row}*'Входные_Данные'!L{source_row},'Тариф'!$B$3))", formats["money"], row.Комиссия_руб)
                ws_calc.write_formula(idx, 9, f"=IF({is_fixed},H{excel_row}*({fix_rate}),{srednya})", formats["money"], row.Логистика_руб)
                ws_calc.write_formula(idx, 10, f"={storage}", formats["money"], row.Хранение_руб)
                ws_calc.write_formula(idx, 11, f"=H{excel_row}*'Тариф'!$B$7", formats["money"], row.Эквайринг_руб)
                ws_calc.write_formula(idx, 12, f"=IF({is_fixed},0,H{excel_row}*'Тариф'!$B$8)", formats["money"], row.Возвраты_руб)
                ws_calc.write_formula(idx, 13, f"='Тариф'!$B$9+'Тариф'!$B$10+'Тариф'!$B$11+H{excel_row}*'Тариф'!$B$12+IF('Входные_Данные'!J{source_row},H{excel_row}*'Тариф'!$B$13,0)+IF('Входные_Данные'!K{source_row},H{excel_row}*'Тариф'!$B$14,0)", formats["money"], row.Спец_расходы_FBS)
                # Итого = Себестоимость + Комиссия+Логистика+Хранение+Эквайринг+Возвраты+Спец
                ws_calc.write_formula(idx, 14, f"=E{excel_row}+I{excel_row}+J{excel_row}+K{excel_row}+L{excel_row}+M{excel_row}+N{excel_row}", formats["money"], row.Итого_расходы)
                ws_calc.write_formula(idx, 15, f"=H{excel_row}-O{excel_row}", formats["money"], float(row.Прибыль_с_наценкой))
                val_margin = float(df["Маржа_с_наценкой_%"].iat[idx-2])
                ws_calc.write_formula(idx, 16, f"=IF(H{excel_row}>0,P{excel_row}/H{excel_row},0)", formats["percent"], val_margin)
                ws_calc.write_formula(idx, 17, f"='Входные_Данные'!I{source_row}", formats["number"], float(row.Продажи_шт))
                ws_calc.write_formula(idx, 18, f"=H{excel_row}*R{excel_row}", formats["money"], float(row.Выручка))
                ws_calc.write_formula(idx, 19, f"=P{excel_row}*R{excel_row}", formats["money"], float(row.Прибыль_период))
                ws_calc.write(idx, 20, str(df["ABC"].iat[idx-2]))
                ws_calc.write(idx, 21, str(df["XYZ"].iat[idx-2]))
                ws_calc.write(idx, 22, str(df["ABC_XYZ"].iat[idx-2]))
                p_val = float(df["Прибыль"].iat[idx-2])
                m_val = float(df["Маржа_%"].iat[idx-2])
                rec_text = "↑ Поднять" if p_val < 0 else ("⚠ Критично" if m_val < 0.05 else ("→ Можно +10%" if m_val < 0.15 else "✓ ОК"))
                ws_calc.write_formula(idx, 23, f"=IF(P{excel_row}<0,\"↑ Поднять до \"&TEXT(G{excel_row},\"#,##0\")&\" ₽\",IF(Q{excel_row}<0.05,\"⚠ Критично\",IF(Q{excel_row}<0.15,\"→ Можно +10%\",\"✓ ОК\")))", None, rec_text)
                ws_calc.write(idx, 24, str(df["Валидация"].iat[idx-2]))

            ws_calc.freeze_panes(2, 1)
            ws_calc.autofilter(0, 0, len(df)+1, len(headers) - 1)
            ws_calc.set_column("A:A", 18)
            ws_calc.set_column("B:C", 20)
            ws_calc.set_column("D:Y", 14)
            # ── Максимальная палитра: все возможные условные форматирования ──
            # Маржа — светофор + 4 градации
            ws_calc.conditional_format(2, 16, len(df)+1, 16, {"type": "icon_set", "icon_style":"3_traffic_lights"})
            ws_calc.conditional_format(2, 16, len(df)+1, 16, {"type": "cell", "criteria": "<", "value": 0, "format": formats["red"]})
            ws_calc.conditional_format(2, 16, len(df)+1, 16, {"type": "cell", "criteria": "between", "minimum": 0, "maximum": 0.05, "format": formats["red_light"]})
            ws_calc.conditional_format(2, 16, len(df)+1, 16, {"type": "cell", "criteria": "between", "minimum": 0.05, "maximum": 0.15, "format": formats["amber"]})
            ws_calc.conditional_format(2, 16, len(df)+1, 16, {"type": "cell", "criteria": ">=", "value": 0.15, "format": formats["green"]})
            # Тепловые карты и гистограммы — прибыль, выручка, продажи
            ws_calc.conditional_format(2, 15, len(df)+1, 15, {"type":"3_color_scale", "min_color":"#FECACA", "mid_color":"#FEF3C7", "max_color":"#A7F3D0"})
            ws_calc.conditional_format(2, 19, len(df)+1, 19, {"type":"3_color_scale", "min_color":"#FECACA", "mid_color":"#FEF3C7", "max_color":"#A7F3D0"})
            ws_calc.conditional_format(2, 14, len(df)+1, 14, {"type":"3_color_scale", "min_color":"#FECACA", "mid_color":"#FEF3C7", "max_color":"#A7F3D0"})
            ws_calc.conditional_format(2, 18, len(df)+1, 18, {"type":"data_bar", "bar_color":"#6366F1"})
            ws_calc.conditional_format(2, 17, len(df)+1, 17, {"type":"data_bar", "bar_color":"#0D9488"})
            ws_calc.conditional_format(2, 3, len(df)+1, 3, {"type":"data_bar", "bar_color":"#0EA5E9"})
            # Рекомендованная > цены
            ws_calc.conditional_format(2, 6, len(df)+1, 6, {"type": "formula", "criteria": "=$G3>$D3", "format": formats["amber"]})
            ws_calc.conditional_format(2, 6, len(df)+1, 6, {"type": "formula", "criteria": "=$G3>$D3*1.2", "format": formats["red"]})
            ws_calc.conditional_format(2, 6, len(df)+1, 6, {"type":"3_color_scale", "min_color":"#ECFDF5", "mid_color":"#FEF3C7", "max_color":"#FECACA"})
            # Комиссия/Логистика/Хранение — чем больше, тем краснее
            for col_idx in [8,9,10,11,12,13]:
                ws_calc.conditional_format(2, col_idx, len(df)+1, col_idx, {"type":"3_color_scale", "min_color":"#FFFFFF", "mid_color":"#FEF3C7", "max_color":"#FECACA"})
            # ABC/XYZ
            ws_calc.conditional_format(2, 20, len(df)+1, 20, {"type":"text","criteria":"containing","value":"A","format":formats["green"]})
            ws_calc.conditional_format(2, 20, len(df)+1, 20, {"type":"text","criteria":"containing","value":"B","format":formats["amber"]})
            ws_calc.conditional_format(2, 20, len(df)+1, 20, {"type":"text","criteria":"containing","value":"C","format":formats["blue"]})
            ws_calc.conditional_format(2, 21, len(df)+1, 21, {"type":"text","criteria":"containing","value":"X","format":formats["green"]})
            ws_calc.conditional_format(2, 21, len(df)+1, 21, {"type":"text","criteria":"containing","value":"Y","format":formats["amber_light"]})
            ws_calc.conditional_format(2, 21, len(df)+1, 21, {"type":"text","criteria":"containing","value":"Z","format":formats["red"]})
            ws_calc.conditional_format(2, 22, len(df)+1, 22, {"type":"text","criteria":"containing","value":"AX","format":formats["green"]})
            ws_calc.conditional_format(2, 22, len(df)+1, 22, {"type":"text","criteria":"containing","value":"CZ","format":formats["red"]})
            # Рекомендация
            ws_calc.conditional_format(2, 23, len(df)+1, 23, {"type": "text", "criteria": "containing", "value": "Поднять", "format": formats["red"]})
            ws_calc.conditional_format(2, 23, len(df)+1, 23, {"type": "text", "criteria": "containing", "value": "ОК", "format": formats["green"]})
            ws_calc.conditional_format(2, 23, len(df)+1, 23, {"type": "text", "criteria": "containing", "value": "Критично", "format": formats["amber"]})
            ws_calc.conditional_format(2, 23, len(df)+1, 23, {"type": "text", "criteria": "containing", "value": "Можно", "format": formats["amber_light"]})
            # Валидация — всё, что не OK
            ws_calc.conditional_format(2, 24, len(df)+1, 24, {"type": "text", "criteria": "containing", "value": "нет габаритов", "format": formats["amber_light"]})
            ws_calc.conditional_format(2, 24, len(df)+1, 24, {"type": "text", "criteria": "containing", "value": "не примет", "format": formats["red"]})
            ws_calc.conditional_format(2, 24, len(df)+1, 24, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["green_light"]})

            write_summary_sheet(writer, "Сводка_Категории", "Категории", category, formats)
            write_summary_sheet(writer, "Сводка_Бренды", "Бренды", brand, formats)
            # ── ABC / XYZ матрица (красочная) ──
            try:
                ws_abc = workbook.add_worksheet("ABC_XYZ")
                ws_abc.write_row(0, 0, ["ABC / XYZ — МАТРИЦА 9 ЯЧЕЕК"], formats["header_warm"])
                ws_abc.write_row(1, 0, ["", "X — быстро ≤30дн", "Y — средне ≤60дн", "Z — медленно >60дн"], formats["amber_light"])
                matrix = df.groupby(["ABC","XYZ"]).size().unstack(fill_value=0)
                # Header
                ws_abc.write_row(2, 0, ["", "X", "Y", "Z", "ИТОГО"], formats["header"])
                for r_idx, abc in enumerate(["A","B","C"], start=3):
                    label = {"A":"A — ядро 80%","B":"B — середняки 15%","C":"C — хвост 5%"}[abc]
                    row = [label]
                    total_r = 0
                    for xyz in ["X","Y","Z"]:
                        val = int(matrix.loc[abc, xyz] if abc in matrix.index and xyz in matrix.columns else 0)
                        row.append(val); total_r += val
                    row.append(total_r)
                    ws_abc.write_row(r_idx, 0, row)
                    # Подсветка ячеек матрицы
                    for c_idx, xyz in enumerate(["X","Y","Z"], start=1):
                        val = row[c_idx+1] if isinstance(row[c_idx+1], int) else 0
                        # need to get value again
                    # раскраска по смыслу
                    fmt_cell = formats["green"] if abc=="A" and r_idx==3 else (formats["red"] if abc=="C" else formats["amber_light"])
                    # apply via conditional? simple static color by row
                    # We'll set row background via conditional_format on range
                # Totals column data_bar
                ws_abc.set_column("A:A", 18)
                ws_abc.set_column("B:E", 16)
                # Смысл матрицы
                ws_abc.write_row(7, 0, ["Расшифровка:"], formats["header"])
                meanings = [
                    ("AX", "★ Ядро — держать сток"),
                    ("AZ", "Дорогие зависающие — риск заморозки"),
                    ("CX", "Дешёвые ходовые — поднять цену"),
                    ("CZ", "⚠ Балласт — вывести"),
                ]
                for i, (k,m) in enumerate(meanings, start=8):
                    ws_abc.write_row(i, 0, [k, m])
                    ws_abc.set_row(i, None, formats["band"] if i%2==0 else None)
                ws_abc.freeze_panes(3, 1)
            except Exception:
                pass
            # ── Валидация (если есть проблемы) ──
            try:
                bad = df[df["Валидация"] != "OK"]
                if not bad.empty:
                    ws_val = workbook.add_worksheet("Валидация")
                    ws_val.write_row(0, 0, ["Валидация габаритов и веса — проблемные SKU (топ 5000)"], formats["header"])
                    cols = ["Артикул","Бренд","Категория","Цена","Вес_кг","Объем_л","Плотность","Сумма_сторон","Валидация"]
                    ws_val.write_row(1, 0, cols, formats["header"])
                    for rn, (_, r) in enumerate(bad.head(5000).iterrows(), start=2):
                        ws_val.write_row(rn, 0, [r["Артикул"], r["Бренд"], r["Категория"], r["Цена"], r["Вес_кг"], r["Объем_л"], r["Плотность"], r["Сумма_сторон"], r["Валидация"]])
                    ws_val.set_column("A:A", 18); ws_val.set_column("B:C", 16); ws_val.set_column("D:H", 12)
                    ws_val.autofilter(1, 0, min(len(bad),5000)+1, len(cols)-1)
                    ws_val.freeze_panes(2, 0)
                    ws_val.conditional_format(2, 8, min(len(bad),5000)+1, 8, {"type":"text","criteria":"containing","value":"не примет","format":formats["red"]})
            except Exception:
                pass
            legend = legend_table()
            legend.to_excel(writer, sheet_name="Легенда", index=False)
            write_dataframe_header(writer.sheets["Легенда"], legend, formats["header"])
            writer.sheets["Легенда"].set_column("A:A", 25)
            writer.sheets["Легенда"].set_column("B:C", 55)
        return read_and_remove(path)
    except Exception:
        try:
            os.remove(path)
        except OSError:
            pass
        raise


def export_csv(df: pd.DataFrame) -> bytes:
    path = temporary_path(".csv")
    try:
        csv_df = df[RESULT_COLUMNS].copy()
        for col in ["Маржа_%","Маржа_с_наценкой_%","Выручка_доля"]:
            if col in csv_df.columns:
                csv_df[col] = csv_df[col] * 100.0
        if "Ставка_комиссии" in csv_df.columns:
            csv_df["Ставка_комиссии"] = csv_df["Ставка_комиссии"] * 100.0
        csv_df.to_csv(
            path,
            index=False,
            sep=";",
            encoding="utf-8-sig",
            decimal=",",
            float_format="%.2f",
        )
        return read_and_remove(path)
    except Exception:
        try:
            os.remove(path)
        except OSError:
            pass
        raise


# =============================================================================
# SIDEBAR
# =============================================================================

with st.sidebar:
    st.markdown("## Unit Economics")
    st.caption(f"Яндекс Маркет · FBS · v{APP_VERSION}")
    st.divider()

    raw_sidebar = st.session_state.raw_df
    result_sidebar = st.session_state.result_df
    if raw_sidebar is None:
        st.info("Каталог не загружен")
    else:
        st.metric("Строк в каталоге", f"{len(raw_sidebar):,}".replace(",", " "))
        st.caption(st.session_state.source_name or "Источник данных")

    if result_sidebar is not None:
        totals_sidebar = calculate_totals(result_sidebar)
        st.metric("Прибыль", money_short(totals_sidebar["profit"]))
        st.metric("Маржа", percent(totals_sidebar["margin"]))
        if not calculation_is_current():
            st.warning("Тарифы изменены. Нужен пересчёт.")

    st.divider()
    st.markdown("**Быстрые действия**")
    if st.button("Пересчитать каталог", use_container_width=True, disabled=raw_sidebar is None):
        run_calculation()
    if st.button("Очистить данные", use_container_width=True, disabled=raw_sidebar is None):
        st.session_state.raw_df = None
        st.session_state.result_df = None
        st.session_state.source_name = ""
        st.session_state.parse_meta = {}
        st.session_state.calculated_settings_hash = ""
        clear_export()
        st.rerun()

    st.divider()
    st.caption("Все расчёты и файлы остаются на вашем сервере Streamlit.")


# =============================================================================
# HEADER AND NAVIGATION
# =============================================================================

st.markdown(
    f"""
    <section class="hero">
      <div class="hero-kicker">YANDEX MARKET · FBS · PREMIUM ANALYTICS</div>
      <h1>{APP_NAME} <span style="background:linear-gradient(90deg,#fbbf24,#f59e0b);-webkit-background-clip:text;-webkit-text-fill-color:transparent">· PRO</span></h1>
      <p>Без лимитов по строкам · Живые формулы · ABC/XYZ · Рекомендованная цена · Умная наценка</p>
      <span class="hero-badge">До 1 048 576 SKU · Красочные Excel · Кастомные категории</span>
    </section>
    """,
    unsafe_allow_html=True,
)

steps = ["1. Тарифы", "2. Данные", "3. Дашборд", "4. Экспорт"]
if "navigation" not in st.session_state:
    st.session_state.navigation = steps[0]

selected_step = st.radio(
    "Этап работы",
    steps,
    horizontal=True,
    key="navigation",
    label_visibility="collapsed",
)


# =============================================================================
# STEP 1: TARIFFS
# =============================================================================

if selected_step == steps[0]:
    st.markdown('<div class="section-title">Тарифы и расходы</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub">Настройки уже заполнены типовыми значениями. '
        'Новичку достаточно проверить себестоимость по умолчанию и перейти к загрузке.</div>',
        unsafe_allow_html=True,
    )
    settings = st.session_state.settings

    # ── Достоверная модель FBS 2026 ──
    st.markdown('<div class="soft-panel" style="border-left:4px solid #6366f1; background: linear-gradient(135deg, #eef2ff, white);">', unsafe_allow_html=True)
    st.markdown("**✅ Достоверная модель FBS 2026 — проверьте перед расчётом**")
    st.caption("С 1 апреля 2026: фиксированный тариф для товаров ≤300₽/≤5л (FBS 42% / FBY 35%), новая объёмная сетка средней мили, приём платежей по графику выплат. Все формулы открыты — сверьте с кабинетом партнёра.")
    acc_c1, acc_c2, acc_c3, acc_c4 = st.columns(4)
    model = acc_c1.selectbox("Модель", ["FBS","FBY"], index=0 if settings.get("model","FBS")=="FBS" else 1, help="FBS — ваш склад (хранение 0₽), FBY — склад Маркета (платное после льготного периода)")
    payout = acc_c2.selectbox("График выплат", ["weekly_4w","weekly_2w","weekly_1w","daily"], index=["weekly_4w","weekly_2w","weekly_1w","daily"].index(settings.get("payout_schedule","weekly_2w")), format_func=lambda x: {"weekly_4w":"Еженедельно 4н — 1,6%","weekly_2w":"Еженедельно 2н — 2,3%","weekly_1w":"Еженедельно 1н — 2,8%","daily":"Ежедневно — 3,3%"}[x], help="Найдите в Финансы → Настройки выплат")
    use_fixed = acc_c3.toggle("Фиксированный ≤300₽/5л", value=settings.get("use_fixed_tariff", True), help="42% FBS / 35% FBY включает размещение+доставку. Выкл — всегда по сетке средней мили.")
    speed = acc_c4.selectbox("Скидка FBS за скорость", [0.0,0.04,0.07], index=[0.0,0.04,0.07].index(float(settings.get("fbs_speed_discount",0.0))), format_func=lambda x: {0.0:"Стандарт",0.04:"≤36ч −4 п.п.",0.07:"≤28ч −7 п.п. (35%)"}[x], help="Только для FBS. Проверьте в Отчёты → Скорость отгрузки", disabled=(model!="FBS"))
    free_days = st.slider("Бесплатно хранение FBY, дней", 0, 120, int(settings.get("storage_free_days",90)), 15, help="Для FBS не влияет (0₽). Для FBY — льготный период.", disabled=(model!="FBY"))
    st.markdown(f'<div class="small-muted">Средняя миля 2026: 80₽ (≤1л) · +9₽/л до 30л · +7₽ до 200л · +5₽ свыше · потолок 5500₽ · объём = Д×Ш×В/1000 ↑</div>', unsafe_allow_html=True)
    if st.button("Сохранить достоверность", type="primary", use_container_width=True):
        st.session_state.settings["model"] = model
        st.session_state.settings["payout_schedule"] = payout
        st.session_state.settings["use_fixed_tariff"] = use_fixed
        st.session_state.settings["fbs_speed_discount"] = float(speed)
        st.session_state.settings["storage_free_days"] = int(free_days)
        # Перезаписываем acquiring_fee для совместимости
        st.session_state.settings["acquiring_fee"] = PAYOUT_SCHEDULE_RATES[payout]
        save_settings(st.session_state.settings)
        mark_dirty()
        st.success("Достоверная модель сохранена. Если каталог рассчитан — пересчитайте.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.write("")

    with st.form("tariff_form", border=False):
        base_tab, special_tab, costs_tab, fallback_tab, cat_tab, price_tab = st.tabs(
            ["Базовый тариф", "Спецтарифы", "Расходы FBS", "Если данных нет", "Категории", "Цены"]
        )

        with base_tab:
            c1, c2, c3, c4 = st.columns(4)
            commission = c1.number_input(
                "Комиссия, %", 0.0, 60.0, settings["commission_rate"] * 100, 0.5,
                help="Процент Маркета с цены продажи.",
            )
            min_commission = c2.number_input(
                "Минимальная комиссия, ₽", 0.0, 10_000.0,
                settings["min_commission"], 5.0,
            )
            logistics_base = c3.number_input(
                "Логистика: база, ₽", 0.0, 100_000.0,
                settings["logistics_base"], 5.0,
            )
            logistics_per_kg = c4.number_input(
                "Логистика: за кг, ₽", 0.0, 10_000.0,
                settings["logistics_per_kg"], 0.5,
            )

            c1, c2, c3 = st.columns(3)
            storage_rate = c1.number_input(
                "Хранение, ₽/л/сутки", 0.0, 1_000.0,
                settings["storage_per_day_per_liter"], 0.05,
            )
            acquiring = c2.number_input(
                "Эквайринг, %", 0.0, 20.0, settings["acquiring_fee"] * 100, 0.1,
            )
            return_fee = c3.number_input(
                "Резерв возвратов, %", 0.0, 30.0, settings["return_fee"] * 100, 0.1,
            )

        with special_tab:
            special_enabled = st.toggle(
                "Применять спецтарифы автозапчастей",
                value=settings.get("special_enabled", True),
            )
            special_rows = []
            for key, rule in settings["special_tariffs"].items():
                special_rows.append(
                    {
                        "Ключ категории": key,
                        "Название": rule["label"],
                        "Комиссия, %": rule["commission_rate"] * 100,
                        "Логистика, ₽": rule["logistics_base"],
                        "Хранение, ₽/л": rule["storage_per_day_per_liter"],
                        "Причина": rule.get("reason", "Спецтариф"),
                    }
                )
            special_editor = st.data_editor(
                pd.DataFrame(special_rows),
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["Ключ категории", "Название", "Причина"],
                column_config={
                    "Комиссия, %": st.column_config.NumberColumn(min_value=0.0, max_value=60.0, step=0.5),
                    "Логистика, ₽": st.column_config.NumberColumn(min_value=0.0, step=5.0),
                    "Хранение, ₽/л": st.column_config.NumberColumn(min_value=0.0, step=0.05),
                },
            )

        with costs_tab:
            c1, c2, c3 = st.columns(3)
            packaging = c1.number_input("Упаковка FBS, ₽", 0.0, 100_000.0, settings["packaging"], 5.0)
            chestny_znak = c2.number_input("Честный знак, ₽", 0.0, 10_000.0, settings["chestny_znak"], 0.5)
            labeling = c3.number_input("Маркировка, ₽", 0.0, 10_000.0, settings["labeling"], 0.5)
            c1, c2, c3 = st.columns(3)
            warranty = c1.number_input("Гарантийный резерв, %", 0.0, 30.0, settings["warranty_reserve"] * 100, 0.5)
            hazard = c2.number_input("Надбавка: опасный груз, %", 0.0, 30.0, settings["hazard_surcharge"] * 100, 0.5)
            fragile = c3.number_input("Надбавка: хрупкий груз, %", 0.0, 30.0, settings["fragile_surcharge"] * 100, 0.5)

        with fallback_tab:
            st.markdown(
                '<div class="warn-box"><b>Если себестоимости нет:</b> она рассчитывается '
                'как процент от цены. Если веса нет, он оценивается по объёму и плотности.</div>',
                unsafe_allow_html=True,
            )
            c1, c2 = st.columns(2)
            cost_fallback = c1.number_input(
                "Себестоимость, % от цены", 0.0, 99.0,
                settings["cost_fallback_rate"] * 100, 1.0,
            )
            density = c2.number_input(
                "Плотность для оценки веса, кг/л", 0.01, 20.0,
                settings["density_kg_per_liter"], 0.05,
            )

        with cat_tab:
            st.caption("Добавьте свои категории: ключ для поиска в названии, габариты, признаки, свои ставки.")
            custom_df = pd.DataFrame(settings.get("custom_categories", []))
            if custom_df.empty:
                custom_df = pd.DataFrame(columns=["key","label","volume_l","weight_kg","is_hazardous","is_fragile","commission_rate","logistics_base","storage_per_day_per_liter"])
            edited_custom = st.data_editor(
                custom_df,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "key": st.column_config.TextColumn("Ключ"),
                    "label": st.column_config.TextColumn("Название"),
                    "volume_l": st.column_config.NumberColumn("Объём, л", min_value=0.1, step=0.5),
                    "weight_kg": st.column_config.NumberColumn("Вес, кг", min_value=0.1, step=0.5),
                    "is_hazardous": st.column_config.CheckboxColumn("Опасный"),
                    "is_fragile": st.column_config.CheckboxColumn("Хрупкий"),
                    "commission_rate": st.column_config.NumberColumn("Комиссия", min_value=0.0, max_value=1.0, step=0.01, help="0.14 = 14%"),
                    "logistics_base": st.column_config.NumberColumn("Логистика база"),
                    "storage_per_day_per_liter": st.column_config.NumberColumn("Хранение"),
                },
            )
        with price_tab:
            pricing = settings.get("pricing", {"mode":"none","markupPercent":15,"targetMargin":0.20})
            mode = st.radio("Режим ценообразования", ["none","markup","targetMargin"], index=["none","markup","targetMargin"].index(pricing.get("mode","none")), format_func=lambda x: {"none":"Без изменений","markup":"Наценка %","targetMargin":"Целевая маржа"}[x], horizontal=True)
            markup = st.slider("Наценка, %", -50, 200, int(pricing.get("markupPercent",15)), 5)
            target = st.slider("Целевая маржа, %", 0, 60, int(float(pricing.get("targetMargin",0.20))*100), 1)

        submitted = st.form_submit_button(
            "Сохранить тарифы", type="primary", use_container_width=True
        )

    if submitted:
        updated = deep_copy_json(settings)
        updated.update(
            {
                "commission_rate": commission / 100,
                "min_commission": min_commission,
                "logistics_base": logistics_base,
                "logistics_per_kg": logistics_per_kg,
                "storage_per_day_per_liter": storage_rate,
                "acquiring_fee": acquiring / 100,
                "return_fee": return_fee / 100,
                "packaging": packaging,
                "chestny_znak": chestny_znak,
                "labeling": labeling,
                "warranty_reserve": warranty / 100,
                "hazard_surcharge": hazard / 100,
                "fragile_surcharge": fragile / 100,
                "cost_fallback_rate": cost_fallback / 100,
                "density_kg_per_liter": density,
                "special_enabled": special_enabled,
            }
        )
        for _, edited_row in special_editor.iterrows():
            key = str(edited_row["Ключ категории"])
            updated["special_tariffs"][key]["commission_rate"] = float(
                edited_row["Комиссия, %"]
            ) / 100
            updated["special_tariffs"][key]["logistics_base"] = float(
                edited_row["Логистика, ₽"]
            )
            updated["special_tariffs"][key]["storage_per_day_per_liter"] = float(
                edited_row["Хранение, ₽/л"]
            )
        # Save custom categories
        custom_list = []
        for _, row in edited_custom.iterrows():
            if not str(row.get("key","")).strip():
                continue
            custom_list.append({
                "key": str(row.get("key","")).strip().lower(),
                "label": str(row.get("label","")).strip() or str(row.get("key","")).strip(),
                "volume_l": float(row.get("volume_l",2) or 2),
                "weight_kg": float(row.get("weight_kg",1) or 1),
                "is_hazardous": bool(row.get("is_hazardous", False)),
                "is_fragile": bool(row.get("is_fragile", False)),
                "commission_rate": float(row.get("commission_rate")) if pd.notna(row.get("commission_rate")) and str(row.get("commission_rate")).strip() != "" else None,
                "logistics_base": float(row.get("logistics_base")) if pd.notna(row.get("logistics_base")) and str(row.get("logistics_base")).strip() != "" else None,
                "storage_per_day_per_liter": float(row.get("storage_per_day_per_liter")) if pd.notna(row.get("storage_per_day_per_liter")) and str(row.get("storage_per_day_per_liter")).strip() != "" else None,
            })
        updated["custom_categories"] = custom_list
        updated["pricing"] = {"mode": mode, "markupPercent": int(markup), "targetMargin": float(target)/100}
        st.session_state.settings = updated
        save_settings(updated)
        mark_dirty()
        st.success("Тарифы, категории и сценарии цен сохранены. Выполните пересчёт.")

    st.divider()
    st.markdown("### API Яндекс Маркета")
    st.caption(
        "Рекомендуется Api-Key. OAuth оставлен для совместимости. Актуальный метод: "
        "POST /v2/tariffs/calculate."
    )
    api_left, api_right = st.columns([1.05, 1])
    with api_left:
        token_type = st.selectbox("Тип токена", ["Api-Key", "OAuth (устаревший)"])
        token = st.text_input("Токен", type="password", placeholder="Токен не сохраняется")
        campaign_id = st.number_input("Campaign ID", min_value=1, value=123456, step=1)

        with st.expander("Параметры тестового товара", expanded=False):
            category_id = st.number_input("ID категории Маркета", min_value=1, value=90401, step=1)
            test_price = st.number_input("Цена, ₽", min_value=1.0, value=5000.0, step=100.0)
            a1, a2, a3, a4 = st.columns(4)
            test_length = a1.number_input("Длина", min_value=0.1, value=30.0)
            test_width = a2.number_input("Ширина", min_value=0.1, value=20.0)
            test_height = a3.number_input("Высота", min_value=0.1, value=15.0)
            test_weight = a4.number_input("Вес", min_value=0.01, value=2.0)

        if st.button("Рассчитать услуги через API", type="primary", use_container_width=True):
            if not token:
                st.warning("Введите токен.")
            else:
                ok, message, api_df = calculate_tariff_via_api(
                    token,
                    "Api-Key" if token_type == "Api-Key" else "OAuth",
                    int(campaign_id),
                    int(category_id),
                    test_price,
                    test_length,
                    test_width,
                    test_height,
                    test_weight,
                )
                st.session_state.last_api_result = api_df
                (st.success if ok else st.error)(message)

    with api_right:
        if st.session_state.last_api_result is not None:
            st.dataframe(
                st.session_state.last_api_result,
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.markdown(
                '<div class="info-box"><b>Для чего нужен API-калькулятор?</b><br>'
                'Он возвращает ориентировочную стоимость услуг для товара с конкретной '
                'категорией, ценой, габаритами и весом. Для каталога 300 000 SKU используйте '
                'локальный векторизованный расчёт: у API есть лимит запросов.</div>',
                unsafe_allow_html=True,
            )

        c1, c2, c3 = st.columns(3)
        if c1.button("Демо-ставки категорий", use_container_width=True):
            st.session_state.settings["category_rates"] = deep_copy_json(DEMO_CATEGORY_RATES)
            st.session_state.settings["use_category_rates"] = True
            save_settings(st.session_state.settings)
            mark_dirty()
            st.success("Демо-ставки загружены.")
        if c2.button("Очистить ставки", use_container_width=True):
            st.session_state.settings["category_rates"] = {}
            save_settings(st.session_state.settings)
            mark_dirty()
            st.success("Индивидуальные ставки удалены.")
        if c3.button("Ставки из API (legacy)", use_container_width=True):
            if not token:
                st.warning("Введите токен.")
            else:
                ok, message, rates = legacy_category_commissions(
                    token,
                    "Api-Key" if token_type == "Api-Key" else "OAuth",
                    int(campaign_id),
                )
                if ok:
                    st.session_state.settings["category_rates"] = rates
                    st.session_state.settings["use_category_rates"] = True
                    save_settings(st.session_state.settings)
                    mark_dirty()
                    st.success(message)
                else:
                    st.warning(message)

    tariff_csv = tariff_table(st.session_state.settings).to_csv(
        index=False, sep=";", encoding="utf-8-sig", decimal=","
    ).encode("utf-8-sig")
    st.download_button(
        "Скачать текущий тариф CSV",
        tariff_csv,
        file_name="current_tariff.csv",
        mime="text/csv",
    )
    tariff_upload = st.file_uploader(
        "Загрузить тариф из CSV",
        type=["csv"],
        key="tariff_csv_upload",
        help="Поддерживается CSV из этого приложения и старый формат с техническими колонками.",
    )
    if st.button(
        "Применить тариф из CSV",
        use_container_width=True,
        disabled=tariff_upload is None,
    ) and tariff_upload is not None:
        ok, message, updated = update_settings_from_tariff_csv(
            tariff_upload.getvalue(), st.session_state.settings
        )
        if ok:
            st.session_state.settings = updated
            save_settings(updated)
            mark_dirty()
            st.success(message)
            st.session_state.navigation = steps[1]
            st.rerun()
        else:
            st.error(message)

    st.divider()
    st.markdown("### Категории через CSV")
    cat_col1, cat_col2 = st.columns([1,1])
    with cat_col1:
        st.download_button("Шаблон категорий CSV", CATEGORY_CSV_TEMPLATE.encode("utf-8-sig"), file_name="categories_template.csv", mime="text/csv")
    cat_upload = st.file_uploader("Загрузить категории CSV", type=["csv"], key="cat_csv_upload")
    if st.button("Применить категории из CSV", use_container_width=True, disabled=cat_upload is None) and cat_upload is not None:
        ok, msg, updated_cats = parse_category_csv(cat_upload.getvalue(), st.session_state.settings.get("custom_categories", []))
        if ok:
            st.session_state.settings["custom_categories"] = updated_cats
            save_settings(st.session_state.settings)
            mark_dirty()
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)


# =============================================================================
# STEP 2: DATA
# =============================================================================

elif selected_step == steps[1]:
    st.markdown('<div class="section-title">Загрузка каталога</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub">Минимум: Артикул, Категория, Цена. Бренд, габариты, '
        'себестоимость и вес необязательны.</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.4, 0.8])
    with left:
        uploaded = st.file_uploader(
            "CSV или Excel",
            type=["csv", "txt", "tsv", "xlsx", "xls"],
            help="Для 100 000+ SKU рекомендуется CSV.",
        )
        if uploaded is not None:
            st.caption(f"{uploaded.name} · {file_size_label(uploaded.size)}")
        load_clicked = st.button(
            "Загрузить и проверить",
            type="primary",
            use_container_width=True,
            disabled=uploaded is None,
        )
        if load_clicked and uploaded is not None:
            try:
                with st.status("Читаем файл и сопоставляем колонки", expanded=True) as status:
                    raw_file, parse_meta = read_uploaded_file(uploaded)
                    st.write(
                        f"Прочитано {len(raw_file):,} строк и {len(raw_file.columns)} колонок".replace(",", " ")
                    )
                    prepared, preparation_meta = prepare_input_frame(raw_file)
                    parse_meta.update(preparation_meta)
                    st.write("Колонки сопоставлены, числовые значения нормализованы")
                    status.update(label="Файл готов к расчёту", state="complete")

                st.session_state.raw_df = prepared
                st.session_state.result_df = None
                st.session_state.source_name = uploaded.name
                st.session_state.source_size = uploaded.size
                st.session_state.parse_meta = parse_meta
                st.session_state.calculated_settings_hash = ""
                clear_export()
                st.success(f"Загружено {len(prepared):,} SKU".replace(",", " "))
            except Exception as exc:
                logger.exception("File parsing failed")
                st.error(f"Не удалось обработать файл: {exc}")

    with right:
        st.markdown('<div class="soft-panel">', unsafe_allow_html=True)
        st.markdown("**Обязательные колонки**")
        st.markdown(
            '<span class="chip">Артикул</span><span class="chip">Категория</span>'
            '<span class="chip">Цена</span>',
            unsafe_allow_html=True,
        )
        st.markdown("**Рекомендуемые колонки**")
        st.markdown(
            '<span class="chip">Бренд</span><span class="chip">Длина</span>'
            '<span class="chip">Ширина</span><span class="chip">Высота</span>'
            '<span class="chip">Себестоимость</span><span class="chip">Вес_кг</span>',
            unsafe_allow_html=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
        st.download_button(
            "Скачать шаблон CSV",
            build_template_csv(),
            file_name="template_fbs.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.divider()
    st.markdown("### Демо-каталог и нагрузочный тест — полный генератор (не урезан)")
    demo_cols = st.columns(5)
    demo_options = [
        (24, "Пример: 24 SKU"),
        (10_000, "Тест: 10 000 SKU"),
        (50_000, "Тест: 50 000 SKU"),
        (300_000, "Нагрузка: 300 000 SKU"),
        (500_000, "Нагрузка: 500 000 SKU"),
    ]
    for column, (count, label) in zip(demo_cols, demo_options):
        if column.button(label, use_container_width=True):
            with st.spinner(f"Генерируем {count:,} SKU...".replace(",", " ")):
                demo = generate_demo_catalog(count, include_cost=True)
                st.session_state.raw_df = demo
                st.session_state.result_df = None
                st.session_state.source_name = f"demo_{count}_sku.csv"
                st.session_state.source_size = int(demo.memory_usage(deep=True).sum())
                st.session_state.parse_meta = {
                    "input_type": "Демо-каталог",
                    "prepared_rows": count,
                    "has_cost": True,
                    "has_brand": True,
                    "has_dimensions": True,
                    "mapping": {c: c for c in demo.columns},
                    "parse_seconds": 0.0,
                    "zero_prices": 0,
                }
                st.session_state.calculated_settings_hash = ""
                clear_export()
            st.success(f"Сгенерировано {count:,} SKU".replace(",", " "))

    raw = st.session_state.raw_df
    if raw is not None:
        meta = st.session_state.parse_meta
        st.divider()
        st.markdown("### Проверка данных")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("SKU", f"{len(raw):,}".replace(",", " "))
        m2.metric("Брендов", f"{raw['Бренд'].nunique():,}".replace(",", " "))
        m3.metric("Категорий", f"{raw['Категория'].nunique():,}".replace(",", " "))
        m4.metric("Нулевых цен", f"{int(raw['Цена'].le(0).sum()):,}".replace(",", " "))

        mapping = meta.get("mapping", {})
        if mapping:
            chips = "".join(
                f'<span class="chip">{canonical} ← {source}</span>'
                for canonical, source in mapping.items()
            )
            st.markdown(chips, unsafe_allow_html=True)

        if not meta.get("has_cost", False):
            st.markdown(
                f'<div class="warn-box"><b>Себестоимость не найдена.</b> Она будет '
                f'оценена как {percent(st.session_state.settings["cost_fallback_rate"], 0)} '
                'от цены. Такие строки помечаются флагом «Себестоимость_оценка».</div>',
                unsafe_allow_html=True,
            )
        if not meta.get("has_brand", False):
            st.markdown(
                '<div class="info-box">Бренд не найден: используется значение «Без бренда». '
                'Добавьте колонку Бренд для отдельного аналитического разреза.</div>',
                unsafe_allow_html=True,
            )
        if not meta.get("has_dimensions", False):
            st.markdown(
                '<div class="info-box">Полный набор габаритов не найден. Объём и вес будут '
                'дополнены из справочника категорий.</div>',
                unsafe_allow_html=True,
            )

        preview_columns = [
            "Артикул", "Бренд", "Категория", "Длина", "Ширина", "Высота",
            "Цена", "Себестоимость", "Вес_кг",
        ]
        st.dataframe(
            raw[preview_columns].head(100),
            use_container_width=True,
            hide_index=True,
            height=330,
        )
        st.caption("Показаны первые 100 строк. Полный каталог используется в расчёте.")

        if st.button("Рассчитать юнит-экономику", type="primary", use_container_width=True):
            run_calculation()

        st.divider()
        st.markdown("### 💾 Сохранённые каталоги")
        st.caption("Сохраните загруженные данные в браузере сервера и подгружайте снова — даже после перезапуска.")
        col_s1, col_s2 = st.columns([2,1])
        save_name = col_s1.text_input("Название для сохранения", placeholder="Напр. Мой каталог 300k", key="save_name_input")
        if col_s2.button("💾 Сохранить текущий", use_container_width=True):
            if raw is not None and len(raw)>0:
                save_dataset(save_name or f"Каталог {len(raw)} SKU", raw, st.session_state.parse_meta | {"fileName": st.session_state.source_name})
                st.success("Сохранено!")
            else:
                st.warning("Нет данных для сохранения")
        datasets = get_saved_datasets()
        if datasets:
            st.markdown(f"Сохранено наборов: {len(datasets)}")
            for ds in datasets[-10:][::-1]:
                c1,c2,c3 = st.columns([3,1,1])
                c1.markdown(f"**{ds['name']}** — {ds['rows']} SKU · {ds['createdAt'][:19]}")
                if c2.button("Загрузить", key=f"load_{ds['id']}", use_container_width=True):
                    df_loaded, meta_loaded, name_loaded = load_dataset(ds['id'])
                    if df_loaded is not None:
                        st.session_state.raw_df = df_loaded
                        st.session_state.parse_meta = meta_loaded
                        st.session_state.source_name = ds.get("fileName","")
                        st.session_state.result_df = None
                        st.session_state.calculated_settings_hash = ""
                        clear_export()
                        st.success(f"Загружено {len(df_loaded)} SKU")
                        st.rerun()
                if c3.button("🗑️", key=f"del_{ds['id']}"):
                    delete_dataset(ds['id'])
                    st.rerun()
        else:
            st.caption("Пока нет сохранённых наборов.")


# =============================================================================
# STEP 3: DASHBOARD
# =============================================================================

elif selected_step == steps[2]:
    st.markdown('<div class="section-title">Дашборд юнит-экономики</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub">Маржа, прибыль, расходы, бренды, категории и убыточные SKU.</div>',
        unsafe_allow_html=True,
    )
    result = st.session_state.result_df

    if result is None:
        st.markdown(
            '<div class="info-box"><b>Нет рассчитанных данных.</b> Перейдите на шаг «Данные», '
            'загрузите каталог и нажмите «Рассчитать юнит-экономику».</div>',
            unsafe_allow_html=True,
        )
    else:
        if not calculation_is_current():
            st.markdown(
                '<div class="warn-box"><b>Тарифы изменены после расчёта.</b> Цифры ниже '
                'относятся к предыдущим настройкам.</div>',
                unsafe_allow_html=True,
            )
            if st.button("Пересчитать по новым тарифам", type="primary"):
                run_calculation()
                st.rerun()

        totals = calculate_totals(result)
        metric_cols = st.columns(4)
        with metric_cols[0]:
            metric_card(
                "Выручка",
                money_short(totals["revenue"]),
                f"{totals['count']:,} SKU".replace(",", " "),
                "#4f46e5",
            )
        with metric_cols[1]:
            metric_card(
                "Расходы",
                money_short(totals["expenses"]),
                f"{percent(totals['expenses'] / totals['revenue'] if totals['revenue'] else 0, 0)} от выручки",
                "#7c3aed",
            )
        with metric_cols[2]:
            metric_card(
                "Чистая прибыль",
                money_short(totals["profit"]),
                "После комиссий, логистики и хранения",
                "#059669" if totals["profit"] >= 0 else "#e11d48",
            )
        with metric_cols[3]:
            metric_card(
                "Маржа по обороту",
                percent(totals["margin"]),
                f"Убыточных SKU: {totals['loss_count']:,}".replace(",", " "),
                "#059669" if totals["margin"] >= 0.15 else "#d97706",
            )

        st.write("")
        stat_cols = st.columns(6)
        avg_rec = float(result["Рекомендованная_цена"].mean()) if len(result) else 0
        stat_cols[0].metric(
            "Эффективная комиссия",
            percent(totals["commission"] / totals["revenue"] if totals["revenue"] else 0),
        )
        stat_cols[1].metric(
            "Средняя логистика", money(totals["logistics"] / max(totals["count"], 1))
        )
        stat_cols[2].metric(
            "Среднее хранение", money(totals["storage"] / max(totals["count"], 1))
        )
        stat_cols[3].metric("Ср. реком. цена", money(avg_rec))
        stat_cols[4].metric(
            "Спецтариф", f"{totals['special_count']:,} SKU".replace(",", " ")
        )
        stat_cols[5].metric(
            "К поднятию цены", f"{int((result['Прибыль']<0).sum()):,}".replace(",", " ")
        )

        with st.expander("💹 Сценарий цен — наценка или целевая маржа (пересчитывает всё)", expanded=False):
            cur_pricing = st.session_state.settings.get("pricing", {"mode":"none","markupPercent":15,"targetMargin":0.2})
            c1,c2,c3 = st.columns([1,1,1])
            mode_p = c1.radio("Режим", ["none","markup","targetMargin"], index=["none","markup","targetMargin"].index(cur_pricing.get("mode","none")), format_func=lambda x: {"none":"Без изменений","markup":"Наценка %","targetMargin":"Целевая маржа"}[x])
            markup_v = c2.slider("Наценка %", -50, 200, int(cur_pricing.get("markupPercent",15)), 5)
            target_v = c3.slider("Целевая маржа %", 0, 60, int(float(cur_pricing.get("targetMargin",0.2))*100), 1)
            if st.button("Применить сценарий цен", type="primary"):
                st.session_state.settings["pricing"] = {"mode": mode_p, "markupPercent": markup_v, "targetMargin": target_v/100}
                save_settings(st.session_state.settings)
                mark_dirty()
                st.success("Сценарий сохранён — пересчитайте каталог")
                if st.button("Пересчитать сейчас"):
                    run_calculation()
                    st.rerun()

        # ABC / XYZ
        try:
            abc_counts = result["ABC"].value_counts()
            xyz_counts = result["XYZ"].value_counts()
            st.markdown("### 🔤 ABC / XYZ анализ")
            abc_col1, abc_col2 = st.columns(2)
            with abc_col1:
                fig_abc = px.pie(values=[abc_counts.get("A",0), abc_counts.get("B",0), abc_counts.get("C",0)], names=["A 80%","B 15%","C 5%"], hole=0.55, color_discrete_sequence=["#10b981","#f59e0b","#6366f1"], title="ABC по выручке")
                st.plotly_chart(fig_abc, use_container_width=True)
            with abc_col2:
                # matrix
                matrix = result.groupby(["ABC","XYZ"]).size().unstack(fill_value=0)
                st.markdown("**Матрица ABC/XYZ (кол-во SKU)**")
                st.dataframe(matrix, use_container_width=True)
                st.caption("X — быстрый оборот, Z — медленный. AX идеально, CZ — кандидаты на вывод.")
        except Exception as e:
            st.caption(f"ABC/XYZ пока недоступен: {e}")

        st.divider()
        mode = st.radio(
            "Разрез аналитики",
            options=["Категория", "Бренд"],
            horizontal=True,
        )
        summary = aggregate_by(result, mode)
        chart_summary = summary.head(15).copy()

        chart_left, chart_right = st.columns(2)
        with chart_left:
            margin_colors = [
                "#10b981" if value >= 0.15 else "#f59e0b" if value >= 0 else "#ef4444"
                for value in chart_summary["Маржа_%"]
            ]
            fig_margin = go.Figure(
                go.Bar(
                    x=chart_summary[mode],
                    y=chart_summary["Маржа_%"] * 100,
                    marker_color=margin_colors,
                    hovertemplate="%{x}<br>Маржа: %{y:.1f}%<extra></extra>",
                )
            )
            fig_margin.update_layout(
                title=f"Маржа: топ-15 по {mode.lower()}",
                yaxis_title="Маржа, %",
                xaxis_title="",
                height=410,
                margin=dict(l=20, r=20, t=55, b=80),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="white",
            )
            st.plotly_chart(fig_margin, use_container_width=True)

        with chart_right:
            costs = cost_structure(totals)
            fig_cost = px.pie(
                costs,
                names="Статья",
                values="Сумма",
                hole=0.58,
                color="Статья",
                color_discrete_sequence=[
                    "#6366f1", "#f59e0b", "#0ea5e9", "#8b5cf6", "#14b8a6", "#f43f5e",
                ],
                title="Структура расходов",
            )
            fig_cost.update_traces(textposition="inside", textinfo="percent")
            fig_cost.update_layout(
                height=410,
                margin=dict(l=20, r=20, t=55, b=20),
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_cost, use_container_width=True)

        profit_colors = np.where(chart_summary["Прибыль"] >= 0, "#10b981", "#ef4444")
        fig_profit = go.Figure(
            go.Bar(
                x=chart_summary[mode],
                y=chart_summary["Прибыль"],
                marker_color=profit_colors,
                hovertemplate="%{x}<br>Прибыль: %{y:,.0f} ₽<extra></extra>",
            )
        )
        fig_profit.update_layout(
            title=f"Прибыль: топ-15 по {mode.lower()}",
            yaxis_title="Прибыль, ₽",
            xaxis_title="",
            height=420,
            margin=dict(l=20, r=20, t=55, b=90),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_profit, use_container_width=True)

        st.markdown(f"### Сводка по {mode.lower()}")
        display_summary = summary.head(100).copy()
        display_summary["Маржа_%"] = display_summary["Маржа_%"] * 100
        st.dataframe(
            display_summary,
            use_container_width=True,
            hide_index=True,
            height=420,
            column_config={
                "Выручка": st.column_config.NumberColumn(format="%.0f ₽"),
                "Расходы": st.column_config.NumberColumn(format="%.0f ₽"),
                "Прибыль": st.column_config.NumberColumn(format="%.0f ₽"),
                "Маржа_%": st.column_config.NumberColumn("Маржа, %", format="%.1f%%"),
            },
        )
        st.caption("На экране показаны первые 100 групп. Полная сводка войдёт в Excel.")

        if totals["loss_count"]:
            st.markdown("### Самые убыточные SKU")
            loss_view = result.nsmallest(50, "Прибыль")[
                [
                    "Артикул", "Бренд", "Категория", "Цена", "Себестоимость",
                    "Логистика_руб", "Хранение_руб", "Прибыль", "Маржа_%",
                ]
            ].copy()
            loss_view["Маржа_%"] *= 100
            st.dataframe(
                loss_view,
                use_container_width=True,
                hide_index=True,
                height=300,
                column_config={
                    "Прибыль": st.column_config.NumberColumn(format="%.0f ₽"),
                    "Маржа_%": st.column_config.NumberColumn("Маржа, %", format="%.1f%%"),
                },
            )

        st.divider()
        st.markdown("### Расчёт по SKU")
        filter_cols = st.columns([1.4, 1, 1, 0.7])
        search = filter_cols[0].text_input(
            "Поиск", placeholder="Артикул, бренд или категория"
        ).strip().lower()
        brands = ["Все"] + result["Бренд"].value_counts().head(500).index.astype(str).tolist()
        categories = ["Все"] + result["Категория"].value_counts().head(500).index.astype(str).tolist()
        selected_brand = filter_cols[1].selectbox("Бренд", brands)
        selected_category = filter_cols[2].selectbox("Категория", categories)
        only_loss = filter_cols[3].toggle("Только убыток")

        mask = np.ones(len(result), dtype=bool)
        if selected_brand != "Все":
            mask &= result["Бренд"].eq(selected_brand).to_numpy()
        if selected_category != "Все":
            mask &= result["Категория"].eq(selected_category).to_numpy()
        if only_loss:
            mask &= result["Прибыль"].lt(0).to_numpy()
        if search:
            mask &= (
                result["Артикул"].astype("string").str.lower().str.contains(search, regex=False, na=False)
                | result["Бренд"].astype("string").str.lower().str.contains(search, regex=False, na=False)
                | result["Категория"].astype("string").str.lower().str.contains(search, regex=False, na=False)
            ).to_numpy()

        filtered = result.loc[mask]
        screen_df = filtered.head(TABLE_PREVIEW_LIMIT).copy()
        screen_df["Маржа_%"] *= 100
        st.dataframe(
            screen_df,
            use_container_width=True,
            hide_index=True,
            height=610,
            column_config={
                "Маржа_%": st.column_config.ProgressColumn(
                    "Маржа, %", min_value=-100, max_value=50, format="%.1f%%"
                ),
                "Ставка_комиссии": st.column_config.NumberColumn(format="%.2f"),
                "Себестоимость_оценка": st.column_config.CheckboxColumn("Себест. оценена"),
            },
        )
        st.caption(
            f"Найдено {len(filtered):,} SKU. На экране до {TABLE_PREVIEW_LIMIT:,}; "
            "в экспорт попадут все строки.".replace(",", " ")
        )


# =============================================================================
# STEP 4: EXPORT
# =============================================================================

else:
    st.markdown('<div class="section-title">Экспорт отчётов</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-sub">Excel с формулами, быстрый Excel для 300 000 SKU, '
        'компактная сводка и полный CSV.</div>',
        unsafe_allow_html=True,
    )
    result = st.session_state.result_df

    if result is None:
        st.markdown(
            '<div class="info-box"><b>Экспортировать пока нечего.</b> Загрузите каталог '
            'и выполните расчёт.</div>',
            unsafe_allow_html=True,
        )
    else:
        totals = calculate_totals(result)
        st.markdown(
            f'<div class="success-box"><b>Отчёт готов:</b> '
            f'{len(result):,} SKU · Выручка {money_short(totals["revenue"])} · '
            f'Прибыль {money_short(totals["profit"])} · '
            f'Маржа {percent(totals["margin"])}</div>'.replace(",", " "),
            unsafe_allow_html=True,
        )

        export_cols = st.columns(4)
        with export_cols[0]:
            st.markdown("#### Excel с формулами")
            st.caption(
                "Живые формулы, входные данные, категории, бренды и легенда. "
                f"Лимит: {FORMULA_EXCEL_LIMIT:,} строк.".replace(",", " ")
            )
            formula_disabled = len(result) > FORMULA_EXCEL_LIMIT
            if formula_disabled:
                st.warning("Для этого объёма используйте быстрый Excel.")
            if st.button(
                "Сформировать с формулами",
                type="primary",
                use_container_width=True,
                disabled=formula_disabled,
            ):
                with st.spinner("Формируем Excel с формулами..."):
                    try:
                        data = export_formula_excel(result, st.session_state.settings)
                        st.session_state.export_bytes = data
                        st.session_state.export_name = (
                            f"unit_economy_formulas_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                        )
                        st.session_state.export_mime = (
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as exc:
                        logger.exception("Formula export failed")
                        st.error(f"Ошибка Excel: {exc}")

        with export_cols[1]:
            st.markdown("#### Быстрый Excel")
            st.caption(
                "Все значения, условное форматирование, сводки и графики. "
                "Рекомендуется для больших каталогов."
            )
            if st.button("Сформировать быстрый Excel", use_container_width=True):
                with st.spinner("Формируем полный Excel. Для 300 000 SKU это займёт время..."):
                    try:
                        data = export_values_excel(result, st.session_state.settings)
                        st.session_state.export_bytes = data
                        st.session_state.export_name = (
                            f"unit_economy_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                        )
                        st.session_state.export_mime = (
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as exc:
                        logger.exception("Values export failed")
                        st.error(f"Ошибка Excel: {exc}")

        with export_cols[2]:
            st.markdown("#### Только сводки")
            st.caption(
                "Компактный Excel: итоги, категории, бренды, тариф и легенда. "
                "Формируется быстро при любом объёме."
            )
            if st.button("Сформировать сводный Excel", use_container_width=True):
                with st.spinner("Формируем сводный отчёт..."):
                    try:
                        data = export_summary_excel(result, st.session_state.settings)
                        st.session_state.export_bytes = data
                        st.session_state.export_name = (
                            f"unit_economy_summary_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
                        )
                        st.session_state.export_mime = (
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as exc:
                        logger.exception("Summary export failed")
                        st.error(f"Ошибка Excel: {exc}")

        with export_cols[3]:
            st.markdown("#### Полный CSV")
            st.caption(
                "Самый надёжный формат для 300 000 SKU. Все строки и расчётные колонки, "
                "UTF-8 BOM, разделитель «;»."
            )
            if st.button("Сформировать CSV", use_container_width=True):
                with st.spinner("Формируем CSV..."):
                    try:
                        data = export_csv(result)
                        st.session_state.export_bytes = data
                        st.session_state.export_name = (
                            f"unit_economy_{datetime.now():%Y%m%d_%H%M%S}.csv"
                        )
                        st.session_state.export_mime = "text/csv"
                    except Exception as exc:
                        logger.exception("CSV export failed")
                        st.error(f"Ошибка CSV: {exc}")

        if st.session_state.export_bytes is not None:
            st.divider()
            st.markdown(
                f'<div class="success-box"><b>Файл сформирован:</b> '
                f'{st.session_state.export_name} · '
                f'{file_size_label(len(st.session_state.export_bytes))}</div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "Скачать готовый файл",
                data=st.session_state.export_bytes,
                file_name=st.session_state.export_name,
                mime=st.session_state.export_mime,
                type="primary",
                use_container_width=True,
            )

        st.divider()
        st.markdown("### Что входит в Excel")
        sheet_info = pd.DataFrame(
            [
                ["Итоги", "Выручка, расходы, прибыль, маржа, проблемные SKU"],
                ["Расчет_FBS", "Полный расчёт по SKU; в формульной версии — живые формулы"],
                ["Входные_Данные", "Исходные параметры; только в формульной версии"],
                ["Сводка_Категории", "SKU, выручка, прибыль, маржа и график по категориям"],
                ["Сводка_Бренды", "Аналогичная аналитика по брендам"],
                ["Тариф", "Все ставки и допущения расчёта"],
                ["Легенда", "Описание колонок и формул"],
            ],
            columns=["Лист", "Содержание"],
        )
        st.dataframe(sheet_info, use_container_width=True, hide_index=True)


# =============================================================================
# FOOTER
# =============================================================================

st.divider()
st.markdown(
    f'<div class="small-muted" style="text-align:center">{APP_NAME} · '
    f'версия {APP_VERSION} · монолитное приложение Streamlit · '
    'расчёты выполняются локально</div>',
    unsafe_allow_html=True,
)
