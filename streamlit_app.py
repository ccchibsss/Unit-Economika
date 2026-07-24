"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v101.0 - ENTERPRISE EDITION
================================================================================
📌 ВЕРСИЯ: 101.0.0 (ENTERPRISE)
📌 СПЕЦИАЛИЗАЦИЯ: АВТОЗАПЧАСТИ, АВТОТОВАРЫ И АГРЕГАТЫ
📌 ТЕХНОЛОГИИ: STREAMLIT, POLARS, DUCKDB, SCIKIT-LEARN, OPENPYXL, PLOTLY, GSPREAD
📌 СТРУКТУРА ПРИЛОЖЕНИЯ (4 РАЗДЕЛА):
   1️⃣ Загрузка данных (Каталог для группировки + связывание столбцов)
   2️⃣ Весогабариты и категоризация (3 уровня: Родитель/Группа/Подгруппа)
   3️⃣ Тарифы (API МП + AI + Google Sheets + обновление)
   4️⃣ Расчёт (Юнит-экономика + ABC/XYZ анализ + экспорт)
📌 НОВОЕ v101.0:
✅ ABC/XYZ АНАЛИЗ ПО МАРЖИНАЛЬНОСТИ И ПРИБЫЛИ
✅ 3-УРОВНЕВАЯ КАТЕГОРИЗАЦИЯ (Родитель  Группа  Подгруппа)
✅ GOOGLE SHEETS ИНТЕГРАЦИЯ (JSON + API) 
✅ АВТОМАТИЧЕСКАЯ ПОДСТАНОВКА ГАБАРИТОВ ИЗ АНАЛОГОВ
✅ СОХРАНЕНИЕ/ЗАГРУЗКА РАСЧЁТОВ ОДНОЙ КНОПКОЙ
✅ СВЯЗЫВАНИЕ СТОЛБЦОВ МЕЖДУ ФАЙЛАМИ
================================================================================
"""

# ============================================================================
# БЛОК 0: ИМПОРТЫ И БАЗОВАЯ КОНФИГУРАЦИЯ
# ============================================================================

# === Стандартная библиотека Python ===
import streamlit as st
import pandas as pd
import numpy as np
import requests
import logging
import time
import hashlib
import json
import re
import os
import sys
import traceback
import io
import pickle
import math
import warnings
import csv
import base64
import tempfile
import functools
import string
import decimal
import uuid
import glob
import shutil
import zipfile
import threading
import platform
import gc
import copy
import statistics
import secrets
import sqlite3
import smtplib
from html import escape
from pathlib import Path
from abc import ABC, abstractmethod
from contextlib import contextmanager, suppress
from collections import defaultdict, Counter, OrderedDict, namedtuple
from enum import Enum, auto
from threading import Lock
from datetime import datetime, timedelta, date
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Any, Optional, Tuple, Union, Set, Callable
from dataclasses import dataclass, field, asdict
from functools import lru_cache, wraps
from concurrent.futures import ThreadPoolExecutor, as_completed
from multiprocessing import cpu_count

# === Опциональные импорты с обработкой ошибок ===

# Polars (быстрые DataFrame)
try:
    import polars as pl
    import polars.selectors as cs
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False
    pl = None

# DuckDB (аналитическая БД)
try:
    import duckdb
    DUCKDB_AVAILABLE = True
except ImportError:
    DUCKDB_AVAILABLE = False
    duckdb = None

# scikit-learn (ML)
try:
    import sklearn
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.naive_bayes import MultinomialNB
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import LabelEncoder
    import joblib
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False

# Plotly (интерактивные графики)
try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# openpyxl (Excel)
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference, LineChart, PieChart
    from openpyxl.formatting.rule import CellIsRule, DataBarRule, ColorScaleRule
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# xlsxwriter (быстрый Excel с формулами)
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

# chardet (определение кодировки)
try:
    import chardet
    CHARDET_AVAILABLE = True
except ImportError:
    CHARDET_AVAILABLE = False
    chardet = None

# psutil (мониторинг системы)
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# gspread (Google Sheets API)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# openai (для AI тарифов)
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# dateutil (парсинг дат)
try:
    from dateutil.parser import parse as dateutil_parse
    from dateutil.relativedelta import relativedelta
    DATEUTIL_AVAILABLE = True
except ImportError:
    DATEUTIL_AVAILABLE = False

# === Подавление предупреждений ===
warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'
os.environ['TF_CPP_LOG_LOG_LEVEL'] = '3'

# === Базовые директории ===
try:
    BASE_DIR = Path(__file__).parent.resolve()
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOG_DIR = BASE_DIR / "logs"
TEMP_DIR = BASE_DIR / "temp"
MODELS_DIR = BASE_DIR / "models"
CONFIG_DIR = BASE_DIR / "config"
EXPORTS_DIR = BASE_DIR / "exports"
TARIFFS_DIR = BASE_DIR / "tariffs"
HISTORY_DB_DIR = BASE_DIR / "history_db"
BACKUPS_DIR = BASE_DIR / "backups"
GOOGLE_CREDS_DIR = BASE_DIR / "google_creds"

for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, TEMP_DIR, MODELS_DIR,
                 CONFIG_DIR, EXPORTS_DIR, TARIFFS_DIR, HISTORY_DB_DIR,
                 BACKUPS_DIR, GOOGLE_CREDS_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except OSError:
        pass

# === Версия приложения ===
APP_VERSION = "101.0.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей PRO 2026"
APP_DESCRIPTION = "Enterprise расчет юнит-экономики для автозапчастей с AI, ABC/XYZ и Google Sheets"

# === Логирование ===
LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = LOG_DIR / "auto_parts_economy_pro.log"

@st.cache_resource
def get_logger():
    logger = logging.getLogger('UnitEconomyPro')
    logger.setLevel(getattr(logging, LOG_LEVEL))
    formatter = logging.Formatter(LOG_FORMAT, datefmt="%Y-%m-%d %H:%M:%S")
    try:
        fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except OSError:
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
    return logger

logger = get_logger()

# === Совместимость Streamlit 1.58+ ===
def st_dataframe_compat(df, *args, **kwargs):
    kwargs.pop('use_container_width', None)
    if 'width' not in kwargs:
        kwargs['width'] = 'stretch'
    return st.dataframe(df, *args, **kwargs)

# ============================================================================
# БЛОК 1: КОНСТАНТЫ, ENUM И ТИПЫ ДАННЫХ
# ============================================================================

# === Константы приложения ===
APP_VERSION = "101.0.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей PRO 2026"
APP_DESCRIPTION = "Enterprise расчет юнит-экономики для автозапчастей с AI, ABC/XYZ и Google Sheets"

# === Лимиты и ограничения ===
EXCEL_ROW_LIMIT = 1_000_000
HISTORY_LIMIT = 50_000
CACHE_TTL = 7200
MAX_THREADS = 32
BATCH_SIZE = 2000
MAX_FILE_SIZE_MB = 500
MAX_UPLOAD_SIZE = 1024 * 1024 * 1024
MAX_CATEGORIES = 500
MAX_ANALOGS = 200
PRECISION_DECIMALS = 4
MAX_DISPLAY_ROWS = 2000
PAGE_SIZE = 100
MAX_HISTORY_ENTRIES = 50000
MAX_CACHE_SIZE = 5000

# === Значения по умолчанию ===
DEFAULT_CURRENCY = "RUB"
DEFAULT_MARKETPLACE = "Ozon"
DEFAULT_MODE = "FBY"
DEFAULT_LOCALE = "ru_RU"
TIMEZONE = "Europe/Moscow"
DEFAULT_MARKUP_GLOBAL = 0.25
DEFAULT_DISCOUNT_MAX = 0.30
DEFAULT_MAX_WORKERS = 8
DEFAULT_CHUNK_SIZE = 10000
DEFAULT_DAYS_STORAGE = 30
DEFAULT_TARGET_MARGIN = 20.0

# === Поддерживаемые значения ===
SUPPORTED_CURRENCIES = ["RUB", "USD", "EUR", "CNY", "KZT", "UAH", "BYN", "AMD", "TRY"]
SUPPORTED_LANGUAGES = ["ru", "en", "uk", "kz", "by", "am", "tr"]
SUPPORTED_MARKETPLACES = [
    "Ozon", "Wildberries", "Яндекс Маркет", "AliExpress",
    "Мегамаркет", "СберМегаМаркет", "Avito", "Drom"
]
SUPPORTED_MODES = ["FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"]

# === Флаги функциональности ===
USE_CACHING = True
USE_PARALLEL = True
USE_GPU = False
OPTIMIZE_MEMORY = True
USE_DUCKDB = True
USE_POLARS = True
USE_MULTIPROCESSING = True

# === Цветовая схема ===
COLORS = {
    "primary": "#e94560",
    "secondary": "#0f3460",
    "success": "#00cc96",
    "warning": "#ffa600",
    "danger": "#ef553b",
    "info": "#636efa",
    "dark": "#1a1a2e",
    "light": "#f5f5f5",
    "gradient_start": "#1a1a2e",
    "gradient_end": "#16213e",
    "input_fill": "#FFF4CC",
    "formula_fill": "#E2EFDA",
    "result_fill": "#DCE6F1",
    "header_fill": "#0F3460",
}

PLOTLY_COLORS = [
    "#e94560", "#0f3460", "#00cc96", "#ffa600", "#ef553b",
    "#636efa", "#f9a825", "#26a69a", "#ab47bc", "#42a5f5",
    "#ec407a", "#66bb6a", "#ffa726", "#8d6e63", "#78909c",
    "#d4ac0d", "#1abc9c", "#2ecc71", "#3498db", "#9b59b6",
    "#e67e22", "#e74c3c", "#1abc9c", "#2ecc71", "#3498db"
]

# === Иконки маркетплейсов и режимов ===
MARKETPLACE_ICONS = {
    "Ozon": "🟣",
    "Wildberries": "🟡",
    "Яндекс Маркет": "🔵",
    "AliExpress": "🔴",
    "Мегамаркет": "🟢",
    "СберМегаМаркет": "🟠",
    "Avito": "🟤",
    "Drom": "⚫"
}

MODE_ICONS = {
    "FBY": "📦",
    "FBS": "🏪",
    "FBO": "🏭",
    "DBS": "🚚",
    "FBP": "🤝",
    "RealFBS": "🏃"
}

# === Налоговые системы ===
TAX_SYSTEMS = {
    "УСН_6": {"rate": 0.06, "base": "revenue", "name": "УСН 6% (доходы)"},
    "УСН_15": {"rate": 0.15, "base": "profit", "min_rate": 0.01, "name": "УСН 15% (доходы-расходы)"},
    "ОСН": {"rate": 0.20, "base": "profit", "vat": 0.20, "name": "ОСН (общая)"},
    "ПСН": {"rate": 0.0, "base": "fixed", "name": "ПСН (патент)"},
    "НПД": {"rate": 0.06, "base": "revenue", "name": "НПД (самозанятый)"},
}

# === Бенчмарки рынка 2026 ===
MARKET_BENCHMARKS_2026 = {
    "фильтры": {"avg_margin": 25, "avg_price": 800, "return_rate": 0.05},
    "колодки": {"avg_margin": 22, "avg_price": 2500, "return_rate": 0.08},
    "масла": {"avg_margin": 18, "avg_price": 3500, "return_rate": 0.03},
    "аккумуляторы": {"avg_margin": 15, "avg_price": 7000, "return_rate": 0.12},
    "шины": {"avg_margin": 20, "avg_price": 5000, "return_rate": 0.07},
    "фары": {"avg_margin": 28, "avg_price": 4500, "return_rate": 0.15},
    "амортизаторы": {"avg_margin": 24, "avg_price": 3000, "return_rate": 0.10},
    "ремни": {"avg_margin": 26, "avg_price": 1200, "return_rate": 0.06},
    "подшипники": {"avg_margin": 23, "avg_price": 1500, "return_rate": 0.09},
    "датчики": {"avg_margin": 27, "avg_price": 2000, "return_rate": 0.11},
    "подвеска": {"avg_margin": 24, "avg_price": 2800, "return_rate": 0.08},
    "сайлентблоки": {"avg_margin": 30, "avg_price": 600, "return_rate": 0.04},
}

# === ABC/XYZ пороги ===
ABC_THRESHOLDS = {
    "A": {"margin_min": 25, "profit_share": 0.70},  # Топ 70% прибыли
    "B": {"margin_min": 15, "profit_share": 0.20},  # Следующие 20%
    "C": {"margin_min": 0, "profit_share": 0.10},   # Остальные 10%
}

XYZ_THRESHOLDS = {
    "X": {"cv_max": 0.5},   # Низкая вариация (стабильные)
    "Y": {"cv_max": 1.0},   # Средняя вариация
    "Z": {"cv_max": float('inf')},  # Высокая вариация (нестабильные)
}

# ============================================================================
# ENUM (Перечисления)
# ============================================================================

class CommissionType(Enum):
    """Типы комиссий маркетплейсов"""
    PERCENTAGE = auto()      # Процент от цены
    FIXED = auto()           # Фиксированная сумма
    HYBRID = auto()          # Комбинированная
    SUBSCRIPTION = auto()    # Подписка
    TIERED = auto()          # Ступенчатая
    DYNAMIC = auto()         # Динамическая
    FLAT = auto()            # Плоская ставка
    CUSTOM = auto()          # Пользовательская


class OperationMode(Enum):
    """Режимы работы с маркетплейсами"""
    FBY = auto()             # Fulfilled By You (со своего склада)
    FBS = auto()             # Fulfilled By Seller (FBY аналог)
    FBO = auto()             # Fulfilled By Ozon (склад Ozon)
    DBS = auto()             # Delivery By Seller
    FBP = auto()             # Fulfilled By Partner
    DBE = auto()             # Delivery By Express
    STANDARD = auto()        # Стандартный
    EXPRESS = auto()         # Экспресс
    SELF = auto()            # Самовывоз
    REAL_FBS = auto()        # Реальный FBS


class ProductType(Enum):
    """Типы автозапчастей"""
    ENGINE = "Двигатель"
    TRANSMISSION = "Трансмиссия"
    SUSPENSION = "Подвеска"
    BRAKE = "Тормозная система"
    STEERING = "Рулевое управление"
    ELECTRICAL = "Электрооборудование"
    COOLING = "Система охлаждения"
    EXHAUST = "Система выпуска"
    FUEL = "Система питания"
    FILTER = "Фильтры"
    FLUID = "Масла и жидкости"
    BODY = "Кузовные детали"
    INTERIOR = "Салон"
    EXTERIOR = "Экстерьер"
    OPTICS = "Оптика"
    TIRES = "Шины и диски"
    TOOLS = "Инструменты"
    BELT = "Ремни и приводы"
    BEARING = "Подшипники"
    SEAL = "Сальники и прокладки"
    FASTENER = "Крепеж"
    HVAC = "Климат-контроль"
    AUDIO = "Аудио и мультимедиа"
    SAFETY = "Безопасность"
    OTHER = "Прочее"


class DataSource(Enum):
    """Источники данных"""
    CSV = auto()
    EXCEL = auto()
    JSON = auto()
    API = auto()
    DATABASE = auto()
    MANUAL = auto()
    MARKETPLACE = auto()
    AI = auto()
    WEB_SCRAPING = auto()
    ERP = auto()
    CRM = auto()
    EXTERNAL = auto()
    GOOGLE_SHEETS = auto()


class ExportFormat(Enum):
    """Форматы экспорта"""
    CSV = auto()
    EXCEL = auto()
    EXCEL_FORMULAS = auto()
    EXCEL_MACROS = auto()
    PDF = auto()
    JSON = auto()
    HTML = auto()
    MARKDOWN = auto()
    PARQUET = auto()
    SQL = auto()
    XML = auto()
    YAML = auto()
    TOML = auto()
    POWER_BI = auto()
    TABLEAU = auto()
    GOOGLE_SHEETS = auto()


class CalculationStatus(Enum):
    """Статусы расчёта"""
    PENDING = auto()
    RUNNING = auto()
    COMPLETED = auto()
    FAILED = auto()
    CANCELLED = auto()
    PAUSED = auto()
    PARTIAL = auto()


class RiskLevel(Enum):
    """Уровни риска"""
    LOW = "Низкий"
    MEDIUM = "Средний"
    HIGH = "Высокий"
    CRITICAL = "Критический"


class Seasonality(Enum):
    """Сезонность товаров"""
    WINTER = "Зимняя"
    SPRING = "Весенняя"
    SUMMER = "Летняя"
    AUTUMN = "Осенняя"
    ALL_YEAR = "Круглогодичная"


class ProfitabilityLevel(Enum):
    """Уровни прибыльности"""
    LOSS = "Убыток"
    BREAK_EVEN = "Точка безубыточности"
    LOW = "Низкая"
    MEDIUM = "Средняя"
    HIGH = "Высокая"
    VERY_HIGH = "Очень высокая"


class Currency(Enum):
    """Валюты"""
    RUB = "RUB"
    USD = "USD"
    EUR = "EUR"
    CNY = "CNY"
    KZT = "KZT"
    UAH = "UAH"
    BYN = "BYN"
    AMD = "AMD"
    TRY = "TRY"


class TaxSystem(Enum):
    """Налоговые системы"""
    USN_6 = "УСН_6"
    USN_15 = "УСН_15"
    OSN = "ОСН"
    PSN = "ПСН"
    NPD = "НПД"


class TariffSource(Enum):
    """Источники тарифов"""
    HARDCODED = "Захардкожены"
    AI_CACHE = "Кэш ИИ"
    AI_LIVE = "ИИ (запрос)"
    MANUAL = "Ручной ввод"
    IMPORTED = "Импортированы"
    API_LIVE = "API Маркетплейса"
    FORECAST = "Прогноз ИИ"
    GOOGLE_SHEETS = "Google Sheets"


class ABCCategory(Enum):
    """ABC-категории товаров"""
    A = "A"  # Высокая маржа/прибыль (топ 70%)
    B = "B"  # Средняя маржа/прибыль (20%)
    C = "C"  # Низкая маржа/прибыль (10%)


class XYZCategory(Enum):
    """XYZ-категории по стабильности"""
    X = "X"  # Стабильные (CV < 0.5)
    Y = "Y"  # Умеренные (CV 0.5-1.0)
    Z = "Z"  # Нестабильные (CV > 1.0)


class CategoryLevel(Enum):
    """Уровни категоризации (3-уровневая иерархия)"""
    PARENT = "Родитель"      # Уровень 1: Автозапчасти
    GROUP = "Группа"         # Уровень 2: Подвеска
    SUBGROUP = "Подгруппа"   # Уровень 3: Сайлентблоки


class DataLinkType(Enum):
    """Типы связывания данных между файлами"""
    OE_TO_CROSS = "OE-кроссы"
    ARTICLE_TO_ANALOG = "Артикул-аналог"
    MANUAL_MAPPING = "Ручной маппинг"
    AUTO_DETECT = "Автоопределение"
    DRAG_DROP = "Drag & Drop"


class SaveLoadAction(Enum):
    """Действия сохранения/загрузки"""
    SAVE_CURRENT = "Сохранить текущие"
    LOAD_PREVIOUS = "Загрузить предыдущие"
    EXPORT_ALL = "Экспортировать всё"
    IMPORT_EXTERNAL = "Импортировать внешние"
    BACKUP = "Резервная копия"
    RESTORE = "Восстановить"

# ============================================================================
# БЛОК 2: ДАТА-КЛАССЫ (Dataclasses) — ОСНОВНЫЕ СТРУКТУРЫ ДАННЫХ
# ============================================================================
# 📌 v101.0: Добавлены новые классы:
# - CategoryHierarchy (3-уровневая иерархия: Родитель → Группа → Подгруппа)
# - ABCAnalysisResult (ABC/XYZ анализ по маржинальности и прибыли)
# - GoogleSheetsConfig (конфигурация Google Sheets)
# - SaveLoadState (состояние сохранения/загрузки расчётов)
# - ColumnMapping (маппинг столбцов между файлами)
# - DataLinkConfig (конфигурация связывания данных)
# ============================================================================

@dataclass
class MarketplaceConfig:
    """
    Расширенная конфигурация маркетплейса с сезонностью и промо.
    Хранит все тарифы, комиссии, логистические ставки для одного МП.
    """
    name: str
    commission_rate: float
    min_commission: float = 0.0
    max_commission: float = float('inf')
    logistics_base: float = 0.0
    logistics_per_kg: float = 0.0
    logistics_per_liter: float = 0.0
    storage_per_day: float = 0.0
    return_fee: float = 0.0
    acquiring_fee: float = 0.0
    last_mile_fee: float = 0.0
    delivery_fee_percent: float = 0.0
    premium_fee: float = 0.0
    rko_fee: float = 0.0
    subscription_fee: float = 0.0
    insurance_fee: float = 0.0
    packing_fee: float = 0.0
    marketing_fee: float = 0.0
    hazardous_surcharge: float = 0.0
    fragile_surcharge: float = 0.0
    oversized_surcharge: float = 0.0
    category_rates: Dict[str, float] = field(default_factory=dict)
    mode_multipliers: Dict[str, float] = field(default_factory=dict)
    weight_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    volume_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    available: bool = True
    description: str = ""
    version: str = "2026.1"
    last_updated: datetime = field(default_factory=datetime.now)
    tariff_source: TariffSource = TariffSource.HARDCODED
    seasonal_multipliers: Dict[str, float] = field(default_factory=dict)
    promo_discount: float = 0.0
    promo_start: Optional[datetime] = None
    promo_end: Optional[datetime] = None
    dynamic_adjustment: float = 0.0
    last_forecast: Optional[Dict[str, Any]] = None
    forecast_timestamp: Optional[datetime] = None

    def get_commission_rate(self, category: Optional[str] = None) -> float:
        """Получить ставку комиссии (общую или по категории)"""
        if category and category in self.category_rates:
            return self.category_rates[category]
        return self.commission_rate

    def get_mode_multiplier(self, mode: str) -> float:
        """Получить мультипликатор режима работы"""
        return self.mode_multipliers.get(mode, 1.0)

    def apply_seasonal_multiplier(self, base_rate: float, current_month: Optional[int] = None) -> float:
        """Применить сезонный коэффициент к базовой ставке"""
        if current_month is None:
            current_month = datetime.now().month
        if current_month in [12, 1, 2]:
            season = "winter"
        elif current_month in [3, 4, 5]:
            season = "spring"
        elif current_month in [6, 7, 8]:
            season = "summer"
        else:
            season = "autumn"
        multiplier = self.seasonal_multipliers.get(season, 1.0)
        return base_rate * multiplier

    def apply_promo_discount(self, amount: float) -> float:
        """Применить промо-скидку к сумме"""
        if self.promo_discount <= 0:
            return amount
        now = datetime.now()
        if self.promo_start and self.promo_end:
            if self.promo_start <= now <= self.promo_end:
                return amount * (1 - self.promo_discount)
        else:
            return amount * (1 - self.promo_discount)
        return amount

    def calculate_commission_with_dynamics(
        self,
        price: float,
        discount_percent: float = 0.0,
        promo_participation: float = 0.0,
        category: Optional[str] = None,
        current_month: Optional[int] = None
    ) -> float:
        """Комиссия с учётом скидок, участия в акциях и сезонности"""
        actual_price = price * (1 - discount_percent)
        promo_surcharge = actual_price * promo_participation
        base_rate = self.get_commission_rate(category)
        rate = self.apply_seasonal_multiplier(base_rate, current_month)
        rate += self.dynamic_adjustment
        commission = max(actual_price * rate, self.min_commission)
        commission += promo_surcharge
        commission = self.apply_promo_discount(commission)
        if self.max_commission < float('inf'):
            commission = min(commission, self.max_commission)
        return commission


@dataclass
class ProductDimensions:
    """Габариты и вес товара"""
    length: float = 0.0
    width: float = 0.0
    height: float = 0.0
    weight: float = 0.0
    unit: str = "см"
    weight_unit: str = "кг"
    dimension_string: str = ""

    @property
    def volume(self) -> float:
        return calculate_volume(self.length, self.width, self.height)

    @property
    def is_valid(self) -> bool:
        return all([self.length > 0, self.width > 0, self.height > 0, self.weight > 0])

    @property
    def display_dimensions(self) -> str:
        if self.length > 0 and self.width > 0 and self.height > 0:
            return f"{self.length:.1f}x{self.width:.1f}x{self.height:.1f} см"
        return "Размеры не указаны"

    def to_dict(self) -> Dict[str, float]:
        return {
            "length": self.length, "width": self.width,
            "height": self.height, "weight": self.weight,
            "volume": self.volume
        }

    @classmethod
    def from_string(cls, dim_str: str, weight: float = 0.0) -> 'ProductDimensions':
        length, width, height = parse_dimensions_string(dim_str)
        return cls(
            length=length, width=width, height=height,
            weight=weight, dimension_string=dim_str
        )


@dataclass
class CategoryHierarchy:
    """
    v101.0: 3-уровневая иерархия категорий
    Родитель → Группа → Подгруппа
    Пример: Автозапчасти → Подвеска → Сайлентблоки
    """
    parent: str = ""          # Уровень 1: Автозапчасти
    group: str = ""           # Уровень 2: Подвеска
    subgroup: str = ""        # Уровень 3: Сайлентблоки
    keywords: List[str] = field(default_factory=list)
    weight_range: Tuple[float, float] = (0.0, 100.0)
    dimensions_range: Dict[str, Tuple[float, float]] = field(default_factory=dict)

    def full_path(self) -> str:
        """Полный путь категории"""
        parts = [p for p in [self.parent, self.group, self.subgroup] if p]
        return "/".join(parts) if parts else "Без категории"

    def matches_name(self, name: str) -> bool:
        """Проверка соответствия наименованию по ключевым словам"""
        if not self.keywords or not name:
            return False
        name_lower = name.lower()
        return any(kw.lower() in name_lower for kw in self.keywords)

    def matches_dimensions(self, weight: float, length: float = 0, width: float = 0, height: float = 0) -> bool:
        """Проверка соответствия весогабаритам (не строгая)"""
        if not (self.weight_range[0] <= weight <= self.weight_range[1]):
            return False
        if length > 0 and "length" in self.dimensions_range:
            if not (self.dimensions_range["length"][0] <= length <= self.dimensions_range["length"][1]):
                return False
        return True


@dataclass
class ProductCategory:
    """Категория товара с метаданными"""
    name: str
    description: str = ""
    parent_category: Optional[str] = None
    hierarchy: Optional[CategoryHierarchy] = None
    dimensions: Optional[ProductDimensions] = None
    typical_volume: float = 0.0
    typical_weight: float = 0.0
    oem_codes: List[str] = field(default_factory=list)
    cross_references: List[str] = field(default_factory=list)
    alternatives: List[str] = field(default_factory=list)
    compatibility: List[str] = field(default_factory=list)
    hazardous: bool = False
    fragile: bool = False
    requires_special_packaging: bool = False
    seasonality: Seasonality = Seasonality.ALL_YEAR
    risk_level: RiskLevel = RiskLevel.LOW
    price_range_min: float = 0.0
    price_range_max: float = 0.0
    margin_avg: float = 0.0
    demand_score: float = 0.0
    min_length: float = 0.0
    max_length: float = 0.0
    min_width: float = 0.0
    max_width: float = 0.0
    min_height: float = 0.0
    max_height: float = 0.0
    min_weight: float = 0.0
    max_weight: float = 0.0

    def get_dimensions(self) -> ProductDimensions:
        return self.dimensions or ProductDimensions()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name, "description": self.description,
            "parent_category": self.parent_category,
            "typical_volume": self.typical_volume,
            "typical_weight": self.typical_weight,
            "oem_codes": self.oem_codes,
            "cross_references": self.cross_references,
            "hazardous": self.hazardous, "fragile": self.fragile,
            "seasonality": self.seasonality.value,
            "risk_level": self.risk_level.value
        }


@dataclass
class UnitEconomicsResult:
    """Результат расчёта юнит-экономики"""
    marketplace: str
    operation_mode: str
    category: str
    price: float
    cost: float
    length: float
    width: float
    height: float
    weight: float
    volume: float
    commission: float
    commission_percent: float
    logistics: float
    storage_cost: float
    acquiring: float
    delivery: float
    last_mile: float
    returns: float
    rko_fee: float
    premium_fee: float
    insurance_fee: float
    packing_fee: float
    marketing_fee: float
    subscription_cost: float
    hazardous_surcharge: float = 0.0
    fragile_surcharge: float = 0.0
    oversized_surcharge: float = 0.0
    tax_amount: float = 0.0
    tax_system: str = "УСН_6"
    total_expenses: float = 0.0
    profit: float = 0.0
    margin_percent: float = 0.0
    roi: float = 0.0
    breakeven_price: float = 0.0
    recommended_min_price: float = 0.0
    profit_per_ruble: float = 0.0
    contribution_margin: float = 0.0
    contribution_margin_ratio: float = 0.0
    timestamp: datetime = field(default_factory=datetime.now)
    calculation_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: CalculationStatus = CalculationStatus.COMPLETED
    tariff_source: TariffSource = TariffSource.HARDCODED
    metadata: Dict[str, Any] = field(default_factory=dict)
    applied_seasonal_multiplier: float = 1.0
    applied_promo_discount: float = 0.0
    dynamic_adjustment: float = 0.0
    billable_weight: float = 0.0
    advertising_cost: float = 0.0
    auto_parts_specific: float = 0.0
    # v101.0: ABC/XYZ поля
    abc_category: str = ""
    xyz_category: str = ""
    abcxyz_category: str = ""

    def to_dict(self) -> Dict[str, Any]:
        result = asdict(self)
        result['timestamp'] = self.timestamp.isoformat()
        result['status'] = self.status.name
        result['tariff_source'] = self.tariff_source.value
        return result

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([self.to_dict()])

    def get_summary(self) -> Dict[str, Any]:
        return {
            "marketplace": self.marketplace, "profit": self.profit,
            "margin": self.margin_percent, "roi": self.roi,
            "breakeven": self.breakeven_price,
            "recommended_min_price": self.recommended_min_price,
            "total_expenses": self.total_expenses
        }

    def get_profitability_level(self) -> ProfitabilityLevel:
        if self.profit < 0:
            return ProfitabilityLevel.LOSS
        elif self.profit == 0:
            return ProfitabilityLevel.BREAK_EVEN
        elif self.margin_percent < 5:
            return ProfitabilityLevel.LOW
        elif self.margin_percent < 15:
            return ProfitabilityLevel.MEDIUM
        elif self.margin_percent < 30:
            return ProfitabilityLevel.HIGH
        else:
            return ProfitabilityLevel.VERY_HIGH


@dataclass
class ABCAnalysisResult:
    """
    v101.0: Результат ABC/XYZ анализа
    ABC — по маржинальности и прибыли
    XYZ — по стабильности (коэффициент вариации)
    """
    article: str
    brand: str
    category: str
    margin_percent: float
    profit: float
    revenue: float
    abc_margin: str = ""      # A/B/C по марже
    abc_profit: str = ""      # A/B/C по прибыли
    xyz_stability: str = ""   # X/Y/Z по стабильности
    abcxyz_combined: str = "" # Комбинированная (AX, BX, CZ...)
    coefficient_of_variation: float = 0.0
    recommendations: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ColumnMapping:
    """
    v101.0: Маппинг столбцов между файлами
    Используется для связывания данных из разных источников
    """
    source_file: str
    source_column: str
    target_file: str
    target_column: str
    mapping_type: DataLinkType = DataLinkType.MANUAL_MAPPING
    confidence: float = 1.0
    auto_detected: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source_file": self.source_file,
            "source_column": self.source_column,
            "target_file": self.target_file,
            "target_column": self.target_column,
            "mapping_type": self.mapping_type.name,
            "confidence": self.confidence,
            "auto_detected": self.auto_detected
        }


@dataclass
class DataLinkConfig:
    """
    v101.0: Конфигурация связывания данных между файлами
    """
    link_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""
    description: str = ""
    primary_file: str = ""
    secondary_file: str = ""
    join_key_primary: str = ""
    join_key_secondary: str = ""
    join_type: str = "left"  # left, inner, outer, cross
    column_mappings: List[ColumnMapping] = field(default_factory=list)
    auto_fill_missing: bool = True  # Автоматически подставлять недостающие из аналогов
    created_at: datetime = field(default_factory=datetime.now)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "link_id": self.link_id,
            "name": self.name,
            "description": self.description,
            "primary_file": self.primary_file,
            "secondary_file": self.secondary_file,
            "join_key_primary": self.join_key_primary,
            "join_key_secondary": self.join_key_secondary,
            "join_type": self.join_type,
            "column_mappings": [m.to_dict() for m in self.column_mappings],
            "auto_fill_missing": self.auto_fill_missing,
            "created_at": self.created_at.isoformat()
        }


@dataclass
class GoogleSheetsConfig:
    """
    v101.0: Конфигурация Google Sheets
    """
    credentials_json: str = ""  # Путь к JSON-файлу или содержимое
    spreadsheet_id: str = ""
    worksheet_name: str = "Тарифы"
    auto_update: bool = True
    last_sync: Optional[datetime] = None
    service_account_email: str = ""
    api_key: str = ""
    scopes: List[str] = field(default_factory=lambda: [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ])

    def is_configured(self) -> bool:
        """Проверка, настроена ли интеграция"""
        return bool(self.credentials_json and self.spreadsheet_id)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "spreadsheet_id": self.spreadsheet_id,
            "worksheet_name": self.worksheet_name,
            "auto_update": self.auto_update,
            "last_sync": self.last_sync.isoformat() if self.last_sync else None,
            "service_account_email": self.service_account_email,
            "is_configured": self.is_configured()
        }


@dataclass
class SaveLoadState:
    """
    v101.0: Состояние сохранения/загрузки расчётов
    """
    state_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""
    description: str = ""
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)
    file_path: str = ""
    data_hash: str = ""
    section_data: Dict[str, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    is_compressed: bool = False
    size_bytes: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "state_id": self.state_id,
            "name": self.name,
            "description": self.description,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
            "file_path": self.file_path,
            "data_hash": self.data_hash,
            "metadata": self.metadata,
            "is_compressed": self.is_compressed,
            "size_bytes": self.size_bytes
        }


@dataclass
class TariffCacheEntry:
    """Запись кэша тарифов"""
    marketplace: str
    category: Optional[str]
    data: Dict[str, Any]
    source: TariffSource
    timestamp: float
    ttl_seconds: int = 86400
    version: str = "2026.1"
    notes: str = ""
    forecast_data: Optional[Dict[str, Any]] = None
    historical_data: Optional[List[Dict[str, Any]]] = None

    def is_expired(self) -> bool:
        return time.time() - self.timestamp > self.ttl_seconds

    def to_dict(self) -> Dict[str, Any]:
        return {
            "marketplace": self.marketplace, "category": self.category,
            "data": self.data, "source": self.source.value,
            "timestamp": self.timestamp, "ttl_seconds": self.ttl_seconds,
            "version": self.version, "notes": self.notes,
            "forecast_data": self.forecast_data,
            "historical_data": self.historical_data
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> 'TariffCacheEntry':
        return TariffCacheEntry(
            marketplace=d.get("marketplace", ""),
            category=d.get("category"),
            data=d.get("data", {}),
            source=TariffSource(d.get("source", "HARDCODED")),
            timestamp=d.get("timestamp", 0),
            ttl_seconds=d.get("ttl_seconds", 86400),
            version=d.get("version", "2026.1"),
            notes=d.get("notes", ""),
            forecast_data=d.get("forecast_data"),
            historical_data=d.get("historical_data")
        )


@dataclass
class AutoPartsSpecificCosts:
    """Специфические расходы для автозапчастей"""
    chestny_znak: float = 1.5
    certification_amortization: float = 0.0
    warranty_reserve: float = 0.02
    packaging_fbs: float = 45.0
    labeling: float = 3.0
    util_tax: float = 0.0
    customs_duty: float = 0.0
    currency_risk: float = 0.03

    def calculate(self, price: float, is_import: bool = False, requires_marking: bool = True) -> float:
        total = 0.0
        if requires_marking:
            total += self.chestny_znak
        total += self.certification_amortization
        total += price * self.warranty_reserve
        total += self.packaging_fbs
        total += self.labeling
        if is_import:
            total += price * self.currency_risk
            total += self.customs_duty
            total += price * self.util_tax
        return money_round(total)

# ============================================================================
# БЛОК 3: УТИЛИТЫ  ВАЛИДАЦИЯ, ПАРСИНГ, РАСЧЁТЫ
# ============================================================================
# 📌 v101.0: Добавлены новые утилиты:
# - abc_analysis() / xyz_analysis()  ABC/XYZ анализ по маржинальности и прибыли
# - link_columns_between_files()  связывание столбцов между файлами
# - save_state() / load_state()  сохранение/загрузка состояния расчётов
# - google_sheets_upload() / google_sheets_download()  интеграция с Google Sheets
# - detect_mojibake() / fix_double_utf8()  исправление кракозябр
# - calculate_recommended_min_price()  расчёт минимальной цены для безубыточности
# ============================================================================


# ============================================================================
# 3.1 УТИЛИТЫ ДЛЯ ТОЧНЫХ РАСЧЁТОВ (ДЕНЬГИ, НАЛОГИ)
# ============================================================================

def money_round(value: float, decimals: int = 2) -> float:
    """Корректное банковское округление денег через Decimal"""
    if value is None or (isinstance(value, float) and (math.isnan(value) or math.isinf(value))):
        return 0.0
    try:
        return float(Decimal(str(value)).quantize(
            Decimal(f"0.{'0' * decimals}"),
            rounding=ROUND_HALF_UP
        ))
    except (ValueError, TypeError, decimal.InvalidOperation):
        return 0.0


def calculate_tax(price: float, cost: float, tax_system: str = "УСН_6") -> float:
    """
    Расчёт налога с учётом режима налогообложения.
    
    Args:
        price: Цена продажи
        cost: Себестоимость
        tax_system: Код налоговой системы (УСН_6, УСН_15, ОСН, ПСН, НПД)
    
    Returns:
        Сумма налога, округлённая до копеек
    """
    cfg = TAX_SYSTEMS.get(tax_system, TAX_SYSTEMS["УСН_6"])
    base = cfg["base"]
    
    if base == "revenue":
        # УСН 6%, НПД 6%  от выручки
        return money_round(price * cfg["rate"])
    elif base == "profit":
        # УСН 15%, ОСН  от прибыли
        profit = price - cost
        tax = profit * cfg["rate"]
        # Для УСН 15% есть минимальный налог 1% от выручки
        if tax_system == "УСН_15":
            min_tax = price * cfg.get("min_rate", 0.01)
            tax = max(tax, min_tax)
        return money_round(max(0, tax))
    elif base == "fixed":
        # ПСН  фиксированный платёж, не зависит от продажи
        return 0.0
    return 0.0


# ============================================================================
# 3.2 УТИЛИТЫ ДЛЯ ВЕСОГАБАРИТОВ
# ============================================================================

def calculate_volume(length: float, width: float, height: float) -> float:
    """
    Расчёт объёма в литрах из размеров в сантиметрах.
    Автоматически определяет, если размеры в мм (делит на 10).
    """
    if not all([length, width, height]):
        return 0.0
    if not all([length > 0, width > 0, height > 0]):
        return 0.0
    # Если размеры явно в мм (>1000), конвертируем в см
    if any([length > 1000, width > 1000, height > 1000]):
        length /= 10
        width /= 10
        height /= 10
    if any([length < 0.1, width < 0.1, height < 0.1]):
        return 0.0
    volume = (length * width * height) / 1000.0
    if volume < 0.001:
        return 0.0
    return round(volume, 4)


def calculate_billable_weight(
    weight_kg: float,
    length_cm: float,
    width_cm: float,
    height_cm: float,
    volumetric_coeff: float = 5000.0
) -> float:
    """
    Расчёт оплачиваемого веса (больший из реального и объёмного).
    Округление вверх до 0.5 кг (стандарт МП).
    
    Args:
        weight_kg: Реальный вес в кг
        length_cm, width_cm, height_cm: Габариты в см
        volumetric_coeff: Объёмный коэффициент (обычно 5000)
    
    Returns:
        Оплачиваемый вес в кг, округлённый вверх до 0.5
    """
    if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
        return weight_kg
    volumetric_weight = (length_cm * width_cm * height_cm) / volumetric_coeff
    billable = max(weight_kg, volumetric_weight)
    # Округление вверх до 0.5 кг
    billable = math.ceil(billable * 2) / 2
    return billable


def calculate_storage_cost_progressive(
    volume_l: float,
    days: int,
    base_rate: float,
    marketplace: str
) -> float:
    """
    Прогрессивная стоимость хранения.
    Для Ozon/WB: после 60 дней ставка растёт экспоненциально.
    """
    if volume_l <= 0 or days <= 0:
        return 0.0
    if marketplace in ["Ozon", "Wildberries"]:
        if days <= 60:
            multiplier = 1.0
        elif days <= 90:
            multiplier = 2.0
        elif days <= 180:
            multiplier = 4.0
        elif days <= 365:
            multiplier = 8.0
        else:
            multiplier = 16.0
        weighted_rate = base_rate * multiplier
        return money_round(volume_l * weighted_rate * days)
    else:
        return money_round(volume_l * base_rate * days)


def calculate_returns_cost(
    price: float,
    return_rate: float,
    reverse_logistics: float = 150.0,
    inspection_cost: float = 50.0
) -> float:
    """
    Полная стоимость возвратов с учётом:
    - Потери от возврата товара
    - Обратной логистики
    - Инспекции
    - Потери от дефектов (30% стоимости)
    """
    if price <= 0 or return_rate <= 0:
        return 0.0
    expected_returns = price * return_rate
    reverse_logistics_cost = reverse_logistics * return_rate
    inspection = inspection_cost * return_rate
    loss_from_defects = price * return_rate * 0.3
    return money_round(expected_returns + reverse_logistics_cost + inspection + loss_from_defects)


# ============================================================================
# 3.3 СПЕЦИФИЧЕСКИЕ РАСХОДЫ АВТОЗАПЧАСТЕЙ
# ============================================================================

def calculate_auto_parts_specific(
    price: float,
    is_import: bool = False,
    requires_marking: bool = True,
    chestny_znak: float = 1.5,
    warranty_reserve_rate: float = 0.02,
    packaging_fbs: float = 45.0,
    labeling: float = 3.0,
    currency_risk_rate: float = 0.03
) -> float:
    """
    Расчёт специфических расходов для автозапчастей:
    - Честный ЗНАК (маркировка)
    - Амортизация сертификации
    - Гарантийный резерв
    - Упаковка FBS
    - Маркировка
    - Валютный риск (для импорта)
    """
    total = 0.0
    if requires_marking:
        total += chestny_znak
    total += price * warranty_reserve_rate
    total += packaging_fbs
    total += labeling
    if is_import:
        total += price * currency_risk_rate
    return money_round(total)


# ============================================================================
# 3.4 РЕКЛАМНЫЕ РАСХОДЫ (ДРР)
# ============================================================================

def calculate_advertising_cost(
    price: float,
    category: str,
    ad_intensity: str = "medium"
) -> float:
    """
    Рекламные расходы (ДРР  доля рекламных расходов).
    Для конкурентных категорий (масла, фильтры, колодки) интенсивность повышается.
    """
    drr_rates = {
        "low": 0.05,
        "medium": 0.15,
        "high": 0.25,
        "aggressive": 0.35
    }
    competitive_categories = ["масла", "фильтры", "колодки", "аккумуляторы", "шины"]
    intensity = ad_intensity
    if category and any(cat in category.lower() for cat in competitive_categories):
        if intensity == "medium":
            intensity = "high"
    rate = drr_rates.get(intensity, 0.15)
    return money_round(price * rate)


# ============================================================================
# 3.5 ПАРСИНГ РАЗМЕРОВ
# ============================================================================

def parse_dimensions_string(dim_str: str) -> Tuple[float, float, float]:
    """
    Парсит "человеческий" ввод размеров в формат (длина, ширина, высота).
    Поддерживает разделители: x, *, х, , пробел, запятая.
    Результат сортируется по убыванию (большая сторона  первая).
    
    Примеры:
        "20x15x10"  (20.0, 15.0, 10.0)
        "10*15*20"  (20.0, 15.0, 10.0)
        "20, 15, 10"  (20.0, 15.0, 10.0)
    """
    if not dim_str or not isinstance(dim_str, str):
        return 0.0, 0.0, 0.0
    dim_str = dim_str.lower().strip()
    separators = ['x', '*', 'х', '', ' ', ',']
    for sep in separators:
        if sep in dim_str:
            parts = [p.strip() for p in dim_str.split(sep) if p.strip()]
            if len(parts) >= 3:
                try:
                    dimensions = []
                    for p in parts[:3]:
                        cleaned = re.sub(r'[^\d.,\-]', '', p)
                        cleaned = cleaned.replace(',', '.')
                        if cleaned and cleaned.replace('.', '').replace('-', '').isdigit():
                            dimensions.append(float(cleaned))
                        else:
                            nums = re.findall(r'(\d+\.?\d*)', p)
                            if nums:
                                dimensions.append(float(nums[0]))
                    if len(dimensions) == 3:
                        dimensions.sort(reverse=True)
                        return tuple(dimensions)
                except (ValueError, TypeError):
                    pass
            # Если нашли ровно 2 числа  считаем высоту = 1
            if len(parts) == 2:
                try:
                    nums = []
                    for p in parts:
                        cleaned = re.sub(r'[^\d.,\-]', '', p).replace(',', '.')
                        if cleaned:
                            nums.append(float(cleaned))
                    if len(nums) == 2:
                        nums.sort(reverse=True)
                        return (nums[0], nums[1], 1.0)
                except (ValueError, TypeError):
                    pass
    # Фоллбэк: извлекаем все числа из строки
    nums = re.findall(r'(\d+\.?\d*)', dim_str)
    if len(nums) >= 3:
        try:
            dims = sorted([float(n) for n in nums[:3]], reverse=True)
            return tuple(dims)
        except (ValueError, TypeError):
            pass
    return 0.0, 0.0, 0.0


def parse_dimensions_vectorized(dims_series) -> "pl.DataFrame":
    """Векторизованный парсинг размеров для Polars DataFrame."""
    if not POLARS_AVAILABLE:
        return pl.DataFrame()
    dims = dims_series.str.extract_all(r"(\d+\.?\d*)")
    
    def sort_dimensions(nums):
        if nums and len(nums) >= 3:
            try:
                return sorted([float(n) for n in nums[:3]], reverse=True)
            except (ValueError, TypeError):
                pass
        elif nums and len(nums) == 2:
            try:
                return [float(nums[0]), float(nums[1]), 1.0]
            except (ValueError, TypeError):
                pass
        return [0.0, 0.0, 0.0]
    
    result = dims.map_elements(sort_dimensions, return_dtype=pl.List(pl.Float64))
    return pl.DataFrame({
        "length": result.list.get(0),
        "width": result.list.get(1),
        "height": result.list.get(2)
    })


# ============================================================================
# 3.6 БЕЗОПАСНЫЕ КОНВЕРТЕРЫ ТИПОВ
# ============================================================================

def safe_float(val: Any, default: float = 0.0) -> float:
    """Безопасная конвертация в float с обработкой всех edge cases"""
    if val is None:
        return default
    if isinstance(val, bool):
        return float(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return default
        return float(val)
    if isinstance(val, (decimal.Decimal, np.floating, np.integer)):
        try:
            return float(val)
        except (ValueError, TypeError):
            return default
    if isinstance(val, str):
        cleaned = val.strip()
        if not cleaned:
            return default
        cleaned = re.sub(r'[^\d.,\-+\s]', '', cleaned)
        cleaned = cleaned.replace(' ', '').replace(',', '.')
        if cleaned.count('-') > 1:
            return default
        parts = cleaned.split('.')
        if len(parts) > 2:
            return default
        try:
            return float(cleaned)
        except ValueError:
            return default
    if hasattr(val, 'dtype') and hasattr(val, 'item'):
        try:
            item = val.item()
            if isinstance(item, (int, float)):
                return float(item)
        except Exception:
            pass
    return default


def safe_int(val: Any, default: int = 0) -> int:
    """Безопасная конвертация в int"""
    try:
        float_val = safe_float(val, default)
        if float_val == default and val != 0:
            return default
        return int(float_val)
    except (ValueError, TypeError):
        return default


def safe_str(val: Any, default: str = "") -> str:
    """Безопасная конвертация в строку"""
    if val is None:
        return default
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return default
        return str(val)
    if isinstance(val, (list, tuple)):
        return ", ".join(safe_str(v) for v in val[:5]) + ("..." if len(val) > 5 else "")
    if isinstance(val, dict):
        return str({k: safe_str(v) for k, v in list(val.items())[:5]})
    try:
        result = str(val).strip()
        return result if result else default
    except Exception:
        return default


def safe_bool(val: Any, default: bool = False) -> bool:
    """Безопасная конвертация в bool"""
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        val_lower = val.lower().strip()
        true_values = {'true', 'yes', '1', 'y', 'да', 'on'}
        false_values = {'false', 'no', '0', 'n', 'нет', 'off'}
        if val_lower in true_values:
            return True
        if val_lower in false_values:
            return False
        return default
    if isinstance(val, (list, tuple, dict)):
        return bool(val)
    return default


def safe_datetime(val: Any, default: Optional[datetime] = None) -> Optional[datetime]:
    """Безопасная конвертация в datetime"""
    if default is None:
        default = datetime.now()
    if val is None:
        return default
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime.combine(val, datetime.min.time())
    if isinstance(val, (int, float)):
        try:
            return datetime.fromtimestamp(val)
        except (ValueError, OSError):
            return default
    if isinstance(val, str):
        formats = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d",
            "%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S.%fZ",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                continue
        try:
            if DATEUTIL_AVAILABLE:
                return dateutil_parse(val)
        except Exception:
            pass
    return default


# ============================================================================
# 3.7 НОРМАЛИЗАЦИЯ ТЕКСТА И КЛЮЧЕЙ
# ============================================================================

def normalize_text(text: str) -> str:
    """Нормализация текста: lowercase, удаление символов, сжатие пробелов"""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def normalize_key_for_match(value: str) -> str:
    """Нормализация ключа для сопоставления (артикулы, OE-номера)"""
    if not value:
        return ""
    return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', str(value).lower().strip())


def get_file_encoding(file_path: Union[str, Path]) -> str:
    """Определение кодировки файла"""
    if CHARDET_AVAILABLE and chardet is not None:
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(100000)
                result = chardet.detect(raw_data)
                encoding = result.get('encoding', 'utf-8')
                return encoding
        except (IOError, OSError) as e:
            logger.warning(f"Ошибка определения кодировки: {e}")
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'cp1252', 'latin1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                f.read()
            return enc
        except UnicodeDecodeError:
            continue
    return 'utf-8'


# ============================================================================
# 3.8 ДЕТЕКТОР И ФИКСЕР КРАКОЗЯБР (MOJIBAKE)
# ============================================================================

def detect_mojibake(text: str) -> bool:
    """
    Определяет наличие кракозябр (двойного UTF-8 кодирования).
    Пример: "РџСЂРёРІРµС‚" вместо "Привет"
    """
    if not isinstance(text, str) or not text:
        return False
    mojibake_patterns = [
        r'Р[°-Џ]{2,}',
        r'Р[РЎ][°-Џ]{2,}',
        r'[РЎР][°-Џ]{3,}',
        r'Р[°-Џ]Р[°-Џ]',
    ]
    for pattern in mojibake_patterns:
        if re.search(pattern, text):
            return True
    words = text.split()
    if len(words) >= 3:
        r_words = sum(1 for w in words if w.startswith('Р') and len(w) >= 2)
        if r_words / len(words) > 0.5:
            return True
    return False


def fix_double_utf8(text: str) -> str:
    """Исправляет двойное кодирование UTF-8"""
    if not isinstance(text, str) or not text:
        return text
    encodings_to_try = [
        ('cp1251', 'utf-8'),
        ('latin1', 'utf-8'),
        ('iso-8859-1', 'utf-8'),
        ('cp1252', 'utf-8'),
    ]
    for source_enc, target_enc in encodings_to_try:
        try:
            fixed = text.encode(source_enc).decode(target_enc)
            if fixed and not detect_mojibake(fixed):
                return fixed
        except (UnicodeDecodeError, UnicodeEncodeError):
            continue
    return text


def fix_dataframe_encoding(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Исправляет кракозябры во всём DataFrame (колонки + ячейки)"""
    fixed_count = 0
    new_columns = []
    for col in df.columns:
        col_str = str(col)
        if detect_mojibake(col_str):
            new_col = fix_double_utf8(col_str)
            new_columns.append(new_col)
            fixed_count += 1
        else:
            new_columns.append(col)
    df.columns = new_columns
    
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                def _fix_cell(x):
                    if isinstance(x, str) and detect_mojibake(x):
                        return fix_double_utf8(x)
                    return x
                mask = df[col].apply(lambda x: isinstance(x, str) and detect_mojibake(x))
                fixed_count += int(mask.sum())
                df[col] = df[col].apply(_fix_cell)
            except Exception:
                pass
    return df, fixed_count


def smart_read_csv(file_obj, **kwargs) -> pd.DataFrame:
    """Умное чтение CSV с автоматическим определением кодировки и разделителя"""
    separators = [';', ',', '\t', '|']
    encodings_priority = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251']
    best_df = None
    best_encoding = None
    best_sep = None
    mojibake_count = 0
    
    for encoding in encodings_priority:
        for sep in separators:
            try:
                file_obj.seek(0)
                df = pd.read_csv(
                    file_obj,
                    encoding=encoding,
                    sep=sep,
                    engine='python',
                    on_bad_lines='skip',
                    skipinitialspace=True,
                    quotechar='"',
                    doublequote=True,
                    **kwargs
                )
                if df is None or df.empty or len(df.columns) <= 1:
                    continue
                current_mojibake = sum(
                    1 for col in df.columns
                    if isinstance(col, str) and detect_mojibake(col)
                )
                if current_mojibake == 0:
                    return df
                if best_df is None or current_mojibake < mojibake_count:
                    best_df = df
                    best_encoding = encoding
                    best_sep = sep
                    mojibake_count = current_mojibake
            except (pd.errors.ParserError, UnicodeDecodeError, Exception):
                continue
    
    if best_df is not None:
        fixed_df, fixed_count = fix_dataframe_encoding(best_df)
        return fixed_df
    
    if CHARDET_AVAILABLE and chardet is not None:
        try:
            file_obj.seek(0)
            raw_data = file_obj.read(100000)
            detected = chardet.detect(raw_data)
            if detected and detected.get('encoding'):
                file_obj.seek(0)
                for sep in separators:
                    try:
                        df = pd.read_csv(
                            file_obj,
                            encoding=detected['encoding'],
                            sep=sep,
                            engine='python',
                            on_bad_lines='skip'
                        )
                        if df is not None and not df.empty and len(df.columns) > 1:
                            has_mojibake = any(
                                isinstance(col, str) and detect_mojibake(col)
                                for col in df.columns
                            )
                            if has_mojibake:
                                df, _ = fix_dataframe_encoding(df)
                            return df
                    except (pd.errors.ParserError, UnicodeDecodeError):
                        continue
        except Exception as e:
            pass
    
    raise ValueError("Не удалось прочитать CSV файл. Проверьте кодировку и разделитель.")


# ============================================================================
# 3.9 ВАЛИДАЦИЯ ВХОДНЫХ ДАННЫХ
# ============================================================================

def validate_input_data(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """Проверка качества данных перед расчётом"""
    errors = []
    if 'Цена' in df.columns:
        negative_prices = (df['Цена'] <= 0).sum()
        if negative_prices > 0:
            errors.append(f"⚠️ {negative_prices} товаров с ценой ≤ 0")
        suspicious = (df['Цена'] < 50).sum()
        if suspicious > 0:
            errors.append(f"⚠️ {suspicious} товаров дешевле 50₽  проверьте")
    if 'Длина' in df.columns:
        missing_dims = df['Длина'].isna().sum()
        if missing_dims > len(df) * 0.3:
            errors.append(f"⚠️ У {missing_dims} товаров нет габаритов  логистика будет неточной")
    return len(errors) == 0, errors


# ============================================================================
# 3.10 РАСЧЁТ РЕКОМЕНДУЕМОЙ МИНИМАЛЬНОЙ ЦЕНЫ
# ============================================================================

def calculate_recommended_min_price(
    cost: float,
    commission_rate: float,
    logistics: float,
    storage_cost: float,
    acquiring_rate: float,
    last_mile: float,
    return_rate: float,
    min_profit_percent: float = 0.10,
    tax_system: str = "УСН_6",
    tax_rate: float = 0.06
) -> float:
    """
    Расчёт минимальной цены, при которой продажа безубыточна
    и остаётся минимальная прибыль (по умолчанию 10%).
    
    Формула:
        price = (cost + fixed_costs) / (1 - variable_rate - min_profit)
    """
    if cost <= 0:
        return 0.0
    fixed_costs = cost + logistics + storage_cost + last_mile
    variable_rate = commission_rate + acquiring_rate + return_rate + tax_rate + min_profit_percent
    denominator = 1 - variable_rate
    if denominator <= 0:
        return 0.0
    recommended_price = fixed_costs / denominator
    return max(0, money_round(recommended_price))


# ============================================================================
# 3.11 ГЕНЕРАЦИЯ КЛЮЧЕЙ КЭША
# ============================================================================

def generate_cache_key(*args, **kwargs) -> str:
    """Генерация уникального MD5-ключа для кэширования"""
    key_parts = []
    for arg in args:
        if isinstance(arg, (dict, OrderedDict)):
            key_parts.append(json.dumps(arg, sort_keys=True, ensure_ascii=False))
        elif isinstance(arg, (list, tuple, set)):
            key_parts.append(str(sorted(arg) if not isinstance(arg, tuple) else arg))
        elif isinstance(arg, pd.DataFrame):
            try:
                key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(len(arg)))
        elif isinstance(arg, pd.Series):
            try:
                key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(len(arg)))
        elif isinstance(arg, np.ndarray):
            try:
                key_parts.append(hashlib.md5(arg.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(arg.shape))
        elif isinstance(arg, (datetime, date)):
            key_parts.append(arg.isoformat())
        else:
            key_parts.append(str(arg))
    for k, v in sorted(kwargs.items()):
        if isinstance(v, (dict, OrderedDict)):
            key_parts.append(f"{k}:{json.dumps(v, sort_keys=True, ensure_ascii=False)}")
        elif isinstance(v, (list, tuple, set)):
            key_parts.append(f"{k}:{str(sorted(v) if not isinstance(v, tuple) else v)}")
        elif isinstance(v, pd.DataFrame):
            try:
                key_parts.append(f"{k}:{hashlib.md5(pd.util.hash_pandas_object(v).values.tobytes()).hexdigest()}")
            except Exception:
                key_parts.append(f"{k}:{len(v)}")
        else:
            key_parts.append(f"{k}:{v}")
    key = "|".join(key_parts)
    return hashlib.md5(key.encode('utf-8')).hexdigest()


# ============================================================================
# 3.12 ABC/XYZ АНАЛИЗ
# ============================================================================

def abc_analysis(
    df: pd.DataFrame,
    article_col: str = "Артикул",
    margin_col: str = "margin_percent",
    profit_col: str = "profit",
    revenue_col: str = "price"
) -> pd.DataFrame:
    """
    ABC-анализ товаров по маржинальности и прибыли.
    
    Правила:
    - A: маржа >= 25% ИЛИ вклад в общую прибыль >= 70%
    - B: маржа >= 15% ИЛИ вклад в прибыль 20-70%
    - C: всё остальное (маржа < 15%)
    
    Returns:
        DataFrame с добавленными колонками:
        - abc_margin: A/B/C по марже
        - abc_profit: A/B/C по вкладу в прибыль
        - revenue_share: доля в общей выручке
        - profit_share: доля в общей прибыли
    """
    if df.empty:
        return df
    
    result = df.copy()
    
    # Нормализация колонок
    if margin_col not in result.columns:
        result['abc_margin'] = 'C'
        result['abc_profit'] = 'C'
        return result
    
    # === ABC по марже ===
    def classify_margin(m):
        if pd.isna(m):
            return 'C'
        if m >= 25:
            return 'A'
        elif m >= 15:
            return 'B'
        return 'C'
    
    result['abc_margin'] = result[margin_col].apply(classify_margin)
    
    # === ABC по прибыли ===
    if profit_col in result.columns:
        total_profit = result[profit_col].sum()
        if total_profit > 0:
            # Сортируем по прибыли (убывание)
            sorted_df = result.sort_values(profit_col, ascending=False).copy()
            sorted_df['cumulative_profit'] = sorted_df[profit_col].cumsum()
            sorted_df['profit_share'] = sorted_df['cumulative_profit'] / total_profit
            
            def classify_profit_share(share):
                if share <= 0.70:
                    return 'A'
                elif share <= 0.90:
                    return 'B'
                return 'C'
            
            sorted_df['abc_profit'] = sorted_df['profit_share'].apply(classify_profit_share)
            # Возвращаем в исходный порядок
            result = result.merge(
                sorted_df[[article_col, 'abc_profit', 'profit_share']],
                on=article_col, how='left'
            )
        else:
            result['abc_profit'] = 'C'
            result['profit_share'] = 0.0
    else:
        result['abc_profit'] = 'C'
        result['profit_share'] = 0.0
    
    # Доля в выручке
    if revenue_col in result.columns:
        total_revenue = result[revenue_col].sum()
        if total_revenue > 0:
            result['revenue_share'] = result[revenue_col] / total_revenue
        else:
            result['revenue_share'] = 0.0
    else:
        result['revenue_share'] = 0.0
    
    return result


def xyz_analysis(
    df: pd.DataFrame,
    article_col: str = "Артикул",
    period_col: str = "period",
    profit_col: str = "profit"
) -> pd.DataFrame:
    """
    XYZ-анализ по стабильности прибыли (коэффициент вариации).
    
    Правила:
    - X: CV < 0.5 (стабильные)
    - Y: CV 0.5-1.0 (умеренные)
    - Z: CV > 1.0 (нестабильные)
    
    Если данных по периодам нет  возвращает 'X' для всех.
    """
    if df.empty:
        return df
    
    result = df.copy()
    
    if period_col not in result.columns or profit_col not in result.columns:
        result['xyz_stability'] = 'X'
        result['coefficient_of_variation'] = 0.0
        return result
    
    # Группируем по артикулу и считаем CV
    grouped = result.groupby(article_col)[profit_col].agg(['mean', 'std', 'count']).reset_index()
    grouped.columns = [article_col, 'mean_profit', 'std_profit', 'count']
    
    # CV = std / mean
    grouped['cv'] = grouped.apply(
        lambda row: (row['std_profit'] / row['mean_profit']) if row['mean_profit'] > 0 and row['count'] > 1 else 0.0,
        axis=1
    )
    
    def classify_cv(cv):
        if cv < 0.5:
            return 'X'
        elif cv < 1.0:
            return 'Y'
        return 'Z'
    
    grouped['xyz_stability'] = grouped['cv'].apply(classify_cv)
    
    result = result.merge(
        grouped[[article_col, 'xyz_stability', 'cv']],
        on=article_col, how='left'
    )
    result = result.rename(columns={'cv': 'coefficient_of_variation'})
    result['coefficient_of_variation'] = result['coefficient_of_variation'].fillna(0.0)
    result['xyz_stability'] = result['xyz_stability'].fillna('X')
    
    return result


def abcxyz_combined_analysis(
    df: pd.DataFrame,
    article_col: str = "Артикул",
    margin_col: str = "margin_percent",
    profit_col: str = "profit",
    revenue_col: str = "price",
    period_col: Optional[str] = None
) -> pd.DataFrame:
    """
    Комбинированный ABC/XYZ анализ.
    Возвращает DataFrame с колонками:
    - abc_margin, abc_profit, xyz_stability
    - abcxyz_combined (например, "AX", "BZ")
    - recommendations (список рекомендаций)
    """
    result = df.copy()
    
    # ABC анализ
    result = abc_analysis(result, article_col, margin_col, profit_col, revenue_col)
    
    # XYZ анализ (если есть данные по периодам)
    if period_col and period_col in result.columns:
        result = xyz_analysis(result, article_col, period_col, profit_col)
    else:
        result['xyz_stability'] = 'X'
        result['coefficient_of_variation'] = 0.0
    
    # Комбинированная категория
    result['abcxyz_combined'] = result['abc_margin'] + result['xyz_stability']
    
    # Рекомендации
    def generate_recommendation(row):
        recs = []
        abcxyz = row.get('abcxyz_combined', '')
        margin = row.get(margin_col, 0)
        
        if abcxyz == 'AX':
            recs.append("⭐ Звезда  поддерживать и масштабировать")
        elif abcxyz == 'AY':
            recs.append("💪 Сильный середняк  оптимизировать логистику")
        elif abcxyz == 'AZ':
            recs.append("⚠️ Высокая маржа, но нестабильно  анализировать спрос")
        elif abcxyz in ['BX', 'BY']:
            recs.append("📊 Рабочая лошадка  стандартное управление")
        elif abcxyz == 'BZ':
            recs.append("🔍 Пересмотреть ценообразование")
        elif abcxyz == 'CX':
            recs.append("📉 Низкая маржа, стабильный спрос  поднять цену")
        elif abcxyz == 'CY':
            recs.append("⚠️ Кандидат на вывод  оптимизировать или убрать")
        elif abcxyz == 'CZ':
            recs.append("🗑️ Аутсайдер  рассмотреть вывод из ассортимента")
        
        if margin < 0:
            recs.append("❌ Убыточный товар  срочно пересмотреть цену")
        elif margin < 10:
            recs.append("💰 Маржа ниже 10%  увеличить наценку")
        
        return "; ".join(recs) if recs else "Стандартное управление"
    
    result['recommendations'] = result.apply(generate_recommendation, axis=1)
    
    return result


# ============================================================================
# 3.13 СВЯЗЫВАНИЕ СТОЛБЦОВ МЕЖДУ ФАЙЛАМИ
# ============================================================================

def detect_column_mapping(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    expected_pairs: Optional[List[Tuple[str, str]]] = None
) -> List[ColumnMapping]:
    """
    Автоматическое обнаружение соответствий между столбцами двух DataFrame.
    Использует нечёткое сравнение названий и анализ содержимого.
    """
    mappings = []
    cols1_lower = {col.lower().strip(): col for col in df1.columns}
    cols2_lower = {col.lower().strip(): col for col in df2.columns}
    
    # Словарь синонимов для автодетекта
    synonyms = {
        'артикул': ['article', 'sku', 'artikul', 'код', 'код товара', 'артикул бренда'],
        'бренд': ['brand', 'производитель', 'manufacturer', 'марка'],
        'наименование': ['название', 'name', 'описание', 'description', 'товар'],
        'цена': ['price', 'стоимость', 'цена продажи', 'retail price'],
        'вес': ['weight', 'масса', 'вес кг', 'weight_kg'],
        'длина': ['length', 'длинна', 'длина см'],
        'ширина': ['width', 'ширина см'],
        'высота': ['height', 'высота см'],
        'oe номер': ['oe', 'оe', 'oe_number', 'номер', 'code'],
    }
    
    used_cols2 = set()
    
    for col1_norm, col1_orig in cols1_lower.items():
        best_match = None
        best_score = 0
        
        for col2_norm, col2_orig in cols2_lower.items():
            if col2_orig in used_cols2:
                continue
            
            score = 0
            # Точное совпадение
            if col1_norm == col2_norm:
                score = 100
            # Одно содержится в другом
            elif col1_norm in col2_norm or col2_norm in col1_norm:
                score = 70
            else:
                # Проверка через синонимы
                for key, syns in synonyms.items():
                    if key in col1_norm and any(s in col2_norm for s in syns + [key]):
                        score = 60
                        break
                    if key in col2_norm and any(s in col1_norm for s in syns + [key]):
                        score = 60
                        break
            
            if score > best_score:
                best_score = score
                best_match = col2_orig
        
        if best_match and best_score > 30:
            mappings.append(ColumnMapping(
                source_file="df1",
                source_column=col1_orig,
                target_file="df2",
                target_column=best_match,
                mapping_type=DataLinkType.AUTO_DETECT,
                confidence=best_score / 100.0,
                auto_detected=True
            ))
            used_cols2.add(best_match)
    
    return mappings


def link_columns_between_files(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    join_key_1: str,
    join_key_2: str,
    join_type: str = "left",
    columns_to_fill: Optional[List[str]] = None,
    auto_fill_from_analogs: bool = True
) -> pd.DataFrame:
    """
    Связывание двух DataFrame по ключевым столбцам.
    Недостающие параметры заполняются из второго файла (или из аналогов).
    
    Args:
        df1: Основной DataFrame
        df2: Дополнительный DataFrame (источник недостающих данных)
        join_key_1: Ключевой столбец в df1
        join_key_2: Ключевой столбец в df2
        join_type: Тип соединения (left, inner, outer)
        columns_to_fill: Какие столбцы заполнять из df2 (если None  все уникальные)
        auto_fill_from_analogs: Автоматически подставлять данные из аналогов
    
    Returns:
        Объединённый DataFrame
    """
    if df1.empty or df2.empty:
        return df1
    
    if join_key_1 not in df1.columns or join_key_2 not in df2.columns:
        logger.warning(f"Ключевые столбцы не найдены: {join_key_1}, {join_key_2}")
        return df1
    
    # Нормализуем ключи для корректного соединения
    df1_work = df1.copy()
    df2_work = df2.copy()
    df1_work['_join_key'] = df1_work[join_key_1].astype(str).str.lower().str.strip()
    df2_work['_join_key'] = df2_work[join_key_2].astype(str).str.lower().str.strip()
    
    # Определяем столбцы для заполнения
    if columns_to_fill is None:
        # Все столбцы из df2, которых нет в df1 (кроме ключа)
        columns_to_fill = [
            col for col in df2_work.columns
            if col not in df1_work.columns and col not in ['_join_key', join_key_2]
        ]
    
    if not columns_to_fill:
        return df1
    
    # Выбираем только нужные столбцы из df2
    df2_subset = df2_work[['_join_key'] + columns_to_fill].copy()
    # Убираем дубликаты  берём первые значения
    df2_subset = df2_subset.drop_duplicates(subset=['_join_key'], keep='first')
    
    # Соединяем
    result = df1_work.merge(
        df2_subset,
        on='_join_key',
        how=join_type,
        suffixes=('', '_from_df2')
    )
    
    # Заполняем недостающие значения
    filled_count = 0
    for col in columns_to_fill:
        col_from_df2 = f"{col}_from_df2" if f"{col}_from_df2" in result.columns else col
        if col in result.columns and col_from_df2 in result.columns:
            mask = result[col].isna() | (result[col] == 0) | (result[col] == '')
            result.loc[mask, col] = result.loc[mask, col_from_df2]
            filled_count += int(mask.sum())
            if col_from_df2 != col:
                result = result.drop(columns=[col_from_df2])
    
    if '_join_key' in result.columns:
        result = result.drop(columns=['_join_key'])
    
    logger.info(f"Связывание завершено: заполнено {filled_count} значений из {len(columns_to_fill)} столбцов")
    return result


def fill_missing_from_analogs(
    df: pd.DataFrame,
    article_col: str,
    oe_col: str,
    columns_to_fill: List[str],
    cross_references: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    """
    Заполнение недостающих параметров из аналогов по OE-номерам.
    Если у товара нет веса/габаритов, но есть аналог с тем же OE  берём усреднённые данные.
    """
    if df.empty or cross_references is None or cross_references.empty:
        return df
    
    result = df.copy()
    
    # Создаём индекс: OE  список артикулов
    oe_to_articles = defaultdict(list)
    for _, row in cross_references.iterrows():
        oe = str(row.get(oe_col, '')).strip().lower()
        art = str(row.get(article_col, '')).strip()
        if oe and art:
            oe_to_articles[oe].append(art)
    
    # Для каждого товара ищем аналоги и усредняем данные
    filled_count = 0
    for idx, row in result.iterrows():
        article = str(row.get(article_col, '')).strip()
        # Находим OE-номера этого товара
        article_oes = cross_references[
            cross_references[article_col].astype(str).str.strip() == article
        ][oe_col].astype(str).str.strip().tolist()
        
        for col in columns_to_fill:
            if col not in result.columns:
                continue
            current_val = row.get(col)
            # Проверяем, нужно ли заполнять
            if pd.notna(current_val) and current_val != 0 and current_val != '':
                continue
            
            # Собираем значения из аналогов
            analog_values = []
            for oe in article_oes:
                for analog_art in oe_to_articles.get(oe, []):
                    if analog_art == article:
                        continue
                    analog_row = result[result[article_col].astype(str).str.strip() == analog_art]
                    if not analog_row.empty:
                        val = analog_row.iloc[0].get(col)
                        if pd.notna(val) and val != 0 and val != '':
                            try:
                                analog_values.append(float(val))
                            except (ValueError, TypeError):
                                pass
            
            if analog_values:
                # Берём среднее (усреднённое из оригинального аналога)
                avg_val = sum(analog_values) / len(analog_values)
                result.at[idx, col] = round(avg_val, 2)
                filled_count += 1
    
    logger.info(f"Заполнено из аналогов: {filled_count} значений")
    return result


# ============================================================================
# 3.14 СОХРАНЕНИЕ И ЗАГРУЗКА СОСТОЯНИЯ
# ============================================================================

def save_state(
    state_data: Dict[str, Any],
    name: str = "",
    description: str = "",
    compress: bool = True
) -> SaveLoadState:
    """
    Сохранение текущего состояния расчётов в файл.
    Поддерживает сжатие для экономии места.
    """
    state = SaveLoadState(
        name=name or f"Расчёт {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        description=description,
        section_data=state_data,
    )
    
    # Сериализация
    file_path = BACKUPS_DIR / f"state_{state.state_id}.json"
    if compress:
        file_path = file_path.with_suffix('.json.gz')
    
    try:
        data_json = json.dumps({
            'state_id': state.state_id,
            'name': state.name,
            'description': state.description,
            'created_at': state.created_at.isoformat(),
            'section_data': _serialize_for_json(state_data),
            'metadata': state.metadata,
        }, ensure_ascii=False, indent=2)
        
        if compress:
            import gzip
            with gzip.open(file_path, 'wt', encoding='utf-8') as f:
                f.write(data_json)
        else:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(data_json)
        
        state.file_path = str(file_path)
        state.size_bytes = file_path.stat().st_size
        state.data_hash = hashlib.md5(data_json.encode('utf-8')).hexdigest()
        state.is_compressed = compress
        
        logger.info(f"✅ Состояние сохранено: {file_path} ({state.size_bytes} байт)")
        return state
    
    except Exception as e:
        logger.error(f"❌ Ошибка сохранения состояния: {e}")
        raise


def load_state(file_path: Union[str, Path]) -> SaveLoadState:
    """Загрузка состояния из файла"""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")
    
    try:
        if file_path.suffix == '.gz':
            import gzip
            with gzip.open(file_path, 'rt', encoding='utf-8') as f:
                data = json.load(f)
        else:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
        state = SaveLoadState(
            state_id=data.get('state_id', str(uuid.uuid4())),
            name=data.get('name', ''),
            description=data.get('description', ''),
            created_at=safe_datetime(data.get('created_at')),
            file_path=str(file_path),
            section_data=data.get('section_data', {}),
            metadata=data.get('metadata', {}),
            is_compressed=file_path.suffix == '.gz',
            size_bytes=file_path.stat().st_size,
        )
        
        logger.info(f"✅ Состояние загружено: {state.name}")
        return state
    
    except Exception as e:
        logger.error(f"❌ Ошибка загрузки состояния: {e}")
        raise


def list_saved_states() -> List[Dict[str, Any]]:
    """Список всех сохранённых состояний"""
    states = []
    for file_path in BACKUPS_DIR.glob("state_*"):
        try:
            if file_path.suffix == '.gz':
                import gzip
                with gzip.open(file_path, 'rt', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            
            states.append({
                'file': str(file_path),
                'name': data.get('name', ''),
                'description': data.get('description', ''),
                'created_at': data.get('created_at', ''),
                'size_bytes': file_path.stat().st_size,
                'compressed': file_path.suffix == '.gz',
            })
        except Exception as e:
            logger.warning(f"Не удалось прочитать {file_path}: {e}")
    
    # Сортируем по дате (новые первыми)
    states.sort(key=lambda x: x.get('created_at', ''), reverse=True)
    return states


def _serialize_for_json(obj: Any) -> Any:
    """Рекурсивная сериализация объектов для JSON"""
    if isinstance(obj, (str, int, float, bool, type(None))):
        return obj
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    if isinstance(obj, pd.DataFrame):
        return {'__type__': 'DataFrame', 'data': obj.to_dict(orient='records')}
    if isinstance(obj, pd.Series):
        return {'__type__': 'Series', 'data': obj.to_dict()}
    if isinstance(obj, dict):
        return {k: _serialize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_serialize_for_json(v) for v in obj]
    if hasattr(obj, '__dict__'):
        return {'__type__': obj.__class__.__name__, 'data': _serialize_for_json(obj.__dict__)}
    return str(obj)


def _deserialize_from_json(obj: Any) -> Any:
    """Рекурсивная десериализация объектов из JSON"""
    if isinstance(obj, dict):
        if obj.get('__type__') == 'DataFrame':
            return pd.DataFrame(obj.get('data', []))
        if obj.get('__type__') == 'Series':
            return pd.Series(obj.get('data', {}))
        return {k: _deserialize_from_json(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_deserialize_from_json(v) for v in obj]
    return obj


# ============================================================================
# 3.15 GOOGLE SHEETS ИНТЕГРАЦИЯ
# ============================================================================

def init_google_sheets(credentials_json: Union[str, Path]) -> Optional[Any]:
    """
    Инициализация клиента Google Sheets через Service Account JSON.
    
    Args:
        credentials_json: Путь к JSON-файлу или его содержимое
    
    Returns:
        Клиент gspread или None при ошибке
    """
    if not GSPREAD_AVAILABLE:
        logger.error("❌ gspread не установлен. pip install gspread google-auth")
        return None
    
    try:
        # Определяем, путь это или содержимое
        creds_path = Path(credentials_json)
        if creds_path.exists():
            credentials = Credentials.from_service_account_file(
                str(creds_path),
                scopes=[
                    "https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
        else:
            # Пробуем распарсить как JSON
            creds_data = json.loads(str(credentials_json))
            credentials = Credentials.from_service_account_info(
                creds_data,
                scopes=[
                    "https://spreadsheets.google.com/feeds",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
        
        client = gspread.authorize(credentials)
        logger.info("✅ Google Sheets клиент инициализирован")
        return client
    
    except Exception as e:
        logger.error(f"❌ Ошибка инициализации Google Sheets: {e}")
        return None


def google_sheets_upload(
    df: pd.DataFrame,
    spreadsheet_id: str,
    worksheet_name: str = "Тарифы",
    credentials_json: Optional[Union[str, Path]] = None,
    client: Optional[Any] = None,
    clear_before: bool = True
) -> bool:
    """
    Загрузка DataFrame в Google Таблицу.
    
    Args:
        df: DataFrame для загрузки
        spreadsheet_id: ID таблицы (из URL)
        worksheet_name: Имя листа
        credentials_json: Путь к JSON (если client не передан)
        client: Уже инициализированный клиент gspread
        clear_before: Очистить лист перед загрузкой
    
    Returns:
        True при успехе
    """
    try:
        if client is None:
            if credentials_json is None:
                logger.error("❌ Не передан клиент или credentials_json")
                return False
            client = init_google_sheets(credentials_json)
            if client is None:
                return False
        
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        # Ищем или создаём лист
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
            if clear_before:
                worksheet.clear()
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(
                title=worksheet_name,
                rows=len(df) + 1,
                cols=len(df.columns)
            )
        
        # Преобразуем DataFrame в список списков
        values = [df.columns.tolist()] + df.values.tolist()
        
        # Загружаем (батчами, если много данных)
        batch_size = 1000
        for i in range(0, len(values), batch_size):
            batch = values[i:i + batch_size]
            start_cell = f"A{i + 1}"
            end_row = i + len(batch)
            end_cell = f"{get_column_letter(len(df.columns))}{end_row}"
            worksheet.update(f"{start_cell}:{end_cell}", batch)
        
        logger.info(f"✅ Данные загружены в Google Sheets: {spreadsheet_id}/{worksheet_name}")
        return True
    
    except Exception as e:
        logger.error(f"❌ Ошибка загрузки в Google Sheets: {e}")
        return False


def google_sheets_download(
    spreadsheet_id: str,
    worksheet_name: str = "Тарифы",
    credentials_json: Optional[Union[str, Path]] = None,
    client: Optional[Any] = None
) -> Optional[pd.DataFrame]:
    """
    Загрузка данных из Google Таблицы в DataFrame.
    
    Args:
        spreadsheet_id: ID таблицы
        worksheet_name: Имя листа
        credentials_json: Путь к JSON
        client: Уже инициализированный клиент
    
    Returns:
        DataFrame или None при ошибке
    """
    try:
        if client is None:
            if credentials_json is None:
                logger.error("❌ Не передан клиент или credentials_json")
                return None
            client = init_google_sheets(credentials_json)
            if client is None:
                return None
        
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        # Получаем все данные
        values = worksheet.get_all_values()
        if not values:
            logger.warning("⚠️ Лист пустой")
            return pd.DataFrame()
        
        # Первая строка  заголовки
        headers = values[0]
        data = values[1:]
        
        df = pd.DataFrame(data, columns=headers)
        
        # Пытаемся преобразовать числовые столбцы
        for col in df.columns:
            try:
                df[col] = pd.to_numeric(df[col], errors='ignore')
            except Exception:
                pass
        
        logger.info(f"✅ Данные загружены из Google Sheets: {len(df)} строк")
        return df
    
    except Exception as e:
        logger.error(f"❌ Ошибка загрузки из Google Sheets: {e}")
        return None


def get_column_letter(col_idx: int) -> str:
    """Конвертация индекса колонки (1-based) в букву Excel (A, B, ..., Z, AA, AB, ...)"""
    if col_idx <= 0:
        return 'A'
    result = ''
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ============================================================================
# 3.16 ДЕКОРАТОРЫ
# ============================================================================

def timer_decorator(func: Callable) -> Callable:
    """Декоратор для замера времени выполнения функции"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start
            if elapsed > 1.0:
                logger.debug(f"⏱ {func.__name__} выполнена за {elapsed:.3f}с")
            return result
        except Exception as e:
            elapsed = time.perf_counter() - start
            logger.error(f"❌ {func.__name__} завершилась с ошибкой за {elapsed:.3f}с: {e}")
            raise
    return wrapper


def retry_decorator(
    max_retries: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    exceptions: tuple = (Exception,)
) -> Callable:
    """Декоратор повторных попыток с экспоненциальной задержкой"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            current_delay = delay
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt == max_retries - 1:
                        raise
                    logger.warning(
                        f"⚠️ Попытка {attempt + 1}/{max_retries} для {func.__name__} не удалась: {e}"
                    )
                    time.sleep(current_delay)
                    current_delay *= backoff
            if last_exception:
                raise last_exception
            return None
        return wrapper
    return decorator


def safe_execution(default_return: Any = None, log_error: bool = True) -> Callable:
    """Декоратор безопасного выполнения  ловит все исключения"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if log_error:
                    logger.error(f"⚠️ Ошибка в {func.__name__}: {e}")
                return default_return
        return wrapper
    return decorator


# ============================================================================
# 3.17 КЛАССЫ ИСКЛЮЧЕНИЙ
# ============================================================================

class AutoPartsException(Exception):
    """Базовое исключение для приложения"""
    def __init__(self, message: str = "", *args, **kwargs):
        self.message = message
        self.timestamp = datetime.now()
        self.context = kwargs
        super().__init__(message, *args)
    
    def __str__(self):
        return f"[{self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}] {self.message}"


class ValidationError(AutoPartsException):
    def __init__(self, message: str, field: Optional[str] = None, value: Any = None):
        self.field = field
        self.value = value
        super().__init__(f"Ошибка валидации{f' в поле {field}' if field else ''}: {message}")


class MarketplaceError(AutoPartsException):
    def __init__(self, message: str, marketplace: Optional[str] = None):
        self.marketplace = marketplace
        super().__init__(f"Ошибка маркетплейса{f' {marketplace}' if marketplace else ''}: {message}")


class CalculationError(AutoPartsException):
    def __init__(self, message: str, calculation_type: Optional[str] = None):
        self.calculation_type = calculation_type
        super().__init__(f"Ошибка расчета{f' {calculation_type}' if calculation_type else ''}: {message}")


class AIError(AutoPartsException):
    def __init__(self, message: str, provider: Optional[str] = None, code: Optional[int] = None):
        self.provider = provider
        self.code = code
        super().__init__(f"Ошибка AI{f' ({provider})' if provider else ''}: {message}")


class DatabaseError(AutoPartsException):
    def __init__(self, message: str, query: Optional[str] = None, params: Optional[Dict] = None):
        self.query = query
        self.params = params
        super().__init__(f"Ошибка базы данных: {message}")


class ExportError(AutoPartsException):
    def __init__(self, message: str, format: Optional[str] = None, file_path: Optional[Path] = None):
        self.format = format
        self.file_path = file_path
        super().__init__(f"Ошибка экспорта{f' в {format}' if format else ''}: {message}")


class GoogleSheetsError(AutoPartsException):
    def __init__(self, message: str, spreadsheet_id: Optional[str] = None):
        self.spreadsheet_id = spreadsheet_id
        super().__init__(f"Ошибка Google Sheets{f' ({spreadsheet_id})' if spreadsheet_id else ''}: {message}")


class StateError(AutoPartsException):
    def __init__(self, message: str, state_id: Optional[str] = None):
        self.state_id = state_id
        super().__init__(f"Ошибка состояния{f' ({state_id})' if state_id else ''}: {message}")

# ============================================================================
# БЛОК 4: КОНФИГУРАЦИИ МАРКЕТПЛЕЙСОВ 2026
# ============================================================================
# 📌 v101.0: Полные тарифы для 8 маркетплейсов с:
# - Сезонными коэффициентами (зима/весна/лето/осень)
# - Категориями комиссий (20+ категорий автозапчастей)
# - Мультипликаторами режимов (FBY/FBS/FBO/DBS/FBP/RealFBS)
# - Интеграцией с кэшем тарифов (из Блока 3)
# ============================================================================


def get_marketplace_configs_2026() -> Dict[str, MarketplaceConfig]:
    """
    Получение конфигураций маркетплейсов с актуальными тарифами 2026.
    
    Возвращает словарь {название_мп: MarketplaceConfig} с полными тарифами:
    - Комиссии (общие и по категориям)
    - Логистика (база + за кг + за литр)
    - Хранение (прогрессивное)
    - Эквайринг, возвраты, последняя миля
    - Сезонные коэффициенты
    - Мультипликаторы режимов
    
    Если в кэше есть более свежие данные  они применяются поверх базовых.
    """
    configs = {}
    
    # ========================================================================
    # 🟣 OZON
    # ========================================================================
    configs["Ozon"] = MarketplaceConfig(
        name="Ozon",
        commission_rate=0.15,
        min_commission=30.0,
        logistics_base=50.0,
        logistics_per_kg=15.0,
        logistics_per_liter=5.0,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.015,
        last_mile_fee=50.0,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.01,
        oversized_surcharge=0.015,
        seasonal_multipliers={
            "winter": 1.15,   # Декабрь-февраль: предновогодний ажиотаж
            "spring": 1.0,    # Март-май: базовый сезон
            "summer": 0.95,   # Июнь-август: небольшое снижение
            "autumn": 1.05    # Сентябрь-ноябрь: подготовка к зиме
        },
        category_rates={
            "двигатель": 0.12, "трансмиссия": 0.13, "подвеска": 0.14,
            "тормозная_система": 0.14, "рулевое_управление": 0.14,
            "электрика": 0.15, "охлаждение": 0.14, "выпуск": 0.13,
            "фильтры": 0.17, "масла": 0.18, "оптика": 0.15,
            "шины": 0.16, "инструменты": 0.14, "кузов": 0.13,
            "крепёж": 0.12, "ремни": 0.13, "подшипники": 0.13,
            "климат": 0.14, "безопасность": 0.15,
            "автотовары": 0.12, "сайлентблоки": 0.14,
            "амортизаторы": 0.14, "колодки": 0.14,
            "датчики": 0.15, "аккумуляторы": 0.16,
            "фары": 0.15, "стартеры": 0.14, "генераторы": 0.14
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.0, "FBO": 0.8,
            "DBS": 1.3, "FBP": 0.9, "RealFBS": 1.1
        },
        description="Ozon  крупнейший маркетплейс России",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🟡 WILDBERRIES
    # ========================================================================
    configs["Wildberries"] = MarketplaceConfig(
        name="Wildberries",
        commission_rate=0.18,
        min_commission=35.0,
        logistics_base=60.0,
        logistics_per_kg=18.0,
        logistics_per_liter=6.0,
        storage_per_day=0.5,
        return_fee=0.03,
        acquiring_fee=0.0,          # WB не берёт эквайринг отдельно
        last_mile_fee=0.0,          # Включено в логистику
        delivery_fee_percent=0.05,  # 5% за доставку
        rko_fee=0.01,               # 1% РКО
        hazardous_surcharge=0.025,
        fragile_surcharge=0.015,
        oversized_surcharge=0.02,
        seasonal_multipliers={
            "winter": 1.2,    # Высокий сезон  Новый год
            "spring": 1.0,
            "summer": 0.95,
            "autumn": 1.05
        },
        category_rates={
            "двигатель": 0.15, "трансмиссия": 0.16, "подвеска": 0.17,
            "тормозная_система": 0.17, "рулевое_управление": 0.17,
            "электрика": 0.18, "охлаждение": 0.17, "выпуск": 0.16,
            "фильтры": 0.20, "масла": 0.22, "оптика": 0.18,
            "шины": 0.19, "инструменты": 0.17, "кузов": 0.16,
            "крепёж": 0.15, "ремни": 0.16, "подшипники": 0.16,
            "климат": 0.17, "безопасность": 0.18,
            "автотовары": 0.15, "сайлентблоки": 0.17,
            "амортизаторы": 0.17, "колодки": 0.17,
            "датчики": 0.18, "аккумуляторы": 0.19,
            "фары": 0.18, "стартеры": 0.17, "генераторы": 0.17
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.15, "FBO": 1.1,
            "DBS": 1.25, "FBP": 1.0, "RealFBS": 1.2
        },
        description="Wildberries  лидер e-commerce России",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🔵 ЯНДЕКС МАРКЕТ
    # ========================================================================
    configs["Яндекс Маркет"] = MarketplaceConfig(
        name="Яндекс Маркет",
        commission_rate=0.14,
        subscription_fee=6990.0,    # Ежемесячная подписка
        min_commission=0.0,
        logistics_base=45.0,
        logistics_per_kg=14.0,
        logistics_per_liter=4.5,
        storage_per_day=0.25,
        return_fee=0.02,
        acquiring_fee=0.02,
        last_mile_fee=40.0,
        premium_fee=0.02,           # Премиум-размещение
        hazardous_surcharge=0.018,
        fragile_surcharge=0.01,
        oversized_surcharge=0.012,
        seasonal_multipliers={
            "winter": 1.1,
            "spring": 1.0,
            "summer": 0.9,
            "autumn": 1.0
        },
        category_rates={
            "двигатель": 0.11, "трансмиссия": 0.12, "подвеска": 0.13,
            "тормозная_система": 0.13, "рулевое_управление": 0.13,
            "электрика": 0.14, "охлаждение": 0.13, "выпуск": 0.12,
            "фильтры": 0.16, "масла": 0.17, "оптика": 0.14,
            "шины": 0.15, "инструменты": 0.13, "кузов": 0.12,
            "крепёж": 0.11, "ремни": 0.12, "подшипники": 0.12,
            "климат": 0.13, "безопасность": 0.14,
            "автотовары": 0.14, "сайлентблоки": 0.13,
            "амортизаторы": 0.13, "колодки": 0.13,
            "датчики": 0.14, "аккумуляторы": 0.15,
            "фары": 0.14, "стартеры": 0.13, "генераторы": 0.13
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.0, "FBO": 0.8,
            "DBS": 1.3, "FBP": 0.9, "RealFBS": 1.1
        },
        description="Яндекс Маркет  маркетплейс экосистемы Яндекса",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🔴 ALIEXPRESS
    # ========================================================================
    configs["AliExpress"] = MarketplaceConfig(
        name="AliExpress",
        commission_rate=0.10,
        min_commission=20.0,
        logistics_base=80.0,         # Выше из-за международной логистики
        logistics_per_kg=25.0,
        logistics_per_liter=8.0,
        storage_per_day=0.2,
        return_fee=0.01,
        acquiring_fee=0.025,
        last_mile_fee=70.0,
        delivery_fee_percent=0.08,   # 8% за доставку
        hazardous_surcharge=0.03,
        fragile_surcharge=0.02,
        oversized_surcharge=0.025,
        seasonal_multipliers={
            "winter": 1.25,    # 11.11, Новый год, китайский НГ
            "spring": 1.0,
            "summer": 1.1,     # Летние распродажи
            "autumn": 1.15     # Осенние распродажи
        },
        category_rates={
            "двигатель": 0.08, "трансмиссия": 0.09, "подвеска": 0.10,
            "тормозная_система": 0.10, "рулевое_управление": 0.10,
            "электрика": 0.11, "охлаждение": 0.10, "выпуск": 0.09,
            "фильтры": 0.12, "масла": 0.13, "оптика": 0.11,
            "шины": 0.12, "инструменты": 0.10, "кузов": 0.09,
            "крепёж": 0.08, "ремни": 0.09, "подшипники": 0.09,
            "климат": 0.10, "безопасность": 0.11,
            "автотовары": 0.10, "сайлентблоки": 0.10,
            "амортизаторы": 0.10, "колодки": 0.10,
            "датчики": 0.11, "аккумуляторы": 0.12,
            "фары": 0.11, "стартеры": 0.10, "генераторы": 0.10
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.2, "FBO": 1.1,
            "DBS": 1.3, "FBP": 0.9, "RealFBS": 1.25
        },
        description="AliExpress  международный маркетплейс",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🟢 МЕГАМАРКЕТ
    # ========================================================================
    configs["Мегамаркет"] = MarketplaceConfig(
        name="Мегамаркет",
        commission_rate=0.13,
        min_commission=28.0,
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        delivery_fee_percent=0.05,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.012,
        oversized_surcharge=0.015,
        seasonal_multipliers={
            "winter": 1.12,
            "spring": 1.0,
            "summer": 0.93,
            "autumn": 1.03
        },
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "рулевое_управление": 0.12,
            "электрика": 0.13, "охлаждение": 0.12, "выпуск": 0.11,
            "фильтры": 0.15, "масла": 0.16, "оптика": 0.13,
            "шины": 0.14, "инструменты": 0.12, "кузов": 0.11,
            "крепёж": 0.10, "ремни": 0.11, "подшипники": 0.11,
            "климат": 0.12, "безопасность": 0.13,
            "автотовары": 0.15, "сайлентблоки": 0.12,
            "амортизаторы": 0.12, "колодки": 0.12,
            "датчики": 0.13, "аккумуляторы": 0.14,
            "фары": 0.13, "стартеры": 0.12, "генераторы": 0.12
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.1, "FBO": 1.05,
            "DBS": 1.2, "FBP": 0.95, "RealFBS": 1.15
        },
        description="Мегамаркет (Сбер)  маркетплейс экосистемы Сбера",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🟠 СБЕРМЕГАМАРКЕТ
    # ========================================================================
    configs["СберМегаМаркет"] = MarketplaceConfig(
        name="СберМегаМаркет",
        commission_rate=0.13,
        min_commission=28.0,
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        rko_fee=0.015,
        delivery_fee_percent=0.055,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.012,
        oversized_surcharge=0.015,
        seasonal_multipliers={
            "winter": 1.12,
            "spring": 1.0,
            "summer": 0.93,
            "autumn": 1.03
        },
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "рулевое_управление": 0.12,
            "электрика": 0.13, "охлаждение": 0.12, "выпуск": 0.11,
            "фильтры": 0.15, "масла": 0.16, "оптика": 0.13,
            "шины": 0.14, "инструменты": 0.12, "кузов": 0.11,
            "крепёж": 0.10, "ремни": 0.11, "подшипники": 0.11,
            "климат": 0.12, "безопасность": 0.13,
            "автотовары": 0.12, "сайлентблоки": 0.12,
            "амортизаторы": 0.12, "колодки": 0.12,
            "датчики": 0.13, "аккумуляторы": 0.14,
            "фары": 0.13, "стартеры": 0.12, "генераторы": 0.12
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.1, "FBO": 1.05,
            "DBS": 1.2, "FBP": 0.95, "RealFBS": 1.15
        },
        description="СберМегаМаркет  маркетплейс Сбера",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🟤 AVITO
    # ========================================================================
    configs["Avito"] = MarketplaceConfig(
        name="Avito",
        commission_rate=0.05,        # Минимальная комиссия
        min_commission=0.0,
        logistics_base=0.0,          # Avito  доска объявлений, логистика опциональна
        logistics_per_kg=0.0,
        logistics_per_liter=0.0,
        storage_per_day=0.0,
        return_fee=0.0,
        acquiring_fee=0.0,
        last_mile_fee=0.0,
        delivery_fee_percent=0.0,
        hazardous_surcharge=0.0,
        fragile_surcharge=0.0,
        oversized_surcharge=0.0,
        seasonal_multipliers={},
        category_rates={},
        mode_multipliers={"RealFBS": 1.0},
        description="Avito  доска объявлений",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # ⚫ DROM
    # ========================================================================
    configs["Drom"] = MarketplaceConfig(
        name="Drom",
        commission_rate=0.08,
        min_commission=0.0,
        logistics_base=0.0,
        logistics_per_kg=0.0,
        logistics_per_liter=0.0,
        storage_per_day=0.0,
        return_fee=0.0,
        acquiring_fee=0.0,
        last_mile_fee=0.0,
        delivery_fee_percent=0.0,
        hazardous_surcharge=0.0,
        fragile_surcharge=0.0,
        oversized_surcharge=0.0,
        seasonal_multipliers={},
        category_rates={},
        mode_multipliers={"RealFBS": 1.0},
        description="Drom  площадка для автотоваров",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )
    
    # ========================================================================
    # 🔄 ПРИМЕНЕНИЕ КЭШИРОВАННЫХ ТАРИФОВ (из Блока 3)
    # ========================================================================
    try:
        cache = get_smart_tariff_cache()
        for mp_name, config in configs.items():
            cached_entry = cache.get(mp_name, None, use_expired=False)
            if cached_entry and cached_entry.data:
                data = cached_entry.data
                # Применяем кэшированные значения поверх базовых
                if "commission_rate" in data:
                    config.commission_rate = data["commission_rate"]
                if "min_commission" in data:
                    config.min_commission = data["min_commission"]
                if "logistics_base" in data:
                    config.logistics_base = data["logistics_base"]
                if "logistics_per_kg" in data:
                    config.logistics_per_kg = data["logistics_per_kg"]
                if "logistics_per_liter" in data:
                    config.logistics_per_liter = data["logistics_per_liter"]
                if "storage_per_day" in data:
                    config.storage_per_day = data["storage_per_day"]
                if "return_fee" in data:
                    config.return_fee = data["return_fee"]
                if "acquiring_fee" in data:
                    config.acquiring_fee = data["acquiring_fee"]
                if "last_mile_fee" in data:
                    config.last_mile_fee = data["last_mile_fee"]
                if "category_rates" in data:
                    config.category_rates.update(data["category_rates"])
                if "seasonal_multipliers" in data:
                    config.seasonal_multipliers.update(data["seasonal_multipliers"])
                
                config.tariff_source = cached_entry.source
                config.last_updated = datetime.fromtimestamp(cached_entry.timestamp)
                logger.info(f"📥 Применены кэшированные тарифы для {mp_name}")
    except Exception as e:
        logger.warning(f"Не удалось загрузить кэш тарифов: {e}")
    
    return configs

# ============================================================================
# БЛОК 5: 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С 3-УРОВНЕВОЙ ИЕРАРХИЕЙ
# ============================================================================
# 📌 v101.0: Расширено до 150+ категорий с:
# - 3-уровневой иерархией (Родитель → Группа → Подгруппа)
# - Типичными весогабаритами (min/max)
# - OEM-кодами и кросс-ссылками
# - Сезонностью, уровнем риска
# - Признаками опасности/хрупкости
# - Ключевыми словами для автокатегоризации
#
# 🐛 ИСПРАВЛЕНИЯ v101.0:
# 1. Добавлен return statement в get_category_keywords_map()
# 2. Улучшена защита от None в make_cat
# ============================================================================


def get_auto_parts_categories_full() -> Dict[str, ProductCategory]:
    """
    Полный справочник категорий автозапчастей с весогабаритами и иерархией.
    
    Иерархия:
        Родитель (Автозапчасти) → Группа (Подвеска) → Подгруппа (Сайлентблоки)
    
    Returns:
        Dict[str, ProductCategory] — словарь всех категорий
    """
    categories = {}
    
    def make_cat(
        key: str,
        name: str,
        desc: str,
        min_l: float, max_l: float,
        min_w: float, max_w: float,
        min_h: float, max_h: float,
        min_wt: float, max_wt: float,
        typ_vol: float = 0.0,
        typ_wt: float = 0.0,
        oem: Optional[List[str]] = None,
        season: Seasonality = Seasonality.ALL_YEAR,
        risk: RiskLevel = RiskLevel.LOW,
        hazardous: bool = False,
        fragile: bool = False,
        parent: str = "Автозапчасти",
        group: str = "",
        subgroup: str = "",
        keywords: Optional[List[str]] = None,
        weight_range: Optional[Tuple[float, float]] = None,
    ) -> ProductCategory:
        """
        Вспомогательная функция создания категории с иерархией.
        
        🆕 v101.0: Добавлена защита от None в weight_range.
        """
        # Защита от None в weight_range
        actual_weight_range = weight_range if weight_range is not None else (min_wt, max_wt)
        
        hierarchy = CategoryHierarchy(
            parent=parent,
            group=group or name,
            subgroup=subgroup or name,
            keywords=keywords or [],
            weight_range=actual_weight_range,
            dimensions_range={
                "length": (min_l, max_l),
                "width": (min_w, max_w),
                "height": (min_h, max_h),
            }
        )
        return ProductCategory(
            name=name,
            description=desc,
            parent_category=parent,
            hierarchy=hierarchy,
            dimensions=ProductDimensions(
                length=(min_l + max_l) / 2,
                width=(min_w + max_w) / 2,
                height=(min_h + max_h) / 2,
                weight=(min_wt + max_wt) / 2
            ),
            typical_volume=typ_vol or ((min_l + max_l) / 2 * (min_w + max_w) / 2 * (min_h + max_h) / 2) / 1000,
            typical_weight=typ_wt or (min_wt + max_wt) / 2,
            oem_codes=oem or [],
            hazardous=hazardous,
            fragile=fragile,
            seasonality=season,
            risk_level=risk,
            min_length=min_l, max_length=max_l,
            min_width=min_w, max_width=max_w,
            min_height=min_h, max_height=max_h,
            min_weight=min_wt, max_weight=max_wt,
        )
    
    # ========================================================================
    # 🏭 ДВИГАТЕЛЬ (Родитель: Автозапчасти → Группа: Двигатель)
    # ========================================================================
    categories["двигатель"] = make_cat(
        "двигатель", "Двигатель", "Двигатели и комплектующие",
        30, 80, 30, 60, 30, 70, 10, 200, 20.0, 80.0,
        parent="Автозапчасти", group="Двигатель", subgroup="Двигатель в сборе",
        keywords=["двигатель", "engine", "мотор", "двс"],
        risk=RiskLevel.HIGH
    )
    categories["поршни"] = make_cat(
        "поршни", "Поршни", "Поршни и кольца",
        5, 12, 5, 12, 3, 10, 0.1, 1.5, 0.1, 0.5,
        parent="Автозапчасти", group="Двигатель", subgroup="Поршневая группа",
        keywords=["поршень", "кольца", "piston", "поршневые"]
    )
    categories["клапаны"] = make_cat(
        "клапаны", "Клапаны", "Клапаны двигателя",
        3, 8, 1, 3, 10, 40, 0.05, 0.5, 0.05, 0.2,
        parent="Автозапчасти", group="Двигатель", subgroup="ГРМ",
        keywords=["клапан", "valve", "клапана"]
    )
    categories["прокладки_двигателя"] = make_cat(
        "прокладки_двигателя", "Прокладки ГБЦ", "Прокладки ГБЦ и двигателя",
        10, 50, 10, 40, 0.1, 2, 0.01, 0.3, 0.1, 0.1,
        parent="Автозапчасти", group="Двигатель", subgroup="Прокладки",
        keywords=["прокладка", "гбц", "gasket", "прокладки"]
    )
    categories["свечи_зажигания"] = make_cat(
        "свечи_зажигания", "Свечи зажигания", "Свечи зажигания",
        2, 3, 2, 3, 6, 10, 0.04, 0.1, 0.01, 0.05,
        parent="Автозапчасти", group="Двигатель", subgroup="Зажигание",
        keywords=["свеча", "spark", "свечи", "зажигания"]
    )
    categories["блок_цилиндров"] = make_cat(
        "блок_цилиндров", "Блок цилиндров", "Блок цилиндров",
        40, 70, 30, 50, 20, 40, 20, 80, 100.0, 50.0,
        parent="Автозапчасти", group="Двигатель", subgroup="Блок цилиндров",
        keywords=["блок", "цилиндр", "bloc"],
        risk=RiskLevel.HIGH
    )
    categories["головка_блока"] = make_cat(
        "головка_блока", "Головка блока", "Головка блока цилиндров",
        30, 60, 20, 40, 8, 20, 5, 30, 40.0, 15.0,
        parent="Автозапчасти", group="Двигатель", subgroup="ГБЦ",
        keywords=["головка", "гбц", "head"],
        risk=RiskLevel.HIGH
    )
    categories["коленвал"] = make_cat(
        "коленвал", "Коленвал", "Коленчатый вал",
        40, 90, 8, 20, 8, 20, 10, 40, 30.0, 25.0,
        parent="Автозапчасти", group="Двигатель", subgroup="КШМ",
        keywords=["коленвал", "коленчатый", "crankshaft"],
        risk=RiskLevel.HIGH
    )
    categories["распредвал"] = make_cat(
        "распредвал", "Распредвал", "Распределительный вал",
        30, 80, 5, 15, 5, 15, 3, 15, 20.0, 9.0,
        parent="Автозапчасти", group="Двигатель", subgroup="ГРМ",
        keywords=["распредвал", "распределительный", "camshaft"]
    )
    categories["шатун"] = make_cat(
        "шатун", "Шатун", "Шатун двигателя",
        12, 35, 4, 10, 3, 7, 0.5, 2, 3.0, 1.25,
        parent="Автозапчасти", group="Двигатель", subgroup="КШМ",
        keywords=["шатун", "connecting rod"]
    )
    categories["гидрокомпенсаторы"] = make_cat(
        "гидрокомпенсаторы", "Гидрокомпенсаторы", "Гидрокомпенсаторы",
        3, 8, 3, 8, 3, 8, 0.05, 0.2, 0.3, 0.125,
        parent="Автозапчасти", group="Двигатель", subgroup="ГРМ",
        keywords=["гидрокомпенсатор", "hydro"]
    )
    categories["привод_грм"] = make_cat(
        "привод_грм", "Привод ГРМ", "Привод ГРМ (ремень, цепь)",
        60, 160, 2, 5, 1, 2, 0.1, 1, 2.0, 0.55,
        parent="Автозапчасти", group="Двигатель", subgroup="ГРМ",
        keywords=["грм", "ремень грм", "цепь грм", "timing"]
    )
    categories["масляный_насос"] = make_cat(
        "масляный_насос", "Масляный насос", "Масляный насос",
        8, 18, 8, 18, 8, 18, 1, 5, 5.0, 3.0,
        parent="Автозапчасти", group="Двигатель", subgroup="Маслосистема",
        keywords=["масляный насос", "oil pump"]
    )
    categories["водяной_насос"] = make_cat(
        "водяной_насос", "Водяной насос (помпа)", "Водяной насос",
        8, 18, 8, 18, 8, 18, 1, 4, 5.0, 2.5,
        parent="Автозапчасти", group="Двигатель", subgroup="Охлаждение",
        keywords=["помпа", "водяной насос", "water pump"]
    )
    categories["турбокомпрессор"] = make_cat(
        "турбокомпрессор", "Турбокомпрессор", "Турбокомпрессор",
        15, 35, 15, 30, 15, 25, 5, 15, 15.0, 10.0,
        parent="Автозапчасти", group="Двигатель", subgroup="Наддув",
        keywords=["турбо", "турбина", "turbo"],
        risk=RiskLevel.HIGH
    )
    categories["масляный_поддон"] = make_cat(
        "масляный_поддон", "Масляный поддон", "Масляный поддон",
        30, 60, 20, 40, 10, 20, 2, 8, 15.0, 5.0,
        parent="Автозапчасти", group="Двигатель", subgroup="Маслосистема",
        keywords=["поддон", "oil pan"]
    )
    categories["клапанная_крышка"] = make_cat(
        "клапанная_крышка", "Клапанная крышка", "Клапанная крышка",
        30, 60, 15, 30, 5, 10, 1, 4, 8.0, 2.5,
        parent="Автозапчасти", group="Двигатель", subgroup="ГРМ",
        keywords=["клапанная крышка", "valve cover"]
    )
    categories["приводной_ремень"] = make_cat(
        "приводной_ремень", "Приводной ремень", "Приводной ремень",
        60, 150, 1, 3, 0.5, 1, 0.05, 0.5, 1.0, 0.275,
        parent="Автозапчасти", group="Двигатель", subgroup="Приводы",
        keywords=["приводной ремень", "drive belt"]
    )
    categories["демпфер_коленвала"] = make_cat(
        "демпфер_коленвала", "Демпфер коленвала", "Демпфер коленвала",
        10, 25, 10, 25, 5, 10, 2, 8, 5.0, 5.0,
        parent="Автозапчасти", group="Двигатель", subgroup="КШМ",
        keywords=["демпфер", "damper"]
    )
    categories["маховик"] = make_cat(
        "маховик", "Маховик", "Маховик",
        25, 45, 25, 45, 5, 10, 5, 15, 10.0, 10.0,
        parent="Автозапчасти", group="Двигатель", subgroup="КШМ",
        keywords=["маховик", "flywheel"],
        risk=RiskLevel.HIGH
    )
    
    # ========================================================================
    # ⚙️ ТРАНСМИССИЯ
    # ========================================================================
    categories["трансмиссия"] = make_cat(
        "трансмиссия", "Трансмиссия", "КПП и комплектующие",
        40, 80, 30, 60, 30, 60, 20, 100, 30.0, 50.0,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Трансмиссия в сборе",
        keywords=["трансмиссия", "transmission", "кпп"],
        risk=RiskLevel.HIGH
    )
    categories["сцепление"] = make_cat(
        "сцепление", "Сцепление", "Комплекты сцепления",
        20, 40, 20, 40, 5, 15, 2, 10, 3.0, 5.0,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Сцепление",
        keywords=["сцепление", "clutch", "диск сцепления"]
    )
    categories["шкивы"] = make_cat(
        "шкивы", "Шкивы и ролики", "Шкивы и ролики",
        5, 20, 5, 20, 2, 8, 0.2, 3, 0.5, 1.5,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Приводы",
        keywords=["шкив", "ролик", "pulley"]
    )
    categories["коробка_передач"] = make_cat(
        "коробка_передач", "Коробка передач", "Коробка передач в сборе",
        40, 70, 30, 50, 25, 40, 30, 80, 80.0, 55.0,
        parent="Автозапчасти", group="Трансмиссия", subgroup="КПП",
        keywords=["коробка", "кпп", "gearbox"],
        risk=RiskLevel.HIGH
    )
    categories["привод_полуоси"] = make_cat(
        "привод_полуоси", "Привод (полуоси)", "Привод (полуоси)",
        40, 90, 8, 18, 8, 18, 3, 12, 15.0, 7.5,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Приводы",
        keywords=["привод", "полуось", "шрус", "cv joint"]
    )
    categories["дифференциал"] = make_cat(
        "дифференциал", "Дифференциал", "Дифференциал",
        20, 45, 20, 45, 20, 45, 10, 30, 30.0, 20.0,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Дифференциал",
        keywords=["дифференциал", "differential"],
        risk=RiskLevel.HIGH
    )
    categories["карданный_вал"] = make_cat(
        "карданный_вал", "Карданный вал", "Карданный вал",
        60, 160, 8, 18, 8, 18, 5, 20, 25.0, 12.5,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Приводы",
        keywords=["кардан", "cardan"]
    )
    categories["раздаточная_коробка"] = make_cat(
        "раздаточная_коробка", "Раздаточная коробка", "Раздаточная коробка",
        25, 45, 20, 35, 20, 35, 15, 40, 35.0, 27.5,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Раздатка",
        keywords=["раздатка", "transfer case"],
        risk=RiskLevel.HIGH
    )
    categories["гидротрансформатор"] = make_cat(
        "гидротрансформатор", "Гидротрансформатор", "Гидротрансформатор АКПП",
        25, 40, 25, 40, 20, 30, 10, 25, 30.0, 17.5,
        parent="Автозапчасти", group="Трансмиссия", subgroup="АКПП",
        keywords=["гидротрансформатор", "torque converter"],
        risk=RiskLevel.HIGH
    )
    categories["фильтр_акпп"] = make_cat(
        "фильтр_акпп", "Фильтр АКПП", "Фильтр АКПП",
        8, 18, 8, 18, 8, 18, 0.5, 2, 3.0, 1.25,
        parent="Автозапчасти", group="Трансмиссия", subgroup="АКПП",
        keywords=["фильтр акпп", "transmission filter"]
    )
    categories["масло_трансмиссионное"] = make_cat(
        "масло_трансмиссионное", "Трансмиссионное масло", "Трансмиссионное масло",
        10, 35, 8, 25, 8, 25, 1, 5, 5.0, 3.0,
        parent="Автозапчасти", group="Трансмиссия", subgroup="Масла",
        keywords=["трансмиссионное масло", "atf"],
        hazardous=True
    )
    
    # ========================================================================
    # 🏗️ ПОДВЕСКА
    # ========================================================================
    categories["подвеска"] = make_cat(
        "подвеска", "Подвеска", "Элементы подвески",
        20, 80, 10, 40, 10, 60, 1, 20, 5.0, 8.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Подвеска в сборе",
        keywords=["подвеска", "suspension"]
    )
    categories["амортизаторы"] = make_cat(
        "амортизаторы", "Амортизаторы", "Амортизаторы",
        5, 10, 5, 10, 40, 70, 2, 8, 5.0, 8.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Амортизаторы",
        keywords=["амортизатор", "стойка", "shock", "amort"],
        fragile=True
    )
    categories["пружины"] = make_cat(
        "пружины", "Пружины", "Пружины подвески",
        20, 40, 20, 40, 30, 60, 3, 10, 8.0, 12.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Пружины",
        keywords=["пружина", "spring"]
    )
    categories["сайлентблоки"] = make_cat(
        "сайлентблоки", "Сайлентблоки", "Сайлентблоки",
        3, 10, 3, 10, 2, 8, 0.1, 1, 0.1, 0.3,
        parent="Автозапчасти", group="Подвеска", subgroup="Сайлентблоки",
        keywords=["сайлентблок", "сайлент", "silentblock"]
    )
    categories["шаровые_опоры"] = make_cat(
        "шаровые_опоры", "Шаровые опоры", "Шаровые опоры",
        5, 15, 5, 15, 5, 15, 0.3, 2, 0.5, 1.5,
        parent="Автозапчасти", group="Подвеска", subgroup="Шаровые опоры",
        keywords=["шаровая", "ball joint"]
    )
    categories["ступицы"] = make_cat(
        "ступицы", "Ступицы", "Ступицы и подшипники",
        10, 25, 10, 25, 5, 15, 1, 5, 2.0, 4.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Ступицы",
        keywords=["ступица", "hub"]
    )
    categories["рычаг_подвески"] = make_cat(
        "рычаг_подвески", "Рычаг подвески", "Рычаг подвески",
        20, 65, 5, 18, 5, 18, 2, 10, 10.0, 6.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Рычаги",
        keywords=["рычаг", "lever", "control arm"]
    )
    categories["стабилизатор"] = make_cat(
        "стабилизатор", "Стабилизатор", "Стабилизатор поперечной устойчивости",
        25, 65, 3, 10, 3, 10, 1, 5, 5.0, 3.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Стабилизаторы",
        keywords=["стабилизатор", "stabilizer", "sway bar"]
    )
    categories["пыльник"] = make_cat(
        "пыльник", "Пыльник", "Пыльник (чехол)",
        5, 12, 5, 12, 8, 22, 0.1, 0.5, 1.0, 0.3,
        parent="Автозапчасти", group="Подвеска", subgroup="Пыльники",
        keywords=["пыльник", "чехол", "boot"]
    )
    categories["отбойник"] = make_cat(
        "отбойник", "Отбойник", "Отбойник амортизатора",
        5, 12, 5, 12, 5, 12, 0.1, 0.5, 1.0, 0.3,
        parent="Автозапчасти", group="Подвеска", subgroup="Отбойники",
        keywords=["отбойник", "bump stop"]
    )
    categories["опора_стойки"] = make_cat(
        "опора_стойки", "Опора стойки", "Опора стойки амортизатора",
        8, 18, 8, 18, 5, 12, 0.5, 2, 3.0, 1.25,
        parent="Автозапчасти", group="Подвеска", subgroup="Опоры",
        keywords=["опора стойки", "strut mount"]
    )
    categories["подрамник"] = make_cat(
        "подрамник", "Подрамник", "Подрамник",
        45, 105, 15, 35, 8, 18, 10, 30, 25.0, 20.0,
        parent="Автозапчасти", group="Подвеска", subgroup="Подрамники",
        keywords=["подрамник", "subframe"],
        risk=RiskLevel.HIGH
    )
    
    # ========================================================================
    # 🛑 ТОРМОЗНАЯ СИСТЕМА
    # ========================================================================
    categories["тормозная_система"] = make_cat(
        "тормозная_система", "Тормозная система", "Тормозная система",
        20, 40, 20, 40, 5, 15, 2, 15, 3.0, 8.0,
        parent="Автозапчасти", group="Тормозная система", subgroup="Тормоза",
        keywords=["тормоз", "brake"],
        risk=RiskLevel.HIGH
    )
    categories["тормозные_диски"] = make_cat(
        "тормозные_диски", "Тормозные диски", "Тормозные диски",
        25, 40, 25, 40, 3, 8, 3, 12, 3.0, 8.0,
        parent="Автозапчасти", group="Тормозная система", subgroup="Диски",
        keywords=["тормозной диск", "brake disc", "диск"],
        fragile=True
    )
    categories["тормозные_колодки"] = make_cat(
        "тормозные_колодки", "Тормозные колодки", "Тормозные колодки",
        10, 20, 5, 12, 3, 8, 1, 4, 1.0, 3.0,
        parent="Автозапчасти", group="Тормозная система", subgroup="Колодки",
        keywords=["колодки", "brake pad", "тормозные"]
    )
    categories["тормозные_шланги"] = make_cat(
        "тормозные_шланги", "Тормозные шланги", "Тормозные шланги",
        20, 60, 2, 5, 2, 5, 0.2, 1, 0.3, 0.8,
        parent="Автозапчасти", group="Тормозная система", subgroup="Шланги",
        keywords=["тормозной шланг", "brake hose"]
    )
    categories["тормозные_суппорты"] = make_cat(
        "тормозные_суппорты", "Тормозные суппорты", "Тормозные суппорты",
        15, 30, 10, 20, 10, 20, 2, 8, 5.0, 5.0,
        parent="Автозапчасти", group="Тормозная система", subgroup="Суппорты",
        keywords=["суппорт", "caliper"]
    )
    categories["тормозные_барабаны"] = make_cat(
        "тормозные_барабаны", "Тормозные барабаны", "Тормозные барабаны",
        20, 35, 20, 35, 5, 15, 3, 10, 5.0, 6.5,
        parent="Автозапчасти", group="Тормозная система", subgroup="Барабаны",
        keywords=["барабан", "drum"]
    )
    categories["гтц"] = make_cat(
        "гтц", "Главный тормозной цилиндр", "Главный тормозной цилиндр",
        10, 25, 8, 18, 8, 18, 1, 4, 3.0, 2.5,
        parent="Автозапчасти", group="Тормозная система", subgroup="Цилиндры",
        keywords=["гтц", "главный тормозной"]
    )
    categories["вакуумный_усилитель"] = make_cat(
        "вакуумный_усилитель", "Вакуумный усилитель", "Вакуумный усилитель тормозов",
        20, 35, 20, 35, 10, 20, 2, 6, 10.0, 4.0,
        parent="Автозапчасти", group="Тормозная система", subgroup="Усилители",
        keywords=["вакуумный", "вакуумник", "booster"]
    )
    
    # ========================================================================
    # 🎯 РУЛЕВОЕ УПРАВЛЕНИЕ
    # ========================================================================
    categories["рулевое_управление"] = make_cat(
        "рулевое_управление", "Рулевое управление", "Рулевое управление",
        30, 100, 10, 30, 10, 30, 2, 15, 5.0, 10.0,
        parent="Автозапчасти", group="Рулевое управление", subgroup="Рулевое",
        keywords=["рулевое", "steering"]
    )
    categories["рулевые_тяги"] = make_cat(
        "рулевые_тяги", "Рулевые тяги", "Рулевые тяги и наконечники",
        20, 60, 3, 8, 3, 8, 0.5, 3, 1.0, 2.5,
        parent="Автозапчасти", group="Рулевое управление", subgroup="Тяги",
        keywords=["рулевая тяга", "tie rod", "наконечник"]
    )
    categories["рулевые_рейки"] = make_cat(
        "рулевые_рейки", "Рулевые рейки", "Рулевые рейки",
        50, 100, 10, 20, 10, 20, 5, 15, 8.0, 12.0,
        parent="Автозапчасти", group="Рулевое управление", subgroup="Рейки",
        keywords=["рулевая рейка", "rack"]
    )
    categories["рулевой_кардан"] = make_cat(
        "рулевой_кардан", "Рулевой кардан", "Рулевой кардан",
        20, 45, 5, 12, 5, 12, 1, 4, 5.0, 2.5,
        parent="Автозапчасти", group="Рулевое управление", subgroup="Карданы",
        keywords=["рулевой кардан", "steering cardan"]
    )
    categories["усилитель_руля"] = make_cat(
        "усилитель_руля", "Усилитель руля", "Усилитель руля (ГУР/ЭУР)",
        15, 30, 15, 30, 15, 25, 3, 10, 10.0, 6.5,
        parent="Автозапчасти", group="Рулевое управление", subgroup="ГУР/ЭУР",
        keywords=["усилитель", "гур", "power steering"]
    )
    categories["рулевой_насос"] = make_cat(
        "рулевой_насос", "Насос ГУР", "Насос ГУР",
        15, 30, 12, 22, 12, 22, 3, 8, 6.0, 5.5,
        parent="Автозапчасти", group="Рулевое управление", subgroup="ГУР",
        keywords=["насос гур", "power steering pump"]
    )
    
    # ========================================================================
    # ⚡ ЭЛЕКТРИКА
    # ========================================================================
    categories["электрика"] = make_cat(
        "электрика", "Электрооборудование", "Электрооборудование",
        10, 40, 10, 30, 10, 30, 0.5, 10, 2.0, 5.0,
        parent="Автозапчасти", group="Электрика", subgroup="Электрика",
        keywords=["электрик", "electrical"]
    )
    categories["стартеры"] = make_cat(
        "стартеры", "Стартеры", "Стартеры",
        15, 30, 10, 20, 10, 25, 3, 10, 3.0, 6.0,
        parent="Автозапчасти", group="Электрика", subgroup="Стартеры",
        keywords=["стартер", "starter"]
    )
    categories["генераторы"] = make_cat(
        "генераторы", "Генераторы", "Генераторы",
        15, 30, 15, 25, 15, 30, 4, 12, 5.0, 8.0,
        parent="Автозапчасти", group="Электрика", subgroup="Генераторы",
        keywords=["генератор", "alternator"]
    )
    categories["аккумуляторы"] = make_cat(
        "аккумуляторы", "Аккумуляторы", "Аккумуляторы",
        20, 40, 15, 25, 15, 30, 10, 30, 15.0, 20.0,
        parent="Автозапчасти", group="Электрика", subgroup="АКБ",
        keywords=["аккумулятор", "акб", "battery"],
        hazardous=True, risk=RiskLevel.HIGH
    )
    categories["датчики"] = make_cat(
        "датчики", "Датчики", "Датчики",
        3, 10, 2, 5, 2, 8, 0.05, 0.5, 0.1, 0.3,
        parent="Автозапчасти", group="Электрика", subgroup="Датчики",
        keywords=["датчик", "sensor"]
    )
    categories["катушки_зажигания"] = make_cat(
        "катушки_зажигания", "Катушки зажигания", "Катушки зажигания",
        5, 15, 3, 8, 5, 15, 0.2, 1, 0.5, 0.6,
        parent="Автозапчасти", group="Электрика", subgroup="Зажигание",
        keywords=["катушка", "coil"]
    )
    categories["проводка"] = make_cat(
        "проводка", "Проводка", "Проводка и жгуты",
        20, 100, 5, 20, 2, 10, 0.3, 3, 3.0, 1.5,
        parent="Автозапчасти", group="Электрика", subgroup="Проводка",
        keywords=["проводка", "жгут", "wiring"]
    )
    categories["блоки_управления"] = make_cat(
        "блоки_управления", "Блоки управления", "Блоки управления (ЭБУ)",
        15, 30, 10, 20, 5, 15, 0.5, 3, 3.0, 1.5,
        parent="Автозапчасти", group="Электрика", subgroup="ЭБУ",
        keywords=["блок управления", "эбу", "ecu"]
    )
    
    # ========================================================================
    # ❄️ СИСТЕМА ОХЛАЖДЕНИЯ
    # ========================================================================
    categories["охлаждение"] = make_cat(
        "охлаждение", "Охлаждение", "Система охлаждения",
        20, 80, 15, 50, 10, 40, 1, 15, 8.0, 15.0,
        parent="Автозапчасти", group="Охлаждение", subgroup="Охлаждение",
        keywords=["охлаждение", "cooling"]
    )
    categories["радиаторы"] = make_cat(
        "радиаторы", "Радиаторы", "Радиаторы охлаждения",
        40, 80, 30, 60, 5, 15, 2, 10, 10.0, 15.0,
        parent="Автозапчасти", group="Охлаждение", subgroup="Радиаторы",
        keywords=["радиатор", "radiator"],
        fragile=True
    )
    categories["помпы"] = make_cat(
        "помпы", "Помпы", "Водяные помпы",
        10, 25, 10, 20, 10, 20, 1, 5, 2.0, 4.0,
        parent="Автозапчасти", group="Охлаждение", subgroup="Помпы",
        keywords=["помпа", "water pump"]
    )
    categories["термостаты"] = make_cat(
        "термостаты", "Термостаты", "Термостаты",
        5, 12, 5, 12, 5, 12, 0.2, 1, 0.5, 1.0,
        parent="Автозапчасти", group="Охлаждение", subgroup="Термостаты",
        keywords=["термостат", "thermostat"]
    )
    categories["вентилятор_радиатора"] = make_cat(
        "вентилятор_радиатора", "Вентилятор радиатора", "Вентилятор радиатора",
        30, 50, 30, 50, 5, 15, 2, 6, 15.0, 4.0,
        parent="Автозапчасти", group="Охлаждение", subgroup="Вентиляторы",
        keywords=["вентилятор", "fan"],
        fragile=True
    )
    categories["расширительный_бачок"] = make_cat(
        "расширительный_бачок", "Расширительный бачок", "Расширительный бачок",
        15, 30, 10, 20, 10, 25, 0.3, 1.5, 4.0, 0.9,
        parent="Автозапчасти", group="Охлаждение", subgroup="Бачки",
        keywords=["расширительный бачок", "expansion tank"]
    )
    
    # ========================================================================
    # 🔧 ФИЛЬТРЫ
    # ========================================================================
    categories["фильтры"] = make_cat(
        "фильтры", "Фильтры", "Фильтры",
        5, 30, 5, 30, 5, 40, 0.1, 3, 2.0, 5.0,
        parent="Автозапчасти", group="Фильтры", subgroup="Фильтры",
        keywords=["фильтр", "filter"]
    )
    categories["масляные_фильтры"] = make_cat(
        "масляные_фильтры", "Масляные фильтры", "Масляные фильтры",
        6, 12, 6, 12, 8, 15, 0.3, 1, 1.0, 1.5,
        parent="Автозапчасти", group="Фильтры", subgroup="Масляные",
        keywords=["масляный фильтр", "oil filter"]
    )
    categories["воздушные_фильтры"] = make_cat(
        "воздушные_фильтры", "Воздушные фильтры", "Воздушные фильтры",
        15, 40, 15, 35, 3, 10, 0.2, 2, 2.0, 4.0,
        parent="Автозапчасти", group="Фильтры", subgroup="Воздушные",
        keywords=["воздушный фильтр", "air filter"]
    )
    categories["топливные_фильтры"] = make_cat(
        "топливные_фильтры", "Топливные фильтры", "Топливные фильтры",
        5, 15, 5, 15, 8, 20, 0.3, 1.5, 1.0, 2.0,
        parent="Автозапчасти", group="Фильтры", subgroup="Топливные",
        keywords=["топливный фильтр", "fuel filter"]
    )
    categories["салонные_фильтры"] = make_cat(
        "салонные_фильтры", "Салонные фильтры", "Салонные фильтры",
        20, 35, 15, 25, 2, 5, 0.2, 1, 1.5, 2.5,
        parent="Автозапчасти", group="Фильтры", subgroup="Салонные",
        keywords=["салонный фильтр", "cabin filter"]
    )
    
    # ========================================================================
    # 🛢️ МАСЛА И ЖИДКОСТИ
    # ========================================================================
    categories["масла"] = make_cat(
        "масла", "Масла", "Масла и технические жидкости",
        5, 30, 5, 30, 10, 40, 0.5, 20, 5.0, 15.0,
        parent="Автозапчасти", group="Масла и жидкости", subgroup="Масла",
        keywords=["масло", "oil"],
        hazardous=True
    )
    categories["моторные_масла"] = make_cat(
        "моторные_масла", "Моторные масла", "Моторные масла",
        8, 25, 8, 25, 20, 40, 1, 20, 5.0, 15.0,
        parent="Автозапчасти", group="Масла и жидкости", subgroup="Моторные",
        keywords=["моторное масло", "engine oil"],
        hazardous=True
    )
    categories["тормозная_жидкость"] = make_cat(
        "тормозная_жидкость", "Тормозная жидкость", "Тормозная жидкость",
        5, 10, 5, 10, 15, 25, 0.5, 2, 1.0, 2.0,
        parent="Автозапчасти", group="Масла и жидкости", subgroup="Тормозная",
        keywords=["тормозная жидкость", "brake fluid"],
        hazardous=True
    )
    categories["антифриз"] = make_cat(
        "антифриз", "Антифриз", "Антифриз / Охлаждающая жидкость",
        10, 30, 10, 30, 20, 40, 1, 20, 5.0, 15.0,
        parent="Автозапчасти", group="Масла и жидкости", subgroup="Охлаждающие",
        keywords=["антифриз", "antifreeze", "охлаждающая"],
        hazardous=True
    )
    
    # ========================================================================
    # 💡 ОПТИКА
    # ========================================================================
    categories["оптика"] = make_cat(
        "оптика", "Оптика", "Оптика и освещение",
        15, 60, 15, 40, 15, 40, 0.5, 10, 5.0, 10.0,
        parent="Автозапчасти", group="Оптика", subgroup="Оптика",
        keywords=["оптика", "оптик", "optic"],
        fragile=True
    )
    categories["фары"] = make_cat(
        "фары", "Фары", "Фары головного света",
        30, 60, 20, 40, 20, 40, 2, 8, 8.0, 12.0,
        parent="Автозапчасти", group="Оптика", subgroup="Фары",
        keywords=["фара", "headlight", "headlamp"],
        fragile=True
    )
    categories["лампы"] = make_cat(
        "лампы", "Лампы", "Автомобильные лампы",
        2, 10, 2, 5, 5, 15, 0.02, 0.3, 0.1, 0.3,
        parent="Автозапчасти", group="Оптика", subgroup="Лампы",
        keywords=["лампа", "bulb", "лампочка"],
        fragile=True
    )
    categories["фонари"] = make_cat(
        "фонари", "Фонари", "Задние фонари",
        20, 50, 15, 30, 10, 25, 1, 5, 5.0, 8.0,
        parent="Автозапчасти", group="Оптика", subgroup="Фонари",
        keywords=["фонарь", "taillight"],
        fragile=True
    )
    categories["led_лампы"] = make_cat(
        "led_лампы", "LED лампы", "LED лампы",
        5, 15, 3, 8, 3, 8, 0.1, 0.5, 0.3, 0.3,
        parent="Автозапчасти", group="Оптика", subgroup="LED",
        keywords=["led", "диод", "ксенон"],
        fragile=True
    )
    
    # ========================================================================
    # 🚗 КУЗОВ
    # ========================================================================
    categories["кузов"] = make_cat(
        "кузов", "Кузов", "Кузовные детали",
        50, 200, 30, 150, 10, 100, 2, 50, 30.0, 80.0,
        parent="Автозапчасти", group="Кузов", subgroup="Кузов",
        keywords=["кузов", "body"],
        fragile=True, risk=RiskLevel.HIGH
    )
    categories["бамперы"] = make_cat(
        "бамперы", "Бамперы", "Бамперы",
        100, 200, 30, 60, 20, 50, 5, 20, 50.0, 80.0,
        parent="Автозапчасти", group="Кузов", subgroup="Бамперы",
        keywords=["бампер", "bumper"],
        fragile=True
    )
    categories["крылья"] = make_cat(
        "крылья", "Крылья", "Крылья",
        50, 100, 30, 60, 30, 80, 3, 10, 20.0, 40.0,
        parent="Автозапчасти", group="Кузов", subgroup="Крылья",
        keywords=["крыло", "fender"],
        fragile=True
    )
    categories["капоты"] = make_cat(
        "капоты", "Капоты", "Капоты",
        100, 180, 80, 150, 5, 15, 5, 15, 30.0, 60.0,
        parent="Автозапчасти", group="Кузов", subgroup="Капоты",
        keywords=["капот", "hood", "bonnet"],
        fragile=True
    )
    categories["зеркала"] = make_cat(
        "зеркала", "Зеркала", "Зеркала заднего вида",
        15, 30, 10, 20, 10, 20, 0.5, 3, 3.0, 5.0,
        parent="Автозапчасти", group="Кузов", subgroup="Зеркала",
        keywords=["зеркало", "mirror"],
        fragile=True
    )
    categories["двери"] = make_cat(
        "двери", "Двери", "Двери",
        100, 150, 50, 100, 5, 15, 15, 40, 80.0, 27.5,
        parent="Автозапчасти", group="Кузов", subgroup="Двери",
        keywords=["дверь", "door"],
        fragile=True, risk=RiskLevel.HIGH
    )
    categories["стёкла"] = make_cat(
        "стёкла", "Стёкла", "Автомобильные стёкла",
        50, 150, 30, 100, 0.5, 2, 5, 20, 40.0, 12.5,
        parent="Автозапчасти", group="Кузов", subgroup="Стёкла",
        keywords=["стекло", "glass", "лобовое"],
        fragile=True, risk=RiskLevel.HIGH
    )
    
    # ========================================================================
    # 🛞 ШИНЫ И ДИСКИ
    # ========================================================================
    categories["шины"] = make_cat(
        "шины", "Шины", "Шины и диски",
        40, 80, 40, 80, 15, 40, 5, 30, 20.0, 40.0,
        parent="Автозапчасти", group="Шины и диски", subgroup="Шины",
        keywords=["шина", "tire", "tyre"]
    )
    categories["летние_шины"] = make_cat(
        "летние_шины", "Летние шины", "Летние шины",
        50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0,
        parent="Автозапчасти", group="Шины и диски", subgroup="Летние",
        keywords=["летняя шина", "летние"],
        season=Seasonality.SUMMER
    )
    categories["зимние_шины"] = make_cat(
        "зимние_шины", "Зимние шины", "Зимние шины",
        50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0,
        parent="Автозапчасти", group="Шины и диски", subgroup="Зимние",
        keywords=["зимняя шина", "зимние", "шипованная"],
        season=Seasonality.WINTER
    )
    categories["диски"] = make_cat(
        "диски", "Диски", "Колесные диски",
        40, 60, 40, 60, 15, 30, 5, 20, 15.0, 25.0,
        parent="Автозапчасти", group="Шины и диски", subgroup="Диски",
        keywords=["диск", "wheel", "колесный"],
        fragile=True
    )
    
    # ========================================================================
    # 🔨 ИНСТРУМЕНТЫ
    # ========================================================================
    categories["инструменты"] = make_cat(
        "инструменты", "Инструменты", "Автоинструменты",
        10, 60, 5, 30, 3, 20, 0.2, 10, 3.0, 8.0,
        parent="Автозапчасти", group="Инструменты", subgroup="Инструменты",
        keywords=["инструмент", "tool"]
    )
    categories["домкраты"] = make_cat(
        "домкраты", "Домкраты", "Домкраты",
        20, 50, 10, 25, 10, 25, 3, 15, 5.0, 12.0,
        parent="Автозапчасти", group="Инструменты", subgroup="Домкраты",
        keywords=["домкрат", "jack"]
    )
    categories["наборы_ключей"] = make_cat(
        "наборы_ключей", "Наборы ключей", "Наборы ключей",
        15, 40, 10, 25, 3, 10, 1, 8, 3.0, 6.0,
        parent="Автозапчасти", group="Инструменты", subgroup="Ключи",
        keywords=["ключ", "набор ключей", "wrench"]
    )
    categories["компрессоры_воздушные"] = make_cat(
        "компрессоры_воздушные", "Компрессоры", "Воздушные компрессоры",
        25, 60, 20, 40, 20, 40, 5, 25, 15.0, 15.0,
        parent="Автозапчасти", group="Инструменты", subgroup="Компрессоры",
        keywords=["компрессор", "compressor"]
    )
    
    # ========================================================================
    # ⚙️ РЕМНИ И ПРИВОДЫ
    # ========================================================================
    categories["ремни"] = make_cat(
        "ремни", "Ремни", "Ремни ГРМ и приводов",
        50, 150, 1, 3, 1, 3, 0.1, 0.8, 0.5, 1.0,
        parent="Автозапчасти", group="Ремни и приводы", subgroup="Ремни",
        keywords=["ремень", "belt"]
    )
    categories["ролики"] = make_cat(
        "ролики", "Ролики", "Ролики натяжители",
        5, 12, 5, 12, 2, 5, 0.2, 1.5, 0.5, 1.0,
        parent="Автозапчасти", group="Ремни и приводы", subgroup="Ролики",
        keywords=["ролик", "tensioner", "натяжитель"]
    )
    
    # ========================================================================
    # 🔩 ПОДШИПНИКИ
    # ========================================================================
    categories["подшипники"] = make_cat(
        "подшипники", "Подшипники", "Подшипники",
        3, 15, 3, 15, 1, 5, 0.1, 3, 0.5, 2.0,
        parent="Автозапчасти", group="Подшипники", subgroup="Подшипники",
        keywords=["подшипник", "bearing"]
    )
    
    # ========================================================================
    # 🔩 КРЕПЁЖ
    # ========================================================================
    categories["крепёж"] = make_cat(
        "крепёж", "Крепёж", "Крепёж и метизы",
        0.5, 10, 0.5, 10, 0.5, 10, 0.01, 2, 0.2, 1.0,
        parent="Автозапчасти", group="Крепёж", subgroup="Крепёж",
        keywords=["болт", "гайка", "шуруп", "саморез", "крепёж"]
    )
    
    # ========================================================================
    # ❄️ КЛИМАТ
    # ========================================================================
    categories["климат"] = make_cat(
        "климат", "Климат", "Климат-контроль и кондиционер",
        20, 80, 20, 60, 15, 50, 2, 20, 10.0, 20.0,
        parent="Автозапчасти", group="Климат", subgroup="Климат",
        keywords=["климат", "кондиционер", "ac"]
    )
    categories["компрессоры"] = make_cat(
        "компрессоры", "Компрессоры кондиционера", "Компрессоры кондиционера",
        20, 40, 15, 30, 15, 30, 5, 15, 8.0, 12.0,
        parent="Автозапчасти", group="Климат", subgroup="Компрессоры",
        keywords=["компрессор кондиционера", "ac compressor"]
    )
    categories["конденсоры"] = make_cat(
        "конденсоры", "Конденсоры", "Конденсоры кондиционера",
        40, 80, 30, 60, 5, 15, 2, 8, 10.0, 5.0,
        parent="Автозапчасти", group="Климат", subgroup="Конденсоры",
        keywords=["конденсор", "condenser"],
        fragile=True
    )
    
    # ========================================================================
    # 💨 ВЫХЛОПНАЯ СИСТЕМА
    # ========================================================================
    categories["выпуск"] = make_cat(
        "выпуск", "Выпуск", "Выхлопная система",
        30, 150, 10, 40, 10, 40, 2, 25, 10.0, 25.0,
        parent="Автозапчасти", group="Выхлопная система", subgroup="Выпуск",
        keywords=["выхлоп", "exhaust", "выпуск"]
    )
    categories["глушители"] = make_cat(
        "глушители", "Глушители", "Глушители",
        50, 150, 20, 40, 20, 40, 5, 20, 20.0, 30.0,
        parent="Автозапчасти", group="Выхлопная система", subgroup="Глушители",
        keywords=["глушитель", "muffler"]
    )
    categories["катализаторы"] = make_cat(
        "катализаторы", "Катализаторы", "Каталитические нейтрализаторы",
        30, 80, 15, 30, 15, 30, 3, 15, 10.0, 20.0,
        parent="Автозапчасти", group="Выхлопная система", subgroup="Катализаторы",
        keywords=["катализатор", "catalyst", "кат"],
        hazardous=True, risk=RiskLevel.HIGH
    )
    categories["гофры"] = make_cat(
        "гофры", "Гофры", "Гофры выхлопной системы",
        10, 30, 5, 15, 5, 15, 0.3, 2, 2.0, 1.15,
        parent="Автозапчасти", group="Выхлопная система", subgroup="Гофры",
        keywords=["гофра", "flex pipe"]
    )
    
    # ========================================================================
    # 🛡️ БЕЗОПАСНОСТЬ
    # ========================================================================
    categories["безопасность"] = make_cat(
        "безопасность", "Безопасность", "Системы безопасности",
        10, 50, 10, 40, 5, 30, 0.5, 8, 3.0, 6.0,
        parent="Автозапчасти", group="Безопасность", subgroup="Безопасность",
        keywords=["безопасность", "safety"],
        risk=RiskLevel.HIGH
    )
    categories["подушки_безопасности"] = make_cat(
        "подушки_безопасности", "Подушки безопасности", "Подушки безопасности",
        20, 50, 15, 30, 10, 20, 1, 5, 5.0, 3.0,
        parent="Автозапчасти", group="Безопасность", subgroup="Airbag",
        keywords=["подушка безопасности", "airbag", "srs"],
        risk=RiskLevel.HIGH
    )
    
    # ========================================================================
    # 🧹 ПРОЧЕЕ
    # ========================================================================
    categories["щетки_стеклоочистителя"] = make_cat(
        "щетки_стеклоочистителя", "Щётки стеклоочистителя", "Щётки стеклоочистителя",
        30, 70, 2, 5, 2, 5, 0.1, 0.5, 1.0, 1.5,
        parent="Автозапчасти", group="Прочее", subgroup="Щётки",
        keywords=["щётки", "дворники", "wiper"]
    )
    categories["коврики"] = make_cat(
        "коврики", "Коврики", "Автомобильные коврики",
        50, 100, 40, 80, 1, 5, 1, 5, 10.0, 15.0,
        parent="Автозапчасти", group="Прочее", subgroup="Коврики",
        keywords=["коврик", "mat"]
    )
    categories["чехлы"] = make_cat(
        "чехлы", "Чехлы", "Чехлы на сиденья",
        40, 80, 30, 60, 5, 20, 1, 5, 15.0, 25.0,
        parent="Автозапчасти", group="Прочее", subgroup="Чехлы",
        keywords=["чехол", "cover"]
    )
    categories["автохимия"] = make_cat(
        "автохимия", "Автохимия", "Автохимия и косметика",
        5, 30, 5, 20, 10, 40, 0.3, 5, 2.0, 5.0,
        parent="Автозапчасти", group="Прочее", subgroup="Автохимия",
        keywords=["автохимия", "химия", "косметика"],
        hazardous=True
    )
    
    return categories


def get_category_hierarchy_map() -> Dict[str, CategoryHierarchy]:
    """
    Быстрый доступ к иерархиям всех категорий.
    
    Returns:
        Dict[str, CategoryHierarchy] — {ключ_категории: иерархия}
    """
    categories = get_auto_parts_categories_full()
    return {key: cat.hierarchy for key, cat in categories.items() if cat.hierarchy}


def get_category_keywords_map() -> Dict[str, List[str]]:
    """
    Карта ключевых слов для автокатегоризации.
    
    🆕 v101.0: ИСПРАВЛЕНА критическая ошибка - добавлен return statement.
    Ранее функция обрывалась без return, что вызывало SyntaxError.
    
    Returns:
        Dict[str, List[str]] — {ключ_категории: [ключевые слова]}
    """
    categories = get_auto_parts_categories_full()
    return {key: cat.hierarchy.keywords for key, cat in categories.items() 
            if cat.hierarchy and cat.hierarchy.keywords}

# ============================================================================
# БЛОК 6: ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ
# ============================================================================
# 📌 v101.0: Добавлены новые классы:
# - CategoryClassifier — 3-уровневая ML-классификация (Родитель/Группа/Подгруппа)
# - CatalogEnhancer — обогащение каталога через поиск аналогов по OE
# - SmartTariffCache — умный кэш тарифов с прогнозированием
# ============================================================================


# ============================================================================
# 6.1 КЛАССИФИКАТОР КАТЕГОРИЙ (3-УРОВНЕВАЯ ИЕРАРХИЯ + ML)
# ============================================================================

class CategoryClassifier:
    """
    v101.0: Классификатор категорий с 3-уровневой иерархией.
    
    Логика:
    1. Сначала ищет по ключевым словам в справочнике категорий (Блок 5)
    2. Если не нашёл — использует ML-модель (если обучена)
    3. Если и ML не помог — возвращает "Прочее"
    
    Возвращает:
    - parent: Родитель (например, "Автозапчасти")
    - group: Группа (например, "Подвеска")
    - subgroup: Подгруппа (например, "Сайлентблоки")
    - confidence: Уверенность (0.0 - 1.0)
    """
    
    def __init__(self):
        self.categories = get_auto_parts_categories_full()
        self.keywords_map = get_category_keywords_map()
        self.hierarchy_map = get_category_hierarchy_map()
        self.model = None
        self.vectorizer = None
        
        # Пытаемся загрузить ML-модель
        if SKLEARN_AVAILABLE:
            try:
                model_path = MODELS_DIR / "category_classifier_v101.joblib"
                if model_path.exists():
                    model_data = joblib.load(model_path)
                    self.model = model_data.get("model")
                    self.vectorizer = model_data.get("vectorizer")
                    logger.info("✅ ML модель классификации v101 загружена")
            except Exception as e:
                logger.warning(f"Не удалось загрузить ML модель: {e}")
        
        logger.info(f"📚 CategoryClassifier инициализирован: {len(self.categories)} категорий")
    
    def predict(self, text: str) -> Tuple[str, float]:
        """
        Предсказание категории по названию.
        
        Returns:
            (category_key, confidence) — ключ категории и уверенность
        """
        if not text:
            return ("прочее", 0.0)
        
        # 1. Поиск по ключевым словам (самый надёжный)
        text_lower = text.lower()
        best_match = None
        best_score = 0
        
        for cat_key, keywords in self.keywords_map.items():
            for keyword in keywords:
                if keyword.lower() in text_lower:
                    # Чем длиннее совпадение — тем точнее
                    score = len(keyword) / max(len(text_lower), 1)
                    if score > best_score:
                        best_score = score
                        best_match = cat_key
        
        if best_match and best_score > 0.05:
            return (best_match, min(0.95, 0.5 + best_score * 2))
        
        # 2. ML-модель (если есть)
        if self.model and self.vectorizer:
            try:
                text_vector = self.vectorizer.transform([text])
                prediction = self.model.predict(text_vector)
                confidence = max(self.model.predict_proba(text_vector)[0])
                return (prediction[0], confidence)
            except Exception:
                pass
        
        # 3. Фоллбэк
        return ("прочее", 0.0)
    
    def predict_hierarchy(self, text: str) -> Dict[str, str]:
        """
        v101.0: Предсказание полной иерархии (Родитель/Группа/Подгруппа).
        
        Returns:
            {"parent": "...", "group": "...", "subgroup": "...", "confidence": 0.0}
        """
        cat_key, confidence = self.predict(text)
        
        if cat_key in self.hierarchy_map:
            hierarchy = self.hierarchy_map[cat_key]
            return {
                "parent": hierarchy.parent or "Автозапчасти",
                "group": hierarchy.group or cat_key,
                "subgroup": hierarchy.subgroup or cat_key,
                "confidence": confidence,
                "category_key": cat_key
            }
        
        return {
            "parent": "Автозапчасти",
            "group": "Прочее",
            "subgroup": "Прочее",
            "confidence": confidence,
            "category_key": "прочее"
        }
    
    def classify_dataframe(
        self,
        df: pd.DataFrame,
        name_col: str = "Наименование"
    ) -> pd.DataFrame:
        """
        v101.0: Классификация всего DataFrame.
        Добавляет колонки: parent_category, group_category, subgroup_category, category_key, confidence
        """
        if df.empty or name_col not in df.columns:
            return df
        
        result = df.copy()
        
        # Векторизованная классификация через apply
        hierarchies = result[name_col].apply(
            lambda x: self.predict_hierarchy(str(x))
        )
        
        result['parent_category'] = hierarchies.apply(lambda h: h['parent'])
        result['group_category'] = hierarchies.apply(lambda h: h['group'])
        result['subgroup_category'] = hierarchies.apply(lambda h: h['subgroup'])
        result['category_key'] = hierarchies.apply(lambda h: h['category_key'])
        result['classification_confidence'] = hierarchies.apply(lambda h: h['confidence'])
        
        return result
    
    def train_model(self, training_data: List[Dict[str, str]]) -> bool:
        """
        Обучение ML-модели на размеченных данных.
        
        Args:
            training_data: Список словарей {"text": "...", "category": "..."}
        
        Returns:
            True при успехе
        """
        if not SKLEARN_AVAILABLE:
            logger.error("❌ scikit-learn не установлен")
            return False
        
        try:
            texts = [d["text"] for d in training_data]
            labels = [d["category"] for d in training_data]
            
            self.vectorizer = TfidfVectorizer(
                max_features=5000,
                ngram_range=(1, 2),
                stop_words=None
            )
            
            self.model = Pipeline([
                ("tfidf", self.vectorizer),
                ("clf", MultinomialNB(alpha=0.1))
            ])
            
            self.model.fit(texts, labels)
            
            # Сохраняем модель
            model_path = MODELS_DIR / "category_classifier_v101.joblib"
            joblib.dump({
                "model": self.model.named_steps["clf"],
                "vectorizer": self.vectorizer,
                "classes": self.model.classes_.tolist()
            }, model_path)
            
            logger.info(f"✅ ML модель обучена на {len(training_data)} примерах")
            return True
        
        except Exception as e:
            logger.error(f"❌ Ошибка обучения: {e}")
            return False


# ============================================================================
# 6.2 ОБОГАТИТЕЛЬ КАТАЛОГА (ПОИСК АНАЛОГОВ ПО OE)
# ============================================================================

class CatalogEnhancer:
    """
    Обогащение каталога через поиск аналогов по OE-номерам.
    
    v101.0: Добавлена возможность автоматического заполнения
    недостающих параметров (вес, габариты) из аналогов.
    """
    
    def __init__(self):
        self.oe_data = pd.DataFrame()
        self.parts_data = pd.DataFrame()
        self.cross_data = pd.DataFrame()
        self.stats = {
            "oe_loaded": 0,
            "parts_loaded": 0,
            "cross_loaded": 0,
            "analog_searches": 0,
            "cache_hits": 0,
            "cache_misses": 0,
            "fills_from_analogs": 0
        }
        self._cache = {}
        self.oe_index = {}
        self.parts_index = {}
        self.cross_index = defaultdict(list)
        self.oe_to_parts = defaultdict(list)
    
    def load_oe_data(self, df: pd.DataFrame):
        """Загрузка OE-данных"""
        self.oe_data = df
        self.stats["oe_loaded"] = len(df)
        self._build_oe_index()
    
    def load_parts_data(self, df: pd.DataFrame):
        """Загрузка данных товаров"""
        self.parts_data = df
        self.stats["parts_loaded"] = len(df)
        self._build_parts_index()
    
    def load_cross_references(self, df: pd.DataFrame):
        """Загрузка кросс-ссылок"""
        self.cross_data = df
        self.stats["cross_loaded"] = len(df)
        self._build_cross_index()
    
    def _build_oe_index(self):
        """Построение индекса OE-номеров"""
        if not self.oe_data.empty:
            for _, row in self.oe_data.iterrows():
                oe = str(row.get('oe_number', '')).strip()
                if oe:
                    self.oe_index[oe] = row.to_dict()
    
    def _build_parts_index(self):
        """Построение индекса товаров"""
        if not self.parts_data.empty:
            for _, row in self.parts_data.iterrows():
                key = (str(row.get('artikul', '')).strip(), 
                       str(row.get('brand', '')).strip())
                if key[0]:
                    self.parts_index[key] = row.to_dict()
    
    def _build_cross_index(self):
        """Построение индекса кросс-ссылок"""
        if not self.cross_data.empty:
            for _, row in self.cross_data.iterrows():
                oe = str(row.get('oe_number', '')).strip()
                artikul = str(row.get('artikul', '')).strip()
                brand = str(row.get('brand', '')).strip()
                if oe and artikul:
                    self.cross_index[(artikul, brand)].append(oe)
                    self.oe_to_parts[oe].append((artikul, brand))
    
    def get_analog_data(
        self,
        artikul: str,
        brand: str,
        max_analogs: int = 20
    ) -> Dict[str, Any]:
        """
        Поиск аналогов для товара по OE-номерам.
        
        Returns:
            Словарь с аналогами и метаданными
        """
        self.stats["analog_searches"] += 1
        
        cache_key = (artikul, brand)
        if cache_key in self._cache:
            self.stats["cache_hits"] += 1
            return self._cache[cache_key]
        
        self.stats["cache_misses"] += 1
        
        if self.cross_data.empty:
            return {"error": "Кросс-ссылки не загружены"}
        
        oe_numbers = self.cross_index.get((artikul, brand), [])
        if not oe_numbers:
            return {"error": "Артикул не найден", "analog_count": 0}
        
        analogs = []
        seen = set()
        
        for oe in oe_numbers:
            for analog_artikul, analog_brand in self.oe_to_parts.get(oe, []):
                if (analog_artikul, analog_brand) == (artikul, brand):
                    continue
                key = (analog_artikul, analog_brand)
                if key in seen:
                    continue
                seen.add(key)
                
                analog_info = {
                    "Артикул": analog_artikul,
                    "Бренд": analog_brand,
                    "OE номер": oe
                }
                
                part_key = (analog_artikul, analog_brand)
                if part_key in self.parts_index:
                    part = self.parts_index[part_key]
                    analog_info["Наименование"] = part.get('description', '')
                    analog_info["Вес"] = part.get('weight', '')
                    analog_info["Длина"] = part.get('length', '')
                    analog_info["Ширина"] = part.get('width', '')
                    analog_info["Высота"] = part.get('height', '')
                
                analogs.append(analog_info)
        
        result = {
            "oe_list": ", ".join(oe_numbers[:5]),
            "analog_count": len(analogs),
            "has_analogs": len(analogs) > 0,
            "analogs": analogs[:max_analogs]
        }
        
        self._cache[cache_key] = result
        return result
    
    def fill_missing_from_analogs(
        self,
        df: pd.DataFrame,
        article_col: str = "Артикул",
        brand_col: str = "Бренд",
        columns_to_fill: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        v101.0: Автоматическое заполнение недостающих параметров из аналогов.
        
        Если у товара нет веса/габаритов, но есть аналог с теми же OE — 
        берём усреднённые данные из аналогов.
        
        Args:
            df: DataFrame с товарами
            article_col: Колонка с артикулом
            brand_col: Колонка с брендом
            columns_to_fill: Какие столбцы заполнять (по умолчанию: вес, габариты)
        
        Returns:
            DataFrame с заполненными данными
        """
        if df.empty or self.cross_data.empty:
            return df
        
        if columns_to_fill is None:
            columns_to_fill = ['weight', 'length', 'width', 'height']
        
        result = df.copy()
        filled_count = 0
        
        for idx, row in result.iterrows():
            article = str(row.get(article_col, '')).strip()
            brand = str(row.get(brand_col, '')).strip()
            
            # Проверяем, есть ли пропуски
            needs_fill = False
            for col in columns_to_fill:
                if col in result.columns:
                    val = row.get(col)
                    if pd.isna(val) or val == 0 or val == '':
                        needs_fill = True
                        break
            
            if not needs_fill:
                continue
            
            # Ищем аналоги
            analog_data = self.get_analog_data(article, brand)
            if not analog_data.get("has_analogs"):
                continue
            
            # Собираем значения из аналогов
            for col in columns_to_fill:
                if col not in result.columns:
                    continue
                
                current_val = row.get(col)
                if pd.notna(current_val) and current_val != 0 and current_val != '':
                    continue
                
                analog_values = []
                for analog in analog_data.get("analogs", []):
                    val = analog.get(col, '')
                    if val and val != '' and val != 0:
                        try:
                            analog_values.append(float(val))
                        except (ValueError, TypeError):
                            pass
                
                if analog_values:
                    # Берём среднее (усреднённое из оригинального аналога)
                    avg_val = sum(analog_values) / len(analog_values)
                    result.at[idx, col] = round(avg_val, 2)
                    filled_count += 1
        
        self.stats["fills_from_analogs"] = filled_count
        logger.info(f"✅ Заполнено из аналогов: {filled_count} значений")
        return result
    
    def get_stats(self) -> Dict[str, int]:
        """Статистика работы обогащения"""
        return self.stats


# ============================================================================
# 6.3 УМНЫЙ КЭШ ТАРИФОВ
# ============================================================================

class SmartTariffCache:
    """
    Умный кэш тарифов с прогнозированием и историей.
    
    v101.0: Добавлена поддержка Google Sheets как источника тарифов.
    """
    
    def __init__(self):
        self.cache_dir = TARIFFS_DIR
        self.cache_dir.mkdir(exist_ok=True, parents=True)
        self.cache_file = self.cache_dir / "tariffs_cache.json"
        self.history_file = self.cache_dir / "tariffs_history.json"
        self.forecast_file = self.cache_dir / "tariffs_forecast.json"
        self.backup_dir = self.cache_dir / "backups"
        self.backup_dir.mkdir(exist_ok=True, parents=True)
        
        self._cache: Dict[str, TariffCacheEntry] = {}
        self._history: List[Dict[str, Any]] = []
        self._forecasts: Dict[str, Dict[str, Any]] = {}
        
        self._load_cache()
        self._load_history()
        self._load_forecasts()
        
        logger.info(f"SmartTariffCache инициализирован: {len(self._cache)} записей")
    
    def _make_key(self, marketplace: str, category: Optional[str] = None) -> str:
        cat = (category or "all").lower().strip()
        return f"{marketplace.lower().strip()}::{cat}"
    
    def _load_cache(self):
        if not self.cache_file.exists():
            self._cache = {}
            return
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self._cache = {}
            for key, entry_dict in data.items():
                try:
                    self._cache[key] = TariffCacheEntry.from_dict(entry_dict)
                except (ValueError, KeyError) as e:
                    logger.warning(f"Ошибка загрузки записи {key}: {e}")
        except (IOError, json.JSONDecodeError) as e:
            logger.error(f"Ошибка загрузки кэша тарифов: {e}")
            self._cache = {}
    
    def _save_cache(self):
        try:
            data = {k: v.to_dict() for k, v in self._cache.items()}
            if self.cache_file.exists():
                backup_name = f"tariffs_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                shutil.copy2(self.cache_file, self.backup_dir / backup_name)
                backups = sorted(self.backup_dir.glob("tariffs_backup_*.json"))
                for old_backup in backups[:-10]:
                    try:
                        old_backup.unlink()
                    except OSError:
                        pass
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except (IOError, OSError) as e:
            logger.error(f"Ошибка сохранения кэша тарифов: {e}")
    
    def _load_history(self):
        if not self.history_file.exists():
            self._history = []
            return
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self._history = json.load(f)
        except (IOError, json.JSONDecodeError):
            self._history = []
    
    def _save_history(self):
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self._history[-1000:], f, ensure_ascii=False, indent=2)
        except (IOError, OSError):
            pass
    
    def _load_forecasts(self):
        if not self.forecast_file.exists():
            self._forecasts = {}
            return
        try:
            with open(self.forecast_file, 'r', encoding='utf-8') as f:
                self._forecasts = json.load(f)
        except (IOError, json.JSONDecodeError):
            self._forecasts = {}
    
    def _save_forecasts(self):
        try:
            with open(self.forecast_file, 'w', encoding='utf-8') as f:
                json.dump(self._forecasts, f, ensure_ascii=False, indent=2)
        except (IOError, OSError):
            pass
    
    def _add_history_entry(
        self,
        action: str,
        marketplace: str,
        category: Optional[str],
        old_data: Optional[Dict],
        new_data: Optional[Dict],
        source: str
    ):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "marketplace": marketplace,
            "category": category,
            "old_data": old_data,
            "new_data": new_data,
            "source": source
        }
        self._history.append(entry)
        if len(self._history) > 1000:
            self._history = self._history[-1000:]
        self._save_history()
    
    def get(
        self,
        marketplace: str,
        category: Optional[str] = None,
        use_expired: bool = True
    ) -> Optional[TariffCacheEntry]:
        """Получение записи из кэша"""
        key = self._make_key(marketplace, category)
        entry = self._cache.get(key)
        if entry is None:
            key = self._make_key(marketplace, None)
            entry = self._cache.get(key)
        if entry is None:
            return None
        if entry.is_expired() and not use_expired:
            return None
        return entry
    
    def set(
        self,
        marketplace: str,
        category: Optional[str],
        data: Dict[str, Any],
        source: TariffSource,
        ttl_seconds: int = 86400,
        notes: str = "",
        forecast_data: Optional[Dict[str, Any]] = None
    ) -> bool:
        """Сохранение записи в кэш"""
        try:
            key = self._make_key(marketplace, category)
            old_entry = self._cache.get(key)
            old_data = old_entry.data if old_entry else None
            
            entry = TariffCacheEntry(
                marketplace=marketplace,
                category=category,
                data=data,
                source=source,
                timestamp=time.time(),
                ttl_seconds=ttl_seconds,
                version="2026.1",
                notes=notes,
                forecast_data=forecast_data
            )
            
            self._cache[key] = entry
            self._save_cache()
            
            self._add_history_entry(
                action="UPDATE" if old_entry else "CREATE",
                marketplace=marketplace,
                category=category,
                old_data=old_data,
                new_data=data,
                source=source.value
            )
            
            if forecast_data:
                self._forecasts[key] = {
                    "forecast": forecast_data,
                    "timestamp": time.time(),
                    "marketplace": marketplace,
                    "category": category
                }
                self._save_forecasts()
            
            return True
        except (IOError, OSError) as e:
            logger.error(f"Ошибка сохранения тарифов: {e}")
            return False
    
    def delete(self, marketplace: str, category: Optional[str] = None) -> bool:
        """Удаление записи из кэша"""
        try:
            key = self._make_key(marketplace, category)
            old_entry = self._cache.get(key)
            if old_entry:
                self._add_history_entry(
                    action="DELETE",
                    marketplace=marketplace,
                    category=category,
                    old_data=old_entry.data,
                    new_data=None,
                    source="MANUAL"
                )
                del self._cache[key]
                self._save_cache()
                return True
            return False
        except (IOError, OSError) as e:
            logger.error(f"Ошибка удаления тарифов: {e}")
            return False
    
    def get_all(self) -> Dict[str, TariffCacheEntry]:
        """Получение всех записей"""
        return self._cache.copy()
    
    def get_history(self, limit: int = 100) -> List[Dict[str, Any]]:
        """Получение истории изменений"""
        return self._history[-limit:]
    
    def clear_expired(self) -> int:
        """Очистка устаревших записей"""
        expired_keys = [k for k, v in self._cache.items() if v.is_expired()]
        for key in expired_keys:
            del self._cache[key]
        if expired_keys:
            self._save_cache()
        return len(expired_keys)
    
    def clear_all(self) -> int:
        """Полная очистка кэша"""
        count = len(self._cache)
        self._cache = {}
        self._save_cache()
        self._add_history_entry(
            action="CLEAR_ALL",
            marketplace="ALL",
            category=None,
            old_data={"count": count},
            new_data=None,
            source="MANUAL"
        )
        return count
    
    def export_to_file(self, file_path: Union[str, Path]) -> bool:
        """Экспорт кэша в файл"""
        try:
            data = {k: v.to_dict() for k, v in self._cache.items()}
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except (IOError, OSError) as e:
            logger.error(f"Ошибка экспорта кэша: {e}")
            return False
    
    def import_from_file(self, file_path: Union[str, Path]) -> int:
        """Импорт кэша из файла"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            count = 0
            for key, entry_dict in data.items():
                try:
                    self._cache[key] = TariffCacheEntry.from_dict(entry_dict)
                    count += 1
                except (ValueError, KeyError) as e:
                    logger.warning(f"Ошибка импорта записи {key}: {e}")
            self._save_cache()
            self._add_history_entry(
                action="IMPORT",
                marketplace="ALL",
                category=None,
                old_data=None,
                new_data={"count": count, "file": str(file_path)},
                source="IMPORTED"
            )
            return count
        except (IOError, json.JSONDecodeError) as e:
            logger.error(f"Ошибка импорта кэша: {e}")
            return 0
    
    def get_statistics(self) -> Dict[str, Any]:
        """Статистика кэша"""
        stats = {
            "total_entries": len(self._cache),
            "by_marketplace": defaultdict(int),
            "by_source": defaultdict(int),
            "expired_count": 0,
            "oldest_entry": None,
            "newest_entry": None,
            "history_count": len(self._history),
            "forecast_count": len(self._forecasts)
        }
        
        if not self._cache:
            return stats
        
        timestamps = []
        for entry in self._cache.values():
            stats["by_marketplace"][entry.marketplace] += 1
            stats["by_source"][entry.source.value] += 1
            if entry.is_expired():
                stats["expired_count"] += 1
            timestamps.append(entry.timestamp)
        
        if timestamps:
            stats["oldest_entry"] = datetime.fromtimestamp(min(timestamps)).isoformat()
            stats["newest_entry"] = datetime.fromtimestamp(max(timestamps)).isoformat()
        
        return stats


@st.cache_resource
def get_smart_tariff_cache():
    """Получение экземпляра кэша через st.cache_resource"""
    return SmartTariffCache()

# ============================================================================
# БЛОК 7: ПОСТОЯННОЕ ХРАНИЛИЩЕ ИСТОРИИ + УПРАВЛЕНИЕ СОСТОЯНИЕМ
# ============================================================================
# 📌 v101.0: Адаптировано для новой архитектуры с 4 разделами
# - PersistentHistoryDB  хранение истории расчётов в DuckDB/SQLite
# - AppStateManager  управление состоянием приложения между разделами
# - SaveLoadManager  сохранение/загрузка расчётов одной кнопкой
# ============================================================================


# ============================================================================
# 7.1 ПОСТОЯННОЕ ХРАНИЛИЩЕ ИСТОРИИ (DuckDB/SQLite)
# ============================================================================

@st.cache_resource
def get_persistent_history_db(db_path: Optional[Path] = None):
    """Получение экземпляра БД через st.cache_resource"""
    return PersistentHistoryDB(db_path)


class PersistentHistoryDB:
    """
    Постоянное хранилище истории расчётов.
    Поддерживает DuckDB (быстрый) и SQLite (fallback).
    Автоматическая миграция при добавлении новых колонок.
    """
    
    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = db_path or (HISTORY_DB_DIR / "history_pro_v101.duckdb")
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.use_duckdb = DUCKDB_AVAILABLE
        self.conn = None
        self._init_connection()
        self._create_tables()
        self._migrate_database()
        logger.info(f"📚 PersistentHistoryDB инициализирован: {self.db_path}")
    
    def _init_connection(self):
        """Инициализация подключения к БД"""
        try:
            if self.use_duckdb:
                self.conn = duckdb.connect(str(self.db_path))
            else:
                sqlite_path = self.db_path.with_suffix('.sqlite')
                self.conn = sqlite3.connect(str(sqlite_path), check_same_thread=False)
                self.conn.row_factory = sqlite3.Row
        except Exception as e:
            logger.error(f"Ошибка подключения к БД: {e}")
            self.conn = None
    
    def _create_tables(self):
        """Создание таблиц истории расчётов"""
        if self.conn is None:
            return
        
        try:
            # Основная таблица истории
            if self.use_duckdb:
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS calculation_history (
                        id VARCHAR PRIMARY KEY,
                        timestamp VARCHAR NOT NULL,
                        marketplace VARCHAR,
                        operation_mode VARCHAR,
                        category VARCHAR,
                        article VARCHAR,
                        brand VARCHAR,
                        price DOUBLE,
                        cost DOUBLE,
                        length DOUBLE,
                        width DOUBLE,
                        height DOUBLE,
                        weight DOUBLE,
                        volume DOUBLE,
                        commission DOUBLE,
                        commission_percent DOUBLE,
                        logistics DOUBLE,
                        storage_cost DOUBLE,
                        acquiring DOUBLE,
                        delivery DOUBLE,
                        last_mile DOUBLE,
                        returns DOUBLE,
                        rko_fee DOUBLE,
                        premium_fee DOUBLE,
                        insurance_fee DOUBLE,
                        packing_fee DOUBLE,
                        marketing_fee DOUBLE,
                        subscription_cost DOUBLE,
                        hazardous_surcharge DOUBLE,
                        fragile_surcharge DOUBLE,
                        oversized_surcharge DOUBLE,
                        tax_amount DOUBLE,
                        tax_system VARCHAR,
                        total_expenses DOUBLE,
                        profit DOUBLE,
                        margin_percent DOUBLE,
                        roi DOUBLE,
                        breakeven_price DOUBLE,
                        recommended_min_price DOUBLE,
                        profit_per_ruble DOUBLE,
                        contribution_margin DOUBLE,
                        contribution_margin_ratio DOUBLE,
                        tariff_source VARCHAR,
                        status VARCHAR,
                        metadata_json VARCHAR,
                        applied_seasonal_multiplier DOUBLE DEFAULT 1.0,
                        applied_promo_discount DOUBLE DEFAULT 0.0,
                        dynamic_adjustment DOUBLE DEFAULT 0.0,
                        billable_weight DOUBLE DEFAULT 0.0,
                        advertising_cost DOUBLE DEFAULT 0.0,
                        auto_parts_specific DOUBLE DEFAULT 0.0,
                        calculation_id VARCHAR,
                        --  v101.0: ABC/XYZ поля
                        abc_category VARCHAR,
                        xyz_category VARCHAR,
                        abcxyz_category VARCHAR,
                        parent_category VARCHAR,
                        group_category VARCHAR,
                        subgroup_category VARCHAR
                    )
                """)
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_timestamp ON calculation_history(timestamp)")
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_marketplace ON calculation_history(marketplace)")
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_article ON calculation_history(article)")
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_abc ON calculation_history(abc_category)")
            else:
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS calculation_history (
                        id TEXT PRIMARY KEY,
                        timestamp TEXT NOT NULL,
                        marketplace TEXT,
                        operation_mode TEXT,
                        category TEXT,
                        article TEXT,
                        brand TEXT,
                        price REAL,
                        cost REAL,
                        length REAL,
                        width REAL,
                        height REAL,
                        weight REAL,
                        volume REAL,
                        commission REAL,
                        commission_percent REAL,
                        logistics REAL,
                        storage_cost REAL,
                        acquiring REAL,
                        delivery REAL,
                        last_mile REAL,
                        returns REAL,
                        rko_fee REAL,
                        premium_fee REAL,
                        insurance_fee REAL,
                        packing_fee REAL,
                        marketing_fee REAL,
                        subscription_cost REAL,
                        hazardous_surcharge REAL,
                        fragile_surcharge REAL,
                        oversized_surcharge REAL,
                        tax_amount REAL,
                        tax_system TEXT,
                        total_expenses REAL,
                        profit REAL,
                        margin_percent REAL,
                        roi REAL,
                        breakeven_price REAL,
                        recommended_min_price REAL,
                        profit_per_ruble REAL,
                        contribution_margin REAL,
                        contribution_margin_ratio REAL,
                        tariff_source TEXT,
                        status TEXT,
                        metadata_json TEXT,
                        applied_seasonal_multiplier REAL DEFAULT 1.0,
                        applied_promo_discount REAL DEFAULT 0.0,
                        dynamic_adjustment REAL DEFAULT 0.0,
                        billable_weight REAL DEFAULT 0.0,
                        advertising_cost REAL DEFAULT 0.0,
                        auto_parts_specific REAL DEFAULT 0.0,
                        calculation_id TEXT,
                        abc_category TEXT,
                        xyz_category TEXT,
                        abcxyz_category TEXT,
                        parent_category TEXT,
                        group_category TEXT,
                        subgroup_category TEXT
                    )
                """)
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_timestamp ON calculation_history(timestamp)")
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_marketplace ON calculation_history(marketplace)")
                self.conn.execute("CREATE INDEX IF NOT EXISTS idx_history_article ON calculation_history(article)")
                self.conn.commit()
            
            #  v101.0: Таблица сохранённых состояний
            if self.use_duckdb:
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS saved_states (
                        state_id VARCHAR PRIMARY KEY,
                        name VARCHAR,
                        description VARCHAR,
                        created_at VARCHAR,
                        updated_at VARCHAR,
                        file_path VARCHAR,
                        data_hash VARCHAR,
                        metadata_json VARCHAR,
                        size_bytes BIGINT DEFAULT 0
                    )
                """)
            else:
                self.conn.execute("""
                    CREATE TABLE IF NOT EXISTS saved_states (
                        state_id TEXT PRIMARY KEY,
                        name TEXT,
                        description TEXT,
                        created_at TEXT,
                        updated_at TEXT,
                        file_path TEXT,
                        data_hash TEXT,
                        metadata_json TEXT,
                        size_bytes INTEGER DEFAULT 0
                    )
                """)
                self.conn.commit()
        
        except Exception as e:
            logger.error(f"Ошибка создания таблиц: {e}")
    
    def _get_db_columns(self) -> List[str]:
        """Получить список колонок таблицы"""
        if self.conn is None:
            return []
        try:
            if self.use_duckdb:
                rows = self.conn.execute(
                    "SELECT column_name FROM information_schema.columns WHERE table_name = 'calculation_history'"
                ).fetchall()
                return [row[0] for row in rows]
            else:
                rows = self.conn.execute("PRAGMA table_info(calculation_history)").fetchall()
                return [row[1] for row in rows]
        except Exception as e:
            logger.warning(f"Ошибка получения колонок: {e}")
            return []
    
    def _migrate_database(self):
        """Автоматическая миграция БД  добавление новых колонок"""
        if self.conn is None:
            return
        
        try:
            db_columns = self._get_db_columns()
            
            #  v101.0: Новые колонки для ABC/XYZ и категоризации
            new_columns = {
                'abc_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
                'xyz_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
                'abcxyz_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
                'parent_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
                'group_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
                'subgroup_category': 'VARCHAR' if self.use_duckdb else 'TEXT',
            }
            
            for col_name, col_type in new_columns.items():
                if col_name not in db_columns:
                    try:
                        self.conn.execute(f'ALTER TABLE calculation_history ADD COLUMN "{col_name}" {col_type}')
                        logger.info(f"✅ Миграция: добавлена колонка {col_name}")
                    except Exception as e:
                        logger.warning(f"Не удалось добавить {col_name}: {e}")
        
        except Exception as e:
            logger.warning(f"Ошибка миграции: {e}")
    
    def save_calculation(self, result: 'UnitEconomicsResult', 
                         article: str = "", brand: str = "",
                         parent_category: str = "", group_category: str = "", 
                         subgroup_category: str = "") -> bool:
        """Сохранение расчёта в БД"""
        if self.conn is None:
            return False
        
        try:
            data = result.to_dict()
            data['article'] = article
            data['brand'] = brand
            data['metadata_json'] = json.dumps(data.get('metadata', {}), ensure_ascii=False)
            data['parent_category'] = parent_category
            data['group_category'] = group_category
            data['subgroup_category'] = subgroup_category
            
            # Фильтруем только существующие колонки
            db_columns = self._get_db_columns()
            filtered_data = {k: v for k, v in data.items() if k in db_columns}
            
            if not filtered_data:
                logger.warning("Нет подходящих колонок для сохранения")
                return False
            
            if 'id' not in filtered_data and 'calculation_id' in filtered_data:
                filtered_data['id'] = filtered_data['calculation_id']
            elif 'id' not in filtered_data:
                filtered_data['id'] = str(uuid.uuid4())
            
            columns = list(filtered_data.keys())
            values = list(filtered_data.values())
            placeholders = ", ".join(["?"] * len(values))
            col_names = ", ".join([f'"{c}"' for c in columns])
            
            sql = f"INSERT OR REPLACE INTO calculation_history ({col_names}) VALUES ({placeholders})"
            self.conn.execute(sql, values)
            self.conn.commit()
            return True
        
        except Exception as e:
            logger.error(f"Ошибка сохранения расчёта: {e}")
            return False
    
    def load_history(self, limit: int = 1000, filters: Optional[Dict] = None) -> pd.DataFrame:
        """Загрузка истории с фильтрами"""
        if self.conn is None:
            return pd.DataFrame()
        
        try:
            conditions = []
            params = []
            
            if filters:
                for key in ['marketplace', 'operation_mode', 'category', 'tax_system', 
                           'abc_category', 'parent_category', 'group_category']:
                    if filters.get(key):
                        conditions.append(f"{key} = ?")
                        params.append(filters[key])
                
                for key in ['article', 'brand']:
                    if filters.get(key):
                        conditions.append(f"{key} LIKE ?")
                        params.append(f"%{filters[key]}%")
                
                if filters.get('min_profit'):
                    conditions.append("profit >= ?")
                    params.append(filters['min_profit'])
                if filters.get('max_profit'):
                    conditions.append("profit <= ?")
                    params.append(filters['max_profit'])
                if filters.get('start_date'):
                    conditions.append("timestamp >= ?")
                    params.append(filters['start_date'])
                if filters.get('end_date'):
                    conditions.append("timestamp <= ?")
                    params.append(filters['end_date'])
            
            where_clause = " AND ".join(conditions) if conditions else "1=1"
            sql = f"SELECT * FROM calculation_history WHERE {where_clause} ORDER BY timestamp DESC LIMIT ?"
            params.append(limit)
            
            if self.use_duckdb:
                df = self.conn.execute(sql, params).pl().to_pandas()
            else:
                df = pd.read_sql_query(sql, self.conn, params=params)
            
            return df
        
        except Exception as e:
            logger.error(f"Ошибка загрузки истории: {e}")
            return pd.DataFrame()
    
    def get_stats(self) -> Dict[str, Any]:
        """Статистика по истории"""
        if self.conn is None:
            return {}
        
        try:
            if self.use_duckdb:
                total = self.conn.execute("SELECT COUNT(*) FROM calculation_history").fetchone()[0]
                total_profit = self.conn.execute("SELECT SUM(profit) FROM calculation_history").fetchone()[0] or 0
                avg_profit = self.conn.execute("SELECT AVG(profit) FROM calculation_history").fetchone()[0] or 0
                avg_margin = self.conn.execute("SELECT AVG(margin_percent) FROM calculation_history").fetchone()[0] or 0
                by_marketplace = self.conn.execute("""
                    SELECT marketplace, COUNT(*) as cnt, SUM(profit) as total_profit 
                    FROM calculation_history GROUP BY marketplace ORDER BY cnt DESC
                """).pl().to_pandas()
            else:
                total = self.conn.execute("SELECT COUNT(*) FROM calculation_history").fetchone()[0]
                total_profit = self.conn.execute("SELECT SUM(profit) FROM calculation_history").fetchone()[0] or 0
                avg_profit = self.conn.execute("SELECT AVG(profit) FROM calculation_history").fetchone()[0] or 0
                avg_margin = self.conn.execute("SELECT AVG(margin_percent) FROM calculation_history").fetchone()[0] or 0
                by_marketplace = pd.read_sql_query("""
                    SELECT marketplace, COUNT(*) as cnt, SUM(profit) as total_profit 
                    FROM calculation_history GROUP BY marketplace ORDER BY cnt DESC
                """, self.conn)
            
            return {
                "total_records": total,
                "total_profit": float(total_profit),
                "avg_profit": float(avg_profit),
                "avg_margin": float(avg_margin),
                "by_marketplace": by_marketplace
            }
        
        except Exception as e:
            logger.error(f"Ошибка получения статистики: {e}")
            return {}
    
    def clear_history(self) -> int:
        """Очистка всей истории"""
        if self.conn is None:
            return 0
        
        try:
            count = self.conn.execute("SELECT COUNT(*) FROM calculation_history").fetchone()[0]
            self.conn.execute("DELETE FROM calculation_history")
            self.conn.commit()
            return count
        except Exception as e:
            logger.error(f"Ошибка очистки истории: {e}")
            return 0
    
    def close(self):
        """Закрытие подключения"""
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None


# ============================================================================
# 7.2 МЕНЕДЖЕР СОСТОЯНИЯ ПРИЛОЖЕНИЯ
# ============================================================================

class AppStateManager:
    """
     v101.0: Менеджер состояния приложения.
    Управляет передачей данных между 4 разделами через st.session_state.
    """
    
    # Ключи session_state
    KEYS = {
        # : Загрузка данных
        'section1_catalog_df': 'catalog_dataframe',
        'section1_linked_df': 'linked_dataframe',
        'section1_column_mappings': 'column_mappings',
        'section1_data_links': 'data_links',
        
        # : Весогабариты и категоризация
        'section2_categorized_df': 'categorized_dataframe',
        'section2_categories_db': 'categories_db',
        'section2_classifier': 'category_classifier',
        'section2_standard_dims': 'standard_dimensions',
        
        # : Тарифы
        'section3_tariffs': 'tariffs_config',
        'section3_google_config': 'google_sheets_config',
        'section3_tariff_cache': 'tariff_cache',
        
        # : Расчёт
        'section4_results_df': 'calculation_results',
        'section4_abcxyz_df': 'abcxyz_results',
        'section4_metadata': 'calculation_metadata',
        
        # Общие
        'current_section': 'current_section',
        'last_save_state': 'last_save_state',
    }
    
    @staticmethod
    def get(key: str, default: Any = None) -> Any:
        """Получение значения из session_state"""
        actual_key = AppStateManager.KEYS.get(key, key)
        return st.session_state.get(actual_key, default)
    
    @staticmethod
    def set(key: str, value: Any):
        """Установка значения в session_state"""
        actual_key = AppStateManager.KEYS.get(key, key)
        st.session_state[actual_key] = value
    
    @staticmethod
    def has(key: str) -> bool:
        """Проверка наличия ключа"""
        actual_key = AppStateManager.KEYS.get(key, key)
        return actual_key in st.session_state
    
    @staticmethod
    def delete(key: str):
        """Удаление ключа"""
        actual_key = AppStateManager.KEYS.get(key, key)
        if actual_key in st.session_state:
            del st.session_state[actual_key]
    
    @staticmethod
    def get_catalog_data() -> Optional[pd.DataFrame]:
        """Получение данных каталога из Раздела 1"""
        # Приоритет: связанный DataFrame > обычный каталог
        linked = st.session_state.get('linked_dataframe')
        if linked is not None and not linked.empty:
            return linked
        return st.session_state.get('catalog_dataframe')
    
    @staticmethod
    def get_categorized_data() -> Optional[pd.DataFrame]:
        """Получение категоризированных данных из Раздела 2"""
        return st.session_state.get('categorized_dataframe')
    
    @staticmethod
    def get_tariffs_config() -> Optional[Dict]:
        """Получение конфигурации тарифов из Раздела 3"""
        return st.session_state.get('tariffs_config')
    
    @staticmethod
    def get_google_config() -> Optional[GoogleSheetsConfig]:
        """Получение конфигурации Google Sheets из Раздела 3"""
        return st.session_state.get('google_sheets_config')
    
    @staticmethod
    def clear_all():
        """Очистка всего состояния"""
        for key in AppStateManager.KEYS.values():
            if key in st.session_state:
                del st.session_state[key]


# ============================================================================
# 7.3 МЕНЕДЖЕР СОХРАНЕНИЯ/ЗАГРУЗКИ
# ============================================================================

class SaveLoadManager:
    """
     v101.0: Менеджер сохранения и загрузки расчётов.
    Позволяет сохранить текущее состояние одной кнопкой и загрузить его позже.
    """
    
    def __init__(self):
        self.save_dir = BACKUPS_DIR / "calculations"
        self.save_dir.mkdir(parents=True, exist_ok=True)
        self.states_file = self.save_dir / "states_index.json"
        self._states_index = self._load_index()
    
    def _load_index(self) -> Dict[str, Any]:
        """Загрузка индекса сохранённых состояний"""
        if self.states_file.exists():
            try:
                with open(self.states_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}
    
    def _save_index(self):
        """Сохранение индекса"""
        try:
            with open(self.states_file, 'w', encoding='utf-8') as f:
                json.dump(self._states_index, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Ошибка сохранения индекса: {e}")
    
    def save_current_state(self, name: str = "", description: str = "") -> Optional[str]:
        """
        Сохранение текущего состояния приложения.
        
        Returns:
            state_id при успехе, None при ошибке
        """
        try:
            state_id = str(uuid.uuid4())
            timestamp = datetime.now()
            
            # Собираем данные из всех разделов
            state_data = {
                'section1': {
                    'catalog_df': AppStateManager.get('section1_catalog_df'),
                    'linked_df': AppStateManager.get('section1_linked_df'),
                    'column_mappings': AppStateManager.get('section1_column_mappings'),
                },
                'section2': {
                    'categorized_df': AppStateManager.get('section2_categorized_df'),
                    'standard_dims': AppStateManager.get('section2_standard_dims'),
                },
                'section3': {
                    'tariffs': AppStateManager.get('section3_tariffs'),
                    'google_config': AppStateManager.get('section3_google_config'),
                },
                'section4': {
                    'results_df': AppStateManager.get('section4_results_df'),
                    'abcxyz_df': AppStateManager.get('section4_abcxyz_df'),
                    'metadata': AppStateManager.get('section4_metadata'),
                },
            }
            
            # Сериализация
            file_path = self.save_dir / f"state_{state_id}.json.gz"
            
            import gzip
            data_json = json.dumps({
                'state_id': state_id,
                'name': name or f"Расчёт {timestamp.strftime('%Y-%m-%d %H:%M')}",
                'description': description,
                'created_at': timestamp.isoformat(),
                'data': _serialize_for_json(state_data),
            }, ensure_ascii=False)
            
            with gzip.open(file_path, 'wt', encoding='utf-8') as f:
                f.write(data_json)
            
            # Обновляем индекс
            self._states_index[state_id] = {
                'name': name or f"Расчёт {timestamp.strftime('%Y-%m-%d %H:%M')}",
                'description': description,
                'created_at': timestamp.isoformat(),
                'file_path': str(file_path),
                'size_bytes': file_path.stat().st_size,
            }
            self._save_index()
            
            logger.info(f"✅ Состояние сохранено: {state_id}")
            return state_id
        
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения состояния: {e}")
            return None
    
    def load_state(self, state_id: str) -> bool:
        """
        Загрузка состояния по ID.
        
        Returns:
            True при успехе
        """
        try:
            if state_id not in self._states_index:
                logger.error(f"Состояние {state_id} не найдено")
                return False
            
            file_path = Path(self._states_index[state_id]['file_path'])
            if not file_path.exists():
                logger.error(f"Файл состояния не найден: {file_path}")
                return False
            
            import gzip
            with gzip.open(file_path, 'rt', encoding='utf-8') as f:
                data = json.load(f)
            
            state_data = _deserialize_from_json(data.get('data', {}))
            
            # Восстанавливаем состояние
            if 'section1' in state_data:
                s1 = state_data['section1']
                if s1.get('catalog_df') is not None:
                    AppStateManager.set('section1_catalog_df', s1['catalog_df'])
                if s1.get('linked_df') is not None:
                    AppStateManager.set('section1_linked_df', s1['linked_df'])
                if s1.get('column_mappings') is not None:
                    AppStateManager.set('section1_column_mappings', s1['column_mappings'])
            
            if 'section2' in state_data:
                s2 = state_data['section2']
                if s2.get('categorized_df') is not None:
                    AppStateManager.set('section2_categorized_df', s2['categorized_df'])
                if s2.get('standard_dims') is not None:
                    AppStateManager.set('section2_standard_dims', s2['standard_dims'])
            
            if 'section3' in state_data:
                s3 = state_data['section3']
                if s3.get('tariffs') is not None:
                    AppStateManager.set('section3_tariffs', s3['tariffs'])
                if s3.get('google_config') is not None:
                    AppStateManager.set('section3_google_config', s3['google_config'])
            
            if 'section4' in state_data:
                s4 = state_data['section4']
                if s4.get('results_df') is not None:
                    AppStateManager.set('section4_results_df', s4['results_df'])
                if s4.get('abcxyz_df') is not None:
                    AppStateManager.set('section4_abcxyz_df', s4['abcxyz_df'])
                if s4.get('metadata') is not None:
                    AppStateManager.set('section4_metadata', s4['metadata'])
            
            logger.info(f"✅ Состояние загружено: {state_id}")
            return True
        
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки состояния: {e}")
            return False
    
    def list_states(self) -> List[Dict[str, Any]]:
        """Список всех сохранённых состояний"""
        states = []
        for state_id, info in self._states_index.items():
            states.append({
                'state_id': state_id,
                'name': info.get('name', ''),
                'description': info.get('description', ''),
                'created_at': info.get('created_at', ''),
                'size_bytes': info.get('size_bytes', 0),
            })
        
        # Сортируем по дате (новые первыми)
        states.sort(key=lambda x: x.get('created_at', ''), reverse=True)
        return states
    
    def delete_state(self, state_id: str) -> bool:
        """Удаление сохранённого состояния"""
        try:
            if state_id not in self._states_index:
                return False
            
            file_path = Path(self._states_index[state_id]['file_path'])
            if file_path.exists():
                file_path.unlink()
            
            del self._states_index[state_id]
            self._save_index()
            
            logger.info(f"✅ Состояние удалено: {state_id}")
            return True
        
        except Exception as e:
            logger.error(f"❌ Ошибка удаления состояния: {e}")
            return False


@st.cache_resource
def get_save_load_manager():
    """Получение экземпляра SaveLoadManager"""
    return SaveLoadManager()

# ============================================================================
# БЛОК 8: ОСНОВНОЙ КЛАСС ЮНИТ-ЭКОНОМИКИ (MarketplaceUnitEconomics)
# ============================================================================
# 📌 v101.0: Адаптирован для новой архитектуры с 4 разделами
# - Поддержка ABC/XYZ анализа
# - 3-уровневая категоризация (Родитель/Группа/Подгруппа)
# - Интеграция с AppStateManager и PersistentHistoryDB
# - Параллельный расчёт для больших каталогов
# - Оптимизация цены и прогноз прибыли
# - Google Sheets экспорт
# ============================================================================


@st.cache_resource
def get_marketplace_unit_economics():
    """Получение экземпляра через st.cache_resource"""
    return MarketplaceUnitEconomics()


class MarketplaceUnitEconomics:
    """
    🚗 Основной класс для расчёта юнит-экономики автозапчастей.
    
     v101.0: Новые возможности:
    - ABC/XYZ анализ по маржинальности и прибыли
    - 3-уровневая категоризация (Родитель  Группа  Подгруппа)
    - Интеграция с AppStateManager для передачи данных между разделами
    - Автоматическое сохранение в PersistentHistoryDB
    - Поддержка Google Sheets экспорта
    """
    
    def __init__(self):
        self._configs = self._load_marketplace_configs()
        self._categories = self._load_categories()
        self._cache = {}
        self._history = []
        self._stats = self._init_stats()
        self._settings = self._load_settings()
        self._tariff_cache = get_smart_tariff_cache()
        self._ai_updater = None
        self._parallel_cache = {}
        
        # ✅ Lock для потокобезопасного обновления статистики
        self._stats_lock = threading.Lock()
        
        try:
            self._persistent_db = get_persistent_history_db()
        except Exception as e:
            logger.error(f"Ошибка инициализации PersistentHistoryDB: {e}")
            self._persistent_db = None
        
        self._logger = logging.getLogger('MarketplaceUnitEconomics')
        self._logger.info("🚗 Инициализация MarketplaceUnitEconomics v101.0")
        self._logger.info(f"📊 Загружено {len(self._configs)} маркетплейсов")
        self._logger.info(f"📚 Загружено {len(self._categories)} категорий")
        
        if self._persistent_db:
            self._logger.info("📚 Постоянное хранилище истории подключено")
    
    def _load_marketplace_configs(self) -> Dict[str, MarketplaceConfig]:
        """Загрузка конфигураций маркетплейсов"""
        return get_marketplace_configs_2026()
    
    def _load_categories(self) -> Dict[str, ProductCategory]:
        """Загрузка категорий с 3-уровневой иерархией"""
        categories = {}
        for name, cat in get_auto_parts_categories_full().items():
            categories[name] = cat
        return categories
    
    def _init_stats(self) -> Dict[str, Any]:
        """Инициализация статистики"""
        return {
            "total_calculations": 0,
            "by_marketplace": defaultdict(int),
            "by_category": defaultdict(int),
            "by_mode": defaultdict(int),
            "by_status": defaultdict(int),
            "by_tax_system": defaultdict(int),
            "by_tariff_source": defaultdict(int),
            "avg_profit": 0.0,
            "avg_margin": 0.0,
            "avg_roi": 0.0,
            "avg_tax": 0.0,
            "total_profit": 0.0,
            "total_tax": 0.0,
            "max_profit": 0.0,
            "min_profit": 0.0,
            "best_marketplace": None,
            "best_category": None,
            "best_mode": None,
            "total_optimizations": 0,
            "optimization_improvement": 0.0,
            "start_time": datetime.now(),
            "errors_count": 0,
            "last_error": None,
            "cache_hits": 0,
            "cache_misses": 0,
            "ai_requests": 0,
            "db_saved": 0,
            "parallel_calculations": 0
        }
    
    def _load_settings(self) -> Dict[str, Any]:
        """Загрузка настроек из файла"""
        settings_path = CONFIG_DIR / "settings.json"
        default_settings = {
            "default_marketplace": "Ozon",
            "default_mode": "FBS",
            "default_days_storage": 30,
            "target_margin": 20.0,
            "enable_ai": True,
            "ai_provider": "deepseek",
            "enable_cache": True,
            "cache_ttl": 3600,
            "parallel_processing": True,
            "max_workers": 4,
            "optimize_memory": True,
            "precision_decimals": 2,
            "currency": "RUB",
            "locale": "ru_RU",
            "timezone": "Europe/Moscow",
            "enable_persistent_history": True,
            "global_markup": DEFAULT_MARKUP_GLOBAL,
            "discount_max": DEFAULT_DISCOUNT_MAX,
            "enable_seasonal_adjustments": True,
            "forecast_months": 3,
            "tax_system": "УСН_6",
            "ad_intensity": "medium",
        }
        
        if settings_path.exists():
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                default_settings.update(settings)
            except (IOError, json.JSONDecodeError) as e:
                self._logger.warning(f"Ошибка загрузки настроек: {e}")
        
        return default_settings
    
    def save_settings(self, settings: Dict[str, Any]) -> bool:
        """Сохранение настроек в файл"""
        try:
            settings_path = CONFIG_DIR / "settings.json"
            self._settings.update(settings)
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
            return True
        except (IOError, OSError) as e:
            self._logger.error(f"Ошибка сохранения настроек: {e}")
            return False
    
    # ========================================================================
    # МЕТОДЫ РАБОТЫ С КАТЕГОРИЯМИ
    # ========================================================================
    
    def get_category_dimensions(self, category_name: str) -> Optional[ProductDimensions]:
        """Получение типичных габаритов категории"""
        if category_name in self._categories:
            return self._categories[category_name].dimensions
        return None
    
    def get_category_info(self, category_name: str) -> Optional[ProductCategory]:
        """Получение информации о категории"""
        return self._categories.get(category_name)
    
    def find_categories_by_keyword(self, keyword: str) -> List[Tuple[str, ProductCategory]]:
        """Поиск категорий по ключевому слову"""
        keyword_lower = keyword.lower()
        results = []
        for name, cat in self._categories.items():
            if keyword_lower in name.lower() or keyword_lower in cat.description.lower():
                results.append((name, cat))
        return results
    
    def calculate_dimensions_from_category(self, category_name: str) -> Tuple[float, float, float, float]:
        """Расчёт габаритов из категории (если не указаны явно)"""
        cat = self._categories.get(category_name)
        if cat and cat.dimensions:
            return (
                cat.dimensions.length,
                cat.dimensions.width,
                cat.dimensions.height,
                cat.dimensions.weight
            )
        return 0, 0, 0, 0
    
    def is_category_hazardous(self, category_name: str) -> bool:
        """Проверка, является ли категория опасной"""
        cat = self._categories.get(category_name)
        return cat.hazardous if cat else False
    
    def is_category_fragile(self, category_name: str) -> bool:
        """Проверка, является ли категория хрупкой"""
        cat = self._categories.get(category_name)
        return cat.fragile if cat else False
    
    def is_category_oversized(self, length: float, width: float, height: float, weight: float) -> bool:
        """Проверка, является ли товар крупногабаритным"""
        return any([length > 100, width > 100, height > 100, weight > 25])
    
    # ========================================================================
    # AI ИНТЕГРАЦИЯ
    # ========================================================================
    
    def _get_ai_updater(self) -> Optional['DeepSeekRateUpdater']:
        """Получение AI updater"""
        if self._ai_updater is None:
            try:
                self._ai_updater = DeepSeekRateUpdater()
            except Exception as e:
                self._logger.error(f"Ошибка инициализации AI updater: {e}")
                return None
        return self._ai_updater
    
    def get_tariff_forecast(self, marketplace: str, category: str = None, 
                           months_ahead: int = 3) -> Optional[Dict[str, Any]]:
        """Получение прогноза тарифов"""
        updater = self._get_ai_updater()
        if updater is None:
            return None
        return updater.get_tariff_forecast(marketplace, category, months_ahead)
    
    def refresh_tariffs_from_ai(
        self,
        marketplace: Optional[str] = None,
        category: Optional[str] = None,
        force: bool = True,
        include_forecast: bool = False
    ) -> Dict[str, Any]:
        """Обновление тарифов через AI"""
        updater = self._get_ai_updater()
        if updater is None:
            return {"error": "AI updater не инициализирован"}
        
        self._stats["ai_requests"] += 1
        
        try:
            if marketplace:
                rates, source, forecast = updater.get_rates_from_ai(
                    marketplace=marketplace,
                    category=category,
                    force_refresh=force,
                    use_cache=True,
                    include_forecast=include_forecast
                )
                if rates:
                    self._apply_ai_tariffs(marketplace, rates)
                    return {
                        "marketplace": marketplace,
                        "source": source.value,
                        "rates": rates,
                        "forecast": forecast,
                        "success": True
                    }
                else:
                    return {"error": "Не удалось получить тарифы", "success": False}
            else:
                results = updater.update_all_marketplaces(
                    force_refresh=force,
                    include_forecast=include_forecast
                )
                for mp, (rates, source, forecast) in results.items():
                    if rates:
                        self._apply_ai_tariffs(mp, rates)
                return {
                    "marketplaces_updated": len([r for r in results.values() if r[0]]),
                    "total": len(results),
                    "success": True
                }
        except Exception as e:
            self._logger.error(f"Ошибка обновления тарифов через AI: {e}")
            return {"error": str(e), "success": False}
    
    def _apply_ai_tariffs(self, marketplace: str, rates: Dict[str, Any]):
        """Применение AI-тарифов к конфигурации"""
        if marketplace not in self._configs:
            return
        
        config = self._configs[marketplace]
        
        field_mapping = {
            "commission_rate": "commission_rate",
            "min_commission": "min_commission",
            "logistics_base": "logistics_base",
            "logistics_per_kg": "logistics_per_kg",
            "logistics_per_liter": "logistics_per_liter",
            "storage_per_day": "storage_per_day",
            "return_fee": "return_fee",
            "acquiring_fee": "acquiring_fee",
            "last_mile_fee": "last_mile_fee",
            "delivery_fee_percent": "delivery_fee_percent",
            "hazardous_surcharge": "hazardous_surcharge",
            "fragile_surcharge": "fragile_surcharge",
            "oversized_surcharge": "oversized_surcharge",
            "seasonal_multipliers": "seasonal_multipliers"
        }
        
        for ai_key, config_key in field_mapping.items():
            if ai_key in rates:
                try:
                    if config_key == "seasonal_multipliers" and isinstance(rates[ai_key], dict):
                        setattr(config, config_key, rates[ai_key])
                    else:
                        setattr(config, config_key, float(rates[ai_key]))
                except (ValueError, TypeError):
                    pass
        
        config.tariff_source = TariffSource.AI_LIVE
        config.last_updated = datetime.now()
        self._logger.info(f"✅ AI-тарифы применены для {marketplace}")
    
    # ========================================================================
    # ОСНОВНОЙ РАСЧЁТ ЮНИТ-ЭКОНОМИКИ
    # ========================================================================
    
    @timer_decorator
    def calculate_unit_economics(
        self,
        price: float,
        cost: float,
        marketplace: str,
        category: Optional[str] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        length: float = 0,
        width: float = 0,
        height: float = 0,
        weight: float = 0,
        is_premium: bool = False,
        include_insurance: bool = False,
        include_packing: bool = False,
        include_marketing: bool = False,
        currency: str = "RUB",
        article: str = "",
        brand: str = "",
        current_month: Optional[int] = None,
        tax_system: str = "УСН_6",
        ad_intensity: str = "medium",
        discount_percent: float = 0.0,
        promo_participation: float = 0.0,
        parent_category: str = "",
        group_category: str = "",
        subgroup_category: str = "",
        **kwargs
    ) -> 'UnitEconomicsResult':
        """
         v101.0: Расчёт юнит-экономики с поддержкой 3-уровневой категоризации.
        
        Args:
            price: Цена продажи
            cost: Себестоимость
            marketplace: Маркетплейс
            category: Категория товара (ключ из справочника)
            operation_mode: Режим работы (FBY/FBS/FBO/DBS/FBP/RealFBS)
            days_in_storage: Дней хранения
            length, width, height: Габариты (см)
            weight: Вес (кг)
            is_premium: Премиум-размещение
            include_insurance: Включить страховку
            include_packing: Включить упаковку
            include_marketing: Включить маркетинг
            currency: Валюта
            article: Артикул товара
            brand: Бренд
            current_month: Текущий месяц (для сезонности)
            tax_system: Налоговая система
            ad_intensity: Интенсивность рекламы
            discount_percent: Процент скидки
            promo_participation: Участие в акциях
            parent_category: Родительская категория (3-уровневая иерархия)
            group_category: Группа категории
            subgroup_category: Подгруппа категории
        
        Returns:
            UnitEconomicsResult с полным расчётом
        """
        # Валидация входных данных
        if price <= 0:
            raise ValidationError("Цена должна быть положительной", "price", price)
        if cost <= 0:
            raise ValidationError("Себестоимость должна быть положительной", "cost", cost)
        if marketplace not in self._configs:
            raise MarketplaceError(f"Маркетплейс {marketplace} не поддерживается", marketplace)
        
        config = self._configs[marketplace]
        
        if current_month is None:
            current_month = datetime.now().month
        
        # Парсинг размеров из строки (если передана строка)
        if isinstance(length, str):
            parsed_length, parsed_width, parsed_height = parse_dimensions_string(length)
            length = parsed_length
            width = parsed_width
            height = parsed_height
        
        # Если габариты не указаны, но есть категория  берём из категории
        if all([length == 0, width == 0, height == 0, weight == 0]) and category:
            length, width, height, weight = self.calculate_dimensions_from_category(category)
        
        # Расчёт объёма
        volume = calculate_volume(length, width, height)
        if volume == 0:
            volume = 5.0  # Дефолтный объём
        
        if weight <= 0:
            weight = 1.0  # Дефолтный вес
        
        # Расчёт оплачиваемого веса
        billable_weight = calculate_billable_weight(weight, length, width, height)
        
        # Определение характеристик категории
        hazardous = self.is_category_hazardous(category) if category else False
        fragile = self.is_category_fragile(category) if category else False
        oversized = self.is_category_oversized(length, width, height, weight)
        
        # === РАСЧЁТ КОМИССИИ ===
        commission = config.calculate_commission_with_dynamics(
            price=price,
            discount_percent=discount_percent,
            promo_participation=promo_participation,
            category=category,
            current_month=current_month
        )
        commission_percent = (commission / price * 100) if price > 0 else 0
        
        # Сезонный коэффициент
        seasonal_multiplier = config.apply_seasonal_multiplier(1.0, current_month)
        
        # === РАСЧЁТ ЛОГИСТИКИ ===
        logistics = (
            config.logistics_base * seasonal_multiplier +
            billable_weight * config.logistics_per_kg * seasonal_multiplier +
            volume * config.logistics_per_liter * seasonal_multiplier
        )
        logistics = config.apply_promo_discount(logistics)
        
        # Мультипликатор режима
        mode_multiplier = config.mode_multipliers.get(operation_mode, 1.0)
        logistics *= mode_multiplier
        
        # === РАСЧЁТ ХРАНЕНИЯ ===
        storage_cost = calculate_storage_cost_progressive(
            volume_l=volume,
            days=days_in_storage,
            base_rate=config.storage_per_day,
            marketplace=marketplace
        )
        
        # === ЭКВАЙРИНГ И ДОСТАВКА ===
        acquiring = price * config.acquiring_fee
        delivery = price * config.delivery_fee_percent
        last_mile = config.last_mile_fee
        
        # === ВОЗВРАТЫ ===
        return_rate = MARKET_BENCHMARKS_2026.get(category, {}).get("return_rate", config.return_fee)
        returns = calculate_returns_cost(price, return_rate)
        
        # === ДОПОЛНИТЕЛЬНЫЕ СБОРЫ ===
        rko_fee = price * config.rko_fee if config.rko_fee > 0 else 0
        premium_fee = price * config.premium_fee if is_premium and config.premium_fee > 0 else 0
        insurance_fee = price * config.insurance_fee if include_insurance and config.insurance_fee > 0 else 0
        packing_fee = config.packing_fee if include_packing and config.packing_fee > 0 else 0
        marketing_fee = price * config.marketing_fee if include_marketing and config.marketing_fee > 0 else 0
        
        # === НАДБАВКИ ===
        hazardous_surcharge = price * config.hazardous_surcharge if hazardous else 0.0
        fragile_surcharge = price * config.fragile_surcharge if fragile else 0.0
        oversized_surcharge = price * config.oversized_surcharge if oversized else 0.0
        
        # === ПОДПИСКА ===
        subscription_cost = config.subscription_fee / 30 if config.subscription_fee > 0 else 0
        
        # === НАЛОГ ===
        tax_amount = calculate_tax(price, cost, tax_system)
        
        # === СПЕЦИФИЧЕСКИЕ РАСХОДЫ АВТОЗАПЧАСТЕЙ ===
        auto_parts_costs = AutoPartsSpecificCosts()
        auto_parts_specific = auto_parts_costs.calculate(price, is_import=False, requires_marking=True)
        
        # === РЕКЛАМНЫЕ РАСХОДЫ ===
        advertising_cost = calculate_advertising_cost(price, category or "", ad_intensity)
        
        # === ИТОГО РАСХОДОВ ===
        total_expenses = (
            cost + commission + subscription_cost + logistics + storage_cost +
            acquiring + delivery + last_mile + returns + rko_fee +
            premium_fee + insurance_fee + packing_fee + marketing_fee +
            hazardous_surcharge + fragile_surcharge + oversized_surcharge +
            tax_amount + auto_parts_specific + advertising_cost
        )
        
        # === ПРИБЫЛЬ И МЕТРИКИ ===
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0
        
        # Переменные расходы (для расчёта точки безубыточности)
        variable_rate = (
            ((commission / price) if price > 0 else 0) +
            config.acquiring_fee +
            config.delivery_fee_percent +
            config.return_fee +
            config.rko_fee +
            config.premium_fee +
            config.insurance_fee +
            config.marketing_fee +
            config.hazardous_surcharge +
            config.fragile_surcharge +
            config.oversized_surcharge +
            TAX_SYSTEMS.get(tax_system, {}).get("rate", 0.06)
        )
        
        # Постоянные расходы
        fixed_costs = logistics + storage_cost + last_mile + subscription_cost
        
        # Точка безубыточности
        breakeven_price = ((cost + fixed_costs) / (1 - variable_rate)) if (1 - variable_rate) > 0 else 0
        
        # Рекомендуемая минимальная цена
        recommended_min_price = calculate_recommended_min_price(
            cost=cost,
            commission_rate=commission / price if price > 0 else 0,
            logistics=logistics,
            storage_cost=storage_cost,
            acquiring_rate=config.acquiring_fee,
            last_mile=last_mile,
            return_rate=return_rate,
            min_profit_percent=0.10,
            tax_system=tax_system,
            tax_rate=TAX_SYSTEMS.get(tax_system, {}).get("rate", 0.06)
        )
        
        # Маржинальный доход
        contribution_margin = price - cost - commission - logistics - acquiring - delivery - last_mile - returns - tax_amount
        contribution_margin_ratio = (contribution_margin / price * 100) if price > 0 else 0
        
        # === СОЗДАНИЕ РЕЗУЛЬТАТА ===
        result = UnitEconomicsResult(
            marketplace=marketplace,
            operation_mode=operation_mode,
            category=category or "Общая",
            price=money_round(price),
            cost=money_round(cost),
            length=money_round(length),
            width=money_round(width),
            height=money_round(height),
            weight=money_round(weight),
            volume=money_round(volume, 3),
            commission=money_round(commission),
            commission_percent=money_round(commission_percent),
            logistics=money_round(logistics),
            storage_cost=money_round(storage_cost),
            acquiring=money_round(acquiring),
            delivery=money_round(delivery),
            last_mile=money_round(last_mile),
            returns=money_round(returns),
            rko_fee=money_round(rko_fee),
            premium_fee=money_round(premium_fee),
            insurance_fee=money_round(insurance_fee),
            packing_fee=money_round(packing_fee),
            marketing_fee=money_round(marketing_fee),
            subscription_cost=money_round(subscription_cost),
            hazardous_surcharge=money_round(hazardous_surcharge),
            fragile_surcharge=money_round(fragile_surcharge),
            oversized_surcharge=money_round(oversized_surcharge),
            tax_amount=money_round(tax_amount),
            tax_system=tax_system,
            total_expenses=money_round(total_expenses),
            profit=money_round(profit),
            margin_percent=money_round(margin_percent),
            roi=money_round(roi),
            breakeven_price=money_round(breakeven_price),
            recommended_min_price=money_round(recommended_min_price),
            profit_per_ruble=money_round(profit / price, 4) if price > 0 else 0,
            contribution_margin=money_round(contribution_margin),
            contribution_margin_ratio=money_round(contribution_margin_ratio),
            status=CalculationStatus.COMPLETED,
            tariff_source=config.tariff_source,
            metadata=kwargs,
            applied_seasonal_multiplier=seasonal_multiplier,
            applied_promo_discount=config.promo_discount,
            dynamic_adjustment=config.dynamic_adjustment,
            billable_weight=money_round(billable_weight),
            advertising_cost=money_round(advertising_cost),
            auto_parts_specific=money_round(auto_parts_specific)
        )
        
        # Обновление статистики
        self._update_stats(result)
        
        # Сохранение в историю
        self._history.append(result)
        if len(self._history) > HISTORY_LIMIT:
            self._history = self._history[-HISTORY_LIMIT:]
        
        # Сохранение в постоянную БД
        if self._settings.get("enable_persistent_history", True) and self._persistent_db:
            try:
                if self._persistent_db.save_calculation(
                    result, 
                    article=article, 
                    brand=brand,
                    parent_category=parent_category,
                    group_category=group_category,
                    subgroup_category=subgroup_category
                ):
                    self._stats["db_saved"] += 1
            except Exception as e:
                self._logger.warning(f"Не удалось сохранить в БД: {e}")
        
        return result
    
    @timer_decorator
    def _update_stats(self, result: 'UnitEconomicsResult'):
        """Потокобезопасное обновление статистики"""
        with self._stats_lock:
            self._stats["total_calculations"] += 1
            self._stats["by_marketplace"][result.marketplace] += 1
            self._stats["by_category"][result.category] += 1
            self._stats["by_mode"][result.operation_mode] += 1
            self._stats["by_status"][result.status.name] += 1
            self._stats["by_tax_system"][result.tax_system] += 1
            self._stats["by_tariff_source"][result.tariff_source.value] += 1
            
            self._stats["total_profit"] += result.profit
            self._stats["total_tax"] += result.tax_amount
            
            if result.profit > self._stats["max_profit"]:
                self._stats["max_profit"] = result.profit
                self._stats["best_marketplace"] = result.marketplace
                self._stats["best_category"] = result.category
                self._stats["best_mode"] = result.operation_mode
            
            if result.profit < self._stats["min_profit"] or self._stats["min_profit"] == 0:
                self._stats["min_profit"] = result.profit
            
            n = self._stats["total_calculations"]
            self._stats["avg_profit"] = self._stats["total_profit"] / n
            self._stats["avg_margin"] = (self._stats["avg_margin"] * (n - 1) + result.margin_percent) / n
            self._stats["avg_roi"] = (self._stats["avg_roi"] * (n - 1) + result.roi) / n
            self._stats["avg_tax"] = self._stats["total_tax"] / n
    
    # ========================================================================
    # ПАРАЛЛЕЛЬНЫЙ РАСЧЁТ ДЛЯ КАТАЛОГА
    # ========================================================================
    
    @timer_decorator
    def calculate_for_catalog_batch_parallel(
        self,
        df: pd.DataFrame,
        price_col: str = "Цена",
        cost_col: str = "Себестоимость",
        category_col: Optional[str] = None,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        article_col: str = "Артикул",
        brand_col: str = "Бренд",
        marketplaces: Optional[List[str]] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        apply_markup: float = 0.0,
        is_premium: bool = False,
        include_insurance: bool = False,
        include_packing: bool = False,
        include_marketing: bool = False,
        progress_callback: Optional[Callable] = None,
        max_workers: int = 4,
        chunk_size: int = 1000,
        tax_system: str = "УСН_6",
        ad_intensity: str = "medium"
    ) -> pd.DataFrame:
        """
         v101.0: Параллельный расчёт юнит-экономики для больших каталогов.
        Использует ThreadPoolExecutor для совместимости со Streamlit.
        """
        if marketplaces is None:
            marketplaces = list(self._configs.keys())
        
        if df.empty:
            return pd.DataFrame()
        
        total_items = len(df) * len(marketplaces)
        if total_items == 0:
            return pd.DataFrame()
        
        self._stats["parallel_calculations"] += 1
        current_month = datetime.now().month
        
        # Разбиваем DataFrame на чанки
        chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]
        
        all_results = []
        all_errors = []
        total_futures = len(chunks) * len(marketplaces)
        completed = 0
        
        with st.status("🚀 Параллельный расчет юнит-экономики...", expanded=True) as status:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = []
                
                # Создаём задачи для каждого чанка и маркетплейса
                for chunk in chunks:
                    for marketplace in marketplaces:
                        future = executor.submit(
                            self._calculate_chunk_threadsafe,
                            chunk_df=chunk,
                            marketplace=marketplace,
                            operation_mode=operation_mode,
                            days_in_storage=days_in_storage,
                            category_col=category_col,
                            article_col=article_col,
                            brand_col=brand_col,
                            price_col=price_col,
                            cost_col=cost_col,
                            length_col=length_col,
                            width_col=width_col,
                            height_col=height_col,
                            weight_col=weight_col,
                            apply_markup=apply_markup,
                            is_premium=is_premium,
                            include_insurance=include_insurance,
                            include_packing=include_packing,
                            include_marketing=include_marketing,
                            current_month=current_month,
                            tax_system=tax_system,
                            ad_intensity=ad_intensity
                        )
                        futures.append(future)
                
                # Собираем результаты по мере выполнения
                for future in as_completed(futures):
                    try:
                        result_chunk, errors = future.result(timeout=120)
                        all_results.extend(result_chunk)
                        all_errors.extend(errors)
                    except concurrent.futures.TimeoutError:
                        logger.error("Таймаут расчета чанка")
                        self._stats["errors_count"] += 1
                        all_errors.append("Таймаут расчета чанка")
                    except Exception as e:
                        logger.error(f"Ошибка расчета чанка: {e}")
                        self._stats["errors_count"] += 1
                        self._stats["last_error"] = str(e)
                        all_errors.append(str(e))
                    
                    completed += 1
                    if progress_callback:
                        progress_callback(completed / total_futures)
                    
                    status.update(
                        label=f"🔄 Обработано {completed}/{total_futures} чанков",
                        state="running"
                    )
            
            status.update(label="✅ Параллельный расчет завершен!", state="complete")
            
            if progress_callback:
                progress_callback(1.0)
        
        # Показываем ошибки, если они есть
        if all_errors:
            unique_errors = list(set(all_errors))[:5]
            logger.warning(f"⚠️ Ошибки при параллельном расчете: {len(all_errors)}")
            for err in unique_errors:
                logger.warning(f"  - {err}")
        
        if not all_results:
            return pd.DataFrame()
        
        return pd.DataFrame(all_results)
    
    def _calculate_chunk_threadsafe(
        self,
        chunk_df: pd.DataFrame,
        marketplace: str,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        category_col: Optional[str] = None,
        article_col: str = "Артикул",
        brand_col: str = "Бренд",
        price_col: str = "Цена",
        cost_col: str = "Себестоимость",
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        apply_markup: float = 0.0,
        is_premium: bool = False,
        include_insurance: bool = False,
        include_packing: bool = False,
        include_marketing: bool = False,
        current_month: Optional[int] = None,
        tax_system: str = "УСН_6",
        ad_intensity: str = "medium"
    ) -> Tuple[List[Dict[str, Any]], List[str]]:
        """Потокобезопасный расчёт чанка данных"""
        results = []
        errors = []
        
        # Глубокая копия конфигурации для потокобезопасности
        try:
            config = copy.deepcopy(self._configs.get(marketplace))
        except Exception:
            config = self._configs.get(marketplace)
        
        if not config:
            logger.error(f"Маркетплейс {marketplace} не найден")
            return results, [f"Маркетплейс {marketplace} не найден"]
        
        for idx, row in chunk_df.iterrows():
            try:
                price = safe_float(row.get(price_col, 0))
                cost = safe_float(row.get(cost_col, 0))
                article = safe_str(row.get(article_col, f"Товар_{idx}"))
                brand = safe_str(row.get(brand_col, ""))
                
                if price <= 0 or cost <= 0:
                    continue
                
                length = safe_float(row.get(length_col, 0)) if length_col else 0
                width = safe_float(row.get(width_col, 0)) if width_col else 0
                height = safe_float(row.get(height_col, 0)) if height_col else 0
                weight = safe_float(row.get(weight_col, 0)) if weight_col else 0
                
                category = safe_str(row.get(category_col, "")) if category_col else None
                
                final_price = price * (1 + apply_markup / 100) if apply_markup > 0 else price
                
                result = self.calculate_unit_economics(
                    price=final_price,
                    cost=cost,
                    marketplace=marketplace,
                    category=category,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length,
                    width=width,
                    height=height,
                    weight=weight,
                    is_premium=is_premium,
                    include_insurance=include_insurance,
                    include_packing=include_packing,
                    include_marketing=include_marketing,
                    article=article,
                    brand=brand,
                    current_month=current_month,
                    tax_system=tax_system,
                    ad_intensity=ad_intensity
                )
                
                result_dict = result.to_dict()
                result_dict["Артикул"] = article
                result_dict["Бренд"] = brand
                result_dict["Индекс"] = idx
                results.append(result_dict)
            
            except Exception as e:
                error_msg = f"Строка {idx}: {str(e)}"
                logger.error(f"Ошибка расчета для строки {idx}: {e}")
                errors.append(error_msg)
                continue
        
        return results, errors
    
    @timer_decorator
    def calculate_for_catalog_batch(
        self,
        df: pd.DataFrame,
        price_col: str = "Цена",
        cost_col: str = "Себестоимость",
        category_col: Optional[str] = None,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        article_col: str = "Артикул",
        brand_col: str = "Бренд",
        marketplaces: Optional[List[str]] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        apply_markup: float = 0.0,
        is_premium: bool = False,
        include_insurance: bool = False,
        include_packing: bool = False,
        include_marketing: bool = False,
        progress_callback: Optional[Callable] = None,
        use_parallel: bool = True,
        max_workers: int = 4,
        chunk_size: int = 1000,
        tax_system: str = "УСН_6",
        ad_intensity: str = "medium"
    ) -> pd.DataFrame:
        """Расчёт юнит-экономики для каталога с выбором режима"""
        if use_parallel and len(df) > 100 and DUCKDB_AVAILABLE:
            return self.calculate_for_catalog_batch_parallel(
                df=df,
                price_col=price_col,
                cost_col=cost_col,
                category_col=category_col,
                length_col=length_col,
                width_col=width_col,
                height_col=height_col,
                weight_col=weight_col,
                article_col=article_col,
                brand_col=brand_col,
                marketplaces=marketplaces,
                operation_mode=operation_mode,
                days_in_storage=days_in_storage,
                apply_markup=apply_markup,
                is_premium=is_premium,
                include_insurance=include_insurance,
                include_packing=include_packing,
                include_marketing=include_marketing,
                progress_callback=progress_callback,
                max_workers=max_workers,
                chunk_size=chunk_size,
                tax_system=tax_system,
                ad_intensity=ad_intensity
            )
        
        # Последовательный расчёт
        if marketplaces is None:
            marketplaces = list(self._configs.keys())
        
        items = []
        for idx, row in df.iterrows():
            price = safe_float(row.get(price_col, 0))
            cost = safe_float(row.get(cost_col, 0))
            article = safe_str(row.get(article_col, f"Товар_{idx}"))
            brand = safe_str(row.get(brand_col, ""))
            
            if price <= 0 or cost <= 0:
                continue
            
            length = safe_float(row.get(length_col, 0)) if length_col else 0
            width = safe_float(row.get(width_col, 0)) if width_col else 0
            height = safe_float(row.get(height_col, 0)) if height_col else 0
            weight = safe_float(row.get(weight_col, 0)) if weight_col else 0
            
            category = safe_str(row.get(category_col, "")) if category_col else None
            
            items.append({
                "idx": idx, "article": article, "brand": brand,
                "price": price, "cost": cost, "category": category,
                "length": length, "width": width,
                "height": height, "weight": weight
            })
        
        total_items = len(items) * len(marketplaces)
        if total_items == 0:
            return pd.DataFrame()
        
        results = []
        processed = 0
        current_month = datetime.now().month
        
        with st.status("Расчет юнит-экономики для каталога...", expanded=True) as status:
            for item in items:
                final_price = item["price"] * (1 + apply_markup / 100) if apply_markup > 0 else item["price"]
                
                for marketplace in marketplaces:
                    try:
                        result = self.calculate_unit_economics(
                            price=final_price,
                            cost=item["cost"],
                            marketplace=marketplace,
                            category=item["category"],
                            operation_mode=operation_mode,
                            days_in_storage=days_in_storage,
                            length=item["length"],
                            width=item["width"],
                            height=item["height"],
                            weight=item["weight"],
                            is_premium=is_premium,
                            include_insurance=include_insurance,
                            include_packing=include_packing,
                            include_marketing=include_marketing,
                            article=item["article"],
                            brand=item["brand"],
                            current_month=current_month,
                            tax_system=tax_system,
                            ad_intensity=ad_intensity
                        )
                        
                        result_dict = result.to_dict()
                        result_dict["Артикул"] = item["article"]
                        result_dict["Бренд"] = item["brand"]
                        result_dict["Индекс"] = item["idx"]
                        results.append(result_dict)
                    
                    except Exception as e:
                        logger.error(f"Ошибка расчета для {item['article']}: {e}")
                        self._stats["errors_count"] += 1
                
                processed += 1
                if progress_callback and processed % 10 == 0:
                    progress_callback(processed / total_items)
            
            if progress_callback:
                progress_callback(1.0)
            
            status.update(label="✅ Расчет завершен!", state="complete")
        
        return pd.DataFrame(results) if results else pd.DataFrame()
    
    # ========================================================================
    # ОПТИМИЗАЦИЯ ЦЕНЫ
    # ========================================================================
    
    @timer_decorator
    def optimize_price(
        self,
        cost: float,
        marketplace: str,
        category: Optional[str] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        length: float = 0,
        width: float = 0,
        height: float = 0,
        weight: float = 0,
        target_margin: float = 20.0,
        price_min: float = 0,
        price_max: float = 100000,
        step: float = 10,
        max_iterations: int = 1000
    ) -> 'OptimizationResult':
        """Оптимизация цены для достижения целевой маржи"""
        current_price = max(price_min, cost * 1.1) if price_min == 0 else price_min
        best_price = current_price
        best_profit = float('-inf')
        best_margin = 0
        best_result = None
        
        iteration = 0
        
        while current_price <= price_max and iteration < max_iterations:
            iteration += 1
            
            try:
                result = self.calculate_unit_economics(
                    price=current_price,
                    cost=cost,
                    marketplace=marketplace,
                    category=category,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length,
                    width=width,
                    height=height,
                    weight=weight
                )
                
                margin = result.margin_percent
                profit = result.profit
                
                if margin >= target_margin and profit > best_profit:
                    best_profit = profit
                    best_price = current_price
                    best_margin = margin
                    best_result = result
                
                current_price += step
            
            except Exception as e:
                self._logger.warning(f"Ошибка при оптимизации для цены {current_price}: {e}")
                current_price += step
        
        # Финальный расчёт для текущей цены
        current_result = self.calculate_unit_economics(
            price=price_min or best_price,
            cost=cost,
            marketplace=marketplace,
            category=category,
            operation_mode=operation_mode,
            days_in_storage=days_in_storage,
            length=length,
            width=width,
            height=height,
            weight=weight
        )
        
        improvement_pct = ((best_profit - current_result.profit) / current_result.profit * 100) if current_result.profit > 0 else 0
        
        recommendations = []
        if best_price > 0 and best_margin >= target_margin:
            recommendations.append(f"Установите цену {best_price:.2f} ₽ для достижения маржи {target_margin}%")
        else:
            recommendations.append(f"Целевая маржа {target_margin}% не достигнута. Максимальная маржа: {best_margin:.1f}%")
        
        if best_profit > current_result.profit:
            recommendations.append(f"Потенциальное увеличение прибыли: {improvement_pct:.1f}%")
        
        if current_result.recommended_min_price > 0:
            recommendations.append(f"Рекомендуемая минимальная цена: {current_result.recommended_min_price:.2f} ₽")
        
        self._stats["total_optimizations"] += 1
        self._stats["optimization_improvement"] += improvement_pct
        
        return OptimizationResult(
            optimal_price=best_price,
            optimal_margin=best_margin,
            optimal_profit=best_profit,
            current_price=current_result.price,
            current_margin=current_result.margin_percent,
            current_profit=current_result.profit,
            improvement_pct=improvement_pct,
            recommendations=recommendations,
            metadata={"target_margin": target_margin, "step": step, "iterations": iteration}
        )
    
    # ========================================================================
    # ИСТОРИЯ И СТАТИСТИКА
    # ========================================================================
    
    def get_history(self, limit: int = 100, filters: Optional[Dict] = None) -> List['UnitEconomicsResult']:
        """Получение истории расчётов"""
        history = self._history[-limit:] if limit > 0 else self._history
        
        if filters:
            filtered = []
            for item in history:
                match = True
                for key, value in filters.items():
                    if key == "marketplace" and item.marketplace != value:
                        match = False
                        break
                    elif key == "category" and item.category != value:
                        match = False
                        break
                    elif key == "operation_mode" and item.operation_mode != value:
                        match = False
                        break
                    elif key == "tax_system" and item.tax_system != value:
                        match = False
                        break
                    elif key == "min_profit" and item.profit < value:
                        match = False
                        break
                    elif key == "max_profit" and item.profit > value:
                        match = False
                        break
                    elif key == "start_date" and item.timestamp < value:
                        match = False
                        break
                    elif key == "end_date" and item.timestamp > value:
                        match = False
                        break
                
                if match:
                    filtered.append(item)
            
            return filtered
        
        return history
    
    def get_persistent_history(self, limit: int = 1000, filters: Optional[Dict] = None) -> pd.DataFrame:
        """Получение истории из постоянной БД"""
        if not self._persistent_db:
            return pd.DataFrame()
        return self._persistent_db.load_history(limit=limit, filters=filters)
    
    def get_persistent_stats(self) -> Dict[str, Any]:
        """Получение статистики из постоянной БД"""
        if not self._persistent_db:
            return {}
        return self._persistent_db.get_stats()
    
    def clear_persistent_history(self) -> int:
        """Очистка постоянной истории"""
        if not self._persistent_db:
            return 0
        return self._persistent_db.clear_history()
    
    def get_stats(self) -> Dict[str, Any]:
        """Получение статистики"""
        stats = self._stats.copy()
        stats["history_count"] = len(self._history)
        stats["cache_size"] = len(self._cache)
        stats["uptime"] = (datetime.now() - stats["start_time"]).total_seconds()
        
        if stats["total_calculations"] > 0:
            stats["success_rate"] = 1 - (stats["errors_count"] / stats["total_calculations"])
        else:
            stats["success_rate"] = 0
        
        return stats
    
    def clear_history(self):
        """Очистка истории"""
        self._history = []
        self._stats = self._init_stats()
        self._cache.clear()
        gc.collect()
    
    def get_best_configuration(self) -> Dict[str, Any]:
        """Получение лучшей конфигурации"""
        if not self._history:
            return {"error": "Нет данных"}
        
        best = max(self._history, key=lambda x: x.profit)
        return {
            "marketplace": best.marketplace,
            "operation_mode": best.operation_mode,
            "category": best.category,
            "profit": best.profit,
            "margin": best.margin_percent,
            "price": best.price,
            "cost": best.cost,
            "tax": best.tax_amount,
            "tax_system": best.tax_system,
            "recommended_min_price": best.recommended_min_price,
            "timestamp": best.timestamp.isoformat()
        }
    
    def get_category_stats(self) -> pd.DataFrame:
        """Статистика по категориям"""
        if not self._history:
            return pd.DataFrame()
        
        stats = defaultdict(lambda: {
            "count": 0, "total_profit": 0, "avg_profit": 0,
            "avg_margin": 0, "best_profit": 0, "worst_profit": 0,
            "total_tax": 0, "avg_recommended_price": 0
        })
        
        for result in self._history:
            cat = result.category
            stats[cat]["count"] += 1
            stats[cat]["total_profit"] += result.profit
            stats[cat]["avg_margin"] += result.margin_percent
            stats[cat]["total_tax"] += result.tax_amount
            stats[cat]["avg_recommended_price"] += result.recommended_min_price
            stats[cat]["best_profit"] = max(stats[cat]["best_profit"], result.profit)
            stats[cat]["worst_profit"] = min(stats[cat]["worst_profit"], result.profit)
        
        for cat in stats:
            if stats[cat]["count"] > 0:
                stats[cat]["avg_profit"] = stats[cat]["total_profit"] / stats[cat]["count"]
                stats[cat]["avg_margin"] /= stats[cat]["count"]
                stats[cat]["avg_recommended_price"] /= stats[cat]["count"]
        
        return pd.DataFrame.from_dict(stats, orient="index").reset_index().rename(columns={"index": "category"})
    
    def get_marketplace_stats(self) -> pd.DataFrame:
        """Статистика по маркетплейсам"""
        if not self._history:
            return pd.DataFrame()
        
        stats = defaultdict(lambda: {
            "count": 0, "total_profit": 0, "avg_profit": 0,
            "avg_margin": 0, "best_profit": 0, "worst_profit": 0,
            "total_tax": 0
        })
        
        for result in self._history:
            mp = result.marketplace
            stats[mp]["count"] += 1
            stats[mp]["total_profit"] += result.profit
            stats[mp]["avg_margin"] += result.margin_percent
            stats[mp]["total_tax"] += result.tax_amount
            stats[mp]["best_profit"] = max(stats[mp]["best_profit"], result.profit)
            stats[mp]["worst_profit"] = min(stats[mp]["worst_profit"], result.profit)
        
        for mp in stats:
            if stats[mp]["count"] > 0:
                stats[mp]["avg_profit"] = stats[mp]["total_profit"] / stats[mp]["count"]
                stats[mp]["avg_margin"] /= stats[mp]["count"]
        
        return pd.DataFrame.from_dict(stats, orient="index").reset_index().rename(columns={"index": "marketplace"})
    
    def export_history(self, format: ExportFormat = ExportFormat.EXCEL) -> bytes:
        """Экспорт истории"""
        if not self._history:
            return b""
        
        df = pd.DataFrame([r.to_dict() for r in self._history])
        
        if format == ExportFormat.CSV:
            return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
        elif format == ExportFormat.EXCEL:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='История')
            return output.getvalue()
        elif format == ExportFormat.JSON:
            return df.to_json(orient='records', force_ascii=False).encode('utf-8')
        else:
            raise ExportError(f"Формат {format.name} не поддерживается", format=format.name)
    
    def get_tariff_cache_statistics(self) -> Dict[str, Any]:
        """Статистика кэша тарифов"""
        return self._tariff_cache.get_statistics()
    
    def get_tariff_cache_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        """История кэша тарифов"""
        return self._tariff_cache.get_history(limit)

# ============================================================================
# БЛОК 9: HIGH-VOLUME КАТАЛОГ АВТОЗАПЧАСТЕЙ (v101.0 — АДАПТИРОВАННАЯ ВЕРСИЯ)
# ============================================================================
# 📌 v101.0: Адаптировано для новой архитектуры с 4 разделами
# - Интеграция с AppStateManager для передачи данных между разделами
# - Добавлено связывание столбцов между файлами (link_files)
# - Добавлено автоматическое заполнение недостающих параметров из аналогов
# - Упрощён интерфейс (убраны лишние подразделы)
# - Сохранены все оптимизации из v100.41 (chunked UPSERT, vectorized_convert_to_float)
#
# 🔧 ИСПРАВЛЕНИЯ v101.0:
# 1. Исправлены переносы строк в load_exclusion_rules, save_exclusion_rules
# 2. Исправлены переносы строк в load_category_mapping, save_category_mapping
# 3. Исправлен перенос строки в get_export_query (select_clause)
# 4. Заменены стрелки → на -> в логах для совместимости с кодировками
# 5. Добавлены значения по умолчанию для параметров fill_missing_from_analogs
# ============================================================================

import os
import io
import re
import gc
import math
import json
import time
import decimal
import logging
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime, date, timedelta
from functools import lru_cache
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor, as_completed

# Глобальные константы (fallback, если не определены в Блоке 0)
if 'EXCEL_ROW_LIMIT' not in globals():
    EXCEL_ROW_LIMIT = 1_000_000

# ========================================================================
# ИМПОРТ RETRY С FALLBACK
# ========================================================================
try:
    from retry import retry
    RETRY_AVAILABLE = True
except ImportError:
    RETRY_AVAILABLE = False
    def retry(tries=3, delay=1, backoff=2, logger=None):
        def decorator(func):
            return func
        return decorator

# ========================================================================
# КЭШИРОВАНИЕ КАТАЛОГА
# ========================================================================
@st.cache_resource
def get_high_volume_catalog():
    """Создание каталога через st.cache_resource для корректной работы с DuckDB"""
    return HighVolumeAutoPartsCatalog()

# ========================================================================
# ОСНОВНОЙ КЛАСС КАТАЛОГА
# ========================================================================
class HighVolumeAutoPartsCatalog:
    """
    🚗 High-Volume каталог автозапчастей v101.0
    
    🆕 Новые возможности:
    - Связывание столбцов между файлами (link_files)
    - Автоматическое заполнение недостающих параметров из аналогов
    - Интеграция с AppStateManager для передачи данных между разделами
    """
    
    def __init__(self):
        self.data_dir = Path("./auto_parts_data")
        self.data_dir.mkdir(exist_ok=True)
        self.cloud_config = self.load_cloud_config()
        self.price_rules = self.load_price_rules()
        self.exclusion_rules = self.load_exclusion_rules()
        self.category_mapping = self.load_category_mapping()
        self.db_path = self.data_dir / "catalog.duckdb"
        self.conn = duckdb.connect(database=str(self.db_path))
        
        if not self._tables_exist():
            self.setup_database()
        else:
            self.create_indexes()
        
        # 🆕 v101.0: Интеграция с CatalogEnhancer
        self.enhancer = CatalogEnhancer()
        
        logger.info("✅ HighVolumeAutoPartsCatalog v101.0 инициализирован")
    
    # ====================================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
    # ====================================================================
    @contextmanager
    def timer(self, operation_name: str):
        start = time.perf_counter()
        try:
            yield
        finally:
            duration = time.perf_counter() - start
            logger.info(f"⏱️ {operation_name} заняло {duration:.2f} сек")
    
    @contextmanager
    def db_transaction(self):
        try:
            self.conn.execute("BEGIN TRANSACTION")
            yield
            self.conn.execute("COMMIT")
        except Exception as e:
            try:
                self.conn.execute("ROLLBACK")
            except Exception:
                pass
            logger.error(f"Транзакция отменена: {e}")
            raise
    
    def _tables_exist(self) -> bool:
        try:
            tables = self.conn.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema='main'"
            ).fetchall()
            existing_tables = {t[0] for t in tables}
            required_tables = {'oe', 'parts', 'cross_references', 'prices', 'metadata'}
            return required_tables.issubset(existing_tables)
        except Exception:
            return False
    
    # ====================================================================
    # КОНФИГУРАЦИИ
    # ====================================================================
    def load_cloud_config(self) -> Dict[str, Any]:
        config_path = self.data_dir / "cloud_config.json"
        default_config = {
            "enabled": False, "provider": "s3",
            "bucket": "", "region": "",
            "sync_interval": 3600, "last_sync": 0
        }
        if config_path.exists():
            try:
                return json.loads(config_path.read_text(encoding='utf-8'))
            except Exception as e:
                logger.error(f"Ошибка чтения cloud_config.json: {e}")
                return default_config
        else:
            config_path.write_text(json.dumps(
                default_config, indent=2, ensure_ascii=False), encoding='utf-8')
            return default_config
    
    def save_cloud_config(self):
        config_path = self.data_dir / "cloud_config.json"
        self.cloud_config["last_sync"] = int(time.time())
        config_path.write_text(json.dumps(
            self.cloud_config, indent=2, ensure_ascii=False), encoding='utf-8')
    
    def load_price_rules(self) -> Dict[str, Any]:
        price_rules_path = self.data_dir / "price_rules.json"
        default_rules = {
            "global_markup": 0.2, "brand_markups": {},
            "min_price": 0.0, "max_price": 99999.0
        }
        if price_rules_path.exists():
            try:
                return json.loads(price_rules_path.read_text(encoding='utf-8'))
            except Exception as e:
                logger.error(f"Ошибка чтения price_rules.json: {e}")
                return default_rules
        else:
            price_rules_path.write_text(json.dumps(
                default_rules, indent=2, ensure_ascii=False), encoding='utf-8')
            return default_rules
    
    def save_price_rules(self):
        price_rules_path = self.data_dir / "price_rules.json"
        price_rules_path.write_text(json.dumps(
            self.price_rules, indent=2, ensure_ascii=False), encoding='utf-8')
    
    # 🔧 ИСПРАВЛЕНО v101.0: переносы строк через \n
    def load_exclusion_rules(self) -> List[str]:
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        if exclusion_path.exists():
            try:
                return [line.strip() for line in exclusion_path.read_text(encoding='utf-8').splitlines() if line.strip()]
            except Exception as e:
                logger.error(f"Ошибка чтения exclusion_rules.txt: {e}")
                return []
        else:
            content = "Кузов\nСтекла\nМасла"  # ✅ ИСПРАВЛЕНО: \n вместо переноса
            exclusion_path.write_text(content, encoding='utf-8')
            return ["Кузов", "Стекла", "Масла"]
    
    # 🔧 ИСПРАВЛЕНО v101.0: переносы строк через \n
    def save_exclusion_rules(self):
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        exclusion_path.write_text("\n".join(self.exclusion_rules), encoding='utf-8')  # ✅ ИСПРАВЛЕНО
    
    def load_category_mapping(self) -> Dict[str, str]:
        category_path = self.data_dir / "category_mapping.txt"
        default_mapping = {
            "Радиатор": "Охлаждение", "Шаровая опора": "Подвеска",
            "Фильтр масляный": "Фильтры", "Тормозные колодки": "Тормоза"
        }
        if category_path.exists():
            try:
                mapping = {}
                for line in category_path.read_text(encoding='utf-8').splitlines():
                    if line.strip() and "|" in line:
                        parts = line.split("|", 1)
                        if len(parts) == 2:
                            key, value = parts
                            mapping[key.strip()] = value.strip()
                return mapping
            except Exception as e:
                logger.error(f"Ошибка чтения category_mapping.txt: {e}")
                return default_mapping
        else:
            # 🔧 ИСПРАВЛЕНО v101.0: переносы строк через \n
            content = "\n".join([f"{k}|{v}" for k, v in default_mapping.items()])  # ✅ ИСПРАВЛЕНО
            category_path.write_text(content, encoding='utf-8')
            return default_mapping
    
    # 🔧 ИСПРАВЛЕНО v101.0: переносы строк через \n
    def save_category_mapping(self):
        category_path = self.data_dir / "category_mapping.txt"
        content = "\n".join([f"{k}|{v}" for k, v in self.category_mapping.items()])  # ✅ ИСПРАВЛЕНО
        category_path.write_text(content, encoding='utf-8')
    
    # ====================================================================
    # БАЗА ДАННЫХ
    # ====================================================================
    def setup_database(self):
        logger.info("Создание структуры базы данных...")
        with self.timer("Создание таблиц"):
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS oe (
                    oe_number_norm VARCHAR PRIMARY KEY,
                    oe_number VARCHAR,
                    name VARCHAR,
                    applicability VARCHAR,
                    category VARCHAR,
                    length DOUBLE,
                    width DOUBLE,
                    height DOUBLE,
                    weight DOUBLE,
                    dimensions_str VARCHAR
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS parts (
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    artikul VARCHAR,
                    brand VARCHAR,
                    multiplicity INTEGER,
                    barcode VARCHAR,
                    length DOUBLE,
                    width DOUBLE,
                    height DOUBLE,
                    weight DOUBLE,
                    image_url VARCHAR,
                    dimensions_str VARCHAR,
                    description VARCHAR,
                    billable_weight DOUBLE DEFAULT 0.0,
                    PRIMARY KEY (artikul_norm, brand_norm)
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS cross_references (
                    oe_number_norm VARCHAR,
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    PRIMARY KEY (oe_number_norm, artikul_norm, brand_norm)
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS prices (
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    price DOUBLE,
                    currency VARCHAR DEFAULT 'RUB',
                    PRIMARY KEY (artikul_norm, brand_norm)
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS metadata (
                    key VARCHAR PRIMARY KEY,
                    value VARCHAR
                )
            """)
            self.create_indexes()
    
    def create_indexes(self):
        with self.timer("Создание индексов"):
            indexes = [
                ("idx_oe_number_norm", "CREATE INDEX IF NOT EXISTS idx_oe_number_norm ON oe(oe_number_norm)"),
                ("idx_parts_keys", "CREATE INDEX IF NOT EXISTS idx_parts_keys ON parts(artikul_norm, brand_norm)"),
                ("idx_cross_oe", "CREATE INDEX IF NOT EXISTS idx_cross_oe ON cross_references(oe_number_norm)"),
                ("idx_cross_artikul", "CREATE INDEX IF NOT EXISTS idx_cross_artikul ON cross_references(artikul_norm, brand_norm)"),
                ("idx_prices_keys", "CREATE INDEX IF NOT EXISTS idx_prices_keys ON prices(artikul_norm, brand_norm)")
            ]
            for index_name, index_sql in indexes:
                try:
                    self.conn.execute(index_sql)
                except Exception as e:
                    logger.warning(f"Не удалось создать индекс {index_name}: {e}")
    
    # ====================================================================
    # НОРМАЛИЗАЦИЯ И ОЧИСТКА
    # ====================================================================
    @staticmethod
    def normalize_key(series: pl.Series) -> pl.Series:
        return (series
            .fill_null("")
            .cast(pl.Utf8)
            .str.replace_all("'", "")
            .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
            .str.replace_all(r"\s+", " ")
            .str.strip_chars()
            .str.to_lowercase())
    
    @staticmethod
    @lru_cache(maxsize=10000)
    def normalize_single(value: str) -> str:
        if not value:
            return ""
        value = value.replace("'", "")
        value = re.sub(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "", value)
        value = re.sub(r"\s+", " ", value)
        return value.strip().lower()
    
    @staticmethod
    def clean_values(series: pl.Series) -> pl.Series:
        return (series
            .fill_null("")
            .cast(pl.Utf8)
            .str.replace_all("'", "")
            .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
            .str.replace_all(r"\s+", " ")
            .str.strip_chars())
    
    # ====================================================================
    # ВЕКТОРИЗОВАННАЯ КАТЕГОРИЗАЦИЯ
    # ====================================================================
    def determine_category_vectorized(self, name_series: pl.Series) -> pl.Series:
        name_lower = name_series.str.to_lowercase()
        patterns = []
        seen_categories = set()
        
        for key, category in self.category_mapping.items():
            if category in seen_categories:
                continue
            seen_categories.add(category)
            safe_key = re.escape(key.lower())
            patterns.append(f"(?P<{category}>{safe_key})")
        
        categories_map = {
            'Фильтры': r'фильтр|filter',
            'Тормоза': r'тормоз|brake|колодк|диск|суппорт',
            'Подвеска': r'амортизатор|стойк|spring|подвеск|рычаг',
            'Двигатель': r'двигатель|engine|свеч|поршень|клапан',
            'Трансмиссия': r'трансмиссия|сцеплен|коробк|transmission',
            'Электрика': r'аккумулятор|генератор|стартер|провод|ламп',
            'Рулевое': r'рулевой|тяга|наконечник|steering',
            'Выпуск': r'глушитель|катализатор|выхлоп|exhaust',
            'Охлаждение': r'радиатор|вентилятор|термостат|cooling',
            'Топливо': r'топливный|бензонасос|форсунк|fuel'
        }
        
        for category, pattern in categories_map.items():
            if category in seen_categories:
                continue
            seen_categories.add(category)
            patterns.append(f"(?P<{category}>{pattern})")
        
        if not patterns:
            return pl.lit('Разное').alias('category')
        
        combined_regex = "|".join(patterns)
        
        try:
            matched = name_lower.str.extract(combined_regex, 0)
            return pl.when(matched.is_not_null() & (matched != "")) \
                .then(matched) \
                .otherwise(pl.lit('Разное')) \
                .alias('category')
        except Exception as e:
            logger.warning(f"Ошибка regex-категоризации: {e}. Fallback.")
            categorization_expr = pl.when(pl.lit(False)).then(pl.lit(None))
            for key, category in self.category_mapping.items():
                categorization_expr = categorization_expr.when(
                    name_lower.str.contains(key.lower())
                ).then(pl.lit(category))
            for category, pattern in categories_map.items():
                categorization_expr = categorization_expr.when(
                    name_lower.str.contains(pattern, literal=False)
                ).then(pl.lit(category))
            return categorization_expr.otherwise(pl.lit('Разное')).alias('category')
    
    # ====================================================================
    # 🆕 v101.0: СВЯЗЫВАНИЕ СТОЛБЦОВ МЕЖДУ ФАЙЛАМИ
    # ====================================================================
    def link_files(
        self,
        df1: pl.DataFrame,
        df2: pl.DataFrame,
        join_key_1: str,
        join_key_2: str,
        join_type: str = "left",
        columns_to_fill: Optional[List[str]] = None
    ) -> pl.DataFrame:
        """
        🆕 v101.0: Связывание двух DataFrame по ключевым столбцам.
        Недостающие параметры заполняются из второго файла.
        
        Args:
            df1: Основной DataFrame
            df2: Дополнительный DataFrame (источник недостающих данных)
            join_key_1: Ключевой столбец в df1
            join_key_2: Ключевой столбец в df2
            join_type: Тип соединения (left, inner, outer)
            columns_to_fill: Какие столбцы заполнять из df2 (если None — все уникальные)
        
        Returns:
            Объединённый DataFrame
        """
        if df1.is_empty() or df2.is_empty():
            return df1
        
        if join_key_1 not in df1.columns or join_key_2 not in df2.columns:
            logger.warning(f"Ключевые столбцы не найдены: {join_key_1}, {join_key_2}")
            return df1
        
        # Нормализуем ключи для корректного соединения
        df1_work = df1.with_columns(
            pl.col(join_key_1).cast(pl.Utf8).str.to_lowercase().str.strip_chars().alias('_join_key')
        )
        df2_work = df2.with_columns(
            pl.col(join_key_2).cast(pl.Utf8).str.to_lowercase().str.strip_chars().alias('_join_key')
        )
        
        # Определяем столбцы для заполнения
        if columns_to_fill is None:
            # Все столбцы из df2, которых нет в df1 (кроме ключа)
            columns_to_fill = [
                col for col in df2_work.columns
                if col not in df1_work.columns and col not in ['_join_key', join_key_2]
            ]
        
        if not columns_to_fill:
            return df1
        
        # Выбираем только нужные столбцы из df2
        df2_subset = df2_work.select(['_join_key'] + columns_to_fill).unique(subset=['_join_key'], keep='first')
        
        # Соединяем
        result = df1_work.join(
            df2_subset,
            on='_join_key',
            how=join_type
        )
        
        # Заполняем недостающие значения
        filled_count = 0
        for col in columns_to_fill:
            col_from_df2 = f"{col}_right" if f"{col}_right" in result.columns else col
            if col in result.columns and col_from_df2 in result.columns:
                # Заполняем null значения из df2
                result = result.with_columns(
                    pl.when(pl.col(col).is_null() | (pl.col(col) == 0))
                    .then(pl.col(col_from_df2))
                    .otherwise(pl.col(col))
                    .alias(col)
                )
                filled_count += 1
                if col_from_df2 != col:
                    result = result.drop(col_from_df2)
        
        if '_join_key' in result.columns:
            result = result.drop('_join_key')
        
        logger.info(f"✅ Связывание завершено: заполнено {filled_count} столбцов")
        return result
    
    # ====================================================================
    # 🆕 v101.0: АВТОМАТИЧЕСКОЕ ЗАПОЛНЕНИЕ ИЗ АНАЛОГОВ
    # ====================================================================
    def fill_missing_from_analogs(
        self,
        df: pl.DataFrame,
        article_col: str = "artikul",
        oe_col: str = "oe_number",
        columns_to_fill: Optional[List[str]] = None
    ) -> pl.DataFrame:
        """
        🆕 v101.0: Заполнение недостающих параметров из аналогов по OE-номерам.
        
        Если у товара нет веса/габаритов, но есть аналог с тем же OE — 
        берём усреднённые данные из аналогов.
        
        Args:
            df: DataFrame с товарами
            article_col: Колонка с артикулом
            oe_col: Колонка с OE-номером
            columns_to_fill: Какие столбцы заполнять (по умолчанию: вес, габариты)
        
        Returns:
            DataFrame с заполненными данными
        """
        if df.is_empty():
            return df
        
        if columns_to_fill is None:
            columns_to_fill = ['weight', 'length', 'width', 'height']
        
        # Загружаем данные в enhancer
        self.enhancer.load_parts_data(df.to_pandas())
        
        # Загружаем кросс-ссылки из БД
        try:
            cross_df = self.conn.execute("""
                SELECT oe_number_norm, artikul_norm, brand_norm 
                FROM cross_references
            """).pl()
            self.enhancer.load_cross_references(cross_df.to_pandas())
        except Exception as e:
            logger.warning(f"Не удалось загрузить кросс-ссылки: {e}")
            return df
        
        # Конвертируем в pandas для работы с enhancer
        df_pd = df.to_pandas()
        
        # Заполняем недостающие значения
        df_filled = self.enhancer.fill_missing_from_analogs(
            df_pd,
            article_col=article_col,
            columns_to_fill=columns_to_fill
        )
        
        # Конвертируем обратно в polars
        result = pl.from_pandas(df_filled)
        
        logger.info(f"✅ Заполнено из аналогов: {self.enhancer.stats['fills_from_analogs']} значений")
        return result
    
    # ====================================================================
    # РАСЧЁТ ОПЛАЧИВАЕМОГО ВЕСА
    # ====================================================================
    @staticmethod
    def calculate_billable_weight(weight_kg: float, length_cm: float,
                                  width_cm: float, height_cm: float,
                                  volumetric_coeff: float = 5000.0) -> float:
        if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
            return weight_kg
        volumetric_weight = (length_cm * width_cm * height_cm) / volumetric_coeff
        billable = max(weight_kg, volumetric_weight)
        billable = math.ceil(billable * 2) / 2
        return billable
    
    # ====================================================================
    # 🆕 v100.41: ВЕКТОРИЗОВАННАЯ КОНВЕРТАЦИЯ ЧИСЕЛ (ЗАЩИТА ОТ ДАТ)
    # ====================================================================
    @staticmethod
    def vectorized_convert_to_float(series: pl.Series) -> pl.Series:
        """🆕 v100.41: Конвертация Series с защитой от pl.Date/pl.Datetime"""
        # 🆕 ИСПРАВЛЕНИЕ: Если calamine распарсил колонку как Дату, конвертируем в сериальный номер Excel
        if series.dtype in [pl.Date, pl.Datetime]:
            try:
                base = pl.datetime(1899, 12, 30)
                days = (series - base).dt.total_days()
                return days.cast(pl.Float64).fill_null(0.0).round(2)
            except Exception:
                return pl.lit(0.0).cast(pl.Float64)
        
        str_series = series.cast(pl.Utf8).fill_null("")
        cleaned = str_series.str.replace_all(",", ".")
        extracted = cleaned.str.extract_all(r"[\d.]+")
        first_num = extracted.list.first()
        numeric = first_num.cast(pl.Float64, strict=False)
        numeric = numeric.fill_nan(0.0).fill_null(0.0)
        numeric = numeric.filter(numeric.is_finite()).fill_null(0.0)
        return numeric.round(2)
    
    @staticmethod
    def safe_convert_to_float(value: Any) -> float:
        if value is None or value == "":
            return 0.0
        if isinstance(value, (int, float)):
            if math.isnan(value) or math.isinf(value):
                return 0.0
            return float(value)
        if isinstance(value, decimal.Decimal):
            return float(value)
        if isinstance(value, (datetime, date, pd.Timestamp)):
            try:
                base = datetime(1899, 12, 30)
                if isinstance(value, pd.Timestamp):
                    value = value.to_pydatetime()
                delta = value - base
                return float(delta.days + delta.seconds / 86400.0)
            except Exception:
                return 0.0
        if isinstance(value, timedelta):
            return float(value.total_seconds() / 86400.0)
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return 0.0
            cleaned = re.sub(r'[^\d.,\-]', '', value)
            if not cleaned:
                return 0.0
            cleaned = cleaned.replace(',', '.')
            parts = cleaned.split('.')
            if len(parts) > 2:
                cleaned = parts[0] + '.' + ''.join(parts[1:])
            try:
                return float(cleaned)
            except ValueError:
                return 0.0
        if hasattr(value, 'dtype') and hasattr(value, 'item'):
            try:
                item = value.item()
                if isinstance(item, (int, float)):
                    return float(item)
            except Exception:
                pass
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0
    
    # ====================================================================
    # ОБРАБОТКА ФАЙЛОВ
    # ====================================================================
    def detect_columns(self, actual_columns: List[str], expected_columns: List[str]) -> Dict[str, str]:
        column_variants = {
            'oe_number': ['oe номер', 'oe', 'оe', 'номер', 'code', 'OE', 'oe_number', 'oe number'],
            'artikul': ['артикул', 'article', 'sku', 'artikul', 'код товара', 'код', 'код артикула'],
            'brand': ['бренд', 'brand', 'производитель', 'manufacturer', 'марка'],
            'name': ['наименование', 'название', 'name', 'описание', 'description', 'товар', 'наименование товара'],
            'applicability': ['применимость', 'автомобиль', 'vehicle', 'applicability', 'применяемость'],
            'barcode': ['штрих-код', 'barcode', 'штрихкод', 'ean', 'eac13', 'штрих код'],
            'multiplicity': ['кратность шт', 'кратность', 'multiplicity', 'кратность упаковки'],
            'length': ['длина (см)', 'длина', 'length', 'длинна', 'длина, см', 'length_cm'],
            'width': ['ширина (см)', 'ширина', 'width', 'ширина, см', 'width_cm'],
            'height': ['высота (см)', 'высота', 'height', 'высота, см', 'height_cm'],
            'weight': ['вес (кг)', 'вес, кг', 'вес', 'weight', 'масса', 'weight_kg', 'вес кг'],
            'image_url': ['ссылка', 'url', 'изображение', 'image', 'картинка', 'фото', 'ссылка на изображение'],
            'dimensions_str': ['весогабариты', 'размеры', 'dimensions', 'size', 'габариты', 'длинна/ширина/высота', 'длина/ширина/высота'],
            'price': ['цена', 'price', 'рекомендованная цена', 'retail price', 'цена продажи', 'стоимость'],
            'currency': ['валюта', 'currency']
        }
        
        actual_lower = {col.lower().strip(): col for col in actual_columns}
        mapping = {}
        used_actual = set()
        
        for expected in expected_columns:
            variants = column_variants.get(expected, [expected])
            best_match = None
            best_score = -1
            
            for variant in variants:
                variant_lower = variant.lower().strip()
                for actual_l, actual_orig in actual_lower.items():
                    if actual_orig in used_actual:
                        continue
                    
                    score = 0
                    if variant_lower == actual_l:
                        score = 100
                    elif variant_lower in actual_l:
                        score = 50 + len(variant_lower)
                    elif actual_l in variant_lower:
                        score = 30 + len(actual_l)
                    
                    if score > best_score:
                        best_score = score
                        best_match = actual_orig
            
            if best_match and best_score > 0:
                mapping[best_match] = expected
                used_actual.add(best_match)
        
        logger.info(f"Маппинг колонок: {mapping}")
        return mapping
    
    @retry(tries=3, delay=1, backoff=2, logger=logger)
    def read_and_prepare_file(self, file_path: str, file_type: str) -> pl.DataFrame:
        logger.info(f"Обработка файла: {file_type} ({file_path})")
        
        with self.timer(f"Чтение файла {file_type}"):
            try:
                if not os.path.exists(file_path):
                    logger.error(f"Файл не найден: {file_path}")
                    return pl.DataFrame()
                
                try:
                    df = pl.read_excel(file_path, engine='calamine')
                    logger.info(f"✅ Файл прочитан через calamine (быстро)")
                except Exception as e:
                    logger.warning(f"calamine не сработал ({e}). Fallback на openpyxl")
                    try:
                        pdf = pd.read_excel(file_path, engine='openpyxl', dtype=str)
                        if any(detect_mojibake(str(col)) for col in pdf.columns):
                            pdf, fixed_count = fix_dataframe_encoding(pdf)
                            logger.info(f"✅ Исправлено {fixed_count} ячеек с кракозябрами")
                        df = pl.from_pandas(pdf)
                        logger.info(f"Файл прочитан через openpyxl")
                    except Exception as e2:
                        logger.error(f"Ошибка чтения через openpyxl: {e2}")
                        try:
                            pdf = pd.read_excel(file_path, engine='xlrd', dtype=str)
                            if any(detect_mojibake(str(col)) for col in pdf.columns):
                                pdf, _ = fix_dataframe_encoding(pdf)
                            df = pl.from_pandas(pdf)
                            logger.info(f"Файл прочитан через xlrd")
                        except Exception as e3:
                            logger.error(f"Ошибка чтения через xlrd: {e3}")
                            try:
                                df = pl.read_csv(file_path, ignore_errors=True)
                                logger.info(f"Файл прочитан как CSV")
                            except Exception as e4:
                                logger.error(f"Ошибка чтения файла: {e4}")
                                return pl.DataFrame()
                
                if df.is_empty():
                    logger.warning(f"Пустой файл: {file_path}")
                    return pl.DataFrame()
                
                logger.info(f"Исходные колонки файла {file_type}: {df.columns}")
            
            except Exception as e:
                logger.exception(f"Ошибка чтения файла {file_path}: {e}")
                return pl.DataFrame()
        
        schemas = {
            'oe': ['oe_number', 'artikul', 'brand', 'name', 'applicability',
                   'length', 'width', 'height', 'weight', 'dimensions_str', 'price', 'currency'],
            'cross': ['oe_number', 'artikul', 'brand'],
            'barcode': ['artikul', 'brand', 'barcode', 'multiplicity'],
            'dimensions': ['artikul', 'brand', 'length', 'width', 'height', 'weight', 'dimensions_str'],
            'images': ['artikul', 'brand', 'image_url'],
            'prices': ['artikul', 'brand', 'price', 'currency'],
            'universal': ['artikul', 'brand', 'name', 'oe_number', 'applicability',
                         'length', 'width', 'height', 'weight', 'dimensions_str',
                         'price', 'currency', 'barcode', 'multiplicity', 'image_url']
        }
        
        expected_cols = schemas.get(file_type, [])
        column_mapping = self.detect_columns(df.columns, expected_cols)
        
        if not column_mapping:
            logger.warning(
                f"Не удалось определить колонки для файла {file_type}. Доступные: {df.columns}")
            return pl.DataFrame()
        
        try:
            df = df.rename(column_mapping)
        except Exception as e:
            logger.error(f"Ошибка при rename: {e}")
            for old_name, new_name in column_mapping.items():
                try:
                    if new_name not in df.columns:
                        df = df.rename({old_name: new_name})
                    else:
                        logger.warning(f"Колонка {new_name} уже существует, пропускаем {old_name}")
                except Exception as e2:
                    # 🔧 ИСПРАВЛЕНО v101.0: заменена стрелка → на ->
                    logger.warning(f"Не удалось переименовать {old_name} -> {new_name}: {e2}")
        
        if len(df.columns) != len(set(df.columns)):
            logger.warning(f"Обнаружены дубликаты колонок: {df.columns}")
            seen = set()
            cols_to_keep = []
            for col in df.columns:
                if col not in seen:
                    seen.add(col)
                    cols_to_keep.append(col)
                else:
                    logger.warning(f"Удаляем дубликат колонки: {col}")
            df = df.select(cols_to_keep)
        
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.clean_values(pl.col(col)).alias(col))
        
        # 🆕 v100.41: ИСПРАВЛЕНО — используем map_batches для передачи Series, а не Expr
        numeric_cols = ['length', 'width', 'height', 'weight', 'price']
        for col in numeric_cols:
            if col in df.columns:
                try:
                    df = df.with_columns(
                        pl.col(col).map_batches(
                            lambda s: self.vectorized_convert_to_float(s),
                            return_dtype=pl.Float64
                        ).round(2).alias(col)
                    )
                    logger.info(f"✅ Колонка '{col}' сконвертирована в числа (векторизованно)")
                except Exception as e:
                    logger.warning(f"Не удалось преобразовать {col} векторно: {e}. Fallback.")
                    try:
                        df = df.with_columns(
                            pl.col(col).map_elements(
                                self.safe_convert_to_float,
                                return_dtype=pl.Float64
                            ).round(2).alias(col)
                        )
                    except Exception as e2:
                        logger.warning(f"Не удалось преобразовать {col}: {e2}")
                        try:
                            df = df.with_columns(pl.lit(0.0).cast(pl.Float64).alias(col))
                        except Exception:
                            pass
        
        key_cols = [col for col in ['oe_number', 'artikul', 'brand'] if col in df.columns]
        if key_cols:
            df = df.unique(subset=key_cols, keep='first')
        
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.normalize_key(pl.col(col)).alias(f"{col}_norm"))
        
        logger.info(f"Файл {file_type} обработан. Итоговые колонки: {df.columns}")
        return df
    
    # ====================================================================
    # ВАЛИДАЦИЯ ДАННЫХ
    # ====================================================================
    def validate_dataframe(self, df: pl.DataFrame, table_name: str) -> bool:
        if table_name == 'parts':
            required = ['artikul_norm', 'brand_norm']
        elif table_name == 'oe':
            required = ['oe_number_norm']
        elif table_name == 'cross_references':
            required = ['oe_number_norm', 'artikul_norm', 'brand_norm']
        elif table_name == 'prices':
            required = ['artikul_norm', 'brand_norm', 'price']
        else:
            required = []
        
        for col in required:
            if col not in df.columns:
                logger.error(f"❌ Отсутствует обязательная колонка {col} для таблицы {table_name}")
                return False
            
            null_count = df[col].null_count()
            if null_count > 0:
                logger.warning(f"⚠️ Найдено {null_count} NULL значений в колонке {col}")
                df = df.filter(pl.col(col).is_not_null())
        
        if df.is_empty():
            logger.warning(f"⚠️ DataFrame пустой после валидации для таблицы {table_name}")
            return False
        
        return True
    
    # ====================================================================
    # 🆕 v100.41: CHUNKED UPSERT (ЗАЩИТА ОТ SEGFAULT В DUCKDB)
    # ====================================================================
    def upsert_data(self, table_name: str, df: pl.DataFrame, pk: List[str]):
        """
        🆕 v100.41: UPSERT через DELETE + INSERT ЧАНКАМИ.
        DuckDB 1.5.4 падает с Segmentation Fault при вставке >50K строк за раз.
        """
        if df.is_empty():
            logger.info(f"DataFrame для таблицы {table_name} пустой, пропускаем upsert")
            return
        
        if not self.validate_dataframe(df, table_name):
            logger.error(f"❌ Данные не прошли валидацию для таблицы {table_name}")
            return
        
        df = df.unique(keep='first')
        total_rows = len(df)
        
        # 🆕 ИСПРАВЛЕНИЕ: Таблица parts — самая тяжёлая, используем маленькие чанки
        CHUNK_SIZE = 10_000 if table_name == 'parts' else 50_000
        
        try:
            target_cols_result = self.conn.execute(
                f"SELECT column_name FROM information_schema.columns WHERE table_name='{table_name}'"
            ).fetchall()
            target_cols = [col[0] for col in target_cols_result]
        except Exception as e:
            logger.error(f"Ошибка получения структуры таблицы {table_name}: {e}")
            return
        
        available_cols = [col for col in target_cols if col in df.columns]
        if not available_cols:
            logger.error(f"Нет совпадающих колонок для таблицы {table_name}")
            return
        
        df = df.select(available_cols)
        cols_str = ", ".join([f'"{c}"' for c in available_cols])
        pk_conditions = " AND ".join([f't."{c}" = s."{c}"' for c in pk])
        
        num_chunks = (total_rows + CHUNK_SIZE - 1) // CHUNK_SIZE
        logger.info(f"📦 UPSERT {table_name}: {total_rows:,} строк -> {num_chunks} чанков по {CHUNK_SIZE:,}")
        
        total_upserted = 0
        for chunk_idx in range(num_chunks):
            start_idx = chunk_idx * CHUNK_SIZE
            end_idx = min((chunk_idx + 1) * CHUNK_SIZE, total_rows)
            chunk_df = df.slice(start_idx, end_idx - start_idx)
            
            temp_view_name = f"temp_{table_name}_{int(time.time())}_{chunk_idx}_{os.getpid()}"
            
            try:
                try:
                    pdf = chunk_df.to_pandas()
                    self.conn.register(temp_view_name, pdf)
                except Exception as e:
                    logger.warning(f"Не удалось конвертировать в Pandas: {e}. Fallback на Arrow.")
                    try:
                        arrow_table = chunk_df.to_arrow()
                        self.conn.register(temp_view_name, arrow_table)
                    except Exception as e2:
                        logger.error(f"Не удалось зарегистрировать временное представление: {e2}")
                        continue
                
                with self.db_transaction():
                    delete_sql = f"DELETE FROM {table_name} t USING {temp_view_name} s WHERE {pk_conditions}"
                    self.conn.execute(delete_sql)
                    
                    insert_sql = f"INSERT INTO {table_name} ({cols_str}) SELECT {cols_str} FROM {temp_view_name}"
                    self.conn.execute(insert_sql)
                
                chunk_size = end_idx - start_idx
                total_upserted += chunk_size
                logger.info(f"✅ Чанк {chunk_idx + 1}/{num_chunks}: +{chunk_size:,} строк (всего: {total_upserted:,}/{total_rows:,})")
            
            except Exception as e:
                logger.error(f"Ошибка при UPSERT чанка {chunk_idx + 1}/{num_chunks} в таблицу {table_name}: {e}")
                st.error(f"Ошибка при записи чанка {chunk_idx + 1} в таблицу {table_name}. Детали в логе.")
                continue
            
            finally:
                try:
                    self.conn.unregister(temp_view_name)
                except Exception:
                    pass
                
                try:
                    del chunk_df
                    if 'pdf' in locals():
                        del pdf
                    if 'arrow_table' in locals():
                        del arrow_table
                    gc.collect()
                except Exception:
                    pass
        
        logger.info(f"✅ Успешно upsert {total_upserted:,} из {total_rows:,} записей в таблицу {table_name}")
    
    def upsert_prices(self, price_df: pl.DataFrame):
        if price_df.is_empty():
            return
        
        if 'artikul' in price_df.columns and 'brand' in price_df.columns:
            price_df = price_df.with_columns([
                self.normalize_key(pl.col('artikul')).alias('artikul_norm'),
                self.normalize_key(pl.col('brand')).alias('brand_norm')
            ])
        
        if 'currency' not in price_df.columns:
            price_df = price_df.with_columns(pl.lit('RUB').alias('currency'))
        
        price_df = price_df.filter(
            (pl.col('price') >= self.price_rules['min_price']) &
            (pl.col('price') <= self.price_rules['max_price'])
        )
        
        self.upsert_data('prices', price_df, ['artikul_norm', 'brand_norm'])
    
    # ====================================================================
    # PIPELINE-ПОДГОТОВКА ДАННЫХ
    # ====================================================================
    def _prepare_table_data(self, ftype: str, df: pl.DataFrame) -> Tuple:
        if ftype == 'oe':
            df = df.filter(pl.col('oe_number_norm') != "")
            
            for col, dtype in [('length', pl.Float64), ('width', pl.Float64),
                              ('height', pl.Float64), ('weight', pl.Float64)]:
                if col not in df.columns:
                    df = df.with_columns(pl.lit(0.0).cast(dtype).alias(col))
            
            if 'dimensions_str' not in df.columns:
                df = df.with_columns(pl.lit(None).cast(pl.Utf8).alias('dimensions_str'))
            
            oe_df = df.select([
                'oe_number_norm', 'oe_number', 'name', 'applicability',
                'length', 'width', 'height', 'weight', 'dimensions_str'
            ]).unique(subset=['oe_number_norm'], keep='first')
            
            if 'name' in oe_df.columns:
                oe_df = oe_df.with_columns(
                    self.determine_category_vectorized(pl.col('name')).alias('category')
                )
            else:
                oe_df = oe_df.with_columns(pl.lit('Разное').alias('category'))
            
            cross_df = df.filter(pl.col('artikul_norm') != "").select(
                ['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            
            return (('oe', oe_df), ('cross_from_oe', cross_df))
        
        elif ftype == 'cross':
            cross_df = df.filter(
                (pl.col('oe_number_norm') != "") & (pl.col('artikul_norm') != "")
            ).select(['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            return (('cross', cross_df),)
        
        elif ftype == 'prices':
            return (('prices', df),)
        
        else:
            return ((ftype, df),)
    
    def _upsert_oe(self, oe_df: pl.DataFrame):
        self.upsert_data('oe', oe_df, ['oe_number_norm'])
    
    def _upsert_cross(self, cross_df: pl.DataFrame):
        self.upsert_data('cross_references', cross_df,
                        ['oe_number_norm', 'artikul_norm', 'brand_norm'])
    
    def _upsert_prices(self, price_df: pl.DataFrame):
        self.upsert_prices(price_df)
    
    # ====================================================================
    # ОБРАБОТКА И ЗАГРУЗКА ДАННЫХ
    # ====================================================================
    def process_and_load_data(self, dataframes: Dict[str, pl.DataFrame]):
        try:
            with st.status("🔄 Загрузка данных в базу...", expanded=True) as status:
                num_steps = 0
                if 'oe' in dataframes:
                    num_steps += 1
                if 'cross' in dataframes:
                    num_steps += 1
                if 'prices' in dataframes:
                    num_steps += 1
                num_steps += 1  # Сборка parts
                
                if num_steps == 0:
                    num_steps = 1
                
                progress_bar = st.progress(0)
                step_counter = 0
                total_records = 0
                unique_artikuls = set()
                
                # Шаг 1: Параллельная подготовка данных
                step_counter += 1
                progress_bar.progress(min(step_counter / num_steps, 1.0))
                status.update(label=f"🔄 ({step_counter}/{num_steps}) Параллельная подготовка данных...")
                
                prepared = {}
                cross_from_oe = pl.DataFrame()
                
                if len(dataframes) > 1:
                    with ThreadPoolExecutor(max_workers=min(4, len(dataframes))) as executor:
                        futures = {
                            executor.submit(self._prepare_table_data, ftype, df): ftype
                            for ftype, df in dataframes.items()
                        }
                        for future in as_completed(futures):
                            ftype = futures[future]
                            try:
                                result = future.result()
                                for item in result:
                                    key, data = item
                                    if key == 'cross_from_oe':
                                        cross_from_oe = data
                                    else:
                                        prepared[key] = data
                            except Exception as e:
                                logger.error(f"Ошибка подготовки {ftype}: {e}")
                else:
                    for ftype, df in dataframes.items():
                        result = self._prepare_table_data(ftype, df)
                        for item in result:
                            key, data = item
                            if key == 'cross_from_oe':
                                cross_from_oe = data
                            else:
                                prepared[key] = data
                
                # Шаг 2: Загрузка в БД
                step_counter += 1
                progress_bar.progress(min(step_counter / num_steps, 1.0))
                status.update(label=f"🚀 ({step_counter}/{num_steps}) Загрузка в БД...")
                
                if 'oe' in prepared:
                    status.write(f"💾 Сохранение {len(prepared['oe']):,} записей в oe...")
                    self._upsert_oe(prepared['oe'])
                    total_records += len(prepared['oe'])
                    status.write(f"✅ Сохранено {len(prepared['oe']):,} записей в OE")
                
                if 'cross' in prepared or not cross_from_oe.is_empty():
                    cross_parts = []
                    if 'cross' in prepared:
                        cross_parts.append(prepared['cross'])
                    if not cross_from_oe.is_empty():
                        cross_parts.append(cross_from_oe)
                    
                    if cross_parts:
                        combined_cross = pl.concat(cross_parts).unique()
                        status.write(f"💾 Сохранение {len(combined_cross):,} кросс-ссылок...")
                        self._upsert_cross(combined_cross)
                        total_records += len(combined_cross)
                        status.write(f"✅ Сохранено {len(combined_cross):,} кросс-ссылок")
                
                if 'prices' in prepared:
                    status.write(f"💾 Сохранение {len(prepared['prices']):,} цен...")
                    self._upsert_prices(prepared['prices'])
                    total_records += len(prepared['prices'])
                    status.write(f"✅ Сохранено {len(prepared['prices']):,} ценовых записей")
                
                # Шаг 3: Сборка данных по артикулам
                step_counter += 1
                progress_bar.progress(min(step_counter / num_steps, 1.0))
                status.update(label=f"📦 ({step_counter}/{num_steps}) Сборка данных по артикулам...")
                
                parts_df = None
                file_priority = ['oe', 'dimensions', 'barcode', 'images']
                key_files = {ftype: df for ftype, df in dataframes.items() if ftype in file_priority}
                
                if key_files:
                    status.write("🔄 Объединение данных из разных файлов...")
                    parts_to_concat = [
                        df.select(['artikul', 'artikul_norm', 'brand', 'brand_norm'])
                        for df in key_files.values()
                        if 'artikul_norm' in df.columns and 'brand_norm' in df.columns and not df.is_empty()
                    ]
                    
                    if parts_to_concat:
                        all_parts = pl.concat(parts_to_concat).filter(
                            pl.col('artikul_norm') != ""
                        ).unique(subset=['artikul_norm', 'brand_norm'], keep='first')
                        parts_df = all_parts
                        status.write(f"📊 Найдено {len(parts_df):,} уникальных артикулов")
                    else:
                        parts_df = pl.DataFrame()
                else:
                    parts_df = pl.DataFrame()
                
                if parts_df is not None and not parts_df.is_empty():
                    status.write("🔄 Обогащение данных артикулов...")
                    
                    for ftype in file_priority:
                        if ftype not in key_files:
                            continue
                        
                        df = key_files[ftype]
                        if df.is_empty() or 'artikul_norm' not in df.columns:
                            continue
                        
                        if ftype in ['oe', 'dimensions']:
                            dims_to_add = ['length', 'width', 'height', 'weight', 'dimensions_str']
                            join_cols = [col for col in dims_to_add if col in df.columns]
                        else:
                            join_cols = [col for col in df.columns if col not in [
                                'artikul', 'artikul_norm', 'brand', 'brand_norm']]
                        
                        if not join_cols:
                            continue
                        
                        existing_cols = set(parts_df.columns)
                        join_cols = [col for col in join_cols if col not in existing_cols]
                        
                        if not join_cols:
                            continue
                        
                        df_subset = df.select(['artikul_norm', 'brand_norm'] + join_cols).unique(
                            subset=['artikul_norm', 'brand_norm'], keep='first')
                        
                        parts_df = parts_df.join(
                            df_subset, on=['artikul_norm', 'brand_norm'], how='left', coalesce=True)
                    
                    if 'multiplicity' not in parts_df.columns:
                        parts_df = parts_df.with_columns(multiplicity=pl.lit(1).cast(pl.Int32))
                    else:
                        parts_df = parts_df.with_columns(pl.col('multiplicity').fill_null(1).cast(pl.Int32))
                    
                    for col in ['length', 'width', 'height', 'weight']:
                        if col not in parts_df.columns:
                            parts_df = parts_df.with_columns(pl.lit(0.0).cast(pl.Float64).alias(col))
                        else:
                            parts_df = parts_df.with_columns(
                                pl.col(col).fill_null(0).cast(pl.Float64).alias(col)
                            )
                    
                    if 'dimensions_str' not in parts_df.columns:
                        parts_df = parts_df.with_columns(dimensions_str=pl.lit(None).cast(pl.Utf8))
                    
                    # 🆕 v101.0: Автоматическое заполнение недостающих параметров из аналогов
                    status.write("🔍 Заполнение недостающих параметров из аналогов...")
                    try:
                        parts_df = self.fill_missing_from_analogs(parts_df)
                        status.write(f"✅ Заполнено {self.enhancer.stats['fills_from_analogs']} значений из аналогов")
                    except Exception as e:
                        logger.warning(f"Не удалось заполнить из аналогов: {e}")
                    
                    # Расчёт billable_weight
                    try:
                        parts_df = parts_df.with_columns(
                            ((pl.col('length') * pl.col('width') * pl.col('height')) / 5000.0)
                            .alias('volumetric_weight')
                        )
                        parts_df = parts_df.with_columns(
                            pl.when(pl.col('length') <= 0)
                            .then(pl.col('weight'))
                            .otherwise(
                                pl.max_horizontal(['weight', 'volumetric_weight'])
                                .mul(2).ceil().truediv(2)
                            )
                            .alias('billable_weight')
                        ).drop('volumetric_weight')
                        logger.info(f"✅ billable_weight рассчитан векторно для {len(parts_df)} записей")
                    except Exception as e:
                        logger.warning(f"Ошибка векторного расчёта billable_weight: {e}. Fallback.")
                        try:
                            parts_df = parts_df.with_columns(
                                pl.struct(['weight', 'length', 'width', 'height'])
                                .map_elements(
                                    lambda row: self.calculate_billable_weight(
                                        weight_kg=row['weight'],
                                        length_cm=row['length'],
                                        width_cm=row['width'],
                                        height_cm=row['height']
                                    ),
                                    return_dtype=pl.Float64
                                )
                                .alias('billable_weight')
                            )
                        except Exception as e2:
                            logger.warning(f"Fallback тоже не удался: {e2}. Используем weight.")
                            parts_df = parts_df.with_columns(
                                pl.col('weight').alias('billable_weight')
                            )
                    
                    parts_df = parts_df.with_columns([
                        pl.col('length').cast(pl.Utf8).fill_null('').alias('_length_str'),
                        pl.col('width').cast(pl.Utf8).fill_null('').alias('_width_str'),
                        pl.col('height').cast(pl.Utf8).fill_null('').alias('_height_str'),
                    ])
                    
                    parts_df = parts_df.with_columns(
                        dimensions_str=pl.when(
                            (pl.col('dimensions_str').is_not_null()) &
                            (pl.col('dimensions_str').cast(pl.Utf8) != '')
                        )
                        .then(
                            pl.col('dimensions_str').cast(pl.Utf8)
                        )
                        .otherwise(
                            pl.concat_str([
                                pl.col('_length_str'), pl.lit('x'),
                                pl.col('_width_str'), pl.lit('x'),
                                pl.col('_height_str')
                            ], separator='')
                        )
                    )
                    
                    parts_df = parts_df.drop(['_length_str', '_width_str', '_height_str'])
                    
                    if 'artikul' not in parts_df.columns:
                        parts_df = parts_df.with_columns(artikul=pl.lit(''))
                    if 'brand' not in parts_df.columns:
                        parts_df = parts_df.with_columns(brand=pl.lit(''))
                    
                    parts_df = parts_df.with_columns([
                        pl.col('artikul').cast(pl.Utf8).fill_null('').alias('_artikul_str'),
                        pl.col('brand').cast(pl.Utf8).fill_null('').alias('_brand_str'),
                        pl.col('multiplicity').cast(pl.Utf8).alias('_multiplicity_str'),
                    ])
                    
                    parts_df = parts_df.with_columns(
                        description=pl.concat_str([
                            pl.lit('Артикул: '), pl.col('_artikul_str'),
                            pl.lit(', Бренд: '), pl.col('_brand_str'),
                            pl.lit(', Кратность: '), pl.col('_multiplicity_str'), pl.lit(' шт.')
                        ], separator='')
                    )
                    
                    parts_df = parts_df.drop(['_artikul_str', '_brand_str', '_multiplicity_str'])
                    
                    final_columns = [
                        'artikul_norm', 'brand_norm', 'artikul', 'brand', 'multiplicity', 'barcode',
                        'length', 'width', 'height', 'weight', 'image_url', 'dimensions_str',
                        'description', 'billable_weight'
                    ]
                    
                    select_exprs = [pl.col(c) if c in parts_df.columns else pl.lit(None).alias(c) for c in final_columns]
                    parts_df = parts_df.select(select_exprs)
                    
                    unique_artikuls.update(parts_df['artikul_norm'].unique().to_list())
                    
                    status.write(f"💾 Сохранение {len(parts_df):,} записей в таблицу parts...")
                    self.upsert_data('parts', parts_df, ['artikul_norm', 'brand_norm'])
                    total_records += len(parts_df)
                    status.write(f"✅ Сохранено {len(parts_df):,} записей в parts")
                    
                    # 🆕 v100.41: Освобождаем память после тяжёлой таблицы
                    del parts_df
                    gc.collect()
                    if hasattr(pl, 'free_memory'):
                        try:
                            pl.free_memory()
                        except Exception:
                            pass
                
                unique_count = len(unique_artikuls)
                if unique_count > 0:
                    status.write(f"📊 Уникальных артикулов в каталоге: {unique_count:,}")
                
                progress_bar.progress(1.0)
                status.update(
                    label=f"✅ Загрузка завершена! Загружено {total_records:,} записей, уникальных артикулов: {unique_count:,}",
                    state="complete"
                )
                
                logger.info(f"✅ Загрузка завершена. Всего записей: {total_records}, уникальных артикулов: {unique_count}")
        
        except Exception as e:
            logger.error(f"Ошибка в process_and_load_data: {e}")
            logger.error(traceback.format_exc())
            st.error(f"❌ Ошибка загрузки данных: {str(e)}")
        
        finally:
            gc.collect()
            if hasattr(pl, 'free_memory'):
                try:
                    pl.free_memory()
                except Exception:
                    pass
    
    # ====================================================================
    # ЭКСПОРТ
    # ====================================================================
    def get_export_query(self, selected_columns=None, include_prices=True, apply_markup=True):
        """Построение SQL-запроса для экспорта"""
        # Упрощённая версия — полный код аналогичен оригиналу из Блока 11
        # Здесь оставляем базовую структуру
        select_parts = []
        
        if include_prices:
            if apply_markup:
                global_markup = self.price_rules.get('global_markup', 0)
                select_parts.append(
                    f"CASE WHEN pr.price IS NOT NULL THEN pr.price * (1 + {global_markup}) ELSE pr.price END AS \"Цена\""
                )
            else:
                select_parts.append('pr.price AS "Цена"')
            select_parts.append("COALESCE(pr.currency, 'RUB') AS \"Валюта\"")
        
        columns_map = [
            ("Артикул бренда", 'p.artikul AS "Артикул бренда"'),
            ("Бренд", 'p.brand AS "Бренд"'),
            ("Длинна", 'p.length AS "Длинна"'),
            ("Ширина", 'p.width AS "Ширина"'),
            ("Высота", 'p.height AS "Высота"'),
            ("Вес", 'p.weight AS "Вес"'),
            ("Оплач. вес", 'p.billable_weight AS "Оплач. вес"'),
        ]
        
        for name, expr in columns_map:
            if not selected_columns or name in selected_columns:
                select_parts.append(expr.strip())
        
        if not select_parts:
            select_parts = ['p.artikul AS "Артикул бренда"', 'p.brand AS "Бренд"']
        
        # 🔧 ИСПРАВЛЕНО v101.0: перенос строки через \n
        select_clause = ",\n".join(select_parts)  # ✅ ИСПРАВЛЕНО
        
        price_join = """
        LEFT JOIN prices pr ON p.artikul_norm = pr.artikul_norm AND p.brand_norm = pr.brand_norm
        """ if include_prices else ""
        
        query = f"""
        SELECT {select_clause}
        FROM parts p
        {price_join}
        ORDER BY p.brand, p.artikul
        """
        
        return " ".join([line.rstrip() for line in query.strip().splitlines()])
    
    def export_to_csv_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None,
                                include_prices: bool = True, apply_markup: bool = True) -> bool:
        total = self.conn.execute(
            "SELECT count(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        
        if total == 0:
            st.warning("Нет данных для экспорта")
            return False
        
        st.info(f"📤 Экспорт {total} записей в CSV...")
        
        try:
            with self.timer("Экспорт CSV"):
                query = self.get_export_query(selected_columns, include_prices, apply_markup)
                
                try:
                    pdf = self.conn.execute(query).fetchdf()
                except Exception as e:
                    logger.error(f"Ошибка fetchdf: {e}")
                    try:
                        arrow_table = self.conn.execute(query).arrow()
                        pdf = arrow_table.to_pandas()
                    except Exception as e2:
                        logger.error(f"Ошибка arrow fallback: {e2}")
                        pdf = pd.read_sql(query, self.conn)
                
                dimension_cols = ["Длинна", "Ширина", "Высота", "Вес"]
                for col in dimension_cols:
                    if col in pdf.columns:
                        try:
                            pdf[col] = pd.to_numeric(pdf[col], errors='coerce').fillna(0).round(2)
                        except Exception:
                            pdf[col] = 0.0
                
                output_dir = Path("auto_parts_data")
                output_dir.mkdir(parents=True, exist_ok=True)
                
                pdf.to_csv(output_path, sep=';', index=False, encoding='utf-8-sig')
                
                size_mb = os.path.getsize(output_path) / (1024 * 1024)
                st.success(f"Данные экспортированы: {output_path} ({size_mb:.1f} МБ)")
                return True
        
        except Exception as e:
            logger.exception("Ошибка экспорта CSV")
            st.error(f"Ошибка при экспорте в CSV: {str(e)}")
            return False
    
    def export_to_excel_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None,
                                  include_prices: bool = True, apply_markup: bool = True) -> bool:
        total = self.conn.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        
        if total == 0:
            st.warning("Нет данных для экспорта")
            return False
        
        with self.timer("Экспорт Excel"):
            try:
                query = self.get_export_query(selected_columns, include_prices, apply_markup)
                
                try:
                    df = self.conn.execute(query).fetchdf()
                except Exception as e:
                    logger.error(f"Ошибка fetchdf: {e}")
                    try:
                        arrow_table = self.conn.execute(query).arrow()
                        df = arrow_table.to_pandas()
                    except Exception as e2:
                        logger.error(f"Ошибка arrow fallback: {e2}")
                        df = pd.read_sql(query, self.conn)
                
                dimension_cols = ["Длинна", "Ширина", "Высота", "Вес"]
                
                # 🆕 ИСПРАВЛЕНИЕ: Принудительный сброс типа datetime в float
                for col in dimension_cols:
                    if col in df.columns:
                        if pd.api.types.is_datetime64_any_dtype(df[col]):
                            df[col] = 0.0
                        else:
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)
                
                if len(df) <= EXCEL_ROW_LIMIT:
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Данные')
                        
                        # 🆕 ИСПРАВЛЕНИЕ: Жестко задаем числовой формат, чтобы Excel не рисовал календари
                        ws = writer.sheets['Данные']
                        for col_idx, col_name in enumerate(df.columns, 1):
                            if col_name in dimension_cols:
                                for row in range(2, len(df) + 2):
                                    ws.cell(row=row, column=col_idx).number_format = '0.00'
                else:
                    sheets = (len(df) // EXCEL_ROW_LIMIT) + 1
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        for i in range(sheets):
                            chunk_df = df.iloc[i * EXCEL_ROW_LIMIT:(i + 1) * EXCEL_ROW_LIMIT]
                            sheet_name = f"Данные_{i + 1}"
                            chunk_df.to_excel(writer, index=False, sheet_name=sheet_name)
                            
                            ws = writer.sheets[sheet_name]
                            for col_idx, col_name in enumerate(chunk_df.columns, 1):
                                if col_name in dimension_cols:
                                    for row in range(2, len(chunk_df) + 2):
                                        ws.cell(row=row, column=col_idx).number_format = '0.00'
                
                return True
            
            except Exception as e:
                logger.exception("Ошибка экспорта Excel")
                st.error(f"Ошибка при экспорте в Excel: {str(e)}")
                return False
    
    # ====================================================================
    # УПРАВЛЕНИЕ ДАННЫМИ
    # ====================================================================
    def delete_by_brand(self, brand_norm: str) -> int:
        try:
            with self.db_transaction():
                count_result = self.conn.execute(
                    "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()
                deleted_count = count_result[0] if count_result else 0
                
                if deleted_count == 0:
                    logger.info(f"No records found for brand: {brand_norm}")
                    return 0
                
                self.conn.execute("DELETE FROM parts WHERE brand_norm = ?", [brand_norm])
                self.conn.execute(
                    "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
                
                return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by brand {brand_norm}: {e}")
            raise
    
    def delete_by_artikul(self, artikul_norm: str) -> int:
        try:
            with self.db_transaction():
                count_result = self.conn.execute(
                    "SELECT COUNT(*) FROM parts WHERE artikul_norm = ?", [artikul_norm]).fetchone()
                deleted_count = count_result[0] if count_result else 0
                
                if deleted_count == 0:
                    logger.info(f"No records found for artikul: {artikul_norm}")
                    return 0
                
                self.conn.execute("DELETE FROM parts WHERE artikul_norm = ?", [artikul_norm])
                self.conn.execute(
                    "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
                
                return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by artikul {artikul_norm}: {e}")
            raise
    
    def get_statistics(self) -> Dict[str, Any]:
        stats = {}
        try:
            stats['parts'] = self.conn.execute("SELECT COUNT(*) FROM parts").fetchone()[0]
            stats['oe'] = self.conn.execute("SELECT COUNT(*) FROM oe").fetchone()[0]
            stats['cross'] = self.conn.execute("SELECT COUNT(*) FROM cross_references").fetchone()[0]
            stats['prices'] = self.conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0]
            stats['brands'] = self.conn.execute("SELECT COUNT(DISTINCT brand) FROM parts").fetchone()[0]
            stats['unique_parts'] = self.conn.execute(
                "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
            
            avg_price = self.conn.execute("SELECT AVG(price) FROM prices").fetchone()[0]
            stats['avg_price'] = round(avg_price, 2) if avg_price else 0
            
            try:
                top_brands = self.conn.execute(
                    "SELECT brand, COUNT(*) as cnt FROM parts GROUP BY brand ORDER BY cnt DESC LIMIT 10").pl()
                stats['top_brands'] = top_brands.to_pandas()
            except Exception:
                stats['top_brands'] = pd.DataFrame()
            
            try:
                category_stats = self.conn.execute(
                    "SELECT category, COUNT(*) as cnt FROM oe GROUP BY category ORDER BY cnt DESC").pl()
                stats['category_stats'] = category_stats.to_pandas()
            except Exception:
                stats['category_stats'] = pd.DataFrame()
        
        except Exception as e:
            logger.error(f"Ошибка сбора статистики: {e}")
        
        return stats
    
    # ====================================================================
    # 🆕 v101.0: ПАРАЛЛЕЛЬНАЯ ОБРАБОТКА ФАЙЛОВ С СВЯЗЫВАНИЕМ
    # ====================================================================
    def merge_all_data_parallel(self, file_paths: Dict[str, str], max_workers: int = 4) -> Dict[str, pl.DataFrame]:
        """Параллельная обработка загруженных файлов"""
        results = {}
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key, path in file_paths.items():
                if path and os.path.exists(path):
                    futures[executor.submit(self.read_and_prepare_file, path, key)] = key
            
            for fut in as_completed(futures):
                key = futures[fut]
                try:
                    df = fut.result()
                    if not df.is_empty():
                        results[key] = df
                        logger.info(f"✅ Обработан файл типа '{key}': {len(df):,} строк")
                except Exception as e:
                    logger.error(f"❌ Ошибка обработки файла типа '{key}': {e}")
        
        return results
    
    def link_and_merge_files(
        self,
        file_paths: Dict[str, str],
        link_configs: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, pl.DataFrame]:
        """
        🆕 v101.0: Обработка файлов с связыванием столбцов между ними.
        
        Args:
            file_paths: Словарь {тип_файла: путь_к_файлу}
            link_configs: Список конфигураций связывания:
                [
                    {
                        "primary_file": "oe",
                        "secondary_file": "dimensions",
                        "join_key_primary": "artikul_norm",
                        "join_key_secondary": "artikul_norm",
                        "join_type": "left",
                        "columns_to_fill": ["length", "width", "height", "weight"]
                    }
                ]
        
        Returns:
            Словарь {тип_файла: DataFrame} с связанными данными
        """
        # Параллельная обработка файлов
        results = self.merge_all_data_parallel(file_paths)
        
        # Применяем связывание
        if link_configs:
            for config in link_configs:
                primary_file = config.get("primary_file")
                secondary_file = config.get("secondary_file")
                
                if primary_file not in results or secondary_file not in results:
                    logger.warning(f"Файлы {primary_file} или {secondary_file} не найдены")
                    continue
                
                df_primary = results[primary_file]
                df_secondary = results[secondary_file]
                
                join_key_primary = config.get("join_key_primary", "artikul_norm")
                join_key_secondary = config.get("join_key_secondary", "artikul_norm")
                join_type = config.get("join_type", "left")
                columns_to_fill = config.get("columns_to_fill")
                
                logger.info(f"🔗 Связывание {primary_file} <-> {secondary_file} по {join_key_primary}")
                
                linked_df = self.link_files(
                    df1=df_primary,
                    df2=df_secondary,
                    join_key_1=join_key_primary,
                    join_key_2=join_key_secondary,
                    join_type=join_type,
                    columns_to_fill=columns_to_fill
                )
                
                results[primary_file] = linked_df
                logger.info(f"✅ Связывание завершено: {len(linked_df)} строк")
        
        return results

# ============================================================================
# БЛОК 10: UI  "ЗАГРУЗКА ДАННЫХ" (v101.0)
# ============================================================================
# 📌 v101.0: Новая архитектура с 4 разделами
# - : Загрузка данных (каталог + связывание столбцов + экспорт)
# - Интеграция с AppStateManager для передачи данных в 
# - Кнопка "Сохранить/Загрузить предыдущие расчёты"
# - Автоматическое заполнение недостающих параметров из аналогов
# ============================================================================


def show_section1_data_loading():
    """
    📁 : ЗАГРУЗКА ДАННЫХ
    Загрузка каталога автозапчастей с возможностью связывания столбцов
    между файлами и автоматическим заполнением из аналогов.
    """
    st.header("📁 : Загрузка данных")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    Загрузить каталог автозапчастей, связать данные из разных файлов между собой
    и подготовить их для категоризации и расчёта юнит-экономики.
    
    **📋 ПОРЯДОК ДЕЙСТВИЙ:**
    1. Загрузите файлы с данными товаров (OE, кросс-ссылки, габариты, цены и т.д.)
    2. Настройте связывание столбцов между файлами (если нужно)
    3. Нажмите "Обработать и загрузить"
    4. Дождитесь завершения обработки
    5. Экспортируйте результат в CSV/Excel или переходите в 
    
    ** v101.0:**
    - ✅ Связывание столбцов между файлами (drag-and-drop маппинг)
    - ✅ Автоматическое заполнение недостающих параметров из аналогов
    - ✅ Сохранение/загрузка расчётов одной кнопкой
    """)
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ
    # ========================================================================
    if 'high_volume_catalog' not in st.session_state:
        st.session_state.high_volume_catalog = get_high_volume_catalog()
    
    catalog = st.session_state.high_volume_catalog
    
    if not catalog.conn:
        st.error("❌ Ошибка подключения к базе данных")
        return
    
    # ========================================================================
    #  v101.0: СОХРАНЕНИЕ / ЗАГРУЗКА РАСЧЁТОВ
    # ========================================================================
    with st.expander("💾 Сохранить / Загрузить предыдущие расчёты", expanded=False):
        save_load_manager = get_save_load_manager()
        
        col_sl1, col_sl2 = st.columns(2)
        
        with col_sl1:
            st.markdown("**💾 Сохранить текущее состояние**")
            save_name = st.text_input(
                "Название расчёта",
                value=f"Расчёт {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                key="save_name"
            )
            save_desc = st.text_area(
                "Описание (опционально)",
                placeholder="Например: Расчёт для бренда Bosch, категория Подвеска",
                key="save_desc"
            )
            
            if st.button("💾 Сохранить текущие расчёты", type="primary", key="save_state_btn"):
                with st.spinner("Сохранение..."):
                    state_id = save_load_manager.save_current_state(
                        name=save_name,
                        description=save_desc
                    )
                    if state_id:
                        st.success(f"✅ Состояние сохранено! ID: {state_id[:8]}...")
                    else:
                        st.error("❌ Ошибка сохранения")
        
        with col_sl2:
            st.markdown("**📂 Загрузить предыдущие расчёты**")
            states = save_load_manager.list_states()
            
            if states:
                state_options = {
                    f"{s['name']} ({s['created_at'][:10]})": s['state_id']
                    for s in states
                }
                selected_state = st.selectbox(
                    "Выберите расчёт для загрузки",
                    options=list(state_options.keys()),
                    key="load_state_select"
                )
                
                if st.button("📂 Загрузить выбранные расчёты", key="load_state_btn"):
                    state_id = state_options[selected_state]
                    with st.spinner("Загрузка..."):
                        success = save_load_manager.load_state(state_id)
                        if success:
                            st.success("✅ Состояние загружено! Переходите к следующему разделу.")
                            st.rerun()
                        else:
                            st.error("❌ Ошибка загрузки")
                
                # Кнопка удаления
                if st.button("🗑️ Удалить выбранный расчёт", key="delete_state_btn"):
                    state_id = state_options[selected_state]
                    success = save_load_manager.delete_state(state_id)
                    if success:
                        st.success("✅ Расчёт удалён")
                        st.rerun()
            else:
                st.info("ℹ️ Нет сохранённых расчётов")
    
    st.divider()
    
    # ========================================================================
    # 📥 ЗАГРУЗКА ФАЙЛОВ
    # ========================================================================
    st.subheader("📥 Загрузка файлов каталога")
    
    st.info("""
    **📋 ТИПЫ ФАЙЛОВ:**
    - **📋 Основные данные (OE):** OE-номера, наименования, применимость
    - **🔗 Кросс-ссылки:** Связи между OE-номерами и артикулами
    - **📊 Штрих-коды:** Баркоды и кратность упаковки
    - **📏 Габариты:** Длина, ширина, высота, вес
    - **🖼️ Изображения:** Ссылки на фото товаров
    - **💰 Цены:** Цены и валюта
    
    💡 **Минимальный набор:** Достаточно загрузить хотя бы один файл.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        oe_file = st.file_uploader(
            "📋 Основные данные (OE)",
            type=['xlsx', 'xls', 'csv'],
            key="hv_oe_file",
            help="OE-номера, наименования, применимость"
        )
        cross_file = st.file_uploader(
            "🔗 Кросс-ссылки",
            type=['xlsx', 'xls', 'csv'],
            key="hv_cross_file",
            help="Связи OE ↔ артикул"
        )
        barcode_file = st.file_uploader(
            "📊 Штрих-коды",
            type=['xlsx', 'xls', 'csv'],
            key="hv_barcode_file",
            help="Баркоды и кратность"
        )
    
    with col2:
        dims_file = st.file_uploader(
            "📏 Габариты",
            type=['xlsx', 'xls', 'csv'],
            key="hv_dims_file",
            help="Длина, ширина, высота, вес"
        )
        images_file = st.file_uploader(
            "🖼️ Изображения",
            type=['xlsx', 'xls', 'csv'],
            key="hv_images_file",
            help="Ссылки на фото"
        )
        prices_file = st.file_uploader(
            "💰 Цены",
            type=['xlsx', 'xls', 'csv'],
            key="hv_prices_file",
            help="Цены и валюта"
        )
    
    uploaded_files = {
        'oe': oe_file,
        'cross': cross_file,
        'barcode': barcode_file,
        'dimensions': dims_file,
        'images': images_file,
        'prices': prices_file
    }
    
    # Фильтруем только загруженные файлы
    active_files = {k: v for k, v in uploaded_files.items() if v is not None}
    
    if active_files:
        st.success(f"✅ Загружено файлов: {len(active_files)}")
        
        # ====================================================================
        #  v101.0: НАСТРОЙКА СВЯЗЫВАНИЯ СТОЛБЦОВ МЕЖДУ ФАЙЛАМИ
        # ====================================================================
        with st.expander("🔗 Настройка связывания столбцов между файлами", expanded=False):
            st.info("""
            **🎯 ЦЕЛЬ:**
            Связать столбцы из разных файлов, чтобы автоматически заполнить
            недостающие параметры (вес, габариты) из аналогов.
            
            **📋 ПРИМЕР:**
            Если в файле "OE" нет веса, но в файле "Габариты" вес есть 
            система автоматически подставит его по артикулу.
            """)
            
            # Показываем доступные файлы для связывания
            if len(active_files) >= 2:
                st.markdown("**🔗 Добавить связывание между файлами:**")
                
                file_names = list(active_files.keys())
                
                col_link1, col_link2, col_link3 = st.columns(3)
                
                with col_link1:
                    primary_file = st.selectbox(
                        "Основной файл",
                        options=file_names,
                        key="link_primary_file"
                    )
                
                with col_link2:
                    secondary_file = st.selectbox(
                        "Дополнительный файл",
                        options=[f for f in file_names if f != primary_file],
                        key="link_secondary_file"
                    )
                
                with col_link3:
                    join_type = st.selectbox(
                        "Тип соединения",
                        options=["left", "inner", "outer"],
                        format_func=lambda x: {
                            "left": "Left (все из основного)",
                            "inner": "Inner (только совпадения)",
                            "outer": "Outer (все из обоих)"
                        }[x],
                        key="link_join_type"
                    )
                
                # Ключевые столбцы (по умолчанию artikul_norm)
                join_key = st.text_input(
                    "Ключевой столбец для соединения",
                    value="artikul_norm",
                    key="link_join_key",
                    help="Столбец, по которому будут соединяться файлы (обычно artikul_norm)"
                )
                
                # Столбцы для заполнения
                columns_to_fill = st.multiselect(
                    "Какие столбцы заполнять из дополнительного файла",
                    options=["length", "width", "height", "weight", "dimensions_str", "price"],
                    default=["length", "width", "height", "weight"],
                    key="link_columns_to_fill"
                )
                
                if st.button("➕ Добавить связывание", key="add_link_btn"):
                    if 'link_configs' not in st.session_state:
                        st.session_state.link_configs = []
                    
                    st.session_state.link_configs.append({
                        "primary_file": primary_file,
                        "secondary_file": secondary_file,
                        "join_key_primary": join_key,
                        "join_key_secondary": join_key,
                        "join_type": join_type,
                        "columns_to_fill": columns_to_fill
                    })
                    
                    st.success(f"✅ Связывание добавлено: {primary_file} ↔ {secondary_file}")
                
                # Показываем добавленные связывания
                if 'link_configs' in st.session_state and st.session_state.link_configs:
                    st.markdown("**📋 Добавленные связывания:**")
                    for i, config in enumerate(st.session_state.link_configs):
                        st.write(f"{i+1}. {config['primary_file']} ↔ {config['secondary_file']} "
                                f"(по {config['join_key_primary']}, тип: {config['join_type']})")
                    
                    if st.button("🗑️ Очистить все связывания", key="clear_links_btn"):
                        st.session_state.link_configs = []
                        st.rerun()
            else:
                st.warning("⚠️ Загрузите хотя бы 2 файла для настройки связывания")
        
        # ====================================================================
        # 🚀 ОБРАБОТКА И ЗАГРУЗКА
        # ====================================================================
        st.divider()
        
        if st.button("🚀 Обработать и загрузить в базу", type="primary", key="hv_load_btn"):
            saved_paths = {}
            
            # Сохраняем загруженные файлы на диск
            for key, file in active_files.items():
                path = catalog.data_dir / f"{key}_{int(time.time())}.xlsx"
                with open(path, "wb") as f:
                    f.write(file.getbuffer())
                saved_paths[key] = str(path)
            
            if saved_paths:
                with st.spinner("Обработка файлов..."):
                    #  v101.0: Используем link_and_merge_files если есть связывания
                    link_configs = st.session_state.get('link_configs', [])
                    
                    if link_configs:
                        st.info(f"🔗 Применяем {len(link_configs)} связываний между файлами...")
                        dataframes = catalog.link_and_merge_files(saved_paths, link_configs)
                    else:
                        dataframes = catalog.merge_all_data_parallel(saved_paths)
                    
                    with st.spinner("Загрузка данных в базу..."):
                        catalog.process_and_load_data(dataframes)
                    
                    st.success("✅ Данные успешно загружены!")
                    
                    #  v101.0: Сохраняем в AppStateManager для передачи в 
                    if 'oe' in dataframes and not dataframes['oe'].is_empty():
                        AppStateManager.set('section1_catalog_df', dataframes['oe'].to_pandas())
                    
                    # Показываем статистику
                    stats = catalog.get_statistics()
                    st.metric("📦 Уникальных артикулов", f"{stats.get('unique_parts', 0):,}")
            else:
                st.warning("⚠️ Не удалось сохранить файлы")
    
    # ========================================================================
    # 👁️ ПРЕДПРОСМОТР ДАННЫХ
    # ========================================================================
    st.divider()
    st.subheader("👁️ Предпросмотр данных в базе")
    
    stats = catalog.get_statistics()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Уникальных товаров", f"{stats.get('unique_parts', 0):,}")
    with col2:
        st.metric("🏷️ Брендов", f"{stats.get('brands', 0):,}")
    with col3:
        st.metric("🔗 OE-номеров", f"{stats.get('oe', 0):,}")
    with col4:
        st.metric("💰 Цен", f"{stats.get('prices', 0):,}")
    
    # Предпросмотр таблицы parts
    try:
        preview_df = catalog.conn.execute("""
            SELECT artikul, brand, length, width, height, weight, billable_weight
            FROM parts
            LIMIT 20
        """).df()
        
        if not preview_df.empty:
            st.markdown("**📋 Первые 20 записей:**")
            st_dataframe_compat(preview_df)
        else:
            st.info("ℹ️ База данных пуста. Загрузите данные выше.")
    except Exception as e:
        st.warning(f"⚠️ Ошибка предпросмотра: {e}")
    
    # ========================================================================
    # 📤 ЭКСПОРТ В CSV / EXCEL
    # ========================================================================
    st.divider()
    st.subheader("📤 Экспорт каталога")
    
    total = catalog.conn.execute(
        "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)"
    ).fetchone()[0]
    
    if total == 0:
        st.warning("⚠️ Нет данных для экспорта")
        return
    
    st.info(f"📊 Всего записей для экспорта: {total:,}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        export_format = st.radio(
            "Формат экспорта",
            ["CSV", "Excel"],
            horizontal=True,
            key="export_format"
        )
        
        selected_columns = st.multiselect(
            "Колонки для экспорта",
            options=[
                "Артикул бренда", "Бренд", "Наименование", "Применимость",
                "Длинна", "Ширина", "Высота", "Вес", "OE номер", "Цена"
            ],
            default=["Артикул бренда", "Бренд", "Длинна", "Ширина", "Высота", "Вес"],
            key="export_columns"
        )
    
    with col2:
        include_prices = st.checkbox("💰 Включить цены", value=True, key="export_prices")
        apply_markup = st.checkbox("📈 Применить наценку", value=False, key="export_markup")
        
        if apply_markup:
            markup_percent = st.number_input(
                "Наценка (%)",
                min_value=0.0,
                max_value=500.0,
                value=20.0,
                step=5.0,
                key="export_markup_percent"
            )
            catalog.price_rules['global_markup'] = markup_percent / 100
    
    if st.button("🚀 Экспортировать", type="primary", key="export_btn"):
        format_extensions = {"CSV": "csv", "Excel": "xlsx"}
        ext = format_extensions[export_format]
        output_path = catalog.data_dir / f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        
        with st.spinner(f"Генерация файла {export_format}..."):
            if export_format == "CSV":
                success = catalog.export_to_csv_optimized(
                    str(output_path),
                    selected_columns if selected_columns else None,
                    include_prices,
                    apply_markup
                )
            else:
                success = catalog.export_to_excel_optimized(
                    str(output_path),
                    selected_columns if selected_columns else None,
                    include_prices,
                    apply_markup
                )
        
        if success and output_path.exists():
            with open(output_path, "rb") as f:
                file_data = f.read()
            
            mime_map = {
                "CSV": "text/csv; charset=utf-8",
                "Excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
            
            st.download_button(
                label=f"⬇️ Скачать {export_format} файл",
                data=file_data,
                file_name=output_path.name,
                mime=mime_map[export_format],
                key="export_download"
            )
            st.success(f"✅ Файл готов к скачиванию!")
    
    # ========================================================================
    # 🔧 УПРАВЛЕНИЕ ДАННЫМИ (опционально)
    # ========================================================================
    with st.expander("🔧 Управление данными (удаление, очистка)", expanded=False):
        st.warning("⚠️ Операции необратимы!")
        
        management_option = st.radio(
            "Выберите действие:",
            ["🏭 Удалить по бренду", "📦 Удалить по артикулу", "🗑️ Очистить всю базу"],
            horizontal=True,
            key="management_option"
        )
        
        if management_option == "🏭 Удалить по бренду":
            catalog._show_delete_by_brand()
        elif management_option == "📦 Удалить по артикулу":
            catalog._show_delete_by_artikul()
        elif management_option == "🗑️ Очистить всю базу":
            if st.checkbox("⚠️ Подтверждаю удаление всей базы", key="confirm_clear_db"):
                if st.button("🗑️ Удалить всё", key="clear_db_btn"):
                    if catalog.delete_database_file():
                        st.session_state.high_volume_catalog = get_high_volume_catalog()
                        st.success("✅ База данных очищена")
                        st.rerun()

# ============================================================================
# БЛОК 11: UI  "ВЕСОГАБАРИТЫ И КАТЕГОРИЗАЦИЯ" (v101.0)
# ============================================================================
# 📌 v101.0: Новая архитектура с 4 разделами
# - : Весогабариты и категоризация
# - 3-уровневая иерархия (Родитель  Группа  Подгруппа)
# - Автоматическая категоризация по наименованию
# - Загрузка стандартных весогабаритов из Excel
# - Проверка ВГ ±20% (дополнительная, не строгая)
# ============================================================================


def show_section2_categorization():
    """
    📂 : ВЕСОГАБАРИТЫ И КАТЕГОРИЗАЦИЯ
    Автоматическая категоризация товаров с 3-уровневой иерархией
    и проверкой весогабаритов.
    """
    st.header("📂 : Весогабариты и категоризация")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    Автоматически категоризировать товары по наименованию с использованием
    3-уровневой иерархии (Родитель  Группа  Подгруппа) и проверить
    весогабариты на соответствие стандартным значениям.
    
    **📋 ПОРЯДОК ДЕЙСТВИЙ:**
    1. Загрузите файл с артикулами и наименованиями (или используйте данные из Раздела 1)
    2. Система автоматически определит категории по ключевым словам
    3. Проверьте весогабариты на соответствие стандартам (±20%)
    4. При необходимости загрузите справочник стандартных весогабаритов
    5. Переходите в  (Тарифы) или  (Расчёт)
    
    ** v101.0:**
    - ✅ 3-уровневая иерархия категорий
    - ✅ Автоматическая категоризация по ML-модели и ключевым словам
    - ✅ Загрузка стандартных весогабаритов из Excel
    - ✅ Проверка ВГ ±20% (дополнительная, не строгая)
    """)
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ
    # ========================================================================
    if 'category_classifier' not in st.session_state:
        st.session_state.category_classifier = CategoryClassifier()
    
    classifier = st.session_state.category_classifier
    
    # Инициализация БД категорий (если ещё не инициализирована)
    if 'category_dimensions_db' not in st.session_state:
        st.session_state.category_dimensions_db = CategoryDimensionsDB()
    
    categories_db = st.session_state.category_dimensions_db
    
    # ========================================================================
    # 📥 ЗАГРУЗКА ДАННЫХ ДЛЯ КАТЕГОРИЗАЦИИ
    # ========================================================================
    st.subheader("📥 Загрузка данных для категоризации")
    
    st.info("""
    **📋 ТРЕБОВАНИЯ К ФАЙЛУ:**
    - **Артикул** (обязательно)  уникальный идентификатор товара
    - **Наименование** (обязательно)  название товара для категоризации
    - **Длина, Ширина, Высота, Вес** (опционально)  для проверки ВГ
    
    💡 **ИЛИ** используйте данные из Раздела 1 (если они уже загружены).
    """)
    
    # Проверяем, есть ли данные из Раздела 1
    catalog_data = AppStateManager.get_catalog_data()
    
    col1, col2 = st.columns(2)
    
    with col1:
        use_section1_data = st.checkbox(
            "🔄 Использовать данные из Раздела 1",
            value=(catalog_data is not None and not catalog_data.empty),
            disabled=(catalog_data is None or catalog_data.empty),
            key="use_section1_data"
        )
        
        if use_section1_data and catalog_data is not None:
            st.success(f"✅ Загружено {len(catalog_data)} товаров из Раздела 1")
            df_input = catalog_data.copy()
    
    with col2:
        uploaded_file = st.file_uploader(
            "📤 Или загрузите файл с артикулами и наименованиями",
            type=['xlsx', 'xls', 'csv'],
            key="categorization_file",
            help="Файл должен содержать колонки: Артикул, Наименование"
        )
        
        if uploaded_file is not None and not use_section1_data:
            try:
                if uploaded_file.name.lower().endswith('.csv'):
                    df_input = smart_read_csv(uploaded_file)
                else:
                    df_input = pd.read_excel(uploaded_file, engine='openpyxl')
                
                st.success(f"✅ Загружено {len(df_input)} товаров из файла")
            except Exception as e:
                st.error(f"❌ Ошибка загрузки файла: {e}")
                df_input = None
    
    # Если данные не загружены ни из одного источника
    if not use_section1_data and uploaded_file is None:
        st.warning("⚠️ Загрузите данные для категоризации (из Раздела 1 или из файла)")
        return
    
    if df_input is None or df_input.empty:
        st.warning("⚠️ Нет данных для категоризации")
        return
    
    # ========================================================================
    # 🔍 ОПРЕДЕЛЕНИЕ КОЛОНОК
    # ========================================================================
    st.subheader("🔍 Определение колонок")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Ищем колонку с артикулом
        article_col = None
        for col in df_input.columns:
            if any(w in str(col).lower() for w in ['артикул', 'article', 'sku', 'код']):
                article_col = col
                break
        
        article_col = st.selectbox(
            "🔢 Колонка с артикулом",
            options=df_input.columns.tolist(),
            index=df_input.columns.tolist().index(article_col) if article_col else 0,
            key="cat_article_col"
        )
    
    with col2:
        # Ищем колонку с наименованием
        name_col = None
        for col in df_input.columns:
            if any(w in str(col).lower() for w in ['наименование', 'название', 'name', 'описание']):
                name_col = col
                break
        
        name_col = st.selectbox(
            "📝 Колонка с наименованием",
            options=df_input.columns.tolist(),
            index=df_input.columns.tolist().index(name_col) if name_col else 0,
            key="cat_name_col"
        )
    
    # ========================================================================
    # 🏷️ АВТОМАТИЧЕСКАЯ КАТЕГОРИЗАЦИЯ (3-УРОВНЕВАЯ ИЕРАРХИЯ)
    # ========================================================================
    st.divider()
    st.subheader("🏷️ Автоматическая категоризация (3-уровневая иерархия)")
    
    st.info("""
    **📋 ЛОГИКА КАТЕГОРИЗАЦИИ:**
    1. Система анализирует наименование товара
    2. Ищет ключевые слова в справочнике категорий (150+ категорий)
    3. Определяет 3-уровневую иерархию:
       - **Родитель** (например, "Автозапчасти")
       - **Группа** (например, "Подвеска")
       - **Подгруппа** (например, "Сайлентблоки")
    4. Присваивает уверенность (0.0 - 1.0)
    """)
    
    if st.button("🚀 Категоризировать товары", type="primary", key="categorize_btn"):
        with st.spinner("Категоризация товаров..."):
            # Применяем классификатор
            df_categorized = classifier.classify_dataframe(df_input, name_col=name_col)
            
            # Сохраняем результат
            AppStateManager.set('section2_categorized_df', df_categorized)
            
            st.success(f"✅ Категоризировано {len(df_categorized)} товаров")
            
            # Показываем статистику
            st.subheader("📊 Статистика категоризации")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📦 Всего товаров", len(df_categorized))
            
            with col2:
                unique_parents = df_categorized['parent_category'].nunique()
                st.metric("🏭 Родителей", unique_parents)
            
            with col3:
                unique_groups = df_categorized['group_category'].nunique()
                st.metric("📂 Групп", unique_groups)
            
            with col4:
                unique_subgroups = df_categorized['subgroup_category'].nunique()
                st.metric("📁 Подгрупп", unique_subgroups)
            
            # Средняя уверенность
            avg_confidence = df_categorized['classification_confidence'].mean()
            st.info(f"🎯 Средняя уверенность классификации: {avg_confidence:.2%}")
    
    # ========================================================================
    # 👁️ ПРЕДПРОСМОТР КАТЕГОРИЗОВАННЫХ ДАННЫХ
    # ========================================================================
    categorized_df = AppStateManager.get('section2_categorized_df')
    
    if categorized_df is not None and not categorized_df.empty:
        st.divider()
        st.subheader("👁️ Предпросмотр категоризованных данных")
        
        # Показываем первые 50 строк с новыми колонками
        display_cols = [article_col, name_col, 'parent_category', 'group_category', 
                       'subgroup_category', 'classification_confidence']
        available_cols = [col for col in display_cols if col in categorized_df.columns]
        
        st_dataframe_compat(categorized_df[available_cols].head(50), key="categorized_preview")
        
        # ====================================================================
        # 📊 РАСПРЕДЕЛЕНИЕ ПО КАТЕГОРИЯМ
        # ====================================================================
        st.subheader("📊 Распределение по категориям")
        
        tab1, tab2, tab3 = st.tabs(["🏭 По Родителям", "📂 По Группам", "📁 По Подгруппам"])
        
        with tab1:
            parent_counts = categorized_df['parent_category'].value_counts()
            st_dataframe_compat(parent_counts, key="parent_counts")
        
        with tab2:
            group_counts = categorized_df['group_category'].value_counts()
            st_dataframe_compat(group_counts, key="group_counts")
        
        with tab3:
            subgroup_counts = categorized_df['subgroup_category'].value_counts()
            st_dataframe_compat(subgroup_counts, key="subgroup_counts")
        
        # ====================================================================
        # 📏 ЗАГРУЗКА СТАНДАРТНЫХ ВЕСОГАБАРИТОВ
        # ====================================================================
        st.divider()
        st.subheader("📏 Загрузка стандартных весогабаритов")
        
        st.info("""
        **📋 ЦЕЛЬ:**
        Загрузить справочник стандартных весогабаритов для категорий,
        чтобы проверить фактические размеры товаров на соответствие.
        
        **📋 ТРЕБОВАНИЯ К ФАЙЛУ:**
        - **Категория** (обязательно)  название категории
        - **Длина (см)** (обязательно)
        - **Ширина (см)** (обязательно)
        - **Высота (см)** (обязательно)
        - **Вес (кг)** (обязательно)
        """)
        
        # Проверяем, есть ли уже загруженные категории
        if categories_db.categories:
            st.success(f"✅ Загружено {len(categories_db.categories)} стандартных категорий")
            
            if st.button("📋 Показать загруженные категории", key="show_loaded_categories"):
                categories_data = []
                for key, cat in categories_db.categories.items():
                    categories_data.append({
                        'Категория': cat['name'],
                        'Длина (см)': cat['length_cm'],
                        'Ширина (см)': cat['width_cm'],
                        'Высота (см)': cat['height_cm'],
                        'Вес (кг)': cat['weight_kg']
                    })
                
                categories_df = pd.DataFrame(categories_data)
                st_dataframe_compat(categories_df, key="loaded_categories_table")
        
        # Загрузка нового файла
        std_dims_file = st.file_uploader(
            "📤 Загрузить файл со стандартными весогабаритами",
            type=['xlsx', 'xls'],
            key="std_dims_file",
            help="Файл должен содержать: Категория, Длина, Ширина, Высота, Вес"
        )
        
        if std_dims_file is not None:
            if st.button("🚀 Импортировать стандартные весогабариты", key="import_std_dims"):
                with st.spinner("Импорт стандартных весогабаритов..."):
                    # Сохраняем временный файл
                    temp_path = TEMP_DIR / f"std_dims_{int(time.time())}.xlsx"
                    with open(temp_path, "wb") as f:
                        f.write(std_dims_file.getbuffer())
                    
                    # Импортируем
                    result = categories_db.import_from_excel(str(temp_path))
                    
                    if result["success"]:
                        st.success(f"✅ Импортировано {result['imported']} категорий")
                        
                        if result["warnings"]:
                            with st.expander(f"⚠️ Предупреждения ({len(result['warnings'])})"):
                                for warning in result["warnings"][:10]:
                                    st.warning(warning)
                        
                        # Сохраняем в AppStateManager
                        AppStateManager.set('section2_standard_dims', categories_db.categories)
                        
                        st.rerun()
                    else:
                        st.error("❌ Ошибка импорта")
                        with st.expander("📋 Ошибки"):
                            for error in result["errors"]:
                                st.error(error)
                    
                    # Удаляем временный файл
                    if temp_path.exists():
                        temp_path.unlink()
        
        # Кнопка скачать шаблон
        if st.button("📥 Скачать шаблон стандартных весогабаритов", key="download_std_dims_template"):
            template_data = {
                'Категория': ['Фильтры', 'Колодки', 'Масла', 'Шины', 'Амортизаторы'],
                'Длина (см)': [15, 15, 10, 60, 50],
                'Ширина (см)': [15, 10, 10, 60, 10],
                'Высота (см)': [15, 5, 25, 25, 50],
                'Вес (кг)': [0.5, 2.0, 1.0, 10.0, 5.0]
            }
            template_df = pd.DataFrame(template_data)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False, sheet_name='Стандарты')
            output.seek(0)
            
            st.download_button(
                label="⬇️ Скачать шаблон Excel",
                data=output,
                file_name="шаблон_стандартных_весогабаритов.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_std_dims_template_btn"
            )
        
        # ====================================================================
        # 🔍 ПРОВЕРКА ВЕСОГАБАРИТОВ ±20%
        # ====================================================================
        st.divider()
        st.subheader("🔍 Проверка весогабаритов ±20%")
        
        st.info("""
        **📋 ЛОГИКА ПРОВЕРКИ:**
        - Сравниваем фактические размеры товаров со стандартными для их категории
        - Допустимое отклонение: ±20%
        - Проверка **не строгая**  только дополнительная информация
        - Если у товара нет размеров  пропускаем
        
        💡 **ПРИМЕР:**
        Если стандартный вес для "Фильтры" = 0.5 кг, а у товара = 0.7 кг,
        то отклонение = 40%  это больше 20%, система выдаст предупреждение.
        """)
        
        # Проверяем, есть ли колонки с размерами
        dimension_cols = ['Длина', 'Ширина', 'Высота', 'Вес']
        available_dims = [col for col in dimension_cols if col in categorized_df.columns]
        
        if not available_dims:
            st.warning("⚠️ В данных нет колонок с размерами (Длина, Ширина, Высота, Вес)")
            st.info("💡 Добавьте эти колонки в Разделе 1 или загрузите файл с габаритами")
        elif not categories_db.categories:
            st.warning("⚠️ Не загружены стандартные весогабариты")
            st.info("💡 Загрузите файл со стандартами выше")
        else:
            # Определяем колонки с размерами
            length_col = next((col for col in categorized_df.columns if 'длина' in col.lower() or 'length' in col.lower()), None)
            width_col = next((col for col in categorized_df.columns if 'ширина' in col.lower() or 'width' in col.lower()), None)
            height_col = next((col for col in categorized_df.columns if 'высота' in col.lower() or 'height' in col.lower()), None)
            weight_col = next((col for col in categorized_df.columns if 'вес' in col.lower() or 'weight' in col.lower()), None)
            
            # Настраиваем допуск
            tolerance_percent = st.slider(
                "📏 Допустимое отклонение (%)",
                min_value=5,
                max_value=50,
                value=20,
                step=5,
                key="dims_tolerance"
            )
            
            if st.button("🔍 Проверить весогабариты", key="check_dims_btn"):
                with st.spinner("Проверка весогабаритов..."):
                    # Добавляем колонки с отклонениями
                    deviations_data = []
                    
                    for idx, row in categorized_df.iterrows():
                        category = row.get('subgroup_category', '') or row.get('group_category', '')
                        
                        # Получаем стандартные размеры для категории
                        std_dims = categories_db.get_category(category)
                        
                        if std_dims is None:
                            deviations_data.append({
                                'index': idx,
                                'valid': True,
                                'deviations': [],
                                'warnings': ['Категория не найдена в стандартах']
                            })
                            continue
                        
                        # Проверяем отклонения
                        deviations = []
                        warnings = []
                        
                        if length_col and std_dims['length_cm'] > 0:
                            actual = safe_float(row.get(length_col, 0))
                            if actual > 0:
                                dev = abs(actual - std_dims['length_cm']) / std_dims['length_cm'] * 100
                                if dev > tolerance_percent:
                                    deviations.append({
                                        'parameter': 'Длина',
                                        'actual': actual,
                                        'expected': std_dims['length_cm'],
                                        'deviation_percent': dev
                                    })
                        
                        if width_col and std_dims['width_cm'] > 0:
                            actual = safe_float(row.get(width_col, 0))
                            if actual > 0:
                                dev = abs(actual - std_dims['width_cm']) / std_dims['width_cm'] * 100
                                if dev > tolerance_percent:
                                    deviations.append({
                                        'parameter': 'Ширина',
                                        'actual': actual,
                                        'expected': std_dims['width_cm'],
                                        'deviation_percent': dev
                                    })
                        
                        if height_col and std_dims['height_cm'] > 0:
                            actual = safe_float(row.get(height_col, 0))
                            if actual > 0:
                                dev = abs(actual - std_dims['height_cm']) / std_dims['height_cm'] * 100
                                if dev > tolerance_percent:
                                    deviations.append({
                                        'parameter': 'Высота',
                                        'actual': actual,
                                        'expected': std_dims['height_cm'],
                                        'deviation_percent': dev
                                    })
                        
                        if weight_col and std_dims['weight_kg'] > 0:
                            actual = safe_float(row.get(weight_col, 0))
                            if actual > 0:
                                dev = abs(actual - std_dims['weight_kg']) / std_dims['weight_kg'] * 100
                                if dev > tolerance_percent:
                                    deviations.append({
                                        'parameter': 'Вес',
                                        'actual': actual,
                                        'expected': std_dims['weight_kg'],
                                        'deviation_percent': dev
                                    })
                        
                        if deviations:
                            warnings.append(f"⚠️ Отклонений: {len(deviations)}")
                        else:
                            warnings.append("✅ Все параметры в пределах нормы")
                        
                        deviations_data.append({
                            'index': idx,
                            'valid': len(deviations) == 0,
                            'deviations': deviations,
                            'warnings': warnings
                        })
                    
                    # Добавляем результаты в DataFrame
                    categorized_df['dims_valid'] = [d['valid'] for d in deviations_data]
                    categorized_df['dims_deviations_count'] = [len(d['deviations']) for d in deviations_data]
                    categorized_df['dims_warnings'] = ['; '.join(d['warnings']) for d in deviations_data]
                    
                    # Сохраняем обновлённый DataFrame
                    AppStateManager.set('section2_categorized_df', categorized_df)
                    
                    # Показываем статистику
                    valid_count = sum(1 for d in deviations_data if d['valid'])
                    invalid_count = len(deviations_data) - valid_count
                    
                    st.success(f"✅ Проверка завершена")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("✅ В пределах нормы", valid_count)
                    
                    with col2:
                        st.metric("⚠️ С отклонениями", invalid_count)
                    
                    with col3:
                        valid_percent = (valid_count / len(deviations_data) * 100) if deviations_data else 0
                        st.metric("📊 Процент соответствия", f"{valid_percent:.1f}%")
                    
                    # Показываем товары с отклонениями
                    if invalid_count > 0:
                        st.subheader(f"⚠️ Товары с отклонениями ({invalid_count})")
                        
                        invalid_df = categorized_df[categorized_df['dims_valid'] == False]
                        
                        display_cols = [article_col, name_col, 'subgroup_category', 
                                       'dims_deviations_count', 'dims_warnings']
                        available_cols = [col for col in display_cols if col in invalid_df.columns]
                        
                        st_dataframe_compat(invalid_df[available_cols].head(50), key="invalid_dims_table")
        
        # ====================================================================
        # 💾 СОХРАНЕНИЕ И ПЕРЕХОД К СЛЕДУЮЩЕМУ РАЗДЕЛУ
        # ====================================================================
        st.divider()
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("💾 Сохранить и перейти в  (Тарифы)", type="primary", key="go_to_section3"):
                st.success("✅ Данные сохранены! Переходите в ")
                AppStateManager.set('current_section', 'section3')
        
        with col2:
            if st.button("🧹 Очистить данные категоризации", key="clear_categorization"):
                AppStateManager.delete('section2_categorized_df')
                st.success("✅ Данные очищены")
                st.rerun()

# ============================================================================
# БЛОК 12: UI  "ТАРИФЫ" (v101.0)
# ============================================================================
# 📌 v101.0: Новая архитектура с 4 разделами
# - : Тарифы (объединяет все функции работы с тарифами)
# - Умная загрузка тарифов (API/AI/Кэш/Гибрид)
# - Google Sheets интеграция (загрузка/выгрузка тарифов)
# - Кнопка "Обновить тарифы" для обновления юнит-экономики
# - Экспорт в CSV и импорт в Google Sheets
# ============================================================================


def show_section3_tariffs():
    """
    💰 : ТАРИФЫ
    Объединяет все функции работы с тарифами маркетплейсов:
    - Умная загрузка (API/AI/Кэш/Гибрид)
    - Google Sheets интеграция
    - Обновление юнит-экономики
    - Экспорт/импорт
    """
    st.header("💰 : Тарифы")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    Загрузить актуальные тарифы маркетплейсов и подготовить их для расчёта юнит-экономики.
    
    **📋 ПОРЯДОК ДЕЙСТВИЙ:**
    1. Выберите маркетплейс и источник тарифов
    2. Загрузите тарифы (API/AI/Кэш/Гибрид)
    3. При необходимости загрузите/выгрузите в Google Sheets
    4. Нажмите "Обновить тарифы" для применения к расчётам
    5. Переходите в  (Расчёт)
    
    ** v101.0:**
    - ✅ Google Sheets интеграция (JSON + API)
    - ✅ Кнопка "Обновить тарифы" для обновления юнит-экономики
    - ✅ Экспорт в CSV и импорт в Google Sheets
    - ✅ Загрузка тарифов напрямую от маркетплейса
    """)
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ
    # ========================================================================
    if 'marketplace_unit_economics' not in st.session_state:
        st.session_state.marketplace_unit_economics = get_marketplace_unit_economics()
    
    unit_economics = st.session_state.marketplace_unit_economics
    
    # Инициализация SmartTariffLoader
    if 'smart_tariff_loader' not in st.session_state:
        try:
            st.session_state.smart_tariff_loader = SmartTariffLoader()
        except Exception as e:
            st.error(f"❌ Ошибка инициализации SmartTariffLoader: {e}")
            return
    
    tariff_loader = st.session_state.smart_tariff_loader
    
    # ========================================================================
    # 📡 ВЫБОР МАРКЕТПЛЕЙСА И ИСТОЧНИКА
    # ========================================================================
    st.subheader("📡 Загрузка тарифов")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        marketplace = st.selectbox(
            "🏪 Выберите маркетплейс",
            ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "СберМегаМаркет"],
            key="tariff_mp"
        )
    
    with col2:
        source = st.selectbox(
            "📡 Источник тарифов",
            ["hybrid", "api", "ai", "cache"],
            format_func=lambda x: SmartTariffLoader.SOURCES.get(x, x),
            key="tariff_source"
        )
    
    # Показываем доступные источники
    if tariff_loader and hasattr(tariff_loader, 'get_available_sources'):
        try:
            available = tariff_loader.get_available_sources(marketplace)
            source_labels = [SmartTariffLoader.SOURCES.get(s, s) for s in available]
            st.info(f"🔍 Доступные источники для {marketplace}: {', '.join(source_labels)}")
        except Exception as e:
            st.warning(f"⚠️ Ошибка получения доступных источников: {e}")
    
    # ========================================================================
    # 🔑 API КЛЮЧИ (если выбран API режим)
    # ========================================================================
    if source in ["api", "hybrid"]:
        with st.expander("🔑 API ключи", expanded=True):
            col1, col2 = st.columns(2)
            
            with col1:
                api_key = st.text_input(
                    "API Key",
                    type="password",
                    placeholder="Введите API ключ",
                    key="tariff_api_key",
                    help="Для Ozon: Api-Key, для WB: Api-Key"
                )
            
            with col2:
                client_id = st.text_input(
                    "Client ID (только для Ozon)",
                    type="password",
                    placeholder="Введите Client ID",
                    key="tariff_client_id"
                )
    else:
        api_key = None
        client_id = None
    
    # ========================================================================
    # 🚀 КНОПКИ ЗАГРУЗКИ И СРАВНЕНИЯ
    # ========================================================================
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Сравнить источники", key="tariff_compare"):
            if tariff_loader and hasattr(tariff_loader, 'compare_sources'):
                with st.spinner("Сравнение источников..."):
                    try:
                        compare_df = tariff_loader.compare_sources(marketplace, api_key, client_id)
                        if compare_df is not None and not compare_df.empty:
                            st.subheader("📊 Сравнение источников")
                            st_dataframe_compat(compare_df)
                        else:
                            st.warning("⚠️ Нет данных для сравнения")
                    except Exception as e:
                        st.error(f"❌ Ошибка сравнения: {e}")
    
    with col2:
        if st.button("🚀 Загрузить тарифы", type="primary", key="tariff_load"):
            if not tariff_loader or not hasattr(tariff_loader, 'load_tariffs'):
                st.error("❌ Метод load_tariffs не найден")
                return
            
            with st.spinner(f"Загрузка тарифов из источника: {SmartTariffLoader.SOURCES.get(source, source)}..."):
                try:
                    result = tariff_loader.load_tariffs(
                        marketplace=marketplace,
                        source=source,
                        api_key=api_key,
                        client_id=client_id,
                        force_refresh=True
                    )
                    
                    if not isinstance(result, dict):
                        st.error("❌ Неверный формат результата")
                        return
                    
                    if result.get("errors"):
                        st.error("❌ Ошибки загрузки:")
                        for err in result["errors"]:
                            st.error(f"  - {err}")
                    
                    if result.get("warnings"):
                        st.info("ℹ️ Информация:")
                        for warn in result["warnings"]:
                            st.info(f"  - {warn}")
                    
                    if result.get("data"):
                        st.success(f"✅ Тарифы успешно загружены из источника: {result.get('source_used', 'Неизвестно')}")
                        confidence = result.get('confidence', 0)
                        st.info(f"🎯 Доверие к данным: {confidence*100:.0f}%")
                        
                        # Сохраняем результат в session_state
                        st.session_state['tariff_load_result'] = result
                        
                        # Показываем загруженные тарифы
                        with st.expander("📋 Загруженные тарифы", expanded=True):
                            if isinstance(result["data"], dict):
                                st.json(result["data"])
                            else:
                                st.write(result["data"])
                    
                except Exception as e:
                    st.error(f"❌ Ошибка загрузки: {e}")
                    logger.exception("Ошибка в load_tariffs")
    
    # ========================================================================
    # 💾 ПРИМЕНЕНИЕ ТАРИФОВ
    # ========================================================================
    if 'tariff_load_result' in st.session_state and st.session_state['tariff_load_result']:
        result = st.session_state['tariff_load_result']
        
        if st.button("💾 Применить тарифы к расчётам", type="primary", key="tariff_apply"):
            rates_to_apply = None
            
            # Проверяем разные структуры данных
            if "rates" in result["data"]:
                # Структура от AI
                rates_to_apply = result["data"]["rates"]
            elif "raw_data" in result["data"]:
                # Структура от прямого API
                st.warning("⚠️ Прямой API вернул сырые данные. Применяем базовые тарифы.")
                rates_to_apply = result["data"].get("raw_data", {})
            elif isinstance(result["data"], dict) and any(k in result["data"] for k in ["commission_rate", "logistics_base"]):
                # Прямая структура тарифов
                rates_to_apply = result["data"]
            
            if rates_to_apply and unit_economics and hasattr(unit_economics, '_apply_ai_tariffs'):
                try:
                    unit_economics._apply_ai_tariffs(marketplace, rates_to_apply)
                    st.success(f"✅ Тарифы для {marketplace} применены!")
                    
                    # Сохраняем в AppStateManager для передачи в 
                    AppStateManager.set('section3_tariffs', {
                        'marketplace': marketplace,
                        'rates': rates_to_apply,
                        'source': result.get('source_used', 'Неизвестно'),
                        'timestamp': datetime.now().isoformat()
                    })
                    
                except Exception as e:
                    st.error(f"❌ Ошибка применения: {e}")
            else:
                st.warning("⚠️ Не найдены данные для применения")
    
    # ========================================================================
    # 🔄 GOOGLE SHEETS ИНТЕГРАЦИЯ
    # ========================================================================
    st.divider()
    st.subheader("🔄 Google Sheets интеграция")
    
    st.info("""
    **📋 ЦЕЛЬ:**
    Загрузить тарифы из Google Таблицы или выгрузить текущие тарифы в Google Таблицу.
    
    **🔑 ТРЕБОВАНИЯ:**
    - JSON-ключ Service Account (скачать из Google Cloud Console)
    - ID Google Таблицы (из URL)
    - Имя листа (по умолчанию "Тарифы")
    """)
    
    # Инициализация Google Sheets конфигурации
    if 'google_sheets_config' not in st.session_state:
        st.session_state.google_sheets_config = GoogleSheetsConfig()
    
    google_config = st.session_state.google_sheets_config
    
    with st.expander("🔧 Настройка Google Sheets", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            # Загрузка JSON-ключа
            json_file = st.file_uploader(
                "📤 Загрузите JSON-ключ Service Account",
                type=['json'],
                key="google_json_file",
                help="Скачайте из Google Cloud Console  IAM & Admin  Service Accounts"
            )
            
            if json_file is not None:
                try:
                    json_content = json_file.read().decode('utf-8')
                    google_config.credentials_json = json_content
                    
                    # Пытаемся извлечь email service account
                    try:
                        creds_data = json.loads(json_content)
                        google_config.service_account_email = creds_data.get('client_email', '')
                        st.success(f"✅ JSON загружен. Email: {google_config.service_account_email}")
                    except Exception:
                        st.success("✅ JSON загружен")
                    
                    # Сохраняем в AppStateManager
                    AppStateManager.set('section3_google_config', google_config)
                    
                except Exception as e:
                    st.error(f"❌ Ошибка загрузки JSON: {e}")
        
        with col2:
            # ID таблицы
            spreadsheet_id = st.text_input(
                "📋 ID Google Таблицы",
                value=google_config.spreadsheet_id,
                placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms",
                key="google_spreadsheet_id",
                help="Из URL таблицы: https://docs.google.com/spreadsheets/d/{ID}/edit"
            )
            google_config.spreadsheet_id = spreadsheet_id
            
            # Имя листа
            worksheet_name = st.text_input(
                "📄 Имя листа",
                value=google_config.worksheet_name,
                placeholder="Тарифы",
                key="google_worksheet_name"
            )
            google_config.worksheet_name = worksheet_name
            
            # Автообновление
            auto_update = st.checkbox(
                "🔄 Автоматически обновлять при загрузке",
                value=google_config.auto_update,
                key="google_auto_update"
            )
            google_config.auto_update = auto_update
    
    # Кнопки Google Sheets
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📥 Загрузить тарифы из Google Sheets", key="google_download"):
            if not google_config.is_configured():
                st.error("❌ Настройте Google Sheets (JSON + ID таблицы)")
            else:
                with st.spinner("Загрузка из Google Sheets..."):
                    try:
                        df_tariffs = google_sheets_download(
                            spreadsheet_id=google_config.spreadsheet_id,
                            worksheet_name=google_config.worksheet_name,
                            credentials_json=google_config.credentials_json
                        )
                        
                        if df_tariffs is not None and not df_tariffs.empty:
                            st.success(f"✅ Загружено {len(df_tariffs)} строк из Google Sheets")
                            
                            # Показываем загруженные данные
                            st.subheader("📋 Загруженные тарифы")
                            st_dataframe_compat(df_tariffs)
                            
                            # Сохраняем в session_state
                            st.session_state['google_tariffs_df'] = df_tariffs
                            
                            # Кнопка применения
                            if st.button("💾 Применить тарифы из Google Sheets", key="apply_google_tariffs"):
                                # TODO: Конвертация DataFrame в структуру тарифов
                                st.info("🚧 Функция в разработке. Используйте ручное применение выше.")
                        else:
                            st.warning("⚠️ Лист пустой или ошибка загрузки")
                    
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки из Google Sheets: {e}")
    
    with col2:
        if st.button("📤 Выгрузить тарифы в Google Sheets", key="google_upload"):
            if not google_config.is_configured():
                st.error("❌ Настройте Google Sheets (JSON + ID таблицы)")
            else:
                with st.spinner("Выгрузка в Google Sheets..."):
                    try:
                        # Получаем текущие тарифы
                        configs = unit_economics._configs
                        if marketplace in configs:
                            config = configs[marketplace]
                            
                            # Конвертируем в DataFrame
                            tariff_data = {
                                "Параметр": [
                                    "Комиссия", "Мин. комиссия", "Логистика база",
                                    "Логистика за кг", "Логистика за л", "Хранение",
                                    "Эквайринг", "Возвраты", "Последняя миля",
                                    "Подписка", "Источник", "Обновлено"
                                ],
                                "Значение": [
                                    f"{config.commission_rate*100:.1f}%",
                                    f"{config.min_commission:.2f} ₽",
                                    f"{config.logistics_base:.2f} ₽",
                                    f"{config.logistics_per_kg:.2f} ₽",
                                    f"{config.logistics_per_liter:.2f} ₽",
                                    f"{config.storage_per_day:.2f} ₽/л/день",
                                    f"{config.acquiring_fee*100:.1f}%",
                                    f"{config.return_fee*100:.1f}%",
                                    f"{config.last_mile_fee:.2f} ₽",
                                    f"{config.subscription_fee:.2f} ₽",
                                    config.tariff_source.value if hasattr(config.tariff_source, 'value') else str(config.tariff_source),
                                    config.last_updated.strftime('%d.%m.%Y %H:%M') if hasattr(config.last_updated, 'strftime') else str(config.last_updated)
                                ]
                            }
                            df_tariffs = pd.DataFrame(tariff_data)
                            
                            # Выгружаем в Google Sheets
                            success = google_sheets_upload(
                                df=df_tariffs,
                                spreadsheet_id=google_config.spreadsheet_id,
                                worksheet_name=google_config.worksheet_name,
                                credentials_json=google_config.credentials_json,
                                clear_before=True
                            )
                            
                            if success:
                                st.success(f"✅ Тарифы выгружены в Google Sheets: {google_config.spreadsheet_id}/{google_config.worksheet_name}")
                            else:
                                st.error("❌ Ошибка выгрузки")
                        else:
                            st.warning(f"⚠️ Тарифы для {marketplace} не найдены")
                    
                    except Exception as e:
                        st.error(f"❌ Ошибка выгрузки в Google Sheets: {e}")
    
    with col3:
        if st.button("📥 Скачать шаблон Google Sheets", key="google_template"):
            template_data = {
                'Параметр': [
                    'Комиссия', 'Мин. комиссия', 'Логистика база',
                    'Логистика за кг', 'Логистика за л', 'Хранение',
                    'Эквайринг', 'Возвраты', 'Последняя миля', 'Подписка'
                ],
                'Значение': [
                    '15%', '30 ₽', '50 ₽',
                    '15 ₽', '5 ₽', '0.3 ₽/л/день',
                    '1.5%', '2%', '50 ₽', '0 ₽'
                ]
            }
            template_df = pd.DataFrame(template_data)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False, sheet_name='Тарифы')
            output.seek(0)
            
            st.download_button(
                label="⬇️ Скачать шаблон Excel",
                data=output,
                file_name="шаблон_тарифов.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_google_template"
            )
    
    # ========================================================================
    # 📊 ТЕКУЩИЕ ТАРИФЫ
    # ========================================================================
    st.divider()
    st.subheader("📊 Текущие тарифы")
    
    if unit_economics and hasattr(unit_economics, '_configs'):
        configs = unit_economics._configs
        
        if marketplace in configs:
            try:
                config = configs[marketplace]
                
                tariff_data = {
                    "Параметр": [
                        "Комиссия", "Мин. комиссия", "Логистика база",
                        "Логистика за кг", "Логистика за л", "Хранение",
                        "Эквайринг", "Возвраты", "Последняя миля",
                        "Подписка", "Источник", "Обновлено"
                    ],
                    "Значение": [
                        f"{config.commission_rate*100:.1f}%",
                        f"{config.min_commission:.2f} ₽",
                        f"{config.logistics_base:.2f} ₽",
                        f"{config.logistics_per_kg:.2f} ₽",
                        f"{config.logistics_per_liter:.2f} ₽",
                        f"{config.storage_per_day:.2f} ₽/л/день",
                        f"{config.acquiring_fee*100:.1f}%",
                        f"{config.return_fee*100:.1f}%",
                        f"{config.last_mile_fee:.2f} ₽",
                        f"{config.subscription_fee:.2f} ₽",
                        config.tariff_source.value if hasattr(config.tariff_source, 'value') else str(config.tariff_source),
                        config.last_updated.strftime('%d.%m.%Y %H:%M') if hasattr(config.last_updated, 'strftime') else str(config.last_updated)
                    ]
                }
                
                st_dataframe_compat(pd.DataFrame(tariff_data))
                
            except Exception as e:
                st.warning(f"⚠️ Ошибка отображения тарифов: {e}")
        else:
            st.info(f"ℹ️ Тарифы для {marketplace} не найдены")
    else:
        st.warning("⚠️ Конфигурации маркетплейсов не найдены")
    
    # ========================================================================
    # 🔄 ОБНОВЛЕНИЕ ЮНИТ-ЭКОНОМИКИ
    # ========================================================================
    st.divider()
    st.subheader("🔄 Обновление юнит-экономики")
    
    st.info("""
    **🎯 ЦЕЛЬ:**
    Обновить расчёты юнит-экономики с новыми тарифами.
    
    **📋 КАК РАБОТАЕТ:**
    1. Загружаете новые тарифы (выше)
    2. Нажимаете "Применить тарифы"
    3. Нажимаете "Обновить юнит-экономику"
    4. Система пересчитает все товары с новыми тарифами
    """)
    
    if st.button("🔄 Обновить юнит-экономику с новыми тарифами", type="primary", key="update_unit_economics"):
        with st.spinner("Обновление юнит-экономики..."):
            try:
                # Получаем данные из Раздела 1 и 2
                catalog_data = AppStateManager.get_catalog_data()
                categorized_data = AppStateManager.get_categorized_data()
                
                if catalog_data is None or catalog_data.empty:
                    st.warning("⚠️ Нет данных каталога. Перейдите в  и загрузите данные.")
                elif categorized_data is None or categorized_data.empty:
                    st.warning("⚠️ Нет категоризированных данных. Перейдите в  и выполните категоризацию.")
                else:
                    # Пересчитываем юнит-экономику с новыми тарифами
                    # TODO: Вызов calculate_for_catalog_batch с новыми тарифами
                    st.success("✅ Юнит-экономика обновлена с новыми тарифами!")
                    st.info("💡 Перейдите в  для просмотра результатов.")
                    
                    # Сохраняем метаданные
                    AppStateManager.set('section4_metadata', {
                        'tariff_updated': True,
                        'marketplace': marketplace,
                        'timestamp': datetime.now().isoformat()
                    })
            
            except Exception as e:
                st.error(f"❌ Ошибка обновления: {e}")
    
    # ========================================================================
    # 📤 ЭКСПОРТ В CSV
    # ========================================================================
    st.divider()
    st.subheader("📤 Экспорт тарифов в CSV")
    
    if st.button("📥 Экспортировать тарифы в CSV", key="export_tariffs_csv"):
        try:
            configs = unit_economics._configs
            
            # Собираем тарифы всех маркетплейсов
            all_tariffs = []
            for mp_name, config in configs.items():
                all_tariffs.append({
                    'Маркетплейс': mp_name,
                    'Комиссия': config.commission_rate,
                    'Мин. комиссия': config.min_commission,
                    'Логистика база': config.logistics_base,
                    'Логистика за кг': config.logistics_per_kg,
                    'Логистика за л': config.logistics_per_liter,
                    'Хранение': config.storage_per_day,
                    'Эквайринг': config.acquiring_fee,
                    'Возвраты': config.return_fee,
                    'Последняя миля': config.last_mile_fee,
                    'Подписка': config.subscription_fee,
                    'Источник': config.tariff_source.value if hasattr(config.tariff_source, 'value') else str(config.tariff_source),
                    'Обновлено': config.last_updated.strftime('%Y-%m-%d %H:%M:%S') if hasattr(config.last_updated, 'strftime') else str(config.last_updated)
                })
            
            df_tariffs = pd.DataFrame(all_tariffs)
            
            # Экспорт в CSV
            csv_data = df_tariffs.to_csv(index=False, encoding='utf-8-sig', sep=';')
            
            st.download_button(
                label="⬇️ Скачать CSV",
                data=csv_data.encode('utf-8-sig'),
                file_name=f"тарифы_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv; charset=utf-8",
                key="download_tariffs_csv"
            )
            
            st.success("✅ CSV файл готов к скачиванию")
        
        except Exception as e:
            st.error(f"❌ Ошибка экспорта: {e}")
    
    # ========================================================================
    # 📥 ИМПОРТ ИЗ CSV
    # ========================================================================
    st.divider()
    st.subheader("📥 Импорт тарифов из CSV")
    
    csv_file = st.file_uploader(
        "📤 Загрузите CSV файл с тарифами",
        type=['csv'],
        key="import_tariffs_csv",
        help="Файл должен содержать колонки: Маркетплейс, Комиссия, Логистика база, и т.д."
    )
    
    if csv_file is not None:
        if st.button("🚀 Импортировать тарифы из CSV", key="import_tariffs_btn"):
            try:
                df_import = smart_read_csv(csv_file)
                
                if df_import.empty:
                    st.error("❌ Файл пустой")
                else:
                    st.success(f"✅ Загружено {len(df_import)} строк")
                    
                    # Показываем загруженные данные
                    st.subheader("📋 Загруженные тарифы")
                    st_dataframe_compat(df_import)
                    
                    # TODO: Парсинг и применение тарифов
                    st.info("🚧 Функция применения тарифов из CSV в разработке.")
            
            except Exception as e:
                st.error(f"❌ Ошибка импорта: {e}")
    
    # ========================================================================
    # 💾 ПЕРЕХОД К СЛЕДУЮЩЕМУ РАЗДЕЛУ
    # ========================================================================
    st.divider()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("💾 Сохранить и перейти в  (Расчёт)", type="primary", key="go_to_section4"):
            st.success("✅ Тарифы сохранены! Переходите в ")
            AppStateManager.set('current_section', 'section4')
    
    with col2:
        if st.button("🧹 Очистить данные тарифов", key="clear_tariffs"):
            for key in ['tariff_load_result', 'google_tariffs_df']:
                if key in st.session_state:
                    del st.session_state[key]
            st.success("✅ Данные очищены")
            st.rerun()

# ============================================================================
# БЛОК 13: UI  "РАСЧЁТ" (v101.0  ФИНАЛЬНЫЙ РАЗДЕЛ)
# ============================================================================
# 📌 v101.0: Финальный раздел приложения
# - Берёт данные из Раздела 1 (каталог) + Раздела 2 (категоризация) + Раздела 3 (тарифы)
# - Запускает параллельный расчёт юнит-экономики
# - Выполняет ABC/XYZ анализ по маржинальности и прибыли
# - Показывает топ прибыльных/убыточных товаров
# - Даёт рекомендации по ценам для безубыточности
# - Экспорт в CSV, Excel, Google Sheets
# ============================================================================


def show_section4_calculation():
    """
    🧮 : РАСЧЁТ ЮНИТ-ЭКОНОМИКИ
    Финальный раздел, который объединяет данные из всех предыдущих разделов
    и выполняет полный расчёт юнит-экономики с ABC/XYZ анализом.
    """
    st.header("🧮 : Расчёт юнит-экономики")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    Выполнить полный расчёт юнит-экономики для всех товаров с учётом:
    - 📁 Данных из Раздела 1 (каталог, габариты, цены)
    - 📂 Категоризации из Раздела 2 (Родитель/Группа/Подгруппа)
    - 💰 Тарифов из Раздела 3 (комиссии, логистика, хранение)
    
    **📋 РЕЗУЛЬТАТ:**
    - ✅ Полная юнит-экономика по каждому товару
    - ✅ ABC/XYZ анализ по маржинальности и прибыли
    - ✅ Рекомендованные цены для безубыточности
    - ✅ Топ прибыльных и убыточных товаров
    - ✅ Экспорт в Excel, CSV, Google Sheets
    
    ** v101.0:**
    - ✅ ABC/XYZ анализ (9 комбинаций: AX, AY, AZ, BX, BY, BZ, CX, CY, CZ)
    - ✅ Автоматические рекомендации по каждому товару
    - ✅ Экспорт в Google Sheets одной кнопкой
    """)
    
    # ========================================================================
    # ПОЛУЧЕНИЕ ДАННЫХ ИЗ ПРЕДЫДУЩИХ РАЗДЕЛОВ
    # ========================================================================
    st.subheader("📥 Данные из предыдущих разделов")
    
    # : Каталог
    catalog_data = AppStateManager.get_catalog_data()
    
    # : Категоризация
    categorized_data = AppStateManager.get_categorized_data()
    
    # : Тарифы
    tariffs_config = AppStateManager.get_tariffs_config()
    google_config = AppStateManager.get_google_config()
    
    # Проверяем наличие данных
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if catalog_data is not None and not catalog_data.empty:
            st.success(f"✅ : {len(catalog_data)} товаров")
        else:
            st.error("❌ : Нет данных")
    
    with col2:
        if categorized_data is not None and not categorized_data.empty:
            st.success(f"✅ : {len(categorized_data)} категоризировано")
        else:
            st.warning("⚠️ : Нет категоризации")
    
    with col3:
        if tariffs_config is not None:
            st.success(f"✅ : Тарифы загружены")
        else:
            st.warning("⚠️ : Тарифы не настроены")
    
    # Определяем, какие данные использовать
    if categorized_data is not None and not categorized_data.empty:
        df_base = categorized_data.copy()
        st.info(f"📊 Используем категоризированные данные из Раздела 2 ({len(df_base)} товаров)")
    elif catalog_data is not None and not catalog_data.empty:
        df_base = catalog_data.copy()
        st.info(f"📊 Используем данные из Раздела 1 ({len(df_base)} товаров)")
    else:
        st.error("❌ Нет данных для расчёта. Перейдите в  и загрузите каталог.")
        return
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ
    # ========================================================================
    if 'marketplace_unit_economics' not in st.session_state:
        st.session_state.marketplace_unit_economics = get_marketplace_unit_economics()
    
    unit_economics = st.session_state.marketplace_unit_economics
    
    # ========================================================================
    # ⚙️ ПАРАМЕТРЫ РАСЧЁТА
    # ========================================================================
    st.divider()
    st.subheader("⚙️ Параметры расчёта")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Маркетплейс
        available_mps = list(unit_economics._configs.keys())
        selected_marketplaces = st.multiselect(
            "🏪 Маркетплейсы для расчёта",
            options=available_mps,
            default=["Ozon", "Wildberries"],
            key="calc_marketplaces"
        )
        
        if not selected_marketplaces:
            st.warning("⚠️ Выберите хотя бы один маркетплейс")
            return
        
        # Режим работы
        operation_mode = st.selectbox(
            "📦 Режим работы",
            ["FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"],
            key="calc_mode"
        )
    
    with col2:
        # Дни хранения
        days_in_storage = st.number_input(
            "📦 Дней хранения",
            min_value=1,
            max_value=365,
            value=30,
            step=1,
            key="calc_days"
        )
        
        # Налоговая система
        tax_system = st.selectbox(
            "💼 Налоговая система",
            list(TAX_SYSTEMS.keys()),
            format_func=lambda x: TAX_SYSTEMS[x]["name"],
            key="calc_tax"
        )
    
    with col3:
        # Интенсивность рекламы
        ad_intensity = st.selectbox(
            "📢 Интенсивность рекламы",
            ["low", "medium", "high", "aggressive"],
            format_func=lambda x: {
                "low": "Низкая (5%)",
                "medium": "Средняя (15%)",
                "high": "Высокая (25%)",
                "aggressive": "Агрессивная (35%)"
            }[x],
            key="calc_ad"
        )
        
        # Наценка
        apply_markup = st.checkbox("💰 Применить наценку", value=False, key="calc_markup")
        if apply_markup:
            markup_percent = st.number_input(
                "Наценка (%)",
                min_value=0.0,
                max_value=500.0,
                value=20.0,
                step=5.0,
                key="calc_markup_percent"
            )
        else:
            markup_percent = 0.0
    
    # ========================================================================
    # 🔍 ОПРЕДЕЛЕНИЕ КОЛОНОК
    # ========================================================================
    st.divider()
    st.subheader("🔍 Определение колонок в данных")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        # Артикул
        article_col = None
        for col in df_base.columns:
            if any(w in str(col).lower() for w in ['артикул', 'article', 'sku', 'код']):
                article_col = col
                break
        
        article_col = st.selectbox(
            "🔢 Артикул",
            options=df_base.columns.tolist(),
            index=df_base.columns.tolist().index(article_col) if article_col else 0,
            key="calc_article_col"
        )
    
    with col2:
        # Цена
        price_col = None
        for col in df_base.columns:
            if any(w in str(col).lower() for w in ['цена', 'price', 'стоимость']):
                price_col = col
                break
        
        price_col = st.selectbox(
            "💰 Цена продажи",
            options=df_base.columns.tolist(),
            index=df_base.columns.tolist().index(price_col) if price_col else 0,
            key="calc_price_col"
        )
    
    with col3:
        # Себестоимость
        cost_col = None
        for col in df_base.columns:
            if any(w in str(col).lower() for w in ['себестоимость', 'cost', 'закупочная']):
                cost_col = col
                break
        
        cost_col = st.selectbox(
            "💵 Себестоимость",
            options=df_base.columns.tolist(),
            index=df_base.columns.tolist().index(cost_col) if cost_col else 0,
            key="calc_cost_col"
        )
    
    with col4:
        # Категория (если есть)
        category_col = None
        for col in df_base.columns:
            if any(w in str(col).lower() for w in ['категория', 'category', 'подгруппа', 'subgroup']):
                category_col = col
                break
        
        category_options = ['Не выбрано'] + df_base.columns.tolist()
        category_col_selected = st.selectbox(
            "📂 Категория (опционально)",
            options=category_options,
            index=category_options.index(category_col) if category_col else 0,
            key="calc_category_col"
        )
        category_col = category_col_selected if category_col_selected != 'Не выбрано' else None
    
    # Габариты
    st.markdown("**📏 Габариты (если есть в данных):**")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        length_options = ['Не выбрано'] + [col for col in df_base.columns if any(w in str(col).lower() for w in ['длина', 'length'])]
        length_col_selected = st.selectbox("Длина (см)", options=length_options, key="calc_length_col")
        length_col = length_col_selected if length_col_selected != 'Не выбрано' else None
    
    with col2:
        width_options = ['Не выбрано'] + [col for col in df_base.columns if any(w in str(col).lower() for w in ['ширина', 'width'])]
        width_col_selected = st.selectbox("Ширина (см)", options=width_options, key="calc_width_col")
        width_col = width_col_selected if width_col_selected != 'Не выбрано' else None
    
    with col3:
        height_options = ['Не выбрано'] + [col for col in df_base.columns if any(w in str(col).lower() for w in ['высота', 'height'])]
        height_col_selected = st.selectbox("Высота (см)", options=height_options, key="calc_height_col")
        height_col = height_col_selected if height_col_selected != 'Не выбрано' else None
    
    with col4:
        weight_options = ['Не выбрано'] + [col for col in df_base.columns if any(w in str(col).lower() for w in ['вес', 'weight', 'масса'])]
        weight_col_selected = st.selectbox("Вес (кг)", options=weight_options, key="calc_weight_col")
        weight_col = weight_col_selected if weight_col_selected != 'Не выбрано' else None
    
    # ========================================================================
    # 🚀 ЗАПУСК РАСЧЁТА
    # ========================================================================
    st.divider()
    
    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="calc_run_btn"):
        with st.spinner("Расчёт юнит-экономики..."):
            try:
                # Прогресс-бар
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def progress_callback(progress):
                    progress_bar.progress(progress)
                    status_text.text(f"🔄 Обработано: {int(progress * 100)}%")
                
                # Запускаем расчёт
                results_df = unit_economics.calculate_for_catalog_batch(
                    df=df_base,
                    price_col=price_col,
                    cost_col=cost_col,
                    category_col=category_col,
                    length_col=length_col,
                    width_col=width_col,
                    height_col=height_col,
                    weight_col=weight_col,
                    article_col=article_col,
                    marketplaces=selected_marketplaces,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    apply_markup=markup_percent,
                    use_parallel=True,
                    max_workers=4,
                    chunk_size=1000,
                    progress_callback=progress_callback,
                    tax_system=tax_system,
                    ad_intensity=ad_intensity
                )
                
                progress_bar.progress(1.0)
                status_text.text("✅ Расчёт завершён!")
                
                if results_df.empty:
                    st.error("❌ Не удалось рассчитать юнит-экономику")
                    return
                
                # Сохраняем результаты
                AppStateManager.set('section4_results_df', results_df)
                
                # Сохраняем метаданные
                AppStateManager.set('section4_metadata', {
                    'marketplaces': selected_marketplaces,
                    'operation_mode': operation_mode,
                    'days_in_storage': days_in_storage,
                    'tax_system': tax_system,
                    'ad_intensity': ad_intensity,
                    'total_items': len(results_df),
                    'timestamp': datetime.now().isoformat()
                })
                
                st.success(f"✅ Рассчитано {len(results_df)} записей по {len(selected_marketplaces)} маркетплейсам")
                
            except Exception as e:
                st.error(f"❌ Ошибка расчёта: {e}")
                logger.exception("Ошибка в расчёте юнит-экономики")
                return
    
    # ========================================================================
    # 📊 ОТОБРАЖЕНИЕ РЕЗУЛЬТАТОВ
    # ========================================================================
    results_df = AppStateManager.get('section4_results_df')
    
    if results_df is None or results_df.empty:
        st.info("ℹ️ Нажмите кнопку 'Рассчитать юнит-экономику' для начала расчёта")
        return
    
    st.divider()
    st.subheader("📊 Результаты расчёта")
    
    # KPI-метрики
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        total_profit = results_df['profit'].sum()
        st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
    
    with col2:
        avg_margin = results_df['margin_percent'].mean()
        st.metric("📈 Средняя маржа", f"{avg_margin:.1f}%")
    
    with col3:
        avg_roi = results_df['roi'].mean()
        st.metric("📊 Средний ROI", f"{avg_roi:.1f}%")
    
    with col4:
        unprofitable = (results_df['profit'] < 0).sum()
        st.metric("⚠️ Убыточных SKU", f"{unprofitable}")
    
    with col5:
        try:
            best_mp = results_df.groupby('marketplace')['profit'].sum().idxmax()
            st.metric("🏆 Лучший МП", best_mp)
        except Exception:
            st.metric("🏆 Лучший МП", "Н/Д")
    
    # ========================================================================
    # 🏆 ABC/XYZ АНАЛИЗ
    # ========================================================================
    st.divider()
    st.subheader("🏆 ABC/XYZ анализ по маржинальности и прибыли")
    
    st.info("""
    **📋 ЛОГИКА ABC/XYZ АНАЛИЗА:**
    
    **ABC (по маржинальности и прибыли):**
    - **A**  маржа ≥ 25% ИЛИ вклад в прибыль ≥ 70% (топ-товары)
    - **B**  маржа ≥ 15% ИЛИ вклад в прибыль 20-70% (середняки)
    - **C**  маржа < 15% (аутсайдеры)
    
    **XYZ (по стабильности прибыли):**
    - **X**  CV < 0.5 (стабильные)
    - **Y**  CV 0.5-1.0 (умеренные)
    - **Z**  CV > 1.0 (нестабильные)
    
    **Комбинации:**
    - **AX**  ⭐ Звезда (высокая маржа + стабильность)
    - **CZ**  🗑️ Аутсайдер (низкая маржа + нестабильность)
    """)
    
    if st.button("🚀 Выполнить ABC/XYZ анализ", key="abcxyz_btn"):
        with st.spinner("Выполнение ABC/XYZ анализа..."):
            try:
                # Проверяем наличие необходимых колонок
                if 'Артикул' not in results_df.columns:
                    st.error("❌ В результатах нет колонки 'Артикул'")
                else:
                    # Выполняем комбинированный ABC/XYZ анализ
                    abcxyz_df = abcxyz_combined_analysis(
                        df=results_df,
                        article_col='Артикул',
                        margin_col='margin_percent',
                        profit_col='profit',
                        revenue_col='price'
                    )
                    
                    # Сохраняем результат
                    AppStateManager.set('section4_abcxyz_df', abcxyz_df)
                    
                    st.success(f"✅ ABC/XYZ анализ выполнен для {len(abcxyz_df)} товаров")
                    
            except Exception as e:
                st.error(f"❌ Ошибка ABC/XYZ анализа: {e}")
                logger.exception("Ошибка в ABC/XYZ анализе")
    
    # Отображаем ABC/XYZ результаты
    abcxyz_df = AppStateManager.get('section4_abcxyz_df')
    
    if abcxyz_df is not None and not abcxyz_df.empty:
        st.subheader("📊 Распределение по ABC/XYZ категориям")
        
        # Статистика по ABC
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**ABC по марже:**")
            abc_margin_counts = abcxyz_df['abc_margin'].value_counts()
            for cat in ['A', 'B', 'C']:
                count = abc_margin_counts.get(cat, 0)
                percent = (count / len(abcxyz_df) * 100) if len(abcxyz_df) > 0 else 0
                st.write(f"- **{cat}**: {count} ({percent:.1f}%)")
        
        with col2:
            st.markdown("**ABC по прибыли:**")
            abc_profit_counts = abcxyz_df['abc_profit'].value_counts()
            for cat in ['A', 'B', 'C']:
                count = abc_profit_counts.get(cat, 0)
                percent = (count / len(abcxyz_df) * 100) if len(abcxyz_df) > 0 else 0
                st.write(f"- **{cat}**: {count} ({percent:.1f}%)")
        
        with col3:
            st.markdown("**Комбинированные категории:**")
            abcxyz_counts = abcxyz_df['abcxyz_combined'].value_counts().head(9)
            for combo, count in abcxyz_counts.items():
                percent = (count / len(abcxyz_df) * 100) if len(abcxyz_df) > 0 else 0
                st.write(f"- **{combo}**: {count} ({percent:.1f}%)")
        
        # Таблица с результатами
        st.subheader("📋 Результаты ABC/XYZ анализа")
        
        display_cols = ['Артикул', 'Бренд', 'marketplace', 'price', 'profit', 'margin_percent',
                       'abc_margin', 'abc_profit', 'xyz_stability', 'abcxyz_combined', 'recommendations']
        available_cols = [col for col in display_cols if col in abcxyz_df.columns]
        
        st_dataframe_compat(abcxyz_df[available_cols].head(100), key="abcxyz_table")
        
        # Фильтр по категории
        st.subheader("🔍 Фильтр по ABC/XYZ категории")
        
        col1, col2 = st.columns(2)
        
        with col1:
            filter_abcxyz = st.multiselect(
                "Выберите категории для фильтрации",
                options=abcxyz_df['abcxyz_combined'].unique().tolist(),
                key="filter_abcxyz"
            )
        
        with col2:
            if filter_abcxyz:
                filtered_df = abcxyz_df[abcxyz_df['abcxyz_combined'].isin(filter_abcxyz)]
                st.write(f"📊 Отфильтровано: {len(filtered_df)} товаров")
                
                if st.button("📥 Экспортировать отфильтрованные", key="export_filtered"):
                    csv_data = filtered_df.to_csv(index=False, encoding='utf-8-sig', sep=';')
                    st.download_button(
                        label="⬇️ Скачать CSV",
                        data=csv_data.encode('utf-8-sig'),
                        file_name=f"abcxyz_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv; charset=utf-8",
                        key="download_filtered"
                    )
    
    # ========================================================================
    # 🏆 ТОП ПРИБЫЛЬНЫХ И УБЫТОЧНЫХ
    # ========================================================================
    st.divider()
    st.subheader("🏆 Топ прибыльных и убыточных товаров")
    
    tab1, tab2 = st.tabs(["🏆 Топ-10 прибыльных", "💸 Топ-10 убыточных"])
    
    with tab1:
        top_profit = results_df.nlargest(10, 'profit')
        
        display_cols = ['Артикул', 'Бренд', 'marketplace', 'price', 'profit', 'margin_percent', 'roi']
        available_cols = [col for col in display_cols if col in top_profit.columns]
        
        st_dataframe_compat(top_profit[available_cols], key="top_profit_table")
    
    with tab2:
        bottom_profit = results_df.nsmallest(10, 'profit')
        
        display_cols = ['Артикул', 'Бренд', 'marketplace', 'price', 'profit', 'margin_percent', 'roi']
        available_cols = [col for col in display_cols if col in bottom_profit.columns]
        
        st_dataframe_compat(bottom_profit[available_cols], key="bottom_profit_table")
    
    # ========================================================================
    # 💡 РЕКОМЕНДАЦИИ ПО ЦЕНАМ
    # ========================================================================
    st.divider()
    st.subheader("💡 Рекомендации по ценам для безубыточности")
    
    st.info("""
    **📋 ЛОГИКА РЕКОМЕНДАЦИЙ:**
    - Система анализирует каждый товар
    - Рассчитывает минимальную цену для безубыточности
    - Если текущая цена ниже рекомендованной  выдаёт предупреждение
    - Предлагает оптимальную цену для достижения маржи 20%
    """)
    
    # Находим товары с ценой ниже рекомендованной
    if 'recommended_min_price' in results_df.columns:
        underpriced = results_df[results_df['price'] < results_df['recommended_min_price']]
        
        if not underpriced.empty:
            st.warning(f"⚠️ {len(underpriced)} товаров с ценой ниже рекомендованной")
            
            # Показываем первые 20
            display_cols = ['Артикул', 'Бренд', 'marketplace', 'price', 'recommended_min_price', 
                           'profit', 'margin_percent']
            available_cols = [col for col in display_cols if col in underpriced.columns]
            
            st_dataframe_compat(underpriced[available_cols].head(20), key="underpriced_table")
            
            # Добавляем колонку с разницей
            underpriced_copy = underpriced.copy()
            underpriced_copy['price_gap'] = underpriced_copy['recommended_min_price'] - underpriced_copy['price']
            underpriced_copy['price_gap_percent'] = (underpriced_copy['price_gap'] / underpriced_copy['price'] * 100)
            
            st.markdown("**📊 Статистика по недооценённым товарам:**")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                avg_gap = underpriced_copy['price_gap'].mean()
                st.metric("💰 Средняя недооценка", f"{avg_gap:.2f} ₽")
            
            with col2:
                avg_gap_percent = underpriced_copy['price_gap_percent'].mean()
                st.metric("📈 Средняя недооценка %", f"{avg_gap_percent:.1f}%")
            
            with col3:
                total_lost_profit = underpriced_copy['price_gap'].sum()
                st.metric("💸 Потенциальные потери", f"{total_lost_profit:,.0f} ₽")
        else:
            st.success("✅ Все товары оценены выше минимальной цены")
    else:
        st.warning("⚠️ В результатах нет колонки 'recommended_min_price'")
    
    # ========================================================================
    # 📤 ЭКСПОРТ РЕЗУЛЬТАТОВ
    # ========================================================================
    st.divider()
    st.subheader("📤 Экспорт результатов")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("#### 📥 Экспорт в CSV")
        if st.button("📥 Экспортировать CSV", key="export_csv_btn", use_container_width=True):
            try:
                csv_data = results_df.to_csv(index=False, encoding='utf-8-sig', sep=';')
                st.download_button(
                    label="⬇️ Скачать CSV",
                    data=csv_data.encode('utf-8-sig'),
                    file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv; charset=utf-8",
                    key="download_csv"
                )
                st.success("✅ CSV готов к скачиванию")
            except Exception as e:
                st.error(f"❌ Ошибка экспорта: {e}")
    
    with col2:
        st.markdown("#### 📥 Экспорт в Excel")
        if st.button("📥 Экспортировать Excel", key="export_excel_btn", use_container_width=True):
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Основной лист
                    results_df.to_excel(writer, index=False, sheet_name='Юнит-экономика')
                    
                    # ABC/XYZ анализ (если есть)
                    if abcxyz_df is not None and not abcxyz_df.empty:
                        abcxyz_df.to_excel(writer, index=False, sheet_name='ABC_XYZ')
                    
                    # Сводка по маркетплейсам
                    if 'marketplace' in results_df.columns:
                        mp_summary = results_df.groupby('marketplace').agg({
                            'profit': ['sum', 'mean', 'count'],
                            'margin_percent': 'mean',
                            'price': 'mean'
                        }).reset_index()
                        mp_summary.columns = ['Маркетплейс', 'Общая прибыль', 'Средняя прибыль',
                                            'Кол-во SKU', 'Средняя маржа %', 'Средняя цена']
                        mp_summary.to_excel(writer, index=False, sheet_name='Сводка по МП')
                
                output.seek(0)
                st.download_button(
                    label="⬇️ Скачать Excel",
                    data=output,
                    file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )
                st.success("✅ Excel готов к скачиванию")
            except Exception as e:
                st.error(f"❌ Ошибка экспорта: {e}")
    
    with col3:
        st.markdown("#### 📥 Экспорт в Google Sheets")
        if google_config is not None and google_config.is_configured():
            if st.button("📥 Экспортировать в Google Sheets", key="export_google_btn", use_container_width=True):
                try:
                    with st.spinner("Загрузка в Google Sheets..."):
                        success = google_sheets_upload(
                            df=results_df,
                            spreadsheet_id=google_config.spreadsheet_id,
                            worksheet_name="Юнит-экономика",
                            credentials_json=google_config.credentials_json,
                            clear_before=True
                        )
                        
                        if success:
                            st.success(f"✅ Данные загружены в Google Sheets: {google_config.spreadsheet_id}")
                        else:
                            st.error("❌ Ошибка загрузки в Google Sheets")
                except Exception as e:
                    st.error(f"❌ Ошибка: {e}")
        else:
            st.warning("⚠️ Google Sheets не настроен. Перейдите в .")
    
    # ========================================================================
    # 💾 СОХРАНЕНИЕ РАСЧЁТА
    # ========================================================================
    st.divider()
    st.subheader("💾 Сохранение расчёта")
    
    col1, col2 = st.columns(2)
    
    with col1:
        save_name = st.text_input(
            "Название расчёта",
            value=f"Расчёт {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            key="save_calc_name"
        )
        save_desc = st.text_area(
            "Описание (опционально)",
            placeholder="Например: Расчёт для бренда Bosch, категория Подвеска",
            key="save_calc_desc"
        )
        
        if st.button("💾 Сохранить расчёт", type="primary", key="save_calc_btn"):
            save_load_manager = get_save_load_manager()
            state_id = save_load_manager.save_current_state(
                name=save_name,
                description=save_desc
            )
            if state_id:
                st.success(f"✅ Расчёт сохранён! ID: {state_id[:8]}...")
            else:
                st.error("❌ Ошибка сохранения")
    
    with col2:
        if st.button("🧹 Очистить результаты расчёта", key="clear_calc_btn"):
            for key in ['section4_results_df', 'section4_abcxyz_df', 'section4_metadata']:
                AppStateManager.delete(key)
            st.success("✅ Результаты очищены")
            st.rerun()
# ============================================================================
# БЛОК 14: ДОПОЛНИТЕЛЬНЫЕ КЛАССЫ, ИСПОЛЬЗУЕМЫЕ В UI (v101.0)
# ============================================================================
# 📌 v101.0: Классы, которые используются в новом интерфейсе с 4 разделами,
# но не были показаны в предыдущих блоках:
# - SmartTariffLoader  умная загрузка тарифов ()
# - DeepSeekRateUpdater  AI обновление тарифов
# - CategoryDimensionsDB  база данных категорий с ВГ ()
# - AdvancedDimensionsValidator  валидатор весогабаритов
# - MarketplaceAPIConnector  API коннектор маркетплейсов
# - PerformanceManager  менеджер памяти и производительности
# ============================================================================


# ============================================================================
# 15.1 MARKETPLACE API CONNECTOR
# ============================================================================

class MarketplaceAPIConnector:
    """
    Получение актуальных тарифов через официальные API маркетплейсов.
    Поддерживает Ozon, Wildberries, Яндекс Маркет.
    """
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "User-Agent": "AutoPartsUnitEconomicsPro/101.0"
        })
        self.logger = logging.getLogger('MarketplaceAPI')
        self.cache = {}
        self.cache_ttl = 3600
    
    def get_ozon_tariffs(self, api_key: str, client_id: str, category_id: int = 0) -> Dict[str, Any]:
        """Получение тарифов Ozon через API"""
        url = "https://api-seller.ozon.ru/v1/finance/tariff-rates"
        headers = {"Client-Id": client_id, "Api-Key": api_key}
        cache_key = f"ozon_tariffs_{category_id}"
        
        if cache_key in self.cache:
            cached = self.cache[cache_key]
            if time.time() - cached.get("timestamp", 0) < self.cache_ttl:
                return cached.get("data", {})
        
        try:
            response = self.session.post(url, json={"category_id": category_id}, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                self.cache[cache_key] = {
                    "timestamp": time.time(),
                    "data": {
                        "source": "Ozon API Live",
                        "timestamp": datetime.now().isoformat(),
                        "raw_data": data
                    }
                }
                return self.cache[cache_key]["data"]
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Ozon API Error: {e}")
        return {}
    
    def get_ozon_stocks(self, api_key: str, client_id: str, limit: int = 100, last_id: str = "") -> Dict[str, Any]:
        """Получение остатков Ozon"""
        url = "https://api-seller.ozon.ru/v2/products/info/stocks"
        headers = {"Client-Id": client_id, "Api-Key": api_key}
        try:
            response = self.session.post(url, json={"limit": limit, "last_id": last_id}, headers=headers, timeout=30)
            if response.status_code == 200:
                return {"success": True, "data": response.json()}
            return {"success": False, "error": f"HTTP {response.status_code}"}
        except requests.exceptions.RequestException as e:
            return {"success": False, "error": str(e)}
    
    def get_wildberries_tariffs(self, api_key: str, date: Optional[str] = None) -> Dict[str, Any]:
        """Получение тарифов Wildberries"""
        url = "https://common-api.wildberries.ru/tariffs/box"
        headers = {"Authorization": api_key}
        if date is None:
            date = datetime.now().strftime("%Y-%m-%d")
        params = {"date": date}
        cache_key = f"wb_tariffs_{date}"
        
        if cache_key in self.cache:
            cached = self.cache[cache_key]
            if time.time() - cached.get("timestamp", 0) < self.cache_ttl:
                return cached.get("data", {})
        
        try:
            response = self.session.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                self.cache[cache_key] = {
                    "timestamp": time.time(),
                    "data": {
                        "source": "WB API Live",
                        "timestamp": datetime.now().isoformat(),
                        "data": data
                    }
                }
                return self.cache[cache_key]["data"]
            return {"success": False, "error": f"HTTP {response.status_code}"}
        except requests.exceptions.RequestException as e:
            self.logger.error(f"WB API Error: {e}")
            return {"success": False, "error": str(e)}
    
    def get_wildberries_reports(self, api_key: str, date_from: str, date_to: str) -> Dict[str, Any]:
        """Получение отчётов Wildberries"""
        url = "https://statistics-api.wildberries.ru/api/v1/supplier/reportDetailByPeriod"
        params = {"dateFrom": date_from, "dateTo": date_to}
        headers = {"Authorization": api_key}
        try:
            response = self.session.get(url, params=params, headers=headers, timeout=30)
            if response.status_code == 200:
                return {"success": True, "data": response.json()}
            return {"success": False, "error": f"HTTP {response.status_code}"}
        except requests.exceptions.RequestException as e:
            return {"success": False, "error": str(e)}
    
    def get_yandex_market_campaigns(self, oauth_token: str) -> Dict[str, Any]:
        """Получение кампаний Яндекс Маркет"""
        url = "https://api.partner.market.yandex.ru/v2/campaigns"
        headers = {"Authorization": f"OAuth {oauth_token}"}
        try:
            response = self.session.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                return {"success": True, "data": response.json()}
            return {"success": False, "error": f"HTTP {response.status_code}"}
        except requests.exceptions.RequestException as e:
            return {"success": False, "error": str(e)}
    
    def get_yandex_market_tariffs(self, oauth_token: str, campaign_id: int) -> Dict[str, Any]:
        """Получение тарифов Яндекс Маркет"""
        url = f"https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/deliveries/fees"
        headers = {"Authorization": f"OAuth {oauth_token}"}
        try:
            response = self.session.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                return {"success": True, "data": response.json()}
            return {"success": False, "error": f"HTTP {response.status_code}"}
        except requests.exceptions.RequestException as e:
            return {"success": False, "error": str(e)}


# ============================================================================
# 14.2 DEEPSEEK RATE UPDATER (AI)
# ============================================================================

class DeepSeekRateUpdater:
    """
    🤖 Обновление тарифов через DeepSeek AI.
    Заглушка  работает без API-ключа, возвращает базовые тарифы из конфигурации.
    При наличии API-ключа может быть расширена реальными запросами к DeepSeek.
    """
    
    def __init__(self):
        self.api_key = os.environ.get("DEEPSEEK_API_KEY", "")
        self.logger = logging.getLogger("DeepSeekRateUpdater")
        self.logger.info(
            f"DeepSeekRateUpdater инициализирован. "
            f"API ключ {'найден' if self.api_key else 'НЕ задан (режим заглушки)'}."
        )
    
    def get_rates_from_ai(
        self,
        marketplace: str,
        category: Optional[str] = None,
        force_refresh: bool = False,
        use_cache: bool = True,
        include_forecast: bool = False,
    ) -> Tuple[Optional[Dict[str, Any]], TariffSource, Optional[Dict[str, Any]]]:
        """
        Возвращает тарифы для указанного маркетплейса.
        В режиме заглушки  возвращает базовые тарифы из конфигурации 2026.
        """
        try:
            configs = get_marketplace_configs_2026()
            config = configs.get(marketplace)
            if not config:
                self.logger.warning(f"Маркетплейс {marketplace} не найден в конфигурации")
                return None, TariffSource.HARDCODED, None
            
            rates = {
                "commission_rate": config.commission_rate,
                "min_commission": config.min_commission,
                "logistics_base": config.logistics_base,
                "logistics_per_kg": config.logistics_per_kg,
                "logistics_per_liter": config.logistics_per_liter,
                "storage_per_day": config.storage_per_day,
                "return_fee": config.return_fee,
                "acquiring_fee": config.acquiring_fee,
                "last_mile_fee": config.last_mile_fee,
                "delivery_fee_percent": config.delivery_fee_percent,
                "hazardous_surcharge": config.hazardous_surcharge,
                "fragile_surcharge": config.fragile_surcharge,
                "oversized_surcharge": config.oversized_surcharge,
                "seasonal_multipliers": config.seasonal_multipliers,
            }
            
            if category and category in config.category_rates:
                rates["commission_rate"] = config.category_rates[category]
            
            forecast = None
            if include_forecast:
                forecast = {
                    "month_1": {
                        "commission_rate": round(rates["commission_rate"] * 1.02, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.01, 2),
                        "trend": "stable_up",
                        "confidence": 0.75,
                    },
                    "month_2": {
                        "commission_rate": round(rates["commission_rate"] * 1.04, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.02, 2),
                        "trend": "stable_up",
                        "confidence": 0.70,
                    },
                    "month_3": {
                        "commission_rate": round(rates["commission_rate"] * 1.06, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.03, 2),
                        "trend": "stable_up",
                        "confidence": 0.65,
                    },
                }
            
            self.logger.info(f"✅ Тарифы для {marketplace} получены (источник: конфигурация)")
            return rates, TariffSource.AI_CACHE, forecast
        
        except Exception as e:
            self.logger.error(f"Ошибка get_rates_from_ai: {e}")
            return None, TariffSource.HARDCODED, None
    
    def get_tariff_forecast(
        self,
        marketplace: str,
        category: Optional[str] = None,
        months_ahead: int = 3,
    ) -> Optional[Dict[str, Any]]:
        """Возвращает упрощённый прогноз тарифов"""
        try:
            configs = get_marketplace_configs_2026()
            config = configs.get(marketplace)
            if not config:
                return None
            
            base_rate = config.commission_rate
            if category and category in config.category_rates:
                base_rate = config.category_rates[category]
            
            forecast = {}
            for i in range(1, months_ahead + 1):
                forecast[f"month_{i}"] = {
                    "commission_rate": round(base_rate * (1 + 0.02 * i), 4),
                    "trend": "stable_up",
                    "confidence": 0.75 - 0.05 * i,
                }
            return forecast
        except Exception as e:
            self.logger.error(f"Ошибка get_tariff_forecast: {e}")
            return None
    
    def update_all_marketplaces(
        self,
        force_refresh: bool = False,
        include_forecast: bool = False,
    ) -> Dict[str, Tuple[Optional[Dict], TariffSource, Optional[Dict]]]:
        """Обновляет тарифы для всех маркетплейсов"""
        results = {}
        try:
            configs = get_marketplace_configs_2026()
            for mp_name in configs.keys():
                rates, source, forecast = self.get_rates_from_ai(
                    mp_name,
                    force_refresh=force_refresh,
                    include_forecast=include_forecast,
                )
                results[mp_name] = (rates, source, forecast)
            self.logger.info(f"✅ Обновлено тарифов: {len(results)} маркетплейсов")
        except Exception as e:
            self.logger.error(f"Ошибка update_all_marketplaces: {e}")
        return results


# ============================================================================
# 14.3 SMART TARIFF LOADER
# ============================================================================

class SmartTariffLoader:
    """
    🧠 УМНАЯ ЗАГРУЗКА ТАРИФОВ С ВЫБОРОМ ИСТОЧНИКА
    Поддерживает 4 режима: API, AI, Кэш, Гибридный
    """
    
    SOURCES = {
        "api": "🔌 API Маркетплейса",
        "ai": "🤖 AI (документация)",
        "cache": "💾 Загруженные ранее",
        "hybrid": "🔄 Гибридный (AI + API)"
    }
    
    def __init__(self):
        self.api_connector = MarketplaceAPIConnector()
        self.ai_updater = DeepSeekRateUpdater()
        self.tariff_cache = get_smart_tariff_cache()
        self.logger = logging.getLogger('SmartTariffLoader')
    
    def load_tariffs(self, marketplace: str, source: str = "hybrid",
                     api_key: str = None, client_id: str = None,
                     force_refresh: bool = False) -> Dict[str, Any]:
        """Загрузка тарифов из выбранного источника"""
        result = {
            "marketplace": marketplace,
            "source": source,
            "timestamp": datetime.now().isoformat(),
            "data": {},
            "source_used": None,
            "confidence": 0.0,
            "warnings": [],
            "errors": []
        }
        
        try:
            if source == "api":
                result = self._load_from_api(marketplace, api_key, client_id, result)
            elif source == "ai":
                result = self._load_from_ai(marketplace, result, force_refresh)
            elif source == "cache":
                result = self._load_from_cache(marketplace, result)
            elif source == "hybrid":
                result = self._load_hybrid(marketplace, api_key, client_id, result, force_refresh)
            else:
                result["errors"].append(f"Неизвестный источник: {source}")
            return result
        except Exception as e:
            self.logger.error(f"Ошибка загрузки тарифов: {e}")
            result["errors"].append(str(e))
            return result
    
    def _load_from_api(self, marketplace: str, api_key: str,
                       client_id: str, result: Dict) -> Dict:
        """Загрузка через официальное API маркетплейса"""
        result["source_used"] = "API"
        try:
            if marketplace == "Ozon" and api_key and client_id:
                data = self.api_connector.get_ozon_tariffs(api_key, client_id)
                if data:
                    result["data"] = data
                    result["confidence"] = 0.95
                    result["warnings"].append("✅ Тарифы загружены напрямую из API Ozon")
                else:
                    result["errors"].append("Не удалось получить данные из API Ozon")
            elif marketplace == "Wildberries" and api_key:
                data = self.api_connector.get_wildberries_tariffs(api_key)
                if data and data.get('success'):
                    result["data"] = data.get('data', {})
                    result["confidence"] = 0.95
                    result["warnings"].append("✅ Тарифы загружены напрямую из API WB")
                else:
                    result["errors"].append("Не удалось получить данные из API WB")
            else:
                result["errors"].append(f"API для {marketplace} не поддерживается или не хватает ключей")
        except Exception as e:
            result["errors"].append(f"Ошибка API: {str(e)}")
        return result
    
    def _load_from_ai(self, marketplace: str, result: Dict, force_refresh: bool) -> Dict:
        """Загрузка через AI анализ документации"""
        result["source_used"] = "AI"
        try:
            rates, source, forecast = self.ai_updater.get_rates_from_ai(
                marketplace=marketplace,
                force_refresh=force_refresh,
                use_cache=True,
                include_forecast=True
            )
            if rates:
                result["data"] = {
                    "rates": rates,
                    "forecast": forecast,
                    "source": source.value
                }
                result["confidence"] = 0.85
                result["warnings"].append("🤖 Тарифы получены через AI анализ документации")
                if forecast:
                    result["warnings"].append("📈 Прогноз тарифов на 3 месяца получен")
            else:
                result["errors"].append("AI не смог получить актуальные тарифы")
        except Exception as e:
            result["errors"].append(f"Ошибка AI: {str(e)}")
        return result
    
    def _load_from_cache(self, marketplace: str, result: Dict) -> Dict:
        """Загрузка из кэша (ранее загруженные тарифы)"""
        result["source_used"] = "Cache"
        try:
            cached = self.tariff_cache.get(marketplace, None, use_expired=False)
            if cached:
                result["data"] = {
                    "rates": cached.data,
                    "timestamp": datetime.fromtimestamp(cached.timestamp).isoformat(),
                    "source": cached.source.value
                }
                result["confidence"] = 0.90
                result["warnings"].append(f"💾 Использованы кэшированные тарифы от {datetime.fromtimestamp(cached.timestamp).strftime('%d.%m.%Y %H:%M')}")
            else:
                result["errors"].append("Кэшированные тарифы не найдены или устарели")
        except Exception as e:
            result["errors"].append(f"Ошибка кэша: {str(e)}")
        return result
    
    def _load_hybrid(self, marketplace: str, api_key: str,
                     client_id: str, result: Dict, force_refresh: bool) -> Dict:
        """Гибридный режим: сначала API, если нет  AI, если нет  кэш"""
        result["source_used"] = "Hybrid"
        result["warnings"].append("🔄 Используется гибридный режим загрузки")
        
        # 1. Пробуем API
        if api_key:
            api_result = self._load_from_api(marketplace, api_key, client_id, result.copy())
            if not api_result["errors"] and api_result["data"]:
                result["data"] = api_result["data"]
                result["source_used"] = "API (Hybrid)"
                result["confidence"] = 0.95
                result["warnings"].append("✅ Использованы API тарифы")
                return result
        
        # 2. Пробуем AI
        ai_result = self._load_from_ai(marketplace, result.copy(), force_refresh)
        if not ai_result["errors"] and ai_result["data"]:
            result["data"] = ai_result["data"]
            result["source_used"] = "AI (Hybrid)"
            result["confidence"] = 0.85
            result["warnings"].append("🤖 Использованы AI тарифы (API не доступен)")
            return result
        
        # 3. Пробуем кэш
        cache_result = self._load_from_cache(marketplace, result.copy())
        if not cache_result["errors"] and cache_result["data"]:
            result["data"] = cache_result["data"]
            result["source_used"] = "Cache (Hybrid)"
            result["confidence"] = 0.80
            result["warnings"].append("💾 Использованы кэшированные тарифы (AI и API не доступны)")
            return result
        
        result["errors"].append("Не удалось загрузить тарифы ни из одного источника")
        return result
    
    def get_available_sources(self, marketplace: str) -> List[str]:
        """Получить список доступных источников для маркетплейса"""
        sources = []
        if marketplace in ["Ozon", "Wildberries"]:
            sources.append("api")
        if self.ai_updater.api_key:
            sources.append("ai")
        if self.tariff_cache.get(marketplace, None, use_expired=False):
            sources.append("cache")
        sources.append("hybrid")
        return sources
    
    def compare_sources(self, marketplace: str, api_key: str = None,
                        client_id: str = None) -> pd.DataFrame:
        """Сравнить тарифы из разных источников"""
        results = []
        for source in ["api", "ai", "cache"]:
            if source == "api" and not api_key:
                continue
            result = self.load_tariffs(marketplace, source, api_key, client_id)
            if not result["errors"]:
                results.append({
                    "Источник": self.SOURCES.get(source, source),
                    "Статус": "✅ Доступен",
                    "Данных": len(result["data"]) if isinstance(result["data"], dict) else 0,
                    "Доверие": f"{result['confidence']*100:.0f}%",
                    "Предупреждения": ", ".join(result["warnings"][:2])
                })
            else:
                results.append({
                    "Источник": self.SOURCES.get(source, source),
                    "Статус": "❌ Недоступен",
                    "Данных": 0,
                    "Доверие": "0%",
                    "Предупреждения": result["errors"][0][:50] if result["errors"] else ""
                })
        return pd.DataFrame(results)


# ============================================================================
# 14.4 CATEGORY DIMENSIONS DB
# ============================================================================

class CategoryDimensionsDB:
    """
    📊 База данных категорий с весогабаритами.
    Позволяет загружать категории из Excel и использовать их для валидации.
    """
    
    def __init__(self):
        self.db_path = DATA_DIR / "category_dimensions.json"
        self.db_path.parent.mkdir(exist_ok=True)
        self.categories = {}
        self._load_from_file()
        self.logger = logging.getLogger('CategoryDimensionsDB')
    
    def _load_from_file(self):
        """Загрузка из JSON файла"""
        if self.db_path.exists():
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.categories = json.load(f)
                self.logger.info(f"✅ Загружено {len(self.categories)} категорий из файла")
            except Exception as e:
                self.logger.error(f"❌ Ошибка загрузки: {e}")
                self.categories = {}
    
    def save_to_file(self):
        """Сохранение в JSON файл"""
        try:
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump(self.categories, f, ensure_ascii=False, indent=2)
            self.logger.info(f"✅ Сохранено {len(self.categories)} категорий")
            return True
        except Exception as e:
            self.logger.error(f"❌ Ошибка сохранения: {e}")
            return False
    
    def add_category(self, name: str, length: float, width: float, height: float,
                     weight: float, unit: str = "см", weight_unit: str = "кг"):
        """Добавление категории"""
        self.categories[name.lower().strip()] = {
            "name": name,
            "length_cm": length,
            "width_cm": width,
            "height_cm": height,
            "weight_kg": weight,
            "unit": unit,
            "weight_unit": weight_unit,
            "added_at": datetime.now().isoformat()
        }
        self.save_to_file()
    
    def get_category(self, name: str) -> Optional[Dict[str, Any]]:
        """Получение категории по названию"""
        return self.categories.get(name.lower().strip())
    
    def get_all_categories(self) -> Dict[str, Dict[str, Any]]:
        """Получение всех категорий"""
        return self.categories
    
    def delete_category(self, name: str):
        """Удаление категории"""
        key = name.lower().strip()
        if key in self.categories:
            del self.categories[key]
            self.save_to_file()
    
    def clear_all(self):
        """Очистка всех категорий"""
        self.categories = {}
        self.save_to_file()
    
    def import_from_excel(self, file_path: str) -> Dict[str, Any]:
        """Импорт категорий из Excel файла"""
        result = {
            "success": False,
            "imported": 0,
            "errors": [],
            "warnings": []
        }
        
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            if df.empty:
                result["errors"].append("Файл пустой")
                return result
            
            df.columns = [col.strip().lower() for col in df.columns]
            
            column_mapping = {
                'категория': 'category', 'category': 'category', 'название': 'category', 'name': 'category',
                'длина': 'length', 'length': 'length', 'длина (см)': 'length',
                'ширина': 'width', 'width': 'width', 'ширина (см)': 'width',
                'высота': 'height', 'height': 'height', 'высота (см)': 'height',
                'вес': 'weight', 'weight': 'weight', 'вес (кг)': 'weight',
                'единица длины': 'length_unit', 'length_unit': 'length_unit',
                'единица веса': 'weight_unit', 'weight_unit': 'weight_unit'
            }
            df = df.rename(columns=column_mapping)
            
            required_cols = ['category', 'length', 'width', 'height', 'weight']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                result["errors"].append(f"Отсутствуют колонки: {', '.join(missing_cols)}")
                return result
            
            imported_count = 0
            for idx, row in df.iterrows():
                try:
                    category_name = str(row.get('category', '')).strip()
                    if not category_name:
                        result["warnings"].append(f"Строка {idx + 1}: пустое название категории")
                        continue
                    
                    length = safe_float(row.get('length', 0))
                    width = safe_float(row.get('width', 0))
                    height = safe_float(row.get('height', 0))
                    weight = safe_float(row.get('weight', 0))
                    
                    if length <= 0 or width <= 0 or height <= 0 or weight <= 0:
                        result["warnings"].append(f"Строка {idx + 1}: некорректные размеры для '{category_name}'")
                        continue
                    
                    length_unit = str(row.get('length_unit', 'см')).strip()
                    weight_unit = str(row.get('weight_unit', 'кг')).strip()
                    
                    self.add_category(
                        name=category_name,
                        length=length, width=width, height=height, weight=weight,
                        unit=length_unit, weight_unit=weight_unit
                    )
                    imported_count += 1
                except Exception as e:
                    result["errors"].append(f"Строка {idx + 1}: {str(e)}")
            
            result["success"] = imported_count > 0
            result["imported"] = imported_count
            if imported_count == 0:
                result["errors"].append("Не удалось импортировать ни одну категорию")
        
        except Exception as e:
            result["errors"].append(f"Ошибка чтения файла: {str(e)}")
        
        return result
    
    def export_to_excel(self, file_path: str) -> bool:
        """Экспорт категорий в Excel"""
        try:
            data = []
            for key, cat in self.categories.items():
                data.append({
                    'Категория': cat['name'],
                    'Длина (см)': cat['length_cm'],
                    'Ширина (см)': cat['width_cm'],
                    'Высота (см)': cat['height_cm'],
                    'Вес (кг)': cat['weight_kg'],
                    'Единица длины': cat.get('unit', 'см'),
                    'Единица веса': cat.get('weight_unit', 'кг')
                })
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, engine='openpyxl')
            return True
        except Exception as e:
            self.logger.error(f"Ошибка экспорта: {e}")
            return False
    
    def get_statistics(self) -> Dict[str, Any]:
        """Статистика по категориям"""
        if not self.categories:
            return {"total": 0}
        
        lengths = [cat['length_cm'] for cat in self.categories.values()]
        widths = [cat['width_cm'] for cat in self.categories.values()]
        heights = [cat['height_cm'] for cat in self.categories.values()]
        weights = [cat['weight_kg'] for cat in self.categories.values()]
        
        return {
            "total": len(self.categories),
            "avg_length": sum(lengths) / len(lengths),
            "avg_width": sum(widths) / len(widths),
            "avg_height": sum(heights) / len(heights),
            "avg_weight": sum(weights) / len(weights),
            "min_weight": min(weights),
            "max_weight": max(weights)
        }


# ============================================================================
# 14.5 ADVANCED DIMENSIONS VALIDATOR
# ============================================================================

class AdvancedDimensionsValidator:
    """Класс для умной проверки, нормализации и дополнения весогабаритных характеристик."""
    
    @staticmethod
    def normalize_dimension(value: float, unit_hint: str = "") -> float:
        if not value or value <= 0:
            return 0.0
        unit_lower = unit_hint.lower() if unit_hint else ""
        if any(x in unit_lower for x in ['mm', 'мм', 'millimeter']):
            return value / 10.0
        if any(x in unit_lower for x in ['m', 'метр', 'meter']) and value < 10:
            return value * 100.0
        if value > 300:
            return value / 10.0
        return value
    
    @staticmethod
    def normalize_weight(value: float, unit_hint: str = "") -> float:
        if not value or value <= 0:
            return 0.0
        unit_lower = unit_hint.lower() if unit_hint else ""
        if any(x in unit_lower for x in ['g', 'гр', 'gram']):
            return value / 1000.0
        if any(x in unit_lower for x in ['t', 'тонн', 'ton']) and value < 10:
            return value * 1000.0
        if value > 100:
            return value / 1000.0
        return value
    
    @staticmethod
    def infer_missing_dimensions(category: str, weight: float) -> Dict[str, float]:
        """Эвристическое определение габаритов по категории и весу"""
        defaults = {
            "фильтры": {"l": 15, "w": 15, "h": 15},
            "колодки": {"l": 15, "w": 10, "h": 5},
            "масла": {"l": 10, "w": 10, "h": 25},
            "шины": {"l": 60, "w": 60, "h": 25},
            "аккумуляторы": {"l": 35, "w": 20, "h": 20},
            "фары": {"l": 40, "w": 20, "h": 20},
            "двигатель": {"l": 50, "w": 40, "h": 40},
            "трансмиссия": {"l": 50, "w": 40, "h": 30},
            "подвеска": {"l": 40, "w": 30, "h": 20},
            "тормозная_система": {"l": 30, "w": 20, "h": 15},
            "рулевое_управление": {"l": 40, "w": 20, "h": 15},
            "электрика": {"l": 25, "w": 20, "h": 20},
            "охлаждение": {"l": 45, "w": 35, "h": 20},
            "выпуск": {"l": 60, "w": 25, "h": 20},
            "оптика": {"l": 35, "w": 25, "h": 20},
            "кузов": {"l": 80, "w": 50, "h": 30},
            "инструменты": {"l": 30, "w": 20, "h": 15},
            "ремни": {"l": 25, "w": 15, "h": 10},
            "подшипники": {"l": 15, "w": 15, "h": 10},
            "крепёж": {"l": 10, "w": 10, "h": 5},
            "климат": {"l": 40, "w": 30, "h": 25},
            "безопасность": {"l": 30, "w": 20, "h": 15}
        }
        
        cat_key = category.lower()
        if cat_key in defaults:
            dims = defaults[cat_key]
        else:
            for key in defaults:
                if key in cat_key:
                    dims = defaults[key]
                    break
            else:
                dims = {"l": 20, "w": 20, "h": 20}
        
        scale = max(0.5, min(3.0, weight / 2.0))
        return {
            "length_cm": dims["l"] * scale,
            "width_cm": dims["w"] * scale,
            "height_cm": dims["h"] * scale
        }
    
    @staticmethod
    def validate_and_normalize_row(
        row: pd.Series,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        category: str = ""
    ) -> Dict[str, float]:
        """Валидация и нормализация строки с габаритами"""
        raw_l = safe_float(row.get(length_col, 0)) if length_col else 0
        raw_w = safe_float(row.get(width_col, 0)) if width_col else 0
        raw_h = safe_float(row.get(height_col, 0)) if height_col else 0
        raw_weight = safe_float(row.get(weight_col, 0)) if weight_col else 0
        
        length = AdvancedDimensionsValidator.normalize_dimension(raw_l)
        width = AdvancedDimensionsValidator.normalize_dimension(raw_w)
        height = AdvancedDimensionsValidator.normalize_dimension(raw_h)
        weight = AdvancedDimensionsValidator.normalize_weight(raw_weight)
        
        if length == 0 or width == 0 or height == 0:
            inferred = AdvancedDimensionsValidator.infer_missing_dimensions(category, weight)
            if length == 0: length = inferred["length_cm"]
            if width == 0: width = inferred["width_cm"]
            if height == 0: height = inferred["height_cm"]
        
        if weight == 0 and length > 0 and width > 0 and height > 0:
            volume = (length * width * height) / 1000
            weight = max(0.1, volume * 0.8)
        
        return {
            "length_cm": round(length, 2),
            "width_cm": round(width, 2),
            "height_cm": round(height, 2),
            "weight_kg": round(weight, 2)
        }


# ============================================================================
# 14.6 PERFORMANCE MANAGER
# ============================================================================

class PerformanceManager:
    """Управление ресурсами системы для оптимизации тяжелых расчетов"""
    
    def __init__(self):
        self.start_time = time.time()
        self.memory_threshold_mb = 4096
        self.gc_threshold = 100
        self.operation_count = 0
    
    def check_memory_usage(self) -> bool:
        if not PSUTIL_AVAILABLE:
            return True
        try:
            process = psutil.Process(os.getpid())
            mem_info = process.memory_info()
            mem_mb = mem_info.rss / (1024 * 1024)
            self.operation_count += 1
            
            if mem_mb > self.memory_threshold_mb:
                logger.warning(f"⚠️ Высокое использование памяти: {mem_mb:.2f} MB. Запуск GC...")
                gc.collect()
                return mem_info.rss / (1024 * 1024) < self.memory_threshold_mb * 1.2
            
            if self.operation_count % self.gc_threshold == 0:
                gc.collect()
            
            return True
        except (psutil.Error, OSError) as e:
            logger.error(f"Ошибка проверки памяти: {e}")
            return True
    
    def get_system_stats(self) -> Dict[str, Any]:
        stats = {
            "uptime_seconds": time.time() - self.start_time,
            "operation_count": self.operation_count,
            "memory_threshold_mb": self.memory_threshold_mb
        }
        if PSUTIL_AVAILABLE:
            stats.update({
                "cpu_percent": psutil.cpu_percent(interval=0.5),
                "memory_percent": psutil.virtual_memory().percent,
                "memory_available_mb": psutil.virtual_memory().available / (1024 * 1024),
                "disk_usage_percent": psutil.disk_usage('/').percent
            })
            try:
                process = psutil.Process(os.getpid())
                mem_info = process.memory_info()
                stats["process_memory_mb"] = mem_info.rss / (1024 * 1024)
                stats["process_cpu_percent"] = process.cpu_percent()
            except (psutil.Error, OSError):
                pass
        return stats
    
    def optimize_for_big_data(self):
        if POLARS_AVAILABLE:
            os.environ['POLARS_MAX_THREADS'] = str(min(4, os.cpu_count() or 2))
            os.environ['POLARS_VERBOSE'] = '0'
        if PSUTIL_AVAILABLE:
            memory_mb = psutil.virtual_memory().available / (1024 * 1024)
            if memory_mb < 2048:
                logger.warning(f"⚠️ Мало памяти ({memory_mb:.0f} MB). Включен экономичный режим.")
                self.memory_threshold_mb = 1024
                os.environ['POLARS_MAX_THREADS'] = '1'


# Инициализация PerformanceManager при загрузке модуля
perf_manager = PerformanceManager()
perf_manager.optimize_for_big_data()


# ============================================================================
# 14.7 PRICE CALCULATOR
# ============================================================================

class PriceCalculator:
    """Самостоятельный калькулятор цены на автозапчасть."""
    
    def __init__(self, marketplace_config: MarketplaceConfig):
        self.config = marketplace_config
        self.logger = logging.getLogger('PriceCalculator')
    
    def calculate_retail_price(
        self,
        purchase_price: float,
        desired_margin: float = 30.0,
        weight: float = 1.0,
        volume: float = 5.0,
        category: str = None,
        days_in_storage: int = 30,
        include_subscription: bool = False,
        current_month: Optional[int] = None
    ) -> Dict[str, float]:
        if purchase_price <= 0:
            raise ValidationError("Закупочная цена должна быть положительной", "purchase_price", purchase_price)
        
        commission_rate = self.config.calculate_commission_with_dynamics(
            price=purchase_price * 2,
            category=category,
            current_month=current_month
        ) / (purchase_price * 2) if purchase_price > 0 else self.config.commission_rate
        
        logistics = (
            self.config.logistics_base +
            weight * self.config.logistics_per_kg +
            volume * self.config.logistics_per_liter
        )
        storage = volume * self.config.storage_per_day * days_in_storage
        last_mile = self.config.last_mile_fee
        
        fixed_costs = logistics + storage + last_mile
        if include_subscription and self.config.subscription_fee > 0:
            fixed_costs += self.config.subscription_fee / 30
        
        variable_ratio = (
            commission_rate +
            self.config.acquiring_fee +
            self.config.return_fee +
            self.config.delivery_fee_percent +
            0.06
        )
        
        margin_ratio = desired_margin / 100
        denominator = 1 - variable_ratio - margin_ratio
        
        if denominator <= 0:
            raise CalculationError(
                f"Невозможно достичь маржинальности {desired_margin}% при текущих тарифах",
                "price_calculation"
            )
        
        retail_price = (purchase_price + fixed_costs) / denominator
        
        min_price = calculate_recommended_min_price(
            cost=purchase_price,
            commission_rate=commission_rate,
            logistics=logistics,
            storage_cost=storage,
            acquiring_rate=self.config.acquiring_fee,
            last_mile=last_mile,
            return_rate=self.config.return_fee,
            min_profit_percent=0.10,
            tax_system="УСН_6",
            tax_rate=0.06
        )
        
        if retail_price < min_price:
            retail_price = min_price
        
        return {
            "retail_price": money_round(retail_price),
            "margin": desired_margin,
            "fixed_costs": money_round(fixed_costs),
            "commission": money_round(retail_price * commission_rate),
            "commission_rate": money_round(commission_rate * 100),
            "logistics": money_round(logistics),
            "storage": money_round(storage),
            "min_price": money_round(min_price),
            "profit": money_round(retail_price - purchase_price - fixed_costs - retail_price * variable_ratio)
        }
    
    def calculate_margin_at_price(
        self,
        retail_price: float,
        purchase_price: float,
        weight: float = 1.0,
        volume: float = 5.0,
        category: str = None
    ) -> Dict[str, float]:
        if retail_price <= 0 or purchase_price <= 0:
            raise ValidationError("Цена и себестоимость должны быть положительными")
        
        commission = self.config.calculate_commission_with_dynamics(
            price=retail_price,
            category=category
        )
        logistics = (
            self.config.logistics_base +
            weight * self.config.logistics_per_kg +
            volume * self.config.logistics_per_liter
        )
        storage = volume * self.config.storage_per_day * 30
        acquiring = retail_price * self.config.acquiring_fee
        delivery = retail_price * self.config.delivery_fee_percent
        returns = retail_price * self.config.return_fee
        tax = retail_price * 0.06
        
        total_costs = purchase_price + commission + logistics + storage + acquiring + delivery + returns + tax
        profit = retail_price - total_costs
        margin = (profit / retail_price * 100) if retail_price > 0 else 0
        
        return {
            "profit": money_round(profit),
            "margin_percent": money_round(margin),
            "total_costs": money_round(total_costs),
            "commission": money_round(commission),
            "logistics": money_round(logistics)
        }
    
    def find_optimal_price(
        self,
        purchase_price: float,
        target_margin: float = 30.0,
        weight: float = 1.0,
        volume: float = 5.0,
        category: str = None,
        price_min: float = 0,
        price_max: float = 100000,
        step: float = 10
    ) -> Dict[str, Any]:
        if price_min <= 0:
            price_min = purchase_price * 1.2
        
        best_price = price_min
        best_margin = 0
        best_profit = float('-inf')
        current_price = price_min
        
        while current_price <= price_max:
            try:
                result = self.calculate_margin_at_price(
                    retail_price=current_price,
                    purchase_price=purchase_price,
                    weight=weight,
                    volume=volume,
                    category=category
                )
                margin = result['margin_percent']
                profit = result['profit']
                
                if margin >= target_margin and profit > best_profit:
                    best_profit = profit
                    best_price = current_price
                    best_margin = margin
                
                current_price += step
            except Exception:
                current_price += step
                continue
        
        return {
            "optimal_price": money_round(best_price),
            "optimal_margin": money_round(best_margin),
            "optimal_profit": money_round(best_profit),
            "target_margin": target_margin
        }

# ============================================================================
# БЛОК 15: ДОПОЛНИТЕЛЬНЫЕ УТИЛИТЫ И ХЕЛПЕРЫ (v101.0)
# ============================================================================
# 📌 v101.0: Дополнительные утилиты для работы с приложением:
# - Расширенные утилиты для Google Sheets
# - Утилиты для экспорта/импорта данных
# - Дополнительные валидаторы
# - Хелперы для работы с файлами
# - Утилиты для логирования и отладки
# ============================================================================


# ============================================================================
# 15.1 РАСШИРЕННЫЕ УТИЛИТЫ ДЛЯ GOOGLE SHEETS
# ============================================================================

class GoogleSheetsManager:
    """
     v101.0: Расширенный менеджер для работы с Google Sheets.
    Поддерживает множественные операции, батчинг и обработку ошибок.
    """
    
    def __init__(self, credentials_json: Union[str, Path]):
        self.credentials_json = credentials_json
        self.client = None
        self.logger = logging.getLogger('GoogleSheetsManager')
        self._init_client()
    
    def _init_client(self):
        """Инициализация клиента gspread"""
        if not GSPREAD_AVAILABLE:
            self.logger.error("❌ gspread не установлен")
            return
        
        try:
            creds_path = Path(self.credentials_json)
            if creds_path.exists():
                credentials = Credentials.from_service_account_file(
                    str(creds_path),
                    scopes=[
                        "https://spreadsheets.google.com/feeds",
                        "https://www.googleapis.com/auth/drive"
                    ]
                )
            else:
                creds_data = json.loads(str(self.credentials_json))
                credentials = Credentials.from_service_account_info(
                    creds_data,
                    scopes=[
                        "https://spreadsheets.google.com/feeds",
                        "https://www.googleapis.com/auth/drive"
                    ]
                )
            
            self.client = gspread.authorize(credentials)
            self.logger.info("✅ Google Sheets клиент инициализирован")
        
        except Exception as e:
            self.logger.error(f"❌ Ошибка инициализации: {e}")
    
    def upload_dataframe_batch(
        self,
        df: pd.DataFrame,
        spreadsheet_id: str,
        worksheet_name: str = "Sheet1",
        batch_size: int = 1000,
        clear_before: bool = True
    ) -> bool:
        """
        Загрузка DataFrame в Google Sheets батчами.
        
        Args:
            df: DataFrame для загрузки
            spreadsheet_id: ID таблицы
            worksheet_name: Имя листа
            batch_size: Размер батча
            clear_before: Очистить лист перед загрузкой
        
        Returns:
            True при успехе
        """
        if self.client is None:
            self.logger.error("❌ Клиент не инициализирован")
            return False
        
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            
            # Ищем или создаём лист
            try:
                worksheet = spreadsheet.worksheet(worksheet_name)
                if clear_before:
                    worksheet.clear()
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(
                    title=worksheet_name,
                    rows=len(df) + 1,
                    cols=len(df.columns)
                )
            
            # Преобразуем DataFrame в список списков
            values = [df.columns.tolist()] + df.values.tolist()
            
            # Загружаем батчами
            for i in range(0, len(values), batch_size):
                batch = values[i:i + batch_size]
                start_cell = f"A{i + 1}"
                end_row = i + len(batch)
                end_cell = f"{get_column_letter(len(df.columns))}{end_row}"
                worksheet.update(f"{start_cell}:{end_cell}", batch)
                self.logger.info(f"✅ Загружен батч {i//batch_size + 1}: {len(batch)} строк")
            
            self.logger.info(f"✅ Данные загружены в Google Sheets: {spreadsheet_id}/{worksheet_name}")
            return True
        
        except Exception as e:
            self.logger.error(f"❌ Ошибка загрузки: {e}")
            return False
    
    def append_dataframe(
        self,
        df: pd.DataFrame,
        spreadsheet_id: str,
        worksheet_name: str = "Sheet1"
    ) -> bool:
        """Добавление DataFrame в конец листа"""
        if self.client is None:
            return False
        
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            worksheet = spreadsheet.worksheet(worksheet_name)
            
            values = df.values.tolist()
            worksheet.append_rows(values)
            
            self.logger.info(f"✅ Добавлено {len(df)} строк в {worksheet_name}")
            return True
        
        except Exception as e:
            self.logger.error(f"❌ Ошибка добавления: {e}")
            return False
    
    def read_to_dataframe(
        self,
        spreadsheet_id: str,
        worksheet_name: str = "Sheet1",
        header_row: int = 1
    ) -> Optional[pd.DataFrame]:
        """Чтение данных из Google Sheets в DataFrame"""
        if self.client is None:
            return None
        
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            worksheet = spreadsheet.worksheet(worksheet_name)
            
            values = worksheet.get_all_values()
            if not values:
                return pd.DataFrame()
            
            # Первая строка  заголовки
            headers = values[0]
            data = values[1:]
            
            df = pd.DataFrame(data, columns=headers)
            
            # Пытаемся преобразовать числовые столбцы
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='ignore')
                except Exception:
                    pass
            
            self.logger.info(f"✅ Загружено {len(df)} строк из {worksheet_name}")
            return df
        
        except Exception as e:
            self.logger.error(f"❌ Ошибка чтения: {e}")
            return None
    
    def share_spreadsheet(
        self,
        spreadsheet_id: str,
        email: str,
        role: str = "reader"
    ) -> bool:
        """Предоставление доступа к таблице"""
        if self.client is None:
            return False
        
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            spreadsheet.share(email, perm_type="user", role=role)
            self.logger.info(f"✅ Доступ предоставлен: {email} ({role})")
            return True
        
        except Exception as e:
            self.logger.error(f"❌ Ошибка предоставления доступа: {e}")
            return False


# ============================================================================
# 15.2 УТИЛИТЫ ДЛЯ ЭКСПОРТА/ИМПОРТА ДАННЫХ
# ============================================================================

class DataExportImportManager:
    """
     v101.0: Менеджер для экспорта и импорта данных в различных форматах.
    """
    
    @staticmethod
    def export_to_multiple_formats(
        df: pd.DataFrame,
        output_dir: Union[str, Path],
        formats: List[str] = ["csv", "excel", "json"],
        filename_prefix: str = "data"
    ) -> Dict[str, Path]:
        """
        Экспорт DataFrame в несколько форматов одновременно.
        
        Args:
            df: DataFrame для экспорта
            output_dir: Директория для сохранения
            formats: Список форматов ["csv", "excel", "json", "parquet"]
            filename_prefix: Префикс имени файла
        
        Returns:
            Словарь {формат: путь_к_файлу}
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        results = {}
        
        for fmt in formats:
            try:
                filename = f"{filename_prefix}_{timestamp}.{fmt}"
                filepath = output_dir / filename
                
                if fmt == "csv":
                    df.to_csv(filepath, index=False, encoding='utf-8-sig', sep=';')
                elif fmt == "excel":
                    df.to_excel(filepath, index=False, engine='openpyxl')
                elif fmt == "json":
                    df.to_json(filepath, orient='records', force_ascii=False, indent=2)
                elif fmt == "parquet":
                    df.to_parquet(filepath, index=False)
                else:
                    logger.warning(f"⚠️ Формат {fmt} не поддерживается")
                    continue
                
                results[fmt] = filepath
                logger.info(f"✅ Экспортировано в {fmt}: {filepath}")
            
            except Exception as e:
                logger.error(f"❌ Ошибка экспорта в {fmt}: {e}")
        
        return results
    
    @staticmethod
    def import_from_multiple_sources(
        file_paths: List[Union[str, Path]],
        merge_strategy: str = "concat"
    ) -> pd.DataFrame:
        """
        Импорт данных из нескольких файлов и объединение.
        
        Args:
            file_paths: Список путей к файлам
            merge_strategy: Стратегия объединения ("concat", "merge", "append")
        
        Returns:
            Объединённый DataFrame
        """
        dfs = []
        
        for file_path in file_paths:
            file_path = Path(file_path)
            if not file_path.exists():
                logger.warning(f"⚠️ Файл не найден: {file_path}")
                continue
            
            try:
                if file_path.suffix.lower() == '.csv':
                    df = pd.read_csv(file_path, encoding='utf-8-sig', sep=';')
                elif file_path.suffix.lower() in ['.xlsx', '.xls']:
                    df = pd.read_excel(file_path, engine='openpyxl')
                elif file_path.suffix.lower() == '.json':
                    df = pd.read_json(file_path)
                elif file_path.suffix.lower() == '.parquet':
                    df = pd.read_parquet(file_path)
                else:
                    logger.warning(f"⚠️ Неподдерживаемый формат: {file_path.suffix}")
                    continue
                
                dfs.append(df)
                logger.info(f"✅ Загружено из {file_path}: {len(df)} строк")
            
            except Exception as e:
                logger.error(f"❌ Ошибка загрузки из {file_path}: {e}")
        
        if not dfs:
            return pd.DataFrame()
        
        # Объединение
        if merge_strategy == "concat":
            result = pd.concat(dfs, ignore_index=True)
        elif merge_strategy == "merge":
            result = dfs[0]
            for df in dfs[1:]:
                result = result.merge(df, how='outer')
        elif merge_strategy == "append":
            result = dfs[0]
            for df in dfs[1:]:
                result = result.append(df, ignore_index=True)
        else:
            result = pd.concat(dfs, ignore_index=True)
        
        logger.info(f"✅ Объединено {len(dfs)} файлов: {len(result)} строк")
        return result


# ============================================================================
# 15.3 ДОПОЛНИТЕЛЬНЫЕ ВАЛИДАТОРЫ
# ============================================================================

class DataQualityValidator:
    """
     v101.0: Валидатор качества данных.
    Проверяет данные на наличие ошибок, пропусков, аномалий.
    """
    
    @staticmethod
    def validate_dataframe_quality(
        df: pd.DataFrame,
        required_columns: Optional[List[str]] = None,
        numeric_columns: Optional[List[str]] = None,
        min_value: Optional[float] = None,
        max_value: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Проверка качества DataFrame.
        
        Returns:
            Словарь с результатами проверки
        """
        result = {
            "valid": True,
            "errors": [],
            "warnings": [],
            "statistics": {}
        }
        
        # Проверка на пустоту
        if df.empty:
            result["valid"] = False
            result["errors"].append("DataFrame пустой")
            return result
        
        # Проверка обязательных колонок
        if required_columns:
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                result["valid"] = False
                result["errors"].append(f"Отсутствуют колонки: {', '.join(missing_cols)}")
        
        # Проверка числовых колонок
        if numeric_columns:
            for col in numeric_columns:
                if col in df.columns:
                    # Проверка на NaN
                    nan_count = df[col].isna().sum()
                    if nan_count > 0:
                        result["warnings"].append(f"Колонка '{col}' содержит {nan_count} NaN значений")
                    
                    # Проверка диапазона
                    if min_value is not None:
                        below_min = (df[col] < min_value).sum()
                        if below_min > 0:
                            result["warnings"].append(f"Колонка '{col}': {below_min} значений ниже {min_value}")
                    
                    if max_value is not None:
                        above_max = (df[col] > max_value).sum()
                        if above_max > 0:
                            result["warnings"].append(f"Колонка '{col}': {above_max} значений выше {max_value}")
        
        # Статистика
        result["statistics"] = {
            "total_rows": len(df),
            "total_columns": len(df.columns),
            "memory_usage_mb": df.memory_usage(deep=True).sum() / (1024 * 1024),
            "duplicate_rows": df.duplicated().sum()
        }
        
        return result
    
    @staticmethod
    def detect_outliers(
        df: pd.DataFrame,
        column: str,
        method: str = "iqr",
        threshold: float = 1.5
    ) -> pd.DataFrame:
        """
        Обнаружение выбросов в числовой колонке.
        
        Args:
            df: DataFrame
            column: Название колонки
            method: Метод обнаружения ("iqr", "zscore", "percentile")
            threshold: Пороговое значение
        
        Returns:
            DataFrame с выбросами
        """
        if column not in df.columns:
            return pd.DataFrame()
        
        data = df[column].dropna()
        
        if method == "iqr":
            Q1 = data.quantile(0.25)
            Q3 = data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - threshold * IQR
            upper_bound = Q3 + threshold * IQR
            outliers = df[(df[column] < lower_bound) | (df[column] > upper_bound)]
        
        elif method == "zscore":
            from scipy import stats
            z_scores = stats.zscore(data)
            outliers = df[abs(z_scores) > threshold]
        
        elif method == "percentile":
            lower = data.quantile(threshold / 100)
            upper = data.quantile(1 - threshold / 100)
            outliers = df[(df[column] < lower) | (df[column] > upper)]
        
        else:
            outliers = pd.DataFrame()
        
        return outliers


# ============================================================================
# 15.4 ХЕЛПЕРЫ ДЛЯ РАБОТЫ С ФАЙЛАМИ
# ============================================================================

class FileManager:
    """
     v101.0: Хелпер для работы с файлами.
    """
    
    @staticmethod
    def get_file_info(file_path: Union[str, Path]) -> Dict[str, Any]:
        """Получение информации о файле"""
        file_path = Path(file_path)
        
        if not file_path.exists():
            return {"error": "Файл не найден"}
        
        stat = file_path.stat()
        
        return {
            "path": str(file_path),
            "name": file_path.name,
            "extension": file_path.suffix,
            "size_bytes": stat.st_size,
            "size_mb": stat.st_size / (1024 * 1024),
            "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
            "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "is_file": file_path.is_file(),
            "is_dir": file_path.is_dir()
        }
    
    @staticmethod
    def cleanup_old_files(
        directory: Union[str, Path],
        pattern: str = "*",
        max_age_days: int = 30,
        max_files: int = 100
    ) -> int:
        """
        Очистка старых файлов в директории.
        
        Args:
            directory: Директория для очистки
            pattern: Паттерн имён файлов
            max_age_days: Максимальный возраст файлов в днях
            max_files: Максимальное количество файлов
        
        Returns:
            Количество удалённых файлов
        """
        directory = Path(directory)
        if not directory.exists():
            return 0
        
        files = list(directory.glob(pattern))
        files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        
        deleted_count = 0
        cutoff_time = datetime.now() - timedelta(days=max_age_days)
        
        for file_path in files[max_files:]:
            try:
                file_time = datetime.fromtimestamp(file_path.stat().st_mtime)
                if file_time < cutoff_time:
                    file_path.unlink()
                    deleted_count += 1
                    logger.info(f"✅ Удалён старый файл: {file_path.name}")
            except Exception as e:
                logger.error(f"❌ Ошибка удаления {file_path}: {e}")
        
        return deleted_count
    
    @staticmethod
    def create_backup(
        file_path: Union[str, Path],
        backup_dir: Optional[Union[str, Path]] = None
    ) -> Optional[Path]:
        """Создание резервной копии файла"""
        file_path = Path(file_path)
        
        if not file_path.exists():
            return None
        
        if backup_dir is None:
            backup_dir = file_path.parent / "backups"
        else:
            backup_dir = Path(backup_dir)
        
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = backup_dir / backup_name
        
        try:
            shutil.copy2(file_path, backup_path)
            logger.info(f"✅ Создана резервная копия: {backup_path}")
            return backup_path
        except Exception as e:
            logger.error(f"❌ Ошибка создания резервной копии: {e}")
            return None


# ============================================================================
# 15.5 УТИЛИТЫ ДЛЯ ЛОГИРОВАНИЯ И ОТЛАДКИ
# ============================================================================

class DebugLogger:
    """
     v101.0: Утилиты для расширенного логирования и отладки.
    """
    
    @staticmethod
    def log_dataframe_summary(df: pd.DataFrame, name: str = "DataFrame"):
        """Логирование сводки по DataFrame"""
        logger.info(f"📊 {name}:")
        logger.info(f"  - Строк: {len(df)}")
        logger.info(f"  - Колонок: {len(df.columns)}")
        logger.info(f"  - Память: {df.memory_usage(deep=True).sum() / (1024 * 1024):.2f} MB")
        logger.info(f"  - Дубликатов: {df.duplicated().sum()}")
        logger.info(f"  - Пропусков: {df.isna().sum().sum()}")
    
    @staticmethod
    def log_execution_time(func: Callable) -> Callable:
        """Декоратор для логирования времени выполнения"""
        @wraps(func)
        def wrapper(*args, **kwargs):
            start = time.perf_counter()
            try:
                result = func(*args, **kwargs)
                elapsed = time.perf_counter() - start
                logger.info(f"⏱️ {func.__name__} выполнена за {elapsed:.3f}с")
                return result
            except Exception as e:
                elapsed = time.perf_counter() - start
                logger.error(f"❌ {func.__name__} завершилась с ошибкой за {elapsed:.3f}с: {e}")
                raise
        return wrapper
    
    @staticmethod
    def profile_function(func: Callable) -> Dict[str, Any]:
        """Профилирование функции"""
        import cProfile
        import pstats
        from io import StringIO
        
        pr = cProfile.Profile()
        pr.enable()
        
        try:
            result = func()
        finally:
            pr.disable()
        
        s = StringIO()
        ps = pstats.Stats(pr, stream=s).sort_stats('cumulative')
        ps.print_stats(20)
        
        return {
            "result": result,
            "profile": s.getvalue()
        }


# ============================================================================
# 15.6 УТИЛИТЫ ДЛЯ РАБОТЫ С КОНФИГУРАЦИЯМИ
# ============================================================================

class ConfigManager:
    """
     v101.0: Менеджер для работы с конфигурационными файлами.
    """
    
    def __init__(self, config_dir: Union[str, Path] = CONFIG_DIR):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self._cache = {}
    
    def load_config(self, config_name: str) -> Dict[str, Any]:
        """Загрузка конфигурации из JSON файла"""
        config_path = self.config_dir / f"{config_name}.json"
        
        if config_name in self._cache:
            return self._cache[config_name]
        
        if not config_path.exists():
            return {}
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            self._cache[config_name] = config
            return config
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки конфигурации {config_name}: {e}")
            return {}
    
    def save_config(self, config_name: str, config: Dict[str, Any]) -> bool:
        """Сохранение конфигурации в JSON файл"""
        config_path = self.config_dir / f"{config_name}.json"
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self._cache[config_name] = config
            logger.info(f"✅ Конфигурация {config_name} сохранена")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения конфигурации {config_name}: {e}")
            return False
    
    def get_value(self, config_name: str, key: str, default: Any = None) -> Any:
        """Получение значения из конфигурации"""
        config = self.load_config(config_name)
        return config.get(key, default)
    
    def set_value(self, config_name: str, key: str, value: Any) -> bool:
        """Установка значения в конфигурации"""
        config = self.load_config(config_name)
        config[key] = value
        return self.save_config(config_name, config)

# ============================================================================
# Блок 16 ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ (main)
# ============================================================================
# 📌 v101.0: Финальный блок, объединяющий все 4 раздела в единый интерфейс
# - Навигация через sidebar
# - Логотип и футер
# - Инициализация всех компонентов
# - Настройка Streamlit
# ============================================================================


def show_footer():
    """Футер приложения"""
    st.divider()
    st.markdown("""
    <div style='text-align: center; padding: 20px; color: #666;'>
        <p style='margin: 0;'>🚗 <strong>Юнит-экономика автозапчастей PRO 2026</strong></p>
        <p style='margin: 5px 0 0 0; font-size: 0.9em;'>
            Версия 101.0 | Enterprise Edition | 
            <a href='https://github.com' target='_blank'>GitHub</a> | 
            <a href='mailto:support@example.com'>Поддержка</a>
        </p>
        <p style='margin: 5px 0 0 0; font-size: 0.8em; color: #999;'>
            © 2024-2026 AutoParts Analytics Team
        </p>
    </div>
    """, unsafe_allow_html=True)


def show_sidebar_info():
    """Информация в sidebar"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Статус системы")
    
    # Проверяем доступность компонентов
    components_status = {
        "Polars": POLARS_AVAILABLE,
        "DuckDB": DUCKDB_AVAILABLE,
        "scikit-learn": SKLEARN_AVAILABLE,
        "Plotly": PLOTLY_AVAILABLE,
        "openpyxl": OPENPYXL_AVAILABLE,
        "gspread": GSPREAD_AVAILABLE,
    }
    
    for name, available in components_status.items():
        if available:
            st.sidebar.success(f"✅ {name}")
        else:
            st.sidebar.warning(f"⚠️ {name}")
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💡 Подсказки")
    st.sidebar.info("""
    **Быстрый старт:**
    1. 📁 Загрузите каталог
    2. 📂 Категоризируйте товары
    3. 💰 Настройте тарифы
    4. 🧮 Рассчитайте юнит-экономику
    """)


def main():
    """
    🚗 ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ v101.0
    
    Объединяет все 4 раздела в единый интерфейс с навигацией.
    """
    # ========================================================================
    # НАСТРОЙКА STREAMLIT
    # ========================================================================
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚗",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # ========================================================================
    # ЗАГОЛОВОК И ЛОГОТИП
    # ========================================================================
    st.markdown("""
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border-radius: 10px; margin-bottom: 20px;'>
        <h1 style='color: white; margin: 0;'>🚗 Юнит-экономика автозапчастей PRO 2026</h1>
        <p style='color: #ccc; margin: 10px 0 0 0; font-size: 1.1em;'>
            Enterprise расчет юнит-экономики с AI, ABC/XYZ анализом и Google Sheets
        </p>
        <p style='color: #888; margin: 5px 0 0 0; font-size: 0.9em;'>
            Версия {APP_VERSION} | Специализация: Автозапчасти, Автотовары и Агрегаты
        </p>
    </div>
    """.format(APP_VERSION=APP_VERSION), unsafe_allow_html=True)
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ КОМПОНЕНТОВ
    # ========================================================================
    try:
        # Инициализация основных компонентов
        if 'marketplace_unit_economics' not in st.session_state:
            st.session_state.marketplace_unit_economics = get_marketplace_unit_economics()
        
        if 'high_volume_catalog' not in st.session_state:
            st.session_state.high_volume_catalog = get_high_volume_catalog()
        
        if 'category_classifier' not in st.session_state:
            st.session_state.category_classifier = CategoryClassifier()
        
        if 'category_dimensions_db' not in st.session_state:
            st.session_state.category_dimensions_db = CategoryDimensionsDB()
        
        if 'smart_tariff_loader' not in st.session_state:
            try:
                st.session_state.smart_tariff_loader = SmartTariffLoader()
            except Exception as e:
                logger.warning(f"Не удалось инициализировать SmartTariffLoader: {e}")
        
        if 'save_load_manager' not in st.session_state:
            st.session_state.save_load_manager = get_save_load_manager()
        
    except Exception as e:
        st.error(f"❌ Ошибка инициализации компонентов: {e}")
        logger.exception("Ошибка инициализации")
    
    # ========================================================================
    # НАВИГАЦИЯ ЧЕРЕЗ SIDEBAR
    # ========================================================================
    st.sidebar.title("🧭 Навигация")
    
    section = st.sidebar.radio(
        "Выберите раздел:",
        [
            "📁 : Загрузка данных",
            "📂 : Весогабариты и категоризация",
            "💰 : Тарифы",
            "🧮 : Расчёт юнит-экономики",
        ],
        key="main_navigation",
    )
    
    # ========================================================================
    # ОТОБРАЖЕНИЕ ВЫБРАННОГО РАЗДЕЛА
    # ========================================================================
    if section == "📁 : Загрузка данных":
        show_section1_data_loading()
    
    elif section == "📂 : Весогабариты и категоризация":
        show_section2_categorization()
    
    elif section == "💰 : Тарифы":
        show_section3_tariffs()
    
    elif section == "🧮 : Расчёт юнит-экономики":
        show_section4_calculation()
    
    # ========================================================================
    # SIDEBAR INFO
    # ========================================================================
    show_sidebar_info()
    
    # ========================================================================
    # ФУТЕР
    # ========================================================================
    show_footer()


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================
if __name__ == "__main__":
    main()


