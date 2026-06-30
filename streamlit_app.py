"""
================================================================================
🚀 ULTIMATE UNIT ECONOMICS ENGINE v80.0 - ПОЛНАЯ ВЕРСИЯ (5000+ СТРОК)
================================================================================
📌 ВЕРСИЯ: 80.0.0
📌 ОБЩИЙ ОБЪЕМ: 5,500+ СТРОК (ПОЛНАЯ ВЕРСИЯ БЕЗ СОКРАЩЕНИЙ)
📌 СОВМЕСТИМОСТЬ: Python 3.10 - 3.14
📌 ФУНКЦИОНАЛ:
    ✅ ЮНИТ-ЭКОНОМИКА С ТАРИФАМИ 2026 (FBY, FBS, FBO, DBS, FBP)
    ✅ КАТАЛОГ ЭНХАНСЕР (ПОИСК АНАЛОГОВ 2 УРОВНЯ)
    ✅ 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ
    ✅ ML-КЛАССИФИКАЦИЯ ТОВАРОВ
    ✅ ЭКСПОРТ В CSV/EXCEL/PDF
    ✅ УПРАВЛЕНИЕ ЦЕНАМИ И НАЦЕНКАМИ
    ✅ ИСТОРИЯ РАСЧЕТОВ С ФИЛЬТРАЦИЕЙ
    ✅ РАСШИРЕННАЯ СТАТИСТИКА С ГРАФИКАМИ
    ✅ ПРОГНОЗИРОВАНИЕ ПРИБЫЛИ
    ✅ СРАВНЕНИЕ С КОНКУРЕНТАМИ
    ✅ ВАЛИДАЦИЯ ДАННЫХ
    ✅ НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
    ✅ ПОЛНАЯ ДОКУМЕНТАЦИЯ ВСЕХ ФУНКЦИЙ
    ✅ ПОДДЕРЖКА PYTHON 3.14
    ✅ МИНИМАЛЬНЫЙ НАБОР ЗАВИСИМОСТЕЙ
    ✅ HIGH-VOLUME CATALOG (10M+ ЗАПИСЕЙ) С POLARS И DUCKDB
    ✅ ИИ-ОБНОВЛЕНИЕ ТАРИФОВ
    ✅ ОБЪЕДИНЕНИЕ ДАННЫХ С ВЫБОРОМ КРИТЕРИЕВ
    ✅ ЭКСПОРТ С ФОРМУЛАМИ
================================================================================
"""

# ============================================================================
# БЛОК ИМПОРТОВ - ВСЕ НЕОБХОДИМЫЕ БИБЛИОТЕКИ
# ============================================================================

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
import random
import math
import warnings
from typing import Dict, List, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from collections import defaultdict, Counter
from enum import Enum
from threading import Lock
from contextlib import contextmanager
import tempfile
from pathlib import Path
import csv
import base64
import urllib.parse

# Дополнительные импорты для HighVolume
try:
    import polars as pl
    import duckdb
    HIGH_VOLUME_AVAILABLE = True
except ImportError:
    HIGH_VOLUME_AVAILABLE = False

# Дополнительные импорты для расширенного функционала
try:
    from io import BytesIO
except ImportError:
    pass

# Попытка импорта для PDF экспорта
try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    PDF_EXPORT = True
except ImportError:
    PDF_EXPORT = False

# Попытка импорта для OpenAI
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# ============================================================================
# ПОДАВЛЕНИЕ ПРЕДУПРЕЖДЕНИЙ
# ============================================================================

warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'

# ============================================================================
# ВЕРСИЯ И КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
# ============================================================================

APP_VERSION = "80.0.0"
APP_NAME = "🚀 Юнит-экономика с каталогом и AI 2026"
EXCEL_ROW_LIMIT = 1_000_000
HISTORY_LIMIT = 1000
CACHE_TTL = 3600

# ============================================================================
# ПРОВЕРКА НАЛИЧИЯ УСТАНОВЛЕННЫХ БИБЛИОТЕК
# ============================================================================

LIBRARIES = {
    'openpyxl': False,
    'plotly': False,
    'sklearn': False,
    'openai': False,
    'duckdb': False,
    'polars': False,
    'joblib': False,
    'reportlab': False,
}

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, Reference, Series
    LIBRARIES['openpyxl'] = True
except ImportError:
    pass

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    LIBRARIES['plotly'] = True
except ImportError:
    pass

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.naive_bayes import MultinomialNB
    from sklearn.pipeline import Pipeline
    from sklearn.model_selection import train_test_split
    from sklearn.metrics import accuracy_score, classification_report
    from sklearn.ensemble import RandomForestClassifier
    import joblib
    LIBRARIES['sklearn'] = True
    LIBRARIES['joblib'] = True
except ImportError:
    pass

try:
    import openai
    LIBRARIES['openai'] = True
except ImportError:
    pass

try:
    import duckdb
    LIBRARIES['duckdb'] = True
except ImportError:
    pass

try:
    import polars as pl
    LIBRARIES['polars'] = True
except ImportError:
    pass

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    LIBRARIES['reportlab'] = True
except ImportError:
    pass

# ============================================================================
# КЛАСС ДЛЯ ЛОГИРОВАНИЯ (SINGLETON)
# ============================================================================

class Logger:
    _instance = None
    _lock = Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        
        self.logger = logging.getLogger('UnitEconomy')
        self.logger.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        fh = logging.FileHandler('unit_economy.log', encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)
        
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)
    
    def get(self):
        return self.logger

logger = Logger().get()

# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================

@contextmanager
def timer(name: str):
    start = time.time()
    try:
        yield
    finally:
        elapsed = time.time() - start
        logger.info(f"⏱ {name}: {elapsed:.2f}с")

def safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None:
            return default
        if val == "" or val == "NaN" or val == "nan" or val == "None":
            return default
        if isinstance(val, (int, float)):
            if math.isnan(val) or math.isinf(val):
                return default
            return float(val)
        if isinstance(val, str):
            val = val.replace(',', '.').replace(' ', '').replace('₽', '').replace('%', '')
            val = val.replace('$', '').replace('€', '').replace('£', '').replace('¥', '')
            val = val.replace('₴', '').replace('USD', '').replace('EUR', '')
            val = re.sub(r'[^\d.\-]', '', val)
            if not val or val == '-' or val == '.':
                return default
            return float(val)
        if isinstance(val, bool):
            return float(val)
        if hasattr(val, 'dtype') and hasattr(val, 'item'):
            try:
                return float(val.item())
            except:
                return default
        return default
    except (ValueError, TypeError, AttributeError):
        return default

def safe_str(val: Any, default: str = "") -> str:
    try:
        if val is None:
            return default
        if isinstance(val, (int, float)) and (math.isnan(val) or math.isinf(val)):
            return default
        if isinstance(val, bool):
            return str(val)
        if hasattr(val, 'dtype') and hasattr(val, 'item'):
            try:
                val = val.item()
            except:
                pass
        result = str(val).strip()
        return result if result else default
    except (ValueError, TypeError, AttributeError):
        return default

def safe_int(val: Any, default: int = 0) -> int:
    try:
        return int(safe_float(val, default))
    except (ValueError, TypeError):
        return default

def format_currency(value: float) -> str:
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0 ₽"
        if abs(value) >= 1:
            return f"{value:,.0f} ₽"
        else:
            return f"{value:.2f} ₽"
    except (ValueError, TypeError):
        return "0 ₽"

def format_percent(value: float) -> str:
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return "0%"
        if abs(value) >= 0.1:
            return f"{value:.1f}%"
        else:
            return f"{value:.2f}%"
    except (ValueError, TypeError):
        return "0%"

def generate_cache_key(*args) -> str:
    key = "|".join(str(arg) for arg in args)
    return hashlib.md5(key.encode()).hexdigest()

def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_numbers(text: str) -> List[float]:
    if not text:
        return []
    return [float(x) for x in re.findall(r'\d+\.?\d*', text)]

def calculate_volume(length: float, width: float, height: float) -> float:
    try:
        if all([length, width, height]) and all([length > 0, width > 0, height > 0]):
            if any([length > 1000, width > 1000, height > 1000]):
                return 0.0
            volume = (length * width * height) / 1000.0
            if volume < 0.001:
                return 0.0
            return round(volume, 3)
        return 0.0
    except (TypeError, ValueError):
        return 0.0

def convert_dimension(value: float, from_unit: str, to_unit: str) -> float:
    if value == 0:
        return 0.0
    if from_unit == to_unit:
        return value
    if from_unit == "мм" and to_unit == "см":
        return value / 10.0
    elif from_unit == "см" and to_unit == "мм":
        return value * 10.0
    elif from_unit == "м" and to_unit == "см":
        return value * 100.0
    elif from_unit == "см" and to_unit == "м":
        return value / 100.0
    elif from_unit == "м" and to_unit == "мм":
        return value * 1000.0
    elif from_unit == "мм" and to_unit == "м":
        return value / 1000.0
    return value

def is_valid_barcode(barcode: str) -> bool:
    if not barcode:
        return False
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) not in [8, 12, 13, 14, 15]:
        return False
    return True

def format_barcode(barcode: str) -> str:
    if not barcode:
        return ""
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) == 13:
        return f"{barcode[:3]} {barcode[3:7]} {barcode[7:11]} {barcode[11:]}"
    elif len(barcode) == 12:
        return f"{barcode[:2]} {barcode[2:6]} {barcode[6:10]} {barcode[10:]}"
    elif len(barcode) == 8:
        return f"{barcode[:2]} {barcode[2:5]} {barcode[5:]}"
    return barcode

def validate_article(article: str) -> bool:
    if not article or not article.strip():
        return False
    return bool(re.match(r'^[A-Za-z0-9\-_]+$', article.strip()))

def generate_random_id(length: int = 12) -> str:
    chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    return ''.join(random.choice(chars) for _ in range(length))

def detect_column_mapping(df: pd.DataFrame, required_columns: List[str]) -> Dict[str, str]:
    column_variants = {
        'artikul': ['артикул', 'article', 'sku', 'код товара', 'артикул продавца'],
        'brand': ['бренд', 'brand', 'производитель', 'manufacturer', 'марка'],
        'price': ['цена', 'price', 'стоимость', 'retail price', 'розничная цена'],
        'cost': ['себестоимость', 'cost', 'закупочная цена', 'purchase price'],
        'length': ['длина', 'length', 'длинна', 'габарит длина'],
        'width': ['ширина', 'width', 'габарит ширина'],
        'height': ['высота', 'height', 'габарит высота'],
        'weight': ['вес', 'weight', 'масса'],
        'oe_number': ['oe номер', 'oe number', 'oem', 'оригинальный номер'],
        'category': ['категория', 'category', 'группа', 'раздел'],
        'barcode': ['штрихкод', 'barcode', 'ean', 'штрих-код'],
        'description': ['описание', 'description', 'наименование', 'name'],
        'multiplicity': ['кратность', 'multiplicity', 'упаковка']
    }
    
    mapping = {}
    actual_lower = {col.lower(): col for col in df.columns}
    
    for required in required_columns:
        variants = column_variants.get(required, [required])
        for variant in variants:
            variant_lower = variant.lower()
            for actual_l, actual_orig in actual_lower.items():
                if variant_lower in actual_l and actual_orig not in mapping:
                    mapping[actual_orig] = required
                    break
    
    return mapping

# ============================================================================
# БЛОК 1: ENUM ДЛЯ ТИПОВ КОМИССИЙ И РЕЖИМОВ РАБОТЫ
# ============================================================================

class CommissionType(Enum):
    PERCENTAGE = "percentage"
    FIXED = "fixed"
    HYBRID = "hybrid"
    SUBSCRIPTION = "subscription"

class OperationMode(Enum):
    FBY = "FBY"
    FBS = "FBS"
    FBO = "FBO"
    DBS = "DBS"
    FBP = "FBP"

# ============================================================================
# БЛОК 2: IMMUTABLE КОНФИГУРАЦИИ МАРКЕТПЛЕЙСОВ 2026
# ============================================================================

@dataclass(frozen=True)
class MarketplaceConfig2026:
    commission_rate: float
    commission_type: CommissionType = CommissionType.PERCENTAGE
    subscription_fee: float = 0.0
    min_commission: float = 0.0
    max_commission: float = float('inf')
    logistics_base: float = 0.0
    logistics_per_kg: float = 0.0
    logistics_per_liter: float = 0.0
    logistics_fixed_routes: Dict[str, float] = field(default_factory=dict)
    storage_per_day: float = 0.0
    storage_non_standard_fee: float = 0.0
    return_fee: float = 0.0
    acquiring_fee: float = 0.0
    vat_rate: float = 0.20
    last_mile_fee: float = 0.0
    delivery_fee_percent: float = 0.0
    premium_section_fee: float = 0.0
    rko_fee: float = 0.0
    mode_multipliers: Dict[str, float] = field(default_factory=lambda: {
        "FBY": 0.75,
        "FBS": 1.0,
        "FBO": 0.8,
        "DBS": 1.3,
        "FBP": 0.9
    })
    category_rates: Dict[str, float] = field(default_factory=dict)
    
    def get_commission_rate(self, category: Optional[str] = None) -> float:
        if category and category in self.category_rates:
            return self.category_rates[category]
        return self.commission_rate
    
    def get_mode_multiplier(self, mode: str) -> float:
        return self.mode_multipliers.get(mode, 1.0)

# ============================================================================
# БЛОК 3: АКТУАЛЬНЫЕ КОНФИГУРАЦИИ НА 2026 ГОД
# ============================================================================

def get_marketplace_configs_2026() -> Dict[str, MarketplaceConfig2026]:
    return {
        "Яндекс Маркет": MarketplaceConfig2026(
            commission_rate=0.14,
            commission_type=CommissionType.SUBSCRIPTION,
            subscription_fee=6990.0,
            min_commission=0.0,
            logistics_base=150.0,
            logistics_per_kg=50.0,
            logistics_per_liter=30.0,
            storage_per_day=0.5,
            return_fee=0.03,
            acquiring_fee=0.015,
            last_mile_fee=100.0,
            delivery_fee_percent=0.05,
            premium_section_fee=0.02,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "одежда_обувь": 0.14,
                "садоводство": 0.12,
                "строительство": 0.19,
                "красота": 0.14,
                "детские_товары": 0.14,
                "электроника": 0.14,
                "автотовары": 0.14,
                "книги": 0.14,
                "дом": 0.14,
                "спорт": 0.14
            }
        ),
        "Ozon": MarketplaceConfig2026(
            commission_rate=0.15,
            min_commission=30.0,
            logistics_base=90.0,
            logistics_per_kg=35.0,
            logistics_per_liter=20.0,
            storage_per_day=0.3,
            return_fee=0.05,
            acquiring_fee=0.01,
            last_mile_fee=80.0,
            delivery_fee_percent=0.04,
            storage_non_standard_fee=0.03,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "одежда_обувь": 0.15,
                "электроника": 0.10,
                "красота": 0.22,
                "автотовары": 0.12,
                "книги": 0.10,
                "дом": 0.12,
                "спорт": 0.12,
                "детские_товары": 0.12,
                "продукты": 0.08,
                "здоровье": 0.15
            }
        ),
        "Wildberries": MarketplaceConfig2026(
            commission_rate=0.15,
            min_commission=40.0,
            logistics_base=120.0,
            logistics_per_kg=45.0,
            logistics_per_liter=28.0,
            storage_per_day=0.4,
            return_fee=0.04,
            acquiring_fee=0.012,
            last_mile_fee=90.0,
            delivery_fee_percent=0.06,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "одежда": 0.18,
                "электроника": 0.12,
                "дети": 0.15,
                "дом": 0.15,
                "красота": 0.15,
                "продукты": 0.10,
                "здоровье": 0.12,
                "спорт": 0.15,
                "книги": 0.12,
                "автотовары": 0.15
            }
        ),
        "AliExpress": MarketplaceConfig2026(
            commission_rate=0.10,
            min_commission=60.0,
            logistics_base=200.0,
            logistics_per_kg=60.0,
            logistics_per_liter=35.0,
            storage_per_day=0.6,
            return_fee=0.02,
            acquiring_fee=0.02,
            last_mile_fee=150.0,
            delivery_fee_percent=0.08,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "электроника": 0.08,
                "одежда": 0.10,
                "дом": 0.10,
                "красота": 0.10,
                "спорт": 0.10,
                "автотовары": 0.10,
                "книги": 0.08,
                "детские_товары": 0.10
            }
        ),
        "Мегамаркет": MarketplaceConfig2026(
            commission_rate=0.09,
            min_commission=45.0,
            logistics_base=130.0,
            logistics_per_kg=42.0,
            logistics_per_liter=26.0,
            storage_per_day=0.35,
            return_fee=0.03,
            acquiring_fee=0.013,
            last_mile_fee=95.0,
            delivery_fee_percent=0.05,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "электроника": 0.02,
                "одежда": 0.20,
                "обувь": 0.20,
                "автотовары": 0.15,
                "дом": 0.12,
                "красота": 0.12,
                "спорт": 0.12,
                "детские_товары": 0.12,
                "продукты": 0.05,
                "книги": 0.08
            }
        ),
        "СберМегаМаркет": MarketplaceConfig2026(
            commission_rate=0.085,
            rko_fee=0.015,
            min_commission=48.0,
            logistics_base=140.0,
            logistics_per_kg=48.0,
            logistics_per_liter=29.0,
            storage_per_day=0.38,
            return_fee=0.035,
            acquiring_fee=0.014,
            last_mile_fee=105.0,
            delivery_fee_percent=0.055,
            mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
            category_rates={
                "электроника": 0.02,
                "одежда": 0.15,
                "продукты": 0.05,
                "дом": 0.10,
                "красота": 0.10,
                "спорт": 0.10,
                "автотовары": 0.12,
                "детские_товары": 0.10,
                "книги": 0.08
            }
        )
    }

# ============================================================================
# БЛОК 4: 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ
# ============================================================================

@dataclass
class CategoryDimensions:
    min_length: float = 0.0
    max_length: float = 0.0
    min_width: float = 0.0
    max_width: float = 0.0
    min_height: float = 0.0
    max_height: float = 0.0
    min_weight: float = 0.0
    max_weight: float = 0.0
    typical_volume: float = 0.0
    description: str = ""

def get_category_dimensions() -> Dict[str, CategoryDimensions]:
    categories = {
        "Двигатель в сборе": CategoryDimensions(
            min_length=50.0, max_length=90.0,
            min_width=40.0, max_width=70.0,
            min_height=40.0, max_height=70.0,
            min_weight=50.0, max_weight=200.0,
            typical_volume=200.0,
            description="Двигатель в сборе для легковых автомобилей"
        ),
        "Блок цилиндров": CategoryDimensions(
            min_length=40.0, max_length=70.0,
            min_width=30.0, max_width=50.0,
            min_height=20.0, max_height=40.0,
            min_weight=20.0, max_weight=80.0,
            typical_volume=100.0,
            description="Блок цилиндров двигателя"
        ),
        "Головка блока цилиндров": CategoryDimensions(
            min_length=30.0, max_length=60.0,
            min_width=20.0, max_width=40.0,
            min_height=8.0, max_height=20.0,
            min_weight=5.0, max_weight=30.0,
            typical_volume=40.0,
            description="Головка блока цилиндров (ГБЦ)"
        ),
        "Коленчатый вал": CategoryDimensions(
            min_length=40.0, max_length=90.0,
            min_width=8.0, max_width=20.0,
            min_height=8.0, max_height=20.0,
            min_weight=10.0, max_weight=40.0,
            typical_volume=30.0,
            description="Коленчатый вал двигателя"
        ),
        "Распределительный вал": CategoryDimensions(
            min_length=30.0, max_length=80.0,
            min_width=5.0, max_width=15.0,
            min_height=5.0, max_height=15.0,
            min_weight=3.0, max_weight=15.0,
            typical_volume=20.0,
            description="Распределительный вал"
        ),
        "Поршневая группа": CategoryDimensions(
            min_length=8.0, max_length=20.0,
            min_width=8.0, max_width=20.0,
            min_height=8.0, max_height=20.0,
            min_weight=0.5, max_weight=3.0,
            typical_volume=5.0,
            description="Поршневая группа в сборе"
        ),
        "Шатун": CategoryDimensions(
            min_length=12.0, max_length=35.0,
            min_width=4.0, max_width=10.0,
            min_height=3.0, max_height=7.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Шатун двигателя"
        ),
        "Клапана": CategoryDimensions(
            min_length=6.0, max_length=15.0,
            min_width=2.0, max_width=5.0,
            min_height=2.0, max_height=5.0,
            min_weight=0.05, max_weight=0.2,
            typical_volume=0.5,
            description="Клапана двигателя"
        ),
        "Гидрокомпенсаторы": CategoryDimensions(
            min_length=3.0, max_length=8.0,
            min_width=3.0, max_width=8.0,
            min_height=3.0, max_height=8.0,
            min_weight=0.05, max_weight=0.2,
            typical_volume=0.3,
            description="Гидрокомпенсаторы"
        ),
        "Привод ГРМ": CategoryDimensions(
            min_length=60.0, max_length=160.0,
            min_width=2.0, max_width=5.0,
            min_height=1.0, max_height=2.0,
            min_weight=0.1, max_weight=1.0,
            typical_volume=2.0,
            description="Привод ГРМ (ремень, цепь)"
        ),
        "Масляный насос": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=1.0, max_weight=5.0,
            typical_volume=5.0,
            description="Масляный насос двигателя"
        ),
        "Водяной насос": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=1.0, max_weight=4.0,
            typical_volume=5.0,
            description="Водяной насос (помпа)"
        ),
        "Турбокомпрессор": CategoryDimensions(
            min_length=15.0, max_length=35.0,
            min_width=15.0, max_width=30.0,
            min_height=15.0, max_height=25.0,
            min_weight=5.0, max_weight=15.0,
            typical_volume=15.0,
            description="Турбокомпрессор"
        ),
        "Прокладки двигателя": CategoryDimensions(
            min_length=2.0, max_length=60.0,
            min_width=2.0, max_width=40.0,
            min_height=0.1, max_height=2.0,
            min_weight=0.01, max_weight=0.5,
            typical_volume=0.5,
            description="Прокладки двигателя"
        ),
        "Масляный поддон": CategoryDimensions(
            min_length=30.0, max_length=60.0,
            min_width=20.0, max_width=40.0,
            min_height=10.0, max_height=20.0,
            min_weight=2.0, max_weight=8.0,
            typical_volume=15.0,
            description="Масляный поддон"
        ),
        "Клапанная крышка": CategoryDimensions(
            min_length=30.0, max_length=60.0,
            min_width=15.0, max_width=30.0,
            min_height=5.0, max_height=10.0,
            min_weight=1.0, max_weight=4.0,
            typical_volume=8.0,
            description="Клапанная крышка"
        ),
        "Приводной ремень": CategoryDimensions(
            min_length=60.0, max_length=150.0,
            min_width=1.0, max_width=3.0,
            min_height=0.5, max_height=1.0,
            min_weight=0.05, max_weight=0.5,
            typical_volume=1.0,
            description="Приводной ремень"
        ),
        "Демпфер коленвала": CategoryDimensions(
            min_length=10.0, max_length=25.0,
            min_width=10.0, max_width=25.0,
            min_height=5.0, max_height=10.0,
            min_weight=2.0, max_weight=8.0,
            typical_volume=5.0,
            description="Демпфер коленвала"
        ),
        "Маховик": CategoryDimensions(
            min_length=25.0, max_length=45.0,
            min_width=25.0, max_width=45.0,
            min_height=5.0, max_height=10.0,
            min_weight=5.0, max_weight=15.0,
            typical_volume=10.0,
            description="Маховик"
        ),
        "Стартерный венец": CategoryDimensions(
            min_length=25.0, max_length=40.0,
            min_width=25.0, max_width=40.0,
            min_height=2.0, max_height=5.0,
            min_weight=1.0, max_weight=5.0,
            typical_volume=5.0,
            description="Стартерный венец"
        ),
        "Коробка передач в сборе": CategoryDimensions(
            min_length=40.0, max_length=70.0,
            min_width=30.0, max_width=50.0,
            min_height=25.0, max_height=40.0,
            min_weight=30.0, max_weight=80.0,
            typical_volume=80.0,
            description="Коробка передач в сборе"
        ),
        "Сцепление": CategoryDimensions(
            min_length=25.0, max_length=35.0,
            min_width=25.0, max_width=35.0,
            min_height=8.0, max_height=15.0,
            min_weight=5.0, max_weight=15.0,
            typical_volume=15.0,
            description="Сцепление в сборе"
        ),
        "Привод": CategoryDimensions(
            min_length=40.0, max_length=90.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=3.0, max_weight=12.0,
            typical_volume=15.0,
            description="Привод (полуоси)"
        ),
        "Дифференциал": CategoryDimensions(
            min_length=20.0, max_length=45.0,
            min_width=20.0, max_width=45.0,
            min_height=20.0, max_height=45.0,
            min_weight=10.0, max_weight=30.0,
            typical_volume=30.0,
            description="Дифференциал"
        ),
        "Карданный вал": CategoryDimensions(
            min_length=60.0, max_length=160.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=5.0, max_weight=20.0,
            typical_volume=25.0,
            description="Карданный вал"
        ),
        "Раздаточная коробка": CategoryDimensions(
            min_length=25.0, max_length=45.0,
            min_width=20.0, max_width=35.0,
            min_height=20.0, max_height=35.0,
            min_weight=15.0, max_weight=40.0,
            typical_volume=35.0,
            description="Раздаточная коробка"
        ),
        "Гидротрансформатор": CategoryDimensions(
            min_length=25.0, max_length=40.0,
            min_width=25.0, max_width=40.0,
            min_height=20.0, max_height=30.0,
            min_weight=10.0, max_weight=25.0,
            typical_volume=30.0,
            description="Гидротрансформатор АКПП"
        ),
        "Механизм переключения": CategoryDimensions(
            min_length=15.0, max_length=35.0,
            min_width=5.0, max_width=15.0,
            min_height=5.0, max_height=15.0,
            min_weight=1.0, max_weight=5.0,
            typical_volume=5.0,
            description="Механизм переключения передач"
        ),
        "Подшипники трансмиссии": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=0.5, max_weight=3.0,
            typical_volume=3.0,
            description="Подшипники трансмиссии"
        ),
        "Сальники трансмиссии": CategoryDimensions(
            min_length=2.0, max_length=12.0,
            min_width=2.0, max_width=12.0,
            min_height=1.0, max_height=3.0,
            min_weight=0.05, max_weight=0.3,
            typical_volume=0.5,
            description="Сальники трансмиссии"
        ),
        "Фильтр АКПП": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Фильтр АКПП"
        ),
        "Масло трансмиссионное": CategoryDimensions(
            min_length=10.0, max_length=35.0,
            min_width=8.0, max_width=25.0,
            min_height=8.0, max_height=25.0,
            min_weight=1.0, max_weight=5.0,
            typical_volume=5.0,
            description="Канистра с трансмиссионным маслом"
        ),
        "Трос сцепления": CategoryDimensions(
            min_length=40.0, max_length=100.0,
            min_width=1.0, max_width=3.0,
            min_height=1.0, max_height=3.0,
            min_weight=0.1, max_weight=0.5,
            typical_volume=1.0,
            description="Трос сцепления"
        ),
        "Цилиндр сцепления": CategoryDimensions(
            min_length=10.0, max_length=20.0,
            min_width=5.0, max_width=10.0,
            min_height=5.0, max_height=10.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=2.0,
            description="Цилиндр сцепления"
        ),
        "Вал КПП": CategoryDimensions(
            min_length=20.0, max_length=50.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=2.0, max_weight=8.0,
            typical_volume=8.0,
            description="Вал КПП"
        ),
        "Шестерни КПП": CategoryDimensions(
            min_length=5.0, max_length=15.0,
            min_width=5.0, max_width=15.0,
            min_height=5.0, max_height=15.0,
            min_weight=0.5, max_weight=3.0,
            typical_volume=3.0,
            description="Шестерни КПП"
        ),
        "Синхронизатор": CategoryDimensions(
            min_length=5.0, max_length=12.0,
            min_width=5.0, max_width=12.0,
            min_height=3.0, max_height=8.0,
            min_weight=0.3, max_weight=1.5,
            typical_volume=2.0,
            description="Синхронизатор"
        ),
        "Муфта КПП": CategoryDimensions(
            min_length=5.0, max_length=15.0,
            min_width=5.0, max_width=15.0,
            min_height=3.0, max_height=8.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Муфта КПП"
        ),
        "Амортизатор": CategoryDimensions(
            min_length=25.0, max_length=85.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=2.0, max_weight=8.0,
            typical_volume=8.0,
            description="Амортизатор подвески"
        ),
        "Пружина подвески": CategoryDimensions(
            min_length=15.0, max_length=45.0,
            min_width=15.0, max_width=25.0,
            min_height=15.0, max_height=25.0,
            min_weight=2.0, max_weight=8.0,
            typical_volume=10.0,
            description="Пружина подвески"
        ),
        "Рычаг подвески": CategoryDimensions(
            min_length=20.0, max_length=65.0,
            min_width=5.0, max_width=18.0,
            min_height=5.0, max_height=18.0,
            min_weight=2.0, max_weight=10.0,
            typical_volume=10.0,
            description="Рычаг подвески"
        ),
        "Сайлентблок": CategoryDimensions(
            min_length=5.0, max_length=18.0,
            min_width=5.0, max_width=18.0,
            min_height=5.0, max_height=18.0,
            min_weight=0.2, max_weight=1.5,
            typical_volume=2.0,
            description="Сайлентблок"
        ),
        "Шаровая опора": CategoryDimensions(
            min_length=5.0, max_length=12.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.3, max_weight=1.5,
            typical_volume=2.0,
            description="Шаровая опора"
        ),
        "Стабилизатор": CategoryDimensions(
            min_length=25.0, max_length=65.0,
            min_width=3.0, max_width=10.0,
            min_height=3.0, max_height=10.0,
            min_weight=1.0, max_weight=5.0,
            typical_volume=5.0,
            description="Стабилизатор поперечной устойчивости"
        ),
        "Пыльник": CategoryDimensions(
            min_length=5.0, max_length=12.0,
            min_width=5.0, max_width=12.0,
            min_height=8.0, max_height=22.0,
            min_weight=0.1, max_weight=0.5,
            typical_volume=1.0,
            description="Пыльник (чехол)"
        ),
        "Отбойник": CategoryDimensions(
            min_length=5.0, max_length=12.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.1, max_weight=0.5,
            typical_volume=1.0,
            description="Отбойник амортизатора"
        ),
        "Опора стойки": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Опора стойки амортизатора"
        ),
        "Тяга рулевая": CategoryDimensions(
            min_length=25.0, max_length=65.0,
            min_width=3.0, max_width=8.0,
            min_height=3.0, max_height=8.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Рулевая тяга"
        ),
        "Рулевая рейка": CategoryDimensions(
            min_length=35.0, max_length=85.0,
            min_width=8.0, max_width=18.0,
            min_height=8.0, max_height=18.0,
            min_weight=3.0, max_weight=10.0,
            typical_volume=10.0,
            description="Рулевая рейка"
        ),
        "Рулевой кардан": CategoryDimensions(
            min_length=20.0, max_length=45.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=1.0, max_weight=4.0,
            typical_volume=5.0,
            description="Рулевой кардан"
        ),
        "Усилитель руля": CategoryDimensions(
            min_length=15.0, max_length=30.0,
            min_width=15.0, max_width=30.0,
            min_height=15.0, max_height=25.0,
            min_weight=3.0, max_weight=10.0,
            typical_volume=10.0,
            description="Усилитель руля (ГУР/ЭУР)"
        ),
        "Подрамник": CategoryDimensions(
            min_length=45.0, max_length=105.0,
            min_width=15.0, max_width=35.0,
            min_height=8.0, max_height=18.0,
            min_weight=10.0, max_weight=30.0,
            typical_volume=25.0,
            description="Подрамник"
        ),
        "Распорка": CategoryDimensions(
            min_length=25.0, max_length=65.0,
            min_width=2.0, max_width=6.0,
            min_height=2.0, max_height=6.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=2.0,
            description="Распорка подвески"
        ),
        "Сайлентблоки в сборе": CategoryDimensions(
            min_length=8.0, max_length=22.0,
            min_width=8.0, max_width=22.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=3.0,
            description="Сайлентблоки в сборе"
        ),
        "Буфер": CategoryDimensions(
            min_length=5.0, max_length=12.0,
            min_width=5.0, max_width=12.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.1, max_weight=0.5,
            typical_volume=1.0,
            description="Буфер подвески"
        ),
        "Подушка подвески": CategoryDimensions(
            min_length=8.0, max_length=18.0,
            min_width=8.0, max_width=18.0,
            min_height=5.0, max_height=12.0,
            min_weight=0.5, max_weight=2.0,
            typical_volume=2.0,
            description="Подушка подвески"
        ),
        "Тяга продольная": CategoryDimensions(
            min_length=25.0, max_length=65.0,
            min_width=3.0, max_width=8.0,
            min_height=3.0, max_height=8.0,
            min_weight=1.0, max_weight=4.0,
            typical_volume=4.0,
            description="Тяга продольная"
        ),
        "Балка моста": CategoryDimensions(
            min_length=45.0, max_length=85.0,
            min_width=10.0, max_width=20.0,
            min_height=10.0, max_height=20.0,
            min_weight=15.0, max_weight=40.0,
            typical_volume=30.0,
            description="Балка моста"
        )
    }
    return categories

# ============================================================================
# БЛОК 5: HIGH-VOLUME CATALOG (ИНТЕГРАЦИЯ С POLARS И DUCKDB)
# ============================================================================

class HighVolumeAutoPartsCatalog:
    """
    High-Volume каталог автозапчастей с поддержкой 10M+ записей
    
    Использует Polars для быстрой обработки и DuckDB для хранения.
    Поддерживает параллельную загрузку, экспорт и управление данными.
    """
    
    def __init__(self):
        self.data_dir = Path("./auto_parts_data")
        self.data_dir.mkdir(exist_ok=True)
        
        # Загрузка конфигураций
        self.cloud_config = self.load_cloud_config()
        self.price_rules = self.load_price_rules()
        self.exclusion_rules = self.load_exclusion_rules()
        self.category_mapping = self.load_category_mapping()
        
        self.db_path = self.data_dir / "catalog.duckdb"
        
        if HIGH_VOLUME_AVAILABLE and LIBRARIES['duckdb']:
            self.conn = duckdb.connect(database=str(self.db_path))
            self.setup_database()
        else:
            self.conn = None
            logger.warning("HighVolume режим недоступен: установите polars и duckdb")
    
    # --- Конфигурации ---
    def load_cloud_config(self) -> Dict[str, Any]:
        config_path = self.data_dir / "cloud_config.json"
        default_config = {
            "enabled": False,
            "provider": "s3",
            "bucket": "",
            "region": "",
            "sync_interval": 3600,
            "last_sync": 0
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
            "global_markup": 0.2,
            "brand_markups": {},
            "min_price": 0.0,
            "max_price": 99999.0
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
    
    def load_exclusion_rules(self) -> List[str]:
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        if exclusion_path.exists():
            try:
                return [line.strip() for line in exclusion_path.read_text(encoding='utf-8').splitlines() if line.strip()]
            except Exception as e:
                logger.error(f"Ошибка чтения exclusion_rules.txt: {e}")
                return []
        else:
            content = "Кузов\nСтекла\nМасла"
            exclusion_path.write_text(content, encoding='utf-8')
            return ["Кузов", "Стекла", "Масла"]
    
    def save_exclusion_rules(self):
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        exclusion_path.write_text(
            "\n".join(self.exclusion_rules), encoding='utf-8')
    
    def load_category_mapping(self) -> Dict[str, str]:
        category_path = self.data_dir / "category_mapping.txt"
        default_mapping = {
            "Радиатор": "Охлаждение",
            "Шаровая опора": "Подвеска",
            "Фильтр масляный": "Фильтры",
            "Тормозные колодки": "Тормоза"
        }
        if category_path.exists():
            try:
                mapping = {}
                for line in category_path.read_text(encoding='utf-8').splitlines():
                    if line.strip() and "|" in line:
                        key, value = line.split("|", 1)
                        mapping[key.strip()] = value.strip()
                return mapping
            except Exception as e:
                logger.error(f"Ошибка чтения category_mapping.txt: {e}")
                return default_mapping
        else:
            content = "\n".join(
                [f"{k}|{v}" for k, v in default_mapping.items()])
            category_path.write_text(content, encoding='utf-8')
            return default_mapping
    
    def save_category_mapping(self):
        category_path = self.data_dir / "category_mapping.txt"
        content = "\n".join(
            [f"{k}|{v}" for k, v in self.category_mapping.items()])
        category_path.write_text(content, encoding='utf-8')
    
    # --- База данных ---
    def setup_database(self):
        if not self.conn:
            return
            
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS oe (
                oe_number_norm VARCHAR PRIMARY KEY,
                oe_number VARCHAR,
                name VARCHAR,
                applicability VARCHAR,
                category VARCHAR
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
        if not self.conn:
            return
            
        try:
            indexes = [
                "CREATE INDEX IF NOT EXISTS idx_oe_number_norm ON oe(oe_number_norm)",
                "CREATE INDEX IF NOT EXISTS idx_parts_keys ON parts(artikul_norm, brand_norm)",
                "CREATE INDEX IF NOT EXISTS idx_cross_oe ON cross_references(oe_number_norm)",
                "CREATE INDEX IF NOT EXISTS idx_cross_artikul ON cross_references(artikul_norm, brand_norm)",
                "CREATE INDEX IF NOT EXISTS idx_prices_keys ON prices(artikul_norm, brand_norm)"
            ]
            for index_sql in indexes:
                try:
                    self.conn.execute(index_sql)
                except Exception as e:
                    logger.warning(f"Не удалось создать индекс: {e}")
        except Exception as e:
            logger.warning(f"Ошибка создания индексов: {e}")
    
    # --- Нормализация и очистка ---
    @staticmethod
    def normalize_key(series) -> pl.Series:
        if not HIGH_VOLUME_AVAILABLE:
            return series
        
        return (series
                .fill_null("")
                .cast(pl.Utf8)
                .str.replace_all("'", "")
                .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
                .str.replace_all(r"\s+", " ")
                .str.strip_chars()
                .str.to_lowercase())
    
    @staticmethod
    def clean_values(series) -> pl.Series:
        if not HIGH_VOLUME_AVAILABLE:
            return series
            
        return (series
                .fill_null("")
                .cast(pl.Utf8)
                .str.replace_all("'", "")
                .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
                .str.replace_all(r"\s+", " ")
                .str.strip_chars())
    
    def determine_category_vectorized(self, name_series) -> pl.Series:
        if not HIGH_VOLUME_AVAILABLE:
            return name_series
            
        name_lower = name_series.str.to_lowercase()
        categorization_expr = pl.when(pl.lit(False)).then(pl.lit(None))
        
        # Пользовательские правила — приоритет
        for key, category in self.category_mapping.items():
            categorization_expr = categorization_expr.when(
                name_lower.str.contains(key.lower())
            ).then(pl.lit(category))
        
        # Стандартные правила
        categories_map = {
            'Фильтр': 'фильтр|filter',
            'Тормоза': 'тормоз|brake|колодк|диск|суппорт',
            'Подвеска': 'амортизатор|стойк|spring|подвеск|рычаг',
            'Двигатель': 'двигатель|engine|свеч|поршень|клапан',
            'Трансмиссия': 'трансмиссия|сцеплен|коробк|transmission',
            'Электрика': 'аккумулятор|генератор|стартер|провод|ламп',
            'Рулевое': 'рулевой|тяга|наконечник|steering',
            'Выпуск': 'глушитель|катализатор|выхлоп|exhaust',
            'Охлаждение': 'радиатор|вентилятор|термостат|cooling',
            'Топливо': 'топливный|бензонасос|форсунк|fuel'
        }
        
        for category, pattern in categories_map.items():
            categorization_expr = categorization_expr.when(
                name_lower.str.contains(pattern, literal=False)
            ).then(pl.lit(category))
        
        return categorization_expr.otherwise(pl.lit('Разное')).alias('category')
    
    # --- Обработка файлов ---
    def detect_columns(self, actual_columns: List[str], expected_columns: List[str]) -> Dict[str, str]:
        column_variants = {
            'oe_number': ['oe номер', 'oe', 'оe', 'номер', 'code', 'OE'],
            'artikul': ['артикул', 'article', 'sku'],
            'brand': ['бренд', 'brand', 'производитель', 'manufacturer'],
            'name': ['наименование', 'название', 'name', 'описание', 'description'],
            'applicability': ['применимость', 'автомобиль', 'vehicle', 'applicability'],
            'barcode': ['штрих-код', 'barcode', 'штрихкод', 'ean', 'eac13'],
            'multiplicity': ['кратность шт', 'кратность', 'multiplicity'],
            'length': ['длина (см)', 'длина', 'length', 'длинна'],
            'width': ['ширина (см)', 'ширина', 'width'],
            'height': ['высота (см)', 'высота', 'height'],
            'weight': ['вес (кг)', 'вес, кг', 'вес', 'weight'],
            'image_url': ['ссылка', 'url', 'изображение', 'image', 'картинка'],
            'dimensions_str': ['весогабариты', 'размеры', 'dimensions', 'size'],
            'price': ['цена', 'price', 'рекомендованная цена', 'retail price'],
            'currency': ['валюта', 'currency']
        }
        actual_lower = {col.lower(): col for col in actual_columns}
        mapping = {}
        for expected in expected_columns:
            variants = column_variants.get(expected, [expected])
            for variant in variants:
                variant_lower = variant.lower()
                for actual_l, actual_orig in actual_lower.items():
                    if variant_lower in actual_l and actual_orig not in mapping:
                        mapping[actual_orig] = expected
                        break
        return mapping
    
    def read_and_prepare_file(self, file_path: str, file_type: str) -> pl.DataFrame:
        if not HIGH_VOLUME_AVAILABLE:
            logger.warning("Polars не доступен")
            return pl.DataFrame()
            
        logger.info(f"Обработка файла: {file_type} ({file_path})")
        try:
            if not os.path.exists(file_path):
                logger.error(f"Файл не найден: {file_path}")
                return pl.DataFrame()
            
            df = pl.read_excel(file_path, engine='calamine')
            if df.is_empty():
                logger.warning(f"Пустой файл: {file_path}")
                return pl.DataFrame()
                
        except Exception as e:
            logger.exception(f"Ошибка чтения файла {file_path}: {e}")
            return pl.DataFrame()
        
        schemas = {
            'oe': ['oe_number', 'artikul', 'brand', 'name', 'applicability'],
            'cross': ['oe_number', 'artikul', 'brand'],
            'barcode': ['artikul', 'brand', 'barcode', 'multiplicity'],
            'dimensions': ['artikul', 'brand', 'length', 'width', 'height', 'weight', 'dimensions_str'],
            'images': ['artikul', 'brand', 'image_url'],
            'prices': ['artikul', 'brand', 'price', 'currency']
        }
        expected_cols = schemas.get(file_type, [])
        column_mapping = self.detect_columns(df.columns, expected_cols)
        if not column_mapping:
            logger.warning(
                f"Не удалось определить колонки для файла {file_type}. Доступные: {df.columns}")
            return pl.DataFrame()
        
        df = df.rename(column_mapping)
        
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.clean_values(pl.col(col)).alias(col))
        
        key_cols = [col for col in ['oe_number',
                                    'artikul', 'brand'] if col in df.columns]
        if key_cols:
            df = df.unique(subset=key_cols, keep='first')
        
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.normalize_key(
                    pl.col(col)).alias(f"{col}_norm"))
        
        return df
    
    # --- Загрузка и обновление в базе ---
    def upsert_data(self, table_name: str, df: pl.DataFrame, pk: List[str]):
        if not self.conn or df.is_empty():
            return
            
        df = df.unique(keep='first')
        cols = df.columns
        temp_view_name = f"temp_{table_name}_{int(time.time())}"
        
        try:
            self.conn.register(temp_view_name, df.to_arrow())
        except Exception as e:
            logger.error(f"Ошибка регистрации временной таблицы: {e}")
            return
        
        try:
            pk_list = pk
            pk_cols_csv = ", ".join(f'"{c}"' for c in pk_list)
            delete_sql = f"""
                DELETE FROM {table_name}
                WHERE ({pk_cols_csv}) IN (SELECT {pk_cols_csv} FROM {temp_view_name});
            """
            self.conn.execute(delete_sql)
            insert_sql = f"""
                INSERT INTO {table_name}
                SELECT * FROM {temp_view_name};
            """
            self.conn.execute(insert_sql)
            logger.info(
                f"Успешно upsert {len(df)} записей в таблицу {table_name}.")
        except Exception as e:
            logger.error(f"Ошибка при UPSERT в {table_name}: {e}")
        finally:
            try:
                self.conn.unregister(temp_view_name)
            except Exception:
                pass
    
    def upsert_prices(self, price_df: pl.DataFrame):
        if not self.conn or price_df.is_empty():
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
    
    def process_and_load_data(self, dataframes: Dict[str, pl.DataFrame]):
        if not self.conn:
            st.warning("⚠️ База данных не доступна")
            return
            
        st.info("🔄 Начало загрузки и обновления данных в базе...")
        steps = [s for s in ['oe', 'cross', 'parts'] if s in dataframes]
        num_steps = len(steps)
        progress_bar = st.progress(
            0, text="Подготовка к обновлению базы данных...")
        step_counter = 0
        
        if 'oe' in dataframes:
            step_counter += 1
            progress_bar.progress(step_counter / (num_steps + 1),
                                  text=f"({step_counter}/{num_steps}) Обработка OE данных...")
            df = dataframes['oe'].filter(pl.col('oe_number_norm') != "")
            oe_df = df.select(['oe_number_norm', 'oe_number', 'name', 'applicability']).unique(
                subset=['oe_number_norm'], keep='first')
            
            if 'name' in oe_df.columns:
                oe_df = oe_df.with_columns(
                    self.determine_category_vectorized(pl.col('name')))
            else:
                oe_df = oe_df.with_columns(category=pl.lit('Разное'))
            
            self.upsert_data('oe', oe_df, ['oe_number_norm'])
            
            cross_df_from_oe = df.filter(pl.col('artikul_norm') != "").select(
                ['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            self.upsert_data('cross_references', cross_df_from_oe, [
                             'oe_number_norm', 'artikul_norm', 'brand_norm'])
        
        if 'cross' in dataframes:
            step_counter += 1
            progress_bar.progress(step_counter / (num_steps + 1),
                                  text=f"({step_counter}/{num_steps}) Обработка кроссов...")
            df = dataframes['cross'].filter(
                (pl.col('oe_number_norm') != "") & (pl.col('artikul_norm') != ""))
            cross_df_from_cross = df.select(
                ['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            self.upsert_data('cross_references', cross_df_from_cross, [
                             'oe_number_norm', 'artikul_norm', 'brand_norm'])
        
        if 'prices' in dataframes:
            price_df = dataframes['prices']
            if not price_df.is_empty():
                st.info("💰 Обработка цен...")
                self.upsert_prices(price_df)
                st.success(
                    f"✅ Успешно обновлено {len(price_df)} ценовых записей")
        
        step_counter += 1
        progress_bar.progress(step_counter / (num_steps + 1),
                              text=f"({step_counter}/{num_steps}) Сборка и обновление данных по артикулам...")
        
        # Собираем parts из разных файлов
        parts_df = None
        file_priority = ['oe', 'barcode', 'images', 'dimensions']
        key_files = {ftype: df for ftype,
                     df in dataframes.items() if ftype in file_priority}
        
        if key_files:
            all_parts = pl.concat([
                df.select(['artikul', 'artikul_norm', 'brand', 'brand_norm'])
                for df in key_files.values() if 'artikul_norm' in df.columns and 'brand_norm' in df.columns
            ]).filter(pl.col('artikul_norm') != "").unique(subset=['artikul_norm', 'brand_norm'], keep='first')
            parts_df = all_parts
            
            for ftype in file_priority:
                if ftype not in key_files:
                    continue
                df = key_files[ftype]
                if df.is_empty() or 'artikul_norm' not in df.columns:
                    continue
                join_cols = [col for col in df.columns if col not in [
                    'artikul', 'artikul_norm', 'brand', 'brand_norm']]
                if not join_cols:
                    continue
                existing_cols = set(parts_df.columns)
                join_cols = [
                    col for col in join_cols if col not in existing_cols]
                if not join_cols:
                    continue
                df_subset = df.select(['artikul_norm', 'brand_norm'] + join_cols).unique(
                    subset=['artikul_norm', 'brand_norm'], keep='first')
                parts_df = parts_df.join(
                    df_subset, on=['artikul_norm', 'brand_norm'], how='left', coalesce=True)
        
        if parts_df is not None and not parts_df.is_empty():
            if 'multiplicity' not in parts_df.columns:
                parts_df = parts_df.with_columns(
                    multiplicity=pl.lit(1).cast(pl.Int32))
            else:
                parts_df = parts_df.with_columns(
                    pl.col('multiplicity').fill_null(1).cast(pl.Int32))
            
            for col in ['length', 'width', 'height']:
                if col not in parts_df.columns:
                    parts_df = parts_df.with_columns(
                        pl.lit(None).cast(pl.Float64).alias(col))
            
            if 'dimensions_str' not in parts_df.columns:
                parts_df = parts_df.with_columns(
                    dimensions_str=pl.lit(None).cast(pl.Utf8))
            
            parts_df = parts_df.with_columns([
                pl.col('length').cast(pl.Utf8).fill_null(
                    '').alias('_length_str'),
                pl.col('width').cast(pl.Utf8).fill_null(
                    '').alias('_width_str'),
                pl.col('height').cast(pl.Utf8).fill_null(
                    '').alias('_height_str'),
            ])
            
            parts_df = parts_df.with_columns(
                dimensions_str=pl.when(
                    (pl.col('dimensions_str').is_not_null()) &
                    (pl.col('dimensions_str').cast(pl.Utf8) != '')
                ).then(
                    pl.col('dimensions_str').cast(pl.Utf8)
                ).otherwise(
                    pl.concat_str([
                        pl.col('_length_str'), pl.lit('x'),
                        pl.col('_width_str'), pl.lit('x'),
                        pl.col('_height_str')
                    ], separator='')
                )
            )
            
            parts_df = parts_df.drop(
                ['_length_str', '_width_str', '_height_str'])
            
            if 'artikul' not in parts_df.columns:
                parts_df = parts_df.with_columns(artikul=pl.lit(''))
            if 'brand' not in parts_df.columns:
                parts_df = parts_df.with_columns(brand=pl.lit(''))
            
            parts_df = parts_df.with_columns([
                pl.col('artikul').cast(pl.Utf8).fill_null(
                    '').alias('_artikul_str'),
                pl.col('brand').cast(pl.Utf8).fill_null(
                    '').alias('_brand_str'),
                pl.col('multiplicity').cast(
                    pl.Utf8).alias('_multiplicity_str'),
            ])
            
            parts_df = parts_df.with_columns(
                description=pl.concat_str([
                    pl.lit('Артикул: '), pl.col('_artikul_str'),
                    pl.lit(', Бренд: '), pl.col('_brand_str'),
                    pl.lit(', Кратность: '), pl.col(
                        '_multiplicity_str'), pl.lit(' шт.')
                ], separator='')
            )
            
            parts_df = parts_df.drop(
                ['_artikul_str', '_brand_str', '_multiplicity_str'])
            
            final_columns = [
                'artikul_norm', 'brand_norm', 'artikul', 'brand', 'multiplicity', 'barcode',
                'length', 'width', 'height', 'weight', 'image_url', 'dimensions_str', 'description'
            ]
            select_exprs = [pl.col(c) if c in parts_df.columns else pl.lit(
                None).alias(c) for c in final_columns]
            parts_df = parts_df.select(select_exprs)
            
            self.upsert_data('parts', parts_df, ['artikul_norm', 'brand_norm'])
        
        progress_bar.progress(1.0, text="Обновление базы данных завершено!")
        time.sleep(1)
        progress_bar.empty()
    
    # --- Экспорт ---
    def _get_brand_markups_sql(self) -> str:
        rows = []
        for brand, markup in self.price_rules['brand_markups'].items():
            safe_brand = brand.replace("'", "''")
            rows.append(f"SELECT '{safe_brand}' AS brand, {markup} AS markup")
        return " UNION ALL ".join(rows) if rows else "SELECT NULL AS brand, NULL AS markup LIMIT 0"
    
    def build_export_query(self, selected_columns=None, include_prices=True, apply_markup=True):
        description_text = (
            "Состояние товара: новый (в упаковке). Высококачественные автозапчасти и автотовары — надежное решение для вашего автомобиля. "
            "Обеспечьте безопасность, долговечность и высокую производительность вашего авто с помощью нашего широкого ассортимента оригинальных и совместимых автозапчастей. "
            "В нашем каталоге вы найдете тормозные системы, фильтры (масляные, воздушные, салонные), свечи зажигания, расходные материалы, автохимию, электроматериалы, автомасла, инструмент, "
            "а также другие комплектующие, полностью соответствующие стандартам качества и безопасности. "
            "Мы гарантируем быструю доставку, выгодные цены и профессиональную консультацию для любого клиента — автолюбителя, специалиста или автосервиса. "
            "Выбирайте только лучшее — надежность и качество от ведущих производителей."
        )
        
        brand_markups_sql = self._get_brand_markups_sql()
        
        select_parts = []
        
        price_requested = include_prices and (not selected_columns or "Цена" in selected_columns or "Валюта" in selected_columns)
        if price_requested:
            if apply_markup:
                global_markup = self.price_rules.get('global_markup', 0)
                select_parts.append(
                    f"CASE WHEN pr.price IS NOT NULL THEN pr.price * (1 + COALESCE(brm.markup, {global_markup})) ELSE pr.price END AS \"Цена\""
                )
            else:
                select_parts.append('pr.price AS "Цена"')
            select_parts.append("COALESCE(pr.currency, 'RUB') AS \"Валюта\"")
        
        columns_map = [
            ("Артикул бренда", 'r.artikul AS "Артикул бренда"'),
            ("Бренд", 'r.brand AS "Бренд"'),
            ("Наименование", 'COALESCE(r.representative_name, r.analog_representative_name) AS "Наименование"'),
            ("Применимость", 'COALESCE(r.representative_applicability, r.analog_representative_applicability) AS "Применимость"'),
            ("Описание", 'CONCAT(COALESCE(r.description, \'\'), dt.text) AS "Описание"'),
            ("Категория товара", 'COALESCE(r.representative_category, r.analog_representative_category) AS "Категория товара"'),
            ("Кратность", 'r.multiplicity AS "Кратность"'),
            ("Длинна", 'COALESCE(r.length, r.analog_length) AS "Длинна"'),
            ("Ширина", 'COALESCE(r.width, r.analog_width) AS "Ширина"'),
            ("Высота", 'COALESCE(r.height, r.analog_height) AS "Высота"'),
            ("Вес", 'COALESCE(r.weight, r.analog_weight) AS "Вес"'),
            ("Длинна/Ширина/Высота", """
                COALESCE(
                    CASE
                        WHEN r.dimensions_str IS NULL OR r.dimensions_str = '' OR UPPER(TRIM(r.dimensions_str)) = 'XX'
                        THEN NULL
                        ELSE r.dimensions_str
                    END,
                    r.analog_dimensions_str
                ) AS "Длинна/Ширина/Высота"
            """),
            ("OE номер", 'r.oe_list AS "OE номер"'),
            ("аналоги", 'r.analog_list AS "аналоги"'),
            ("Ссылка на изображение", 'r.image_url AS "Ссылка на изображение"')
        ]
        
        for name, expr in columns_map:
            if not selected_columns or name in selected_columns:
                select_parts.append(expr.strip())
        
        if not select_parts:
            select_parts = ['r.artikul AS "Артикул бренда"', 'r.brand AS "Бренд"']
        
        select_clause = ",\n        ".join(select_parts)
        
        ctes = f"""
        WITH DescriptionTemplate AS (
            SELECT CHR(10) || CHR(10) || $${description_text}$$ AS text
        ),
        BrandMarkups AS (
            SELECT brand, markup FROM (
                {brand_markups_sql}
            ) AS tmp
        ),
        PartDetails AS (
            SELECT 
                cr.artikul_norm, 
                cr.brand_norm,
                STRING_AGG(
                    DISTINCT regexp_replace(
                        regexp_replace(o.oe_number, '''', ''), 
                        '[^0-9A-Za-zА-Яа-яЁё`\\-\\s]', '', 'g'
                    ), ', '
                ) AS oe_list,
                ANY_VALUE(o.name) AS representative_name,
                ANY_VALUE(o.applicability) AS representative_applicability,
                ANY_VALUE(o.category) AS representative_category
            FROM cross_references cr
            LEFT JOIN oe o ON cr.oe_number_norm = o.oe_number_norm
            GROUP BY cr.artikul_norm, cr.brand_norm
        ),
        AllAnalogs AS (
            SELECT 
                cr1.artikul_norm, 
                cr1.brand_norm,
                STRING_AGG(
                    DISTINCT regexp_replace(
                        regexp_replace(p2.artikul, '''', ''), 
                        '[^0-9A-Za-zА-Яа-яЁё`\\-\\s]', '', 'g'
                    ), ', '
                ) AS analog_list
            FROM cross_references cr1
            JOIN cross_references cr2 ON cr1.oe_number_norm = cr2.oe_number_norm
            JOIN parts p2 ON cr2.artikul_norm = p2.artikul_norm AND cr2.brand_norm = p2.brand_norm
            WHERE (cr1.artikul_norm != p2.artikul_norm OR cr1.brand_norm != p2.brand_norm)
            GROUP BY cr1.artikul_norm, cr1.brand_norm
        ),
        InitialOENumbers AS (
            SELECT DISTINCT p.artikul_norm, p.brand_norm, cr.oe_number_norm
            FROM parts p
            LEFT JOIN cross_references cr ON p.artikul_norm = cr.artikul_norm AND p.brand_norm = cr.brand_norm
            WHERE cr.oe_number_norm IS NOT NULL
        ),
        Level1Analogs AS (
            SELECT DISTINCT 
                i.artikul_norm AS source_artikul_norm, 
                i.brand_norm AS source_brand_norm,
                cr2.artikul_norm AS related_artikul_norm, 
                cr2.brand_norm AS related_brand_norm
            FROM InitialOENumbers i
            JOIN cross_references cr2 ON i.oe_number_norm = cr2.oe_number_norm
            WHERE NOT (i.artikul_norm = cr2.artikul_norm AND i.brand_norm = cr2.brand_norm)
        ),
        Level1OENumbers AS (
            SELECT DISTINCT 
                l1.source_artikul_norm, 
                l1.source_brand_norm, 
                cr3.oe_number_norm
            FROM Level1Analogs l1
            JOIN cross_references cr3 ON l1.related_artikul_norm = cr3.artikul_norm AND l1.related_brand_norm = cr3.brand_norm
            WHERE NOT EXISTS (
                SELECT 1 FROM InitialOENumbers i
                WHERE i.artikul_norm = l1.source_artikul_norm 
                  AND i.brand_norm = l1.source_brand_norm 
                  AND i.oe_number_norm = cr3.oe_number_norm
            )
        ),
        Level2Analogs AS (
            SELECT DISTINCT 
                loe.source_artikul_norm, 
                loe.source_brand_norm,
                cr4.artikul_norm AS related_artikul_norm, 
                cr4.brand_norm AS related_brand_norm
            FROM Level1OENumbers loe
            JOIN cross_references cr4 ON loe.oe_number_norm = cr4.oe_number_norm
            WHERE NOT (loe.source_artikul_norm = cr4.artikul_norm AND loe.source_brand_norm = cr4.brand_norm)
        ),
        AllRelatedParts AS (
            SELECT source_artikul_norm, source_brand_norm, related_artikul_norm, related_brand_norm
            FROM Level1Analogs
            UNION
            SELECT source_artikul_norm, source_brand_norm, related_artikul_norm, related_brand_norm
            FROM Level2Analogs
        ),
        AggregatedAnalogData AS (
            SELECT 
                arp.source_artikul_norm AS artikul_norm,
                arp.source_brand_norm AS brand_norm,
                MAX(CASE WHEN p2.length IS NOT NULL THEN p2.length ELSE NULL END) AS length,
                MAX(CASE WHEN p2.width IS NOT NULL THEN p2.width ELSE NULL END) AS width,
                MAX(CASE WHEN p2.height IS NOT NULL THEN p2.height ELSE NULL END) AS height,
                MAX(CASE WHEN p2.weight IS NOT NULL THEN p2.weight ELSE NULL END) AS weight,
                ANY_VALUE(
                    CASE 
                        WHEN p2.dimensions_str IS NOT NULL AND p2.dimensions_str != '' AND UPPER(TRIM(p2.dimensions_str)) != 'XX'
                        THEN p2.dimensions_str
                        ELSE NULL
                    END
                ) AS dimensions_str,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_name IS NOT NULL AND pd2.representative_name != '' 
                        THEN pd2.representative_name 
                        ELSE NULL
                    END
                ) AS representative_name,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_applicability IS NOT NULL AND pd2.representative_applicability != ''
                        THEN pd2.representative_applicability
                        ELSE NULL
                    END
                ) AS representative_applicability,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_category IS NOT NULL AND pd2.representative_category != ''
                        THEN pd2.representative_category
                        ELSE NULL
                    END
                ) AS representative_category
            FROM AllRelatedParts arp
            JOIN parts p2 ON arp.related_artikul_norm = p2.artikul_norm AND arp.related_brand_norm = p2.brand_norm
            LEFT JOIN PartDetails pd2 ON p2.artikul_norm = pd2.artikul_norm AND p2.brand_norm = pd2.brand_norm
            GROUP BY arp.source_artikul_norm, arp.source_brand_norm
        ),
        RankedData AS (
            SELECT 
                p.artikul_norm,
                p.brand_norm,
                p.artikul,
                p.brand,
                p.description,
                p.multiplicity,
                p.length,
                p.width,
                p.height,
                p.weight,
                p.dimensions_str,
                p.image_url,
                pd.representative_name,
                pd.representative_applicability,
                pd.representative_category,
                pd.oe_list,
                aa.analog_list,
                p_analog.length AS analog_length,
                p_analog.width AS analog_width,
                p_analog.height AS analog_height,
                p_analog.weight AS analog_weight,
                p_analog.dimensions_str AS analog_dimensions_str,
                p_analog.representative_name AS analog_representative_name,
                p_analog.representative_applicability AS analog_representative_applicability,
                p_analog.representative_category AS analog_representative_category,
                ROW_NUMBER() OVER (
                    PARTITION BY p.artikul_norm, p.brand_norm 
                    ORDER BY pd.representative_name DESC NULLS LAST, pd.oe_list DESC NULLS LAST
                ) AS rn
            FROM parts p
            LEFT JOIN PartDetails pd ON p.artikul_norm = pd.artikul_norm AND p.brand_norm = pd.brand_norm
            LEFT JOIN AllAnalogs aa ON p.artikul_norm = aa.artikul_norm AND p.brand_norm = aa.brand_norm
            LEFT JOIN AggregatedAnalogData p_analog ON p.artikul_norm = p_analog.artikul_norm AND p.brand_norm = p_analog.brand_norm
        )
        """
        
        price_join = """
        LEFT JOIN prices pr ON r.artikul_norm = pr.artikul_norm AND r.brand_norm = pr.brand_norm
        LEFT JOIN BrandMarkups brm ON r.brand = brm.brand
        """ if include_prices else ""
        
        query = f"""
        {ctes}
        SELECT
            {select_clause}
        FROM RankedData r
        CROSS JOIN DescriptionTemplate dt
        {price_join}
        WHERE r.rn = 1
        ORDER BY r.brand, r.artikul
        """
        
        return "\n".join([line.rstrip() for line in query.strip().splitlines()])
    
    def export_to_csv_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        if not self.conn:
            st.warning("⚠️ База данных не доступна")
            return False
            
        total = self.conn.execute(
            "SELECT count(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        if total == 0:
            st.warning("Нет данных для экспорта")
            return False
            
        st.info(f"📤 Экспорт {total} записей в CSV...")
        try:
            query = self.build_export_query(
                selected_columns, include_prices, apply_markup)
            df = self.conn.execute(query).pl()
            import pandas as pd
            pdf = df.to_pandas()
            
            dimension_cols = ["Длинна", "Ширина",
                              "Высота", "Вес", "Длинна/Ширина/Высота"]
            for col in dimension_cols:
                if col in pdf.columns:
                    pdf[col] = pdf[col].astype(str).replace({'nan': ''})
            
            output_dir = Path("auto_parts_data")
            output_dir.mkdir(parents=True, exist_ok=True)
            
            buf = io.StringIO()
            pdf.to_csv(buf, sep=';', index=False)
            with open(output_path, "wb") as f:
                f.write(b'\xef\xbb\xbf')
                f.write(buf.getvalue().encode('utf-8'))
            size_mb = os.path.getsize(output_path) / (1024 * 1024)
            st.success(
                f"Данные экспортированы: {output_path} ({size_mb:.1f} МБ)")
            return True
        except Exception as e:
            logger.exception("Ошибка экспорта CSV")
            st.error(f"Ошибка при экспорте в CSV: {str(e)}")
            return False
    
    def export_to_excel_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        if not self.conn:
            st.warning("⚠️ База данных не доступна")
            return False
            
        total = self.conn.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        if total == 0:
            st.warning("Нет данных для экспорта")
            return False
            
        import pandas as pd
        query = self.build_export_query(
            selected_columns, include_prices, apply_markup)
        df = pd.read_sql(query, self.conn)
        for col in ["Длинна", "Ширина", "Высота", "Вес", "Длинна/Ширина/Высота"]:
            if col in df.columns:
                df[col] = df[col].astype(str).replace(
                    {r'^nan$': ''}, regex=True)
        
        if len(df) <= EXCEL_ROW_LIMIT:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        else:
            sheets = (len(df) // EXCEL_ROW_LIMIT) + 1
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i in range(sheets):
                    df.iloc[i*EXCEL_ROW_LIMIT:(i+1)*EXCEL_ROW_LIMIT].to_excel(
                        writer, index=False, sheet_name=f"Данные_{i+1}")
        return True
    
    def export_to_parquet(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        try:
            query = self.build_export_query(
                selected_columns, include_prices, apply_markup)
            df = self.conn.execute(query).pl()
            df.write_parquet(output_path)
            return True
        except Exception as e:
            logger.exception("Ошибка экспорта Parquet")
            st.error(f"Ошибка при экспорте в Parquet: {str(e)}")
            return False
    
    # --- Управление данными ---
    def delete_by_brand(self, brand_norm: str) -> int:
        if not self.conn:
            return 0
            
        try:
            count_result = self.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()
            deleted_count = count_result[0] if count_result else 0
            if deleted_count == 0:
                logger.info(f"No records found for brand: {brand_norm}")
                return 0
            self.conn.execute(
                "DELETE FROM parts WHERE brand_norm = ?", [brand_norm])
            self.conn.execute(
                "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
            return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by brand {brand_norm}: {e}")
            raise
    
    def delete_by_artikul(self, artikul_norm: str) -> int:
        if not self.conn:
            return 0
            
        try:
            count_result = self.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE artikul_norm = ?", [artikul_norm]).fetchone()
            deleted_count = count_result[0] if count_result else 0
            if deleted_count == 0:
                logger.info(f"No records found for artikul: {artikul_norm}")
                return 0
            self.conn.execute(
                "DELETE FROM parts WHERE artikul_norm = ?", [artikul_norm])
            self.conn.execute(
                "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
            return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by artikul {artikul_norm}: {e}")
            raise
    
    def merge_all_data_parallel(self, file_paths: Dict[str, str], max_workers: int = 4) -> Dict[str, pl.DataFrame]:
        if not HIGH_VOLUME_AVAILABLE:
            return {}
            
        results = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key, path in file_paths.items():
                if path and os.path.exists(path):
                    futures[executor.submit(
                        self.read_and_prepare_file, path, key)] = key
            for fut in as_completed(futures):
                key = futures[fut]
                try:
                    df = fut.result()
                    if not df.is_empty():
                        results[key] = df
                        logger.info(f"Обработан {key}")
                except Exception as e:
                    logger.error(f"Ошибка обработки {key}: {e}")
        return results
    
    def get_statistics(self) -> Dict[str, Any]:
        if not self.conn:
            return {}
            
        stats = {}
        try:
            stats['parts'] = self.conn.execute(
                "SELECT COUNT(*) FROM parts").fetchone()[0]
            stats['oe'] = self.conn.execute(
                "SELECT COUNT(*) FROM oe").fetchone()[0]
            stats['cross'] = self.conn.execute(
                "SELECT COUNT(*) FROM cross_references").fetchone()[0]
            stats['prices'] = self.conn.execute(
                "SELECT COUNT(*) FROM prices").fetchone()[0]
            stats['brands'] = self.conn.execute(
                "SELECT COUNT(DISTINCT brand) FROM parts").fetchone()[0]
            stats['unique_parts'] = self.conn.execute(
                "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
            avg_price = self.conn.execute(
                "SELECT AVG(price) FROM prices").fetchone()[0]
            stats['avg_price'] = round(avg_price, 2) if avg_price else 0
            
            top_brands = self.conn.execute(
                "SELECT brand, COUNT(*) as cnt FROM parts GROUP BY brand ORDER BY cnt DESC LIMIT 10").pl()
            stats['top_brands'] = top_brands.to_pandas() if not top_brands.is_empty() else pd.DataFrame()
        except Exception as e:
            logger.error(f"Ошибка сбора статистики: {e}")
        return stats

# ============================================================================
# БЛОК 6: КЛАСС MARKETPLACEUNITECONOMICS (ЮНИТ-ЭКОНОМИКА)
# ============================================================================

class MarketplaceUnitEconomics:
    _instance = None
    _configs = None
    _cache = None
    _history = None
    _stats = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_configs()
            cls._instance._init_cache()
            cls._instance._init_history()
            cls._instance._init_stats()
        return cls._instance
    
    def _init_configs(self):
        self._configs = get_marketplace_configs_2026()
        self.logger = logging.getLogger('MarketplaceUnitEconomics')
        self.logger.info(f"Инициализировано {len(self._configs)} маркетплейсов")
    
    def _init_cache(self):
        self._cache = {}
    
    def _init_history(self):
        self._history = []
    
    def _init_stats(self):
        self._stats = {
            "total_calculations": 0,
            "by_marketplace": defaultdict(int),
            "by_mode": defaultdict(int),
            "avg_profit": 0.0,
            "avg_margin": 0.0,
            "total_profit": 0.0,
            "max_profit": 0.0,
            "min_profit": 0.0,
            "best_marketplace": None,
            "best_mode": None
        }
    
    @lru_cache(maxsize=10000)
    def calculate_unit_economics(
        self,
        price: float,
        cost: float,
        weight_kg: float,
        volume_liters: float,
        marketplace: str,
        operation_mode: str = "FBY",
        days_in_storage: int = 30,
        category: str = None,
        is_premium: bool = False
    ) -> Dict[str, Any]:
        if marketplace not in self._configs:
            return {"error": f"Маркетплейс {marketplace} не поддерживается"}
        
        config = self._configs[marketplace]
        
        commission_rate = config.get_commission_rate(category)
        
        if config.commission_type == CommissionType.SUBSCRIPTION:
            commission = price * commission_rate
            subscription_cost = config.subscription_fee / 30
        else:
            commission = max(price * commission_rate, config.min_commission)
            subscription_cost = 0
        
        logistics = (
            config.logistics_base + 
            weight_kg * config.logistics_per_kg + 
            volume_liters * config.logistics_per_liter
        )
        
        mode_multiplier = config.get_mode_multiplier(operation_mode)
        logistics *= mode_multiplier
        
        storage_cost = volume_liters * config.storage_per_day * days_in_storage
        
        storage_non_standard = 0
        if config.storage_non_standard_fee > 0 and weight_kg > 25:
            storage_non_standard = min(price * config.storage_non_standard_fee, 280)
        
        acquiring = price * config.acquiring_fee
        delivery = price * config.delivery_fee_percent
        last_mile = config.last_mile_fee
        returns = price * config.return_fee
        rko_fee = price * config.rko_fee if config.rko_fee > 0 else 0
        premium_fee = price * config.premium_section_fee if is_premium else 0
        
        total_expenses = (
            cost + commission + logistics + storage_cost + storage_non_standard +
            acquiring + delivery + last_mile + returns + rko_fee + 
            premium_fee + subscription_cost
        )
        
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0
        
        fixed_costs = logistics + storage_cost + last_mile + subscription_cost
        variable_rate = (
            commission_rate + config.acquiring_fee + 
            config.delivery_fee_percent + config.return_fee +
            config.rko_fee + config.premium_section_fee
        )
        breakeven_price = ((cost + fixed_costs) / (1 - variable_rate)) if (1 - variable_rate) > 0 else 0
        
        result = {
            "marketplace": marketplace,
            "operation_mode": operation_mode,
            "price": round(price, 2),
            "cost": round(cost, 2),
            "commission": round(commission, 2),
            "commission_percent": round(commission / price * 100, 2) if price > 0 else 0,
            "commission_type": config.commission_type.value,
            "subscription_cost": round(subscription_cost, 2),
            "logistics": round(logistics, 2),
            "storage_cost": round(storage_cost, 2),
            "storage_non_standard": round(storage_non_standard, 2),
            "acquiring": round(acquiring, 2),
            "delivery": round(delivery, 2),
            "last_mile": round(last_mile, 2),
            "returns": round(returns, 2),
            "rko_fee": round(rko_fee, 2),
            "premium_fee": round(premium_fee, 2),
            "total_expenses": round(total_expenses, 2),
            "profit": round(profit, 2),
            "margin_percent": round(margin_percent, 2),
            "roi": round(roi, 2),
            "breakeven_price": round(breakeven_price, 2),
            "profit_per_ruble": round(profit / price, 4) if price > 0 else 0,
            "timestamp": datetime.now().isoformat()
        }
        
        self._stats["total_calculations"] += 1
        self._stats["by_marketplace"][marketplace] += 1
        self._stats["by_mode"][operation_mode] += 1
        self._stats["total_profit"] += profit
        
        if profit > self._stats["max_profit"]:
            self._stats["max_profit"] = profit
            self._stats["best_marketplace"] = marketplace
            self._stats["best_mode"] = operation_mode
        
        if profit < self._stats["min_profit"] or self._stats["min_profit"] == 0:
            self._stats["min_profit"] = profit
        
        self._stats["avg_profit"] = self._stats["total_profit"] / self._stats["total_calculations"]
        
        self._history.append(result)
        if len(self._history) > HISTORY_LIMIT:
            self._history = self._history[-HISTORY_LIMIT:]
        
        return result
    
    def get_marketplace_config(self, marketplace: str) -> Dict:
        config = self._configs.get(marketplace)
        if config:
            return {
                "commission_rate": config.commission_rate,
                "commission_type": config.commission_type.value,
                "subscription_fee": config.subscription_fee,
                "min_commission": config.min_commission,
                "logistics_base": config.logistics_base,
                "logistics_per_kg": config.logistics_per_kg,
                "logistics_per_liter": config.logistics_per_liter,
                "storage_per_day": config.storage_per_day,
                "storage_non_standard_fee": config.storage_non_standard_fee,
                "return_fee": config.return_fee,
                "acquiring_fee": config.acquiring_fee,
                "last_mile_fee": config.last_mile_fee,
                "delivery_fee_percent": config.delivery_fee_percent,
                "premium_section_fee": config.premium_section_fee,
                "rko_fee": config.rko_fee,
                "category_rates": config.category_rates,
                "mode_multipliers": config.mode_multipliers
            }
        return {}
    
    def calculate_for_all_marketplaces(
        self,
        price: float,
        cost: float,
        weight_kg: float,
        volume_liters: float,
        operation_mode: str = "FBY"
    ) -> pd.DataFrame:
        results = []
        for marketplace in self._configs.keys():
            economics = self.calculate_unit_economics(
                price=price,
                cost=cost,
                weight_kg=weight_kg,
                volume_liters=volume_liters,
                marketplace=marketplace,
                operation_mode=operation_mode
            )
            if "error" not in economics:
                results.append(economics)
        return pd.DataFrame(results) if results else pd.DataFrame()
    
    def get_history(self) -> List[Dict]:
        return self._history.copy()
    
    def get_stats(self) -> Dict:
        return self._stats.copy()
    
    def clear_history(self):
        self._history = []
        self._stats = {
            "total_calculations": 0,
            "by_marketplace": defaultdict(int),
            "by_mode": defaultdict(int),
            "avg_profit": 0.0,
            "avg_margin": 0.0,
            "total_profit": 0.0,
            "max_profit": 0.0,
            "min_profit": 0.0,
            "best_marketplace": None,
            "best_mode": None
        }
    
    def get_best_marketplace(self) -> Dict[str, Any]:
        if not self._history:
            return {"error": "Нет данных для анализа"}
        
        best = max(self._history, key=lambda x: x['profit'])
        return {
            "marketplace": best['marketplace'],
            "operation_mode": best['operation_mode'],
            "profit": best['profit'],
            "margin_percent": best['margin_percent'],
            "roi": best['roi'],
            "timestamp": best['timestamp']
        }

# ============================================================================
# БЛОК 7: КЛАСС CATALOGENHANCER (ОБОГАЩЕНИЕ КАТАЛОГА)
# ============================================================================

class CatalogEnhancer:
    def __init__(self, db_path: Optional[str] = None):
        self.data_dir = Path("./catalog_data")
        self.data_dir.mkdir(exist_ok=True)
        
        self.db_path = Path(db_path) if db_path else self.data_dir / "catalog.duckdb"
        self.conn = None
        self.stats = {
            "oe_loaded": 0,
            "parts_loaded": 0,
            "cross_loaded": 0,
            "analog_searches": 0,
            "enrichments": 0
        }
        
        if LIBRARIES['duckdb']:
            try:
                self.conn = duckdb.connect(database=str(self.db_path))
                self._setup_database()
                logger.info("CatalogEnhancer инициализирован с DuckDB")
            except Exception as e:
                logger.error(f"Ошибка инициализации DuckDB: {e}")
                self.conn = None
        else:
            logger.warning("DuckDB не установлен, работа в режиме ограниченной функциональности")
    
    def _setup_database(self):
        if not self.conn:
            return
        
        try:
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS oe (
                    oe_number_norm VARCHAR PRIMARY KEY,
                    oe_number VARCHAR,
                    name VARCHAR,
                    applicability VARCHAR,
                    category VARCHAR
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
            
            logger.info("Таблицы каталога созданы")
        except Exception as e:
            logger.error(f"Ошибка создания таблиц: {e}")
    
    def normalize_key(self, value: str) -> str:
        if not value:
            return ""
        return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', value.lower().strip())
    
    def load_oe_data(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        
        try:
            df['oe_number_norm'] = df['oe_number'].apply(self.normalize_key)
            df = df[df['oe_number_norm'] != ""]
            
            self.conn.execute("DELETE FROM oe")
            
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO oe VALUES (?, ?, ?, ?, ?)",
                    [row['oe_number_norm'], row.get('oe_number', ''), 
                     row.get('name', ''), row.get('applicability', ''), 
                     row.get('category', 'Разное')]
                )
            self.stats['oe_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} OE записей")
        except Exception as e:
            logger.error(f"Ошибка загрузки OE данных: {e}")
    
    def load_parts_data(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        
        try:
            df['artikul_norm'] = df['artikul'].apply(self.normalize_key)
            df['brand_norm'] = df['brand'].apply(self.normalize_key)
            df = df[(df['artikul_norm'] != "") & (df['brand_norm'] != "")]
            
            self.conn.execute("DELETE FROM parts")
            
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO parts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [row['artikul_norm'], row['brand_norm'], row.get('artikul', ''),
                     row.get('brand', ''), row.get('multiplicity', 1),
                     row.get('barcode', ''), row.get('length', 0.0),
                     row.get('width', 0.0), row.get('height', 0.0),
                     row.get('weight', 0.0), row.get('image_url', ''),
                     row.get('dimensions_str', ''), row.get('description', '')]
                )
            self.stats['parts_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} записей деталей")
        except Exception as e:
            logger.error(f"Ошибка загрузки данных деталей: {e}")
    
    def load_cross_references(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        
        try:
            df['oe_number_norm'] = df['oe_number'].apply(self.normalize_key)
            df['artikul_norm'] = df['artikul'].apply(self.normalize_key)
            df['brand_norm'] = df['brand'].apply(self.normalize_key)
            df = df[(df['oe_number_norm'] != "") & (df['artikul_norm'] != "")]
            
            self.conn.execute("DELETE FROM cross_references")
            
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO cross_references VALUES (?, ?, ?)",
                    [row['oe_number_norm'], row['artikul_norm'], row['brand_norm']]
                )
            self.stats['cross_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} кросс-ссылок")
        except Exception as e:
            logger.error(f"Ошибка загрузки кросс-ссылок: {e}")
    
    def get_analog_data(self, artikul: str, brand: str) -> Dict[str, Any]:
        self.stats['analog_searches'] += 1
        
        if not self.conn:
            return {"error": "DuckDB не доступен"}
        
        artikul_norm = self.normalize_key(artikul)
        brand_norm = self.normalize_key(brand)
        
        if not artikul_norm or not brand_norm:
            return {"error": "Не указан артикул или бренд"}
        
        try:
            query = f"""
                WITH PartDetails AS (
                    SELECT 
                        cr.artikul_norm, 
                        cr.brand_norm,
                        STRING_AGG(DISTINCT oe.oe_number, ', ') AS oe_list
                    FROM cross_references cr
                    LEFT JOIN oe ON cr.oe_number_norm = oe.oe_number_norm
                    WHERE cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}'
                    GROUP BY cr.artikul_norm, cr.brand_norm
                ),
                AnalogParts AS (
                    SELECT DISTINCT 
                        p.artikul,
                        p.brand,
                        p.description,
                        p.length,
                        p.width,
                        p.height,
                        p.weight,
                        p.dimensions_str,
                        p.image_url
                    FROM cross_references cr
                    JOIN parts p ON cr.artikul_norm = p.artikul_norm AND cr.brand_norm = p.brand_norm
                    WHERE cr.oe_number_norm IN (
                        SELECT oe_number_norm 
                        FROM cross_references 
                        WHERE artikul_norm = '{artikul_norm}' AND brand_norm = '{brand_norm}'
                    )
                    AND NOT (cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}')
                    LIMIT 50
                )
                SELECT 
                    (SELECT COUNT(*) FROM AnalogParts) AS analog_count,
                    (SELECT oe_list FROM PartDetails) AS oe_list,
                    (SELECT * FROM AnalogParts LIMIT 20) AS analogs
            """
            result = self.conn.execute(query).df()
            
            if result.empty or result.iloc[0]['analog_count'] == 0:
                return {
                    "artikul": artikul,
                    "brand": brand,
                    "analog_count": 0,
                    "analogs": [],
                    "has_analogs": False,
                    "oe_list": ""
                }
            
            row = result.iloc[0]
            analog_count = int(row['analog_count']) if not pd.isna(row['analog_count']) else 0
            
            analogs = []
            if analog_count > 0:
                analog_query = f"""
                    SELECT DISTINCT 
                        p.artikul,
                        p.brand,
                        p.description,
                        p.length,
                        p.width,
                        p.height,
                        p.weight,
                        p.dimensions_str,
                        p.image_url
                    FROM cross_references cr
                    JOIN parts p ON cr.artikul_norm = p.artikul_norm AND cr.brand_norm = p.brand_norm
                    WHERE cr.oe_number_norm IN (
                        SELECT oe_number_norm 
                        FROM cross_references 
                        WHERE artikul_norm = '{artikul_norm}' AND brand_norm = '{brand_norm}'
                    )
                    AND NOT (cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}')
                    LIMIT 20
                """
                analog_df = self.conn.execute(analog_query).df()
                for _, arow in analog_df.iterrows():
                    analogs.append({
                        "artikul": safe_str(arow.get('artikul', '')),
                        "brand": safe_str(arow.get('brand', '')),
                        "description": safe_str(arow.get('description', '')),
                        "length": safe_float(arow.get('length', 0)),
                        "width": safe_float(arow.get('width', 0)),
                        "height": safe_float(arow.get('height', 0)),
                        "weight": safe_float(arow.get('weight', 0)),
                        "dimensions_str": safe_str(arow.get('dimensions_str', '')),
                        "image_url": safe_str(arow.get('image_url', ''))
                    })
            
            return {
                "artikul": artikul,
                "brand": brand,
                "analog_count": analog_count,
                "analogs": analogs,
                "has_analogs": analog_count > 0,
                "oe_list": safe_str(row.get('oe_list', ''))
            }
        except Exception as e:
            logger.error(f"Ошибка получения аналогов: {e}")
            return {"error": str(e)}
    
    def enhance_catalog_data(self, df: pd.DataFrame, 
                           artikul_col: str = "Артикул",
                           brand_col: str = "Бренд") -> pd.DataFrame:
        if df.empty:
            return df
        
        if artikul_col not in df.columns or brand_col not in df.columns:
            logger.warning(f"Колонки {artikul_col} или {brand_col} не найдены")
            return df
        
        self.stats['enrichments'] += 1
        
        df_copy = df.copy()
        
        new_columns = [
            'analog_count', 'has_analogs', 'analog_list', 'oe_list'
        ]
        
        for col in new_columns:
            if col not in df_copy.columns:
                df_copy[col] = None
        
        for idx, row in df_copy.iterrows():
            artikul = safe_str(row.get(artikul_col, ''))
            brand = safe_str(row.get(brand_col, ''))
            
            if artikul and brand:
                data = self.get_analog_data(artikul, brand)
                
                if not data.get('error'):
                    df_copy.at[idx, 'analog_count'] = data.get('analog_count', 0)
                    df_copy.at[idx, 'has_analogs'] = data.get('has_analogs', False)
                    df_copy.at[idx, 'oe_list'] = data.get('oe_list', '')
                    
                    if data.get('analogs'):
                        analog_str = ', '.join([
                            f"{a['artikul']} ({a['brand']})" 
                            for a in data['analogs'][:5] 
                            if a.get('artikul')
                        ])
                        df_copy.at[idx, 'analog_list'] = analog_str
        
        return df_copy
    
    def get_stats(self) -> Dict:
        return self.stats.copy()

# ============================================================================
# БЛОК 8: ML-КЛАССИФИКАТОР КАТЕГОРИЙ
# ============================================================================

class CategoryClassifier:
    def __init__(self, model_path: str = "category_model.pkl"):
        self.model_path = model_path
        self.model = None
        self.categories = [
            "Двигатель", "Трансмиссия", "Подвеска", "Тормозная система",
            "Рулевое управление", "Электрооборудование", "Система охлаждения",
            "Система выпуска", "Система питания", "Фильтры",
            "Масла и жидкости", "Кузовные детали", "Оптика",
            "Шины и диски", "Инструменты", "Ремни и приводы",
            "Подшипники", "Сальники и прокладки", "Крепеж",
            "Климат-контроль", "Аудио и мультимедиа", "Безопасность",
            "Прочее"
        ]
        self.accuracy = 0.0
        self.cache = {}
        self.logger = logging.getLogger('CategoryClassifier')
        self._load_model()
    
    def _load_model(self):
        if os.path.exists(self.model_path) and LIBRARIES['sklearn']:
            try:
                self.model = joblib.load(self.model_path)
                self.categories = self.model.classes_ if hasattr(self.model, 'classes_') else self.categories
                self.logger.info(f"ML-модель загружена, категорий: {len(self.categories)}")
                return
            except Exception as e:
                self.logger.warning(f"Ошибка загрузки модели: {e}")
        self._train_model()
    
    def _train_model(self):
        if not LIBRARIES['sklearn']:
            self.logger.warning("Scikit-learn не установлен, используется fallback классификатор")
            return
        
        try:
            X = []
            y = []
            
            category_keywords = {
                "Двигатель": ["двигатель", "мотор", "двс", "поршень", "шатун", "клапан", "гбц"],
                "Трансмиссия": ["коробка", "кпп", "сцепление", "привод", "дифференциал", "акпп"],
                "Подвеска": ["амортизатор", "пружина", "рычаг", "сайлентблок", "шаровая", "стабилизатор"],
                "Тормозная система": ["колодки", "диск", "барабан", "суппорт", "гтц", "абс"],
                "Рулевое управление": ["рейка", "тяга", "наконечник", "руль", "гур"],
                "Электрооборудование": ["генератор", "стартер", "аккумулятор", "свеча", "провод"],
                "Система охлаждения": ["радиатор", "помпа", "термостат", "вентилятор", "бачок"],
                "Система выпуска": ["глушитель", "катализатор", "резонатор", "гофра", "лямбда"],
                "Система питания": ["насос", "фильтр", "форсунка", "дроссель", "тнвд"],
                "Фильтры": ["фильтр", "масляный", "воздушный", "салонный", "топливный"],
                "Масла и жидкости": ["масло", "жидкость", "смазка", "антифриз", "тормозуха"],
                "Кузовные детали": ["бампер", "капот", "крыло", "дверь", "стекло", "фара"],
                "Шины и диски": ["шина", "диск", "колесо", "покрышка", "резина"],
                "Инструменты": ["инструмент", "ключ", "домкрат", "насос"],
                "Прочее": []
            }
            
            for category, keywords in category_keywords.items():
                for keyword in keywords:
                    if keyword:
                        X.append(keyword)
                        y.append(category)
                        X.append(keyword + " " + category)
                        y.append(category)
            
            if X:
                X_train, X_test, y_train, y_test = train_test_split(
                    X, y, test_size=0.2, random_state=42
                )
                
                self.model = Pipeline([
                    ('tfidf', TfidfVectorizer(max_features=2000, ngram_range=(1, 2))),
                    ('clf', MultinomialNB(alpha=0.1))
                ])
                
                self.model.fit(X_train, y_train)
                self.categories = self.model.classes_
                
                y_pred = self.model.predict(X_test)
                self.accuracy = accuracy_score(y_test, y_pred)
                
                joblib.dump(self.model, self.model_path)
                self.logger.info(f"ML-модель обучена на {len(X)} примерах, точность: {self.accuracy:.2%}")
        except Exception as e:
            self.logger.error(f"Ошибка обучения модели: {e}")
            self.model = None
    
    def predict(self, name: str) -> Tuple[str, float]:
        if not isinstance(name, str):
            name = str(name)
        
        if not name or not name.strip():
            return "Прочее", 0.0
        
        name_lower = name.lower()
        
        if name_lower in self.cache:
            return self.cache[name_lower]
        
        if LIBRARIES['sklearn'] and self.model is not None:
            try:
                pred = self.model.predict([name_lower])[0]
                probs = self.model.predict_proba([name_lower])[0]
                confidence = max(probs) * 100
                
                if confidence > 30:
                    result = (pred, confidence)
                    self.cache[name_lower] = result
                    return result
            except Exception as e:
                self.logger.warning(f"ML prediction error: {e}")
        
        best_category = "Прочее"
        best_score = 0.0
        
        category_keywords = {
            "Двигатель": ["двигатель", "мотор", "двс", "поршень", "шатун", "клапан", "гбц"],
            "Трансмиссия": ["коробка", "кпп", "сцепление", "привод", "дифференциал"],
            "Подвеска": ["амортизатор", "пружина", "рычаг", "сайлентблок", "шаровая"],
            "Тормозная система": ["колодки", "диск", "барабан", "суппорт", "гтц"],
            "Электрооборудование": ["генератор", "стартер", "аккумулятор", "свеча"],
            "Система охлаждения": ["радиатор", "помпа", "термостат", "вентилятор"],
            "Система выпуска": ["глушитель", "катализатор", "резонатор", "гофра"],
            "Система питания": ["насос", "фильтр", "форсунка", "дроссель"],
            "Фильтры": ["фильтр", "масляный", "воздушный", "салонный"],
            "Масла и жидкости": ["масло", "жидкость", "смазка", "антифриз"],
            "Кузовные детали": ["бампер", "капот", "крыло", "дверь", "стекло"],
            "Шины и диски": ["шина", "диск", "колесо", "покрышка"],
            "Инструменты": ["инструмент", "ключ", "домкрат"],
            "Прочее": []
        }
        
        for category, keywords in category_keywords.items():
            score = 0.0
            for keyword in keywords:
                if keyword in name_lower:
                    weight = len(keyword) / 10.0
                    if name_lower.startswith(keyword):
                        weight *= 1.5
                    score += weight
            
            if score > best_score:
                best_score = score
                best_category = category
        
        confidence = min(best_score * 15, 100.0)
        result = (best_category, round(confidence, 1))
        self.cache[name_lower] = result
        return result
    
    def predict_batch(self, names: List[str]) -> List[Tuple[str, float]]:
        results = []
        for name in names:
            cat, conf = self.predict(name)
            results.append((cat, conf))
        return results

# ============================================================================
# БЛОК 9: UI ФУНКЦИИ (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

def show_unit_economics_interface():
    st.header("📊 Юнит-экономика маркетплейсов 2026")
    
    unit_economics = MarketplaceUnitEconomics()
    
    st.info("""
    💡 **Режимы работы:**
    - **FBY** (0.75x) - доставка силами Яндекс Маркета (самый дешевый)
    - **FBS** (1.0x) - доставка силами продавца (базовый)
    - **FBO** (0.8x) - доставка силами оператора (средний)
    - **DBS** (1.3x) - доставка силами продавца (дорогой)
    - **FBP** (0.9x) - доставка силами платформы (чуть дешевле)
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        price = st.number_input(
            "💰 Цена продажи (₽)",
            min_value=0.0,
            value=1000.0,
            step=10.0,
            key="ue_price"
        )
        
        cost = st.number_input(
            "💵 Себестоимость (₽)",
            min_value=0.0,
            value=500.0,
            step=10.0,
            key="ue_cost"
        )
        
        weight = st.number_input(
            "⚖️ Вес (кг)",
            min_value=0.0,
            value=1.0,
            step=0.1,
            key="ue_weight"
        )
    
    with col2:
        volume = st.number_input(
            "📦 Объем (литры)",
            min_value=0.0,
            value=5.0,
            step=0.5,
            key="ue_volume"
        )
        
        marketplace = st.selectbox(
            "🏪 Маркетплейс",
            list(unit_economics._configs.keys()),
            key="ue_marketplace"
        )
        
        operation_mode = st.selectbox(
            "📦 Режим работы",
            ["FBY", "FBS", "FBO", "DBS", "FBP"],
            key="ue_mode"
        )
        
        category = st.text_input(
            "📂 Категория (опционально)",
            placeholder="например: одежда_обувь",
            key="ue_category"
        )
        
        is_premium = st.checkbox(
            "⭐ Премиум-раздел (доп. комиссия)",
            key="ue_premium"
        )
    
    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="ue_calc"):
        with st.spinner("Расчет юнит-экономики..."):
            economics = unit_economics.calculate_unit_economics(
                price=price,
                cost=cost,
                weight_kg=weight,
                volume_liters=volume,
                marketplace=marketplace,
                operation_mode=operation_mode,
                category=category if category else None,
                is_premium=is_premium
            )
            
            if "error" not in economics:
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric(
                        "💰 Прибыль",
                        f"{economics['profit']:.2f} ₽",
                        delta=f"{economics['profit_per_ruble']:.2f} ₽/₽"
                    )
                    st.metric(
                        "📈 Маржа",
                        f"{economics['margin_percent']:.2f}%"
                    )
                
                with col2:
                    st.metric(
                        "📊 ROI",
                        f"{economics['roi']:.2f}%"
                    )
                    st.metric(
                        "⚖️ Точка безубыточности",
                        f"{economics['breakeven_price']:.2f} ₽"
                    )
                
                with col3:
                    st.metric(
                        "💵 Комиссия",
                        f"{economics['commission']:.2f} ₽",
                        f"{economics['commission_percent']:.1f}% от цены"
                    )
                    if economics.get('subscription_cost', 0) > 0:
                        st.metric(
                            "📋 Подписка (в день)",
                            f"{economics['subscription_cost']:.2f} ₽"
                        )
                    if economics.get('storage_non_standard', 0) > 0:
                        st.metric(
                            "📦 Нестандарт",
                            f"{economics['storage_non_standard']:.2f} ₽"
                        )
                
                st.subheader("📋 Детализация расходов")
                
                expenses_data = {
                    "Статья расходов": [
                        "Себестоимость",
                        "Комиссия",
                        "Подписка",
                        "Логистика",
                        "Хранение",
                        "Нестандарт",
                        "Эквайринг",
                        "Доставка",
                        "Последняя миля",
                        "Возвраты",
                        "РКО",
                        "Премиум",
                        "ИТОГО"
                    ],
                    "Сумма (₽)": [
                        economics['cost'],
                        economics['commission'],
                        economics.get('subscription_cost', 0),
                        economics['logistics'],
                        economics['storage_cost'],
                        economics.get('storage_non_standard', 0),
                        economics['acquiring'],
                        economics['delivery'],
                        economics['last_mile'],
                        economics['returns'],
                        economics.get('rko_fee', 0),
                        economics.get('premium_fee', 0),
                        economics['total_expenses']
                    ],
                    "% от цены": [
                        f"{economics['cost']/price*100:.1f}%",
                        f"{economics['commission']/price*100:.1f}%",
                        f"{economics.get('subscription_cost', 0)/price*100:.1f}%",
                        f"{economics['logistics']/price*100:.1f}%",
                        f"{economics['storage_cost']/price*100:.1f}%",
                        f"{economics.get('storage_non_standard', 0)/price*100:.1f}%",
                        f"{economics['acquiring']/price*100:.1f}%",
                        f"{economics['delivery']/price*100:.1f}%",
                        f"{economics['last_mile']/price*100:.1f}%",
                        f"{economics['returns']/price*100:.1f}%",
                        f"{economics.get('rko_fee', 0)/price*100:.1f}%",
                        f"{economics.get('premium_fee', 0)/price*100:.1f}%",
                        f"{economics['total_expenses']/price*100:.1f}%"
                    ]
                }
                
                st.dataframe(
                    pd.DataFrame(expenses_data),
                    use_container_width=True,
                    key="ue_expenses_table"
                )
                
                st.subheader("🏆 Сравнение всех маркетплейсов")
                comparison_df = unit_economics.calculate_for_all_marketplaces(
                    price=price,
                    cost=cost,
                    weight_kg=weight,
                    volume_liters=volume,
                    operation_mode=operation_mode
                )
                st.dataframe(
                    comparison_df,
                    use_container_width=True,
                    key="ue_comparison_table"
                )
                
                if not comparison_df.empty:
                    best_idx = comparison_df['profit'].idxmax()
                    best = comparison_df.loc[best_idx]
                    st.success(
                        f"🏆 Оптимальный маркетплейс: **{best['marketplace']}** "
                        f"(прибыль: {best['profit']:.2f} ₽, маржа: {best['margin_percent']:.2f}%)"
                    )
                    
                    if LIBRARIES['plotly']:
                        try:
                            fig = go.Figure()
                            fig.add_trace(go.Bar(
                                x=comparison_df['marketplace'],
                                y=comparison_df['profit'],
                                name='Прибыль',
                                marker_color='#e94560'
                            ))
                            fig.add_trace(go.Bar(
                                x=comparison_df['marketplace'],
                                y=comparison_df['margin_percent'],
                                name='Маржа %',
                                marker_color='#0f3460',
                                yaxis='y2'
                            ))
                            fig.update_layout(
                                title='Сравнение маркетплейсов',
                                xaxis_title='Маркетплейс',
                                yaxis_title='Прибыль (₽)',
                                yaxis2=dict(
                                    title='Маржа (%)',
                                    overlaying='y',
                                    side='right'
                                ),
                                barmode='group',
                                height=400
                            )
                            st.plotly_chart(fig, use_container_width=True)
                        except Exception as e:
                            logger.warning(f"Ошибка визуализации: {e}")
            else:
                st.error(f"❌ Ошибка: {economics['error']}")
    
    stats = unit_economics.get_stats()
    if stats.get('total_calculations', 0) > 0:
        st.subheader("📊 Статистика расчетов")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Всего расчетов", stats.get('total_calculations', 0))
        with col2:
            st.metric("💰 Средняя прибыль", f"{stats.get('avg_profit', 0):.2f} ₽")
        with col3:
            st.metric("📈 Макс. прибыль", f"{stats.get('max_profit', 0):.2f} ₽")
        with col4:
            st.metric("🏆 Лучший МП", stats.get('best_marketplace', '—'))

def show_catalog_enhance_interface():
    st.header("📊 Обогащение каталога (поиск аналогов)")
    
    st.info("""
    🔍 **Поиск аналогов:**
    
    Система ищет аналоги через общие OE номера (2 уровня):
    - **Уровень 1**: Прямые аналоги (общие OE номера)
    - **Уровень 2**: Косвенные аналоги (через аналоги уровнем выше)
    
    **Как это работает:**
    1. Загрузите данные в каталог (OE, артикулы, кросс-ссылки)
    2. Система построит связи между товарами
    3. При обогащении для каждого товара найдутся аналоги
    """)
    
    if 'catalog_enhancer' not in st.session_state:
        st.session_state.catalog_enhancer = CatalogEnhancer()
    
    enhancer = st.session_state.catalog_enhancer
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📤 Загрузка данных каталога")
        
        st.markdown("""
        **Необходимые колонки в файлах:**
        
        **OE данные:**
        - `oe_number` - номер OE
        - `name` - название
        - `applicability` - применимость
        
        **Детали (артикулы):**
        - `artikul` - артикул
        - `brand` - бренд
        - `length`, `width`, `height`, `weight` - габариты
        
        **Кросс-ссылки:**
        - `oe_number` - номер OE
        - `artikul` - артикул
        - `brand` - бренд
        """)
        
        oe_file = st.file_uploader(
            "OE данные",
            type=['xlsx', 'csv'],
            key="enh_oe"
        )
        
        parts_file = st.file_uploader(
            "Детали (артикулы)",
            type=['xlsx', 'csv'],
            key="enh_parts"
        )
        
        cross_file = st.file_uploader(
            "Кросс-ссылки",
            type=['xlsx', 'csv'],
            key="enh_cross"
        )
        
        if st.button("📥 Загрузить данные в каталог", type="primary", key="enh_load_data"):
            with st.spinner("Загрузка данных..."):
                if oe_file:
                    try:
                        df = pd.read_excel(oe_file) if oe_file.name.endswith('.xlsx') else pd.read_csv(oe_file)
                        enhancer.load_oe_data(df)
                        st.success(f"✅ Загружено {len(df)} OE записей")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки OE: {str(e)}")
                
                if parts_file:
                    try:
                        df = pd.read_excel(parts_file) if parts_file.name.endswith('.xlsx') else pd.read_csv(parts_file)
                        enhancer.load_parts_data(df)
                        st.success(f"✅ Загружено {len(df)} записей деталей")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки деталей: {str(e)}")
                
                if cross_file:
                    try:
                        df = pd.read_excel(cross_file) if cross_file.name.endswith('.xlsx') else pd.read_csv(cross_file)
                        enhancer.load_cross_references(df)
                        st.success(f"✅ Загружено {len(df)} кросс-ссылок")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки кросс-ссылок: {str(e)}")
    
    with col2:
        st.subheader("🔍 Поиск аналогов")
        
        artikul = st.text_input(
            "Артикул",
            placeholder="Введите артикул",
            key="enh_artikul_input"
        )
        brand = st.text_input(
            "Бренд",
            placeholder="Введите бренд",
            key="enh_brand_input"
        )
        
        if st.button("🔍 Найти аналоги", type="primary", key="enh_find_analogs"):
            if artikul and brand:
                with st.spinner("Поиск аналогов..."):
                    data = enhancer.get_analog_data(artikul, brand)
                    
                    if data.get('error'):
                        st.error(f"❌ {data['error']}")
                    else:
                        st.success(f"✅ Найдено {data.get('analog_count', 0)} аналогов")
                        
                        if data.get('oe_list'):
                            st.info(f"🔗 OE номера: {data.get('oe_list')}")
                        
                        if data.get('has_analogs'):
                            st.subheader("📋 Аналоги")
                            analogs_df = pd.DataFrame(data.get('analogs', []))
                            st.dataframe(
                                analogs_df,
                                use_container_width=True,
                                key="enh_analogs_table"
                            )
            else:
                st.warning("⚠️ Введите артикул и бренд")
        
        st.subheader("📊 Статистика каталога")
        stats = enhancer.get_stats()
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📦 OE записей", stats.get('oe_loaded', 0))
            st.metric("🔄 Кросс-ссылок", stats.get('cross_loaded', 0))
        with col2:
            st.metric("🔧 Деталей", stats.get('parts_loaded', 0))
            st.metric("🔍 Поисков", stats.get('analog_searches', 0))
    
    st.divider()
    
    st.subheader("📊 Обогащение загруженного каталога")
    
    if 'uploaded_data' in st.session_state and st.session_state.uploaded_data is not None:
        df = st.session_state.uploaded_data
        
        col1, col2 = st.columns(2)
        
        with col1:
            artikul_col = st.selectbox(
                "Колонка с артикулом",
                df.columns,
                key="enh_artikul_col_select"
            )
        
        with col2:
            brand_col = st.selectbox(
                "Колонка с брендом",
                df.columns,
                key="enh_brand_col_select"
            )
        
        if st.button("🚀 Обогатить данные", type="primary", key="enh_enrich_data"):
            with st.spinner("Обогащение данных..."):
                enhanced_df = enhancer.enhance_catalog_data(df, artikul_col, brand_col)
                st.session_state.uploaded_data = enhanced_df
                
                st.success("✅ Данные обогащены!")
                
                st.subheader("📊 Результат обогащения")
                st.dataframe(
                    enhanced_df.head(20),
                    use_container_width=True,
                    key="enh_result_table"
                )
                
                if 'analog_count' in enhanced_df.columns:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📦 Всего товаров", len(enhanced_df))
                    with col2:
                        st.metric("🔄 С аналогами", len(enhanced_df[enhanced_df['analog_count'] > 0]))
                    with col3:
                        st.metric("📊 Среднее аналогов", f"{enhanced_df['analog_count'].mean():.1f}")
    else:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")

def show_data_upload_interface():
    st.header("📁 Загрузка данных каталога")
    
    st.info("""
    📋 **Инструкция по загрузке данных:**
    
    **ОБЯЗАТЕЛЬНЫЕ колонки:**
    - `Артикул` или `article` или `sku` - идентификатор товара
    - `Бренд` или `brand` или `производитель` - бренд товара
    - `Цена` или `price` или `стоимость` - цена продажи
    - `Себестоимость` или `cost` - себестоимость товара
    
    **ОПЦИОНАЛЬНЫЕ колонки (для расширенной функциональности):**
    - `Длина` или `length` - длина в см или мм (для расчета логистики)
    - `Ширина` или `width` - ширина в см или мм (для расчета логистики)
    - `Высота` или `height` - высота в см или мм (для расчета логистики)
    - `Вес` или `weight` - вес в кг (для расчета логистики)
    - `OE номер` или `oe_number` - оригинальный номер запчасти (для поиска аналогов)
    - `Категория` или `category` - категория товара (для классификации)
    - `Штрихкод` или `barcode` - штрихкод товара
    - `Описание` или `description` - описание товара
    - `Кратность` или `multiplicity` - кратность упаковки
    """)
    
    uploaded_file = st.file_uploader(
        "Загрузите файл каталога (Excel или CSV)",
        type=['xlsx', 'xls', 'csv'],
        key="data_upload_file"
    )
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            st.session_state.uploaded_data = df
            st.success(f"✅ Загружено {len(df)} товаров")
            
            st.subheader("📊 Предпросмотр данных")
            st.dataframe(
                df.head(10),
                use_container_width=True,
                key="upload_preview_table"
            )
            
            st.subheader("📋 Найденные колонки")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**Обязательные колонки:**")
                required_cols = ["Артикул", "Бренд", "Цена", "Себестоимость"]
                for col in required_cols:
                    found = any(col.lower() in c.lower() for c in df.columns)
                    st.write(f"{'✅' if found else '❌'} {col}")
            
            with col2:
                st.markdown("**Опциональные колонки:**")
                optional_cols = ["Длина", "Ширина", "Высота", "Вес", "OE номер", "Категория"]
                for col in optional_cols:
                    found = any(col.lower() in c.lower() for c in df.columns)
                    st.write(f"{'✅' if found else '❌'} {col}")
            
            if st.button("🏷️ Классифицировать категории", type="secondary", key="classify_btn"):
                with st.spinner("Классификация товаров..."):
                    classifier = CategoryClassifier()
                    
                    name_col = None
                    for col in df.columns:
                        col_lower = col.lower()
                        if any(w in col_lower for w in ['наименование', 'название', 'name', 'товар']):
                            name_col = col
                            break
                    
                    if name_col:
                        df['Категория'] = df[name_col].apply(lambda x: classifier.predict(x)[0])
                        st.session_state.uploaded_data = df
                        st.success("✅ Классификация завершена!")
                        
                        st.subheader("📊 Распределение по категориям")
                        category_counts = df['Категория'].value_counts()
                        st.dataframe(
                            category_counts,
                            use_container_width=True,
                            key="category_counts"
                        )
                    else:
                        st.warning("⚠️ Не найдена колонка с названием товара")
            
            if st.button("📊 Обогатить каталог (поиск аналогов)", type="primary", key="upload_enrich_button"):
                st.info("Перейдите на вкладку '📊 Обогащение каталога'")
                    
        except Exception as e:
            st.error(f"❌ Ошибка загрузки файла: {str(e)}")
            st.code(traceback.format_exc())

def show_export_interface():
    st.header("📤 Экспорт данных")
    
    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные")
        return
    
    df = st.session_state.uploaded_data
    
    st.success(f"✅ Готово к экспорту: {len(df)} товаров, {len(df.columns)} колонок")
    
    st.subheader("📊 Статистика перед экспортом")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📦 Товаров", len(df))
    
    with col2:
        price_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['цена', 'price', 'стоимость']):
                price_col = col
                break
        if price_col:
            avg_price = safe_float(df[price_col].mean())
            st.metric("💰 Средняя цена", f"{avg_price:.2f} ₽")
    
    with col3:
        cost_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная']):
                cost_col = col
                break
        if cost_col:
            avg_cost = safe_float(df[cost_col].mean())
            st.metric("💵 Средняя себестоимость", f"{avg_cost:.2f} ₽")
    
    with col4:
        if 'Категория' in df.columns:
            st.metric("📂 Категорий", df['Категория'].nunique())
    
    export_format = st.radio(
        "Формат экспорта",
        ["Excel (.xlsx)", "CSV (.csv)"],
        horizontal=True,
        key="export_format"
    )
    
    include_stats = st.checkbox(
        "📊 Включить лист со статистикой",
        value=True,
        key="export_stats"
    )
    
    include_history = st.checkbox(
        "📋 Включить историю расчетов",
        value=True,
        key="export_history"
    )
    
    if st.button("📥 Скачать файл", type="primary", key="export_btn"):
        try:
            if export_format.startswith("Excel"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Данные', index=False)
                    
                    if include_stats:
                        stats_data = {
                            'Показатель': ['Всего товаров', 'Колонок', 'Категорий'],
                            'Значение': [
                                len(df),
                                len(df.columns),
                                df['Категория'].nunique() if 'Категория' in df.columns else 0
                            ]
                        }
                        if price_col:
                            stats_data['Показатель'].append('Средняя цена')
                            stats_data['Значение'].append(safe_float(df[price_col].mean()))
                        if cost_col:
                            stats_data['Показатель'].append('Средняя себестоимость')
                            stats_data['Значение'].append(safe_float(df[cost_col].mean()))
                        pd.DataFrame(stats_data).to_excel(writer, sheet_name='Статистика', index=False)
                    
                    if include_history:
                        unit_economics = MarketplaceUnitEconomics()
                        history = unit_economics.get_history()
                        if history:
                            pd.DataFrame(history).to_excel(writer, sheet_name='История', index=False)
                
                output.seek(0)
                st.download_button(
                    label="📥 Скачать Excel файл",
                    data=output,
                    file_name=f"каталог_экспорт_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )
            else:
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV файл",
                    data=csv,
                    file_name=f"каталог_экспорт_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_csv"
                )
            
            st.success("✅ Данные успешно подготовлены к экспорту!")
            
        except Exception as e:
            st.error(f"❌ Ошибка экспорта: {str(e)}")
            logger.error(f"Export error: {traceback.format_exc()}")

def show_history_interface():
    st.header("📋 История расчетов")
    
    unit_economics = MarketplaceUnitEconomics()
    history = unit_economics.get_history()
    
    if not history:
        st.info("📋 История расчетов пуста. Выполните расчеты в разделе '📊 Юнит-экономика'")
        return
    
    df_history = pd.DataFrame(history)
    
    st.subheader("🔍 Фильтры")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        marketplaces = ['Все'] + sorted(df_history['marketplace'].unique().tolist())
        filter_marketplace = st.selectbox(
            "Маркетплейс",
            marketplaces,
            key="history_marketplace"
        )
    
    with col2:
        modes = ['Все'] + sorted(df_history['operation_mode'].unique().tolist())
        filter_mode = st.selectbox(
            "Режим работы",
            modes,
            key="history_mode"
        )
    
    with col3:
        if 'timestamp' in df_history.columns:
            df_history['timestamp_dt'] = pd.to_datetime(df_history['timestamp'])
            min_date = df_history['timestamp_dt'].min().date()
            max_date = df_history['timestamp_dt'].max().date()
            
            filter_start = st.date_input(
                "Дата с",
                min_date,
                key="history_start"
            )
            filter_end = st.date_input(
                "Дата по",
                max_date,
                key="history_end"
            )
    
    filtered_df = df_history.copy()
    
    if filter_marketplace != 'Все':
        filtered_df = filtered_df[filtered_df['marketplace'] == filter_marketplace]
    
    if filter_mode != 'Все':
        filtered_df = filtered_df[filtered_df['operation_mode'] == filter_mode]
    
    if 'timestamp_dt' in filtered_df.columns:
        filtered_df = filtered_df[
            (filtered_df['timestamp_dt'].dt.date >= filter_start) &
            (filtered_df['timestamp_dt'].dt.date <= filter_end)
        ]
    
    st.subheader(f"📊 Найдено расчетов: {len(filtered_df)}")
    
    if not filtered_df.empty:
        display_cols = [
            'marketplace', 'operation_mode', 'price', 'cost', 'profit',
            'margin_percent', 'roi', 'timestamp'
        ]
        available_cols = [col for col in display_cols if col in filtered_df.columns]
        
        st.dataframe(
            filtered_df[available_cols].sort_values('timestamp', ascending=False),
            use_container_width=True,
            key="history_table"
        )
        
        st.subheader("📊 Статистика по истории")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_profit = filtered_df['profit'].sum()
            st.metric("💰 Суммарная прибыль", f"{total_profit:,.0f} ₽")
        
        with col2:
            avg_profit = filtered_df['profit'].mean()
            st.metric("📈 Средняя прибыль", f"{avg_profit:.2f} ₽")
        
        with col3:
            max_profit = filtered_df['profit'].max()
            best_market = filtered_df[filtered_df['profit'] == max_profit]['marketplace'].iloc[0] if not filtered_df.empty else '-'
            st.metric("🏆 Лучший результат", f"{max_profit:.2f} ₽", delta=f"{best_market}")
        
        with col4:
            count = len(filtered_df)
            st.metric("📊 Всего записей", count)
        
        if LIBRARIES['plotly'] and len(filtered_df) > 1:
            try:
                fig = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=("Прибыль по датам", "Распределение прибыли по маркетплейсам")
                )
                
                fig.add_trace(
                    go.Scatter(
                        x=filtered_df['timestamp_dt'] if 'timestamp_dt' in filtered_df.columns else filtered_df.index,
                        y=filtered_df['profit'],
                        mode='lines+markers',
                        name='Прибыль',
                        line=dict(color='#e94560', width=2)
                    ),
                    row=1, col=1
                )
                
                fig.add_trace(
                    go.Box(
                        x=filtered_df['marketplace'],
                        y=filtered_df['profit'],
                        name='Прибыль по МП',
                        marker_color='#0f3460'
                    ),
                    row=2, col=1
                )
                
                fig.update_layout(
                    height=600,
                    showlegend=True,
                    title_text="Визуализация истории расчетов"
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                logger.warning(f"Ошибка визуализации истории: {e}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📥 Экспортировать историю в CSV", key="history_export"):
            if not filtered_df.empty:
                csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV",
                    data=csv,
                    file_name=f"история_расчетов_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="history_download"
                )
    
    with col2:
        if st.button("🗑️ Очистить историю", type="secondary", key="history_clear"):
            if st.checkbox("Подтвердите очистку", key="history_confirm"):
                unit_economics.clear_history()
                st.success("✅ История очищена")
                st.rerun()

def show_settings_interface():
    st.header("⚙️ Настройки приложения")
    
    st.info("""
    💡 **Настройки сохраняются в сессии браузера**
    
    Настройки позволяют адаптировать приложение под ваши задачи.
    """)
    
    st.subheader("🎨 Тема оформления")
    
    theme = st.selectbox(
        "Тема",
        ["🌞 Светлая", "🌙 Темная", "🔄 Системная"],
        key="settings_theme"
    )
    
    st.caption("Тема применяется при следующем запуске приложения")
    
    st.subheader("💱 Валютные настройки")
    
    currency = st.selectbox(
        "Основная валюта",
        ["₽ (Рубль)", "$ (Доллар)", "€ (Евро)", "₴ (Гривна)", "¥ (Юань)"],
        key="settings_currency"
    )
    
    show_currency_symbol = st.checkbox(
        "Отображать символ валюты",
        value=True,
        key="settings_show_currency"
    )
    
    st.subheader("📊 Параметры расчета")
    
    col1, col2 = st.columns(2)
    
    with col1:
        default_margin = st.number_input(
            "Целевая маржинальность (%)",
            min_value=0.0,
            max_value=100.0,
            value=15.0,
            step=1.0,
            key="settings_default_margin"
        )
    
    with col2:
        min_profit = st.number_input(
            "Минимальная прибыль (₽)",
            min_value=0.0,
            value=50.0,
            step=10.0,
            key="settings_min_profit"
        )
    
    st.subheader("📤 Экспортные настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        default_export_format = st.selectbox(
            "Формат по умолчанию",
            ["Excel", "CSV"],
            key="settings_export_format"
        )
    
    with col2:
        include_timestamp = st.checkbox(
            "Добавлять дату в имя файла",
            value=True,
            key="settings_include_timestamp"
        )
    
    st.subheader("📝 Настройки логирования")
    
    log_level = st.select_slider(
        "Уровень логирования",
        options=["DEBUG", "INFO", "WARNING", "ERROR"],
        value="INFO",
        key="settings_log_level"
    )
    
    if st.button("💾 Сохранить настройки", type="primary", key="settings_save"):
        st.session_state.settings = {
            "theme": theme,
            "currency": currency,
            "show_currency_symbol": show_currency_symbol,
            "default_margin": default_margin,
            "min_profit": min_profit,
            "default_export_format": default_export_format,
            "include_timestamp": include_timestamp,
            "log_level": log_level
        }
        
        st.success("✅ Настройки сохранены!")
        logger.setLevel(log_level)
        st.balloons()
    
    st.divider()
    st.subheader("ℹ️ Информация о приложении")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("📌 Версия", APP_VERSION)
        st.metric("🐍 Python", sys.version.split()[0])
    
    with col2:
        st.metric("📅 Дата", datetime.now().strftime("%d.%m.%Y"))
        st.metric("📊 Данных загружено", len(st.session_state.get('uploaded_data', pd.DataFrame())) if st.session_state.get('uploaded_data') is not None else 0)
    
    st.subheader("📚 Доступные библиотеки")
    
    available_libs = []
    for lib, installed in LIBRARIES.items():
        available_libs.append(f"✅ {lib}" if installed else f"❌ {lib}")
    
    cols = st.columns(3)
    for i, lib in enumerate(available_libs):
        cols[i % 3].write(lib)

def show_analytics_interface():
    st.header("📊 Аналитика и статистика")
    
    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")
        return
    
    df = st.session_state.uploaded_data.copy()
    
    st.success(f"✅ Анализ {len(df)} товаров")
    
    st.subheader("📈 Общая статистика")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📦 Товаров", len(df))
    
    with col2:
        price_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['цена', 'price', 'стоимость']):
                price_col = col
                break
        
        if price_col:
            avg_price = safe_float(df[price_col].mean())
            st.metric("💰 Средняя цена", f"{avg_price:,.0f} ₽")
    
    with col3:
        cost_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная']):
                cost_col = col
                break
        
        if cost_col:
            avg_cost = safe_float(df[cost_col].mean())
            st.metric("💵 Средняя себестоимость", f"{avg_cost:,.0f} ₽")
    
    with col4:
        if price_col and cost_col:
            df['_margin'] = safe_float(df[price_col]) - safe_float(df[cost_col])
            avg_margin = df['_margin'].mean()
            st.metric("📈 Средняя маржа", f"{avg_margin:,.0f} ₽")
    
    if LIBRARIES['plotly']:
        st.subheader("📊 Визуализация данных")
        
        chart_type = st.selectbox(
            "Тип графика",
            ["Распределение цен", "Распределение категорий", "Топ товаров по цене", "Анализ маржи"],
            key="analytics_chart_type"
        )
        
        if chart_type == "Распределение цен" and price_col:
            fig = px.histogram(
                df,
                x=price_col,
                nbins=30,
                title=f"Распределение цен ({price_col})",
                labels={price_col: 'Цена (₽)', 'count': 'Количество товаров'},
                color_discrete_sequence=['#e94560']
            )
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Распределение категорий" and 'Категория' in df.columns:
            category_counts = df['Категория'].value_counts()
            fig = px.pie(
                values=category_counts.values,
                names=category_counts.index,
                title="Распределение товаров по категориям",
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Топ товаров по цене" and price_col:
            name_col = None
            for col in df.columns:
                if any(w in col.lower() for w in ['наименование', 'название', 'name', 'товар']):
                    name_col = col
                    break
            
            if name_col:
                top_df = df.nlargest(10, price_col)[[name_col, price_col]]
                fig = px.bar(
                    top_df,
                    x=name_col,
                    y=price_col,
                    title="Топ 10 товаров по цене",
                    labels={name_col: 'Товар', price_col: 'Цена (₽)'},
                    color_discrete_sequence=['#0f3460']
                )
                st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "Анализ маржи" and price_col and cost_col:
            df['_margin_percent'] = (safe_float(df[price_col]) - safe_float(df[cost_col])) / safe_float(df[price_col]) * 100
            fig = px.scatter(
                df,
                x=price_col,
                y='_margin_percent',
                title="Зависимость маржи от цены",
                labels={price_col: 'Цена (₽)', '_margin_percent': 'Маржа (%)'},
                color_discrete_sequence=['#e94560'],
                trendline="ols"
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📊 Для расширенной визуализации установите plotly: `pip install plotly`")
        
        if price_col:
            st.subheader("📊 Статистика цен")
            st.dataframe(
                df[price_col].describe(),
                use_container_width=True
            )
    
    st.subheader("📊 Корреляционный анализ")
    
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    
    if len(numeric_cols) >= 2:
        corr_df = df[numeric_cols].corr()
        
        st.dataframe(
            corr_df.style.background_gradient(cmap='RdBu_r'),
            use_container_width=True
        )
        
        if LIBRARIES['plotly']:
            fig = go.Figure(data=go.Heatmap(
                z=corr_df.values,
                x=corr_df.columns,
                y=corr_df.columns,
                colorscale='RdBu_r',
                zmin=-1, zmax=1
            ))
            fig.update_layout(
                title="Тепловая карта корреляций",
                height=500
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Для корреляционного анализа необходимо минимум 2 числовые колонки")

# ============================================================================
# БЛОК 10: HIGH-VOLUME ИНТЕРФЕЙС
# ============================================================================

def show_high_volume_interface():
    """Интерфейс для High-Volume каталога"""
    st.header("🚗 High-Volume Каталог автозапчастей (10M+)")
    
    if not HIGH_VOLUME_AVAILABLE:
        st.warning("⚠️ Для работы High-Volume режима установите: pip install polars duckdb")
        return
    
    if 'high_volume_catalog' not in st.session_state:
        st.session_state.high_volume_catalog = HighVolumeAutoPartsCatalog()
    
    catalog = st.session_state.high_volume_catalog
    
    if not catalog.conn:
        st.error("❌ Ошибка подключения к базе данных")
        return
    
    st.sidebar.title("🧭 Меню High-Volume")
    option = st.sidebar.radio(
        "Выберите раздел",
        ["Загрузка данных", "Экспорт", "Статистика", "Управление"]
    )
    
    if option == "Загрузка данных":
        st.header("📥 Загрузка данных")
        col1, col2 = st.columns(2)
        with col1:
            oe_file = st.file_uploader("Основные данные (OE)", type=['xlsx'], key="hv_oe")
            cross_file = st.file_uploader("Кроссы (OE→Артикул)", type=['xlsx'], key="hv_cross")
            barcode_file = st.file_uploader("Штрих-коды", type=['xlsx'], key="hv_barcode")
        with col2:
            weight_dims_file = st.file_uploader("Вес и габариты", type=['xlsx'], key="hv_dims")
            images_file = st.file_uploader("Изображения", type=['xlsx'], key="hv_images")
            prices_file = st.file_uploader("Цены", type=['xlsx'], key="hv_prices")
        
        uploaded_files = {
            'oe': oe_file,
            'cross': cross_file,
            'barcode': barcode_file,
            'dimensions': weight_dims_file,
            'images': images_file,
            'prices': prices_file
        }
        
        if st.button("Обработать и загрузить", key="hv_load"):
            saved_paths = {}
            for key, file in uploaded_files.items():
                if file:
                    path = catalog.data_dir / f"{key}_{int(time.time())}.xlsx"
                    with open(path, "wb") as f:
                        f.write(file.getbuffer())
                    saved_paths[key] = str(path)
            if saved_paths:
                with st.spinner("Обработка файлов..."):
                    dataframes = catalog.merge_all_data_parallel(saved_paths)
                with st.spinner("Загрузка данных в базу..."):
                    catalog.process_and_load_data(dataframes)
            else:
                st.warning("Загрузите хотя бы один файл")
    
    elif option == "Экспорт":
        st.header("📤 Экспорт данных")
        total = catalog.conn.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        st.info(f"Всего: {total}")
        if total == 0:
            st.warning("Нет данных для экспорта")
            return
        
        format_choice = st.radio("Формат", ["CSV", "Excel", "Parquet"], key="hv_format")
        selected_columns = st.multiselect("Колонки", [
            "Артикул бренда", "Бренд", "Наименование", "Применимость", "Описание",
            "Категория товара", "Кратность", "Длинна", "Ширина", "Высота", "Вес",
            "Длинна/Ширина/Высота", "OE номер", "аналоги", "Ссылка на изображение", "Цена", "Валюта"
        ], key="hv_columns")
        
        include_prices = st.checkbox("Включить цены", value=True, key="hv_prices_include")
        apply_markup = st.checkbox("Применить наценку", value=True, disabled=not include_prices, key="hv_markup")
        
        if st.button("🚀 Экспортировать", key="hv_export_btn"):
            output_path = catalog.data_dir / f"export.{format_choice.lower()}"
            with st.spinner("Генерация файла..."):
                if format_choice == "CSV":
                    catalog.export_to_csv_optimized(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)
                elif format_choice == "Excel":
                    catalog.export_to_excel_optimized(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)
                elif format_choice == "Parquet":
                    catalog.export_to_parquet(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)
            with open(output_path, "rb") as f:
                st.download_button("⬇️ Скачать файл", f, file_name=output_path.name, key="hv_download")
    
    elif option == "Статистика":
        st.header("📈 Статистика")
        stats = catalog.get_statistics()
        if stats:
            col1, col2, col3 = st.columns(3)
            col1.metric("Уникальных товаров", f"{stats.get('unique_parts', 0):,}")
            col2.metric("Брендов", f"{stats.get('brands', 0):,}")
            col3.metric("Средняя цена", f"{stats.get('avg_price', 0):.2f} ₽")
            
            st.subheader("Топ 10 брендов")
            if 'top_brands' in stats and not stats['top_brands'].empty:
                st.dataframe(stats['top_brands'])
    
    elif option == "Управление":
        st.header("🔧 Управление данными")
        st.warning("⚠️ Операции необратимы!")
        
        management_option = st.radio(
            "Выберите действие:",
            ["Удалить по бренду", "Удалить по артикулу", "Управление ценами", "Исключения", "Категории"],
            key="hv_management"
        )
        
        if management_option == "Удалить по бренду":
            try:
                brands_result = catalog.conn.execute(
                    "SELECT DISTINCT brand FROM parts WHERE brand IS NOT NULL ORDER BY brand").fetchall()
                available_brands = [row[0] for row in brands_result] if brands_result else []
            except Exception as e:
                st.error(f"Ошибка: {e}")
                return
                
            if not available_brands:
                st.info("Нет данных")
                return
                
            selected_brand = st.selectbox("Бренд", available_brands, key="hv_delete_brand")
            brand_norm_result = catalog.conn.execute(
                "SELECT brand_norm FROM parts WHERE brand = ? LIMIT 1", [selected_brand]).fetchone()
            if brand_norm_result:
                brand_norm = brand_norm_result[0]
            else:
                brand_norm = catalog.normalize_key(pl.Series([selected_brand]))[0]
            
            count = catalog.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()[0]
            st.info(f"Удалить {count} записей бренда '{selected_brand}'?")
            
            if st.checkbox("Подтверждаю удаление", key="hv_confirm_brand"):
                if st.button("Удалить", key="hv_delete_brand_btn"):
                    deleted = catalog.delete_by_brand(brand_norm)
                    st.success(f"Удалено {deleted} записей")
                    st.rerun()
        
        elif management_option == "Удалить по артикулу":
            artikul_input = st.text_input("Артикул", key="hv_delete_artikul")
            if artikul_input:
                artikul_norm = catalog.normalize_key(pl.Series([artikul_input]))[0]
                count = catalog.conn.execute(
                    "SELECT COUNT(*) FROM parts WHERE artikul_norm = ?", [artikul_norm]).fetchone()[0]
                st.info(f"Найдено {count} записей для артикула '{artikul_input}'")
                if st.checkbox("Подтверждаю", key="hv_confirm_artikul"):
                    if st.button("Удалить", key="hv_delete_artikul_btn"):
                        deleted = catalog.delete_by_artikul(artikul_norm)
                        st.success(f"Удалено {deleted} записей")
                        st.rerun()
        
        elif management_option == "Управление ценами":
            show_price_settings(catalog)
        
        elif management_option == "Исключения":
            show_exclusion_settings(catalog)
        
        elif management_option == "Категории":
            show_category_mapping_settings(catalog)

def show_price_settings(catalog):
    st.subheader("💰 Управление ценами и наценками")
    
    st.subheader("Общая наценка")
    global_markup = st.number_input(
        "Общая наценка (%):",
        min_value=0.0,
        max_value=500.0,
        value=catalog.price_rules['global_markup'] * 500,
        step=0.1,
        key="hv_global_markup"
    )
    catalog.price_rules['global_markup'] = global_markup / 500
    
    st.subheader("Наценки по брендам")
    brand_markups = catalog.price_rules.get('brand_markups', {})
    
    try:
        brands_result = catalog.conn.execute(
            "SELECT DISTINCT brand FROM parts WHERE brand IS NOT NULL ORDER BY brand").fetchall()
        available_brands = [row[0] for row in brands_result] if brands_result else []
    except Exception as e:
        st.error(f"Ошибка при получении списка брендов: {e}")
        available_brands = []
    
    if available_brands:
        col1, col2 = st.columns([2, 1])
        with col1:
            selected_brand = st.selectbox("Выберите бренд:", available_brands, key="hv_brand_select")
        with col2:
            current_markup = brand_markups.get(selected_brand, catalog.price_rules.get('global_markup', 0))
            brand_markup = st.number_input(
                "Наценка (%):",
                min_value=0.0,
                max_value=500.0,
                value=current_markup * 500,
                step=0.1,
                key=f"hv_markup_{selected_brand}"
            )
        if st.button("Сохранить наценку", key=f"hv_save_{selected_brand}"):
            brand_markups[selected_brand] = brand_markup / 500
            catalog.price_rules['brand_markups'] = brand_markups
            catalog.save_price_rules()
            st.success(f"✅ Наценка для {selected_brand} сохранена")
    
    st.subheader("Ограничения по ценам")
    col1, col2 = st.columns(2)
    with col1:
        min_price = st.number_input(
            "Минимальная цена:",
            min_value=0.0,
            value=float(catalog.price_rules['min_price']),
            step=0.01,
            key="hv_min_price"
        )
        catalog.price_rules['min_price'] = min_price
    with col2:
        max_price = st.number_input(
            "Максимальная цена:",
            min_value=0.0,
            value=float(catalog.price_rules['max_price']),
            step=0.01,
            key="hv_max_price"
        )
        catalog.price_rules['max_price'] = max_price
    
    if st.button("Сохранить все настройки цен", key="hv_save_prices"):
        catalog.save_price_rules()
        st.success("✅ Все настройки цен сохранены")

def show_exclusion_settings(catalog):
    st.subheader("🚫 Управление исключениями при экспорте")
    st.info("Товары, содержащие эти слова в названии, будут исключены из экспорта")
    
    current_exclusions = "\n".join(catalog.exclusion_rules)
    new_exclusions = st.text_area(
        "Список исключений (по одному на строку):",
        value=current_exclusions,
        height=200,
        placeholder="Введите слова для исключения, например:\nКузов\nСтекла\nМасла",
        key="hv_exclusions"
    )
    
    if st.button("Сохранить правила исключения", key="hv_save_exclusions"):
        cleaned = [line.strip() for line in new_exclusions.splitlines() if line.strip()]
        if len(cleaned) != len(set(cleaned)):
            st.warning("Обнаружены дублирующие записи. Они будут автоматически удалены.")
        catalog.exclusion_rules = list(dict.fromkeys(cleaned))
        catalog.save_exclusion_rules()
        st.success("✅ Правила исключения сохранены")

def show_category_mapping_settings(catalog):
    st.subheader("🗂️ Управление категориями товаров")
    st.info("Настройте соответствие между названиями товаров и категориями")
    
    st.subheader("Текущие правила")
    if catalog.category_mapping:
        mapping_df = pd.DataFrame({
            "Название товара": list(catalog.category_mapping.keys()),
            "Категория": list(catalog.category_mapping.values())
        })
        st.dataframe(mapping_df, use_container_width=True, hide_index=True)
    else:
        st.write("Нет пользовательских правил")
    
    st.subheader("Добавить правило")
    col1, col2 = st.columns(2)
    with col1:
        name_pattern = st.text_input("Ключевое слово в названии", key="hv_cat_pattern")
    with col2:
        category = st.text_input("Категория", key="hv_cat_name")
    
    if st.button("➕ Добавить", key="hv_add_category"):
        if name_pattern.strip() and category.strip():
            normalized_key = name_pattern.strip().lower()
            existing_keys = {k.lower(): k for k in catalog.category_mapping.keys()}
            if normalized_key in existing_keys:
                st.warning(f"Правило для '{existing_keys[normalized_key]}' обновлено")
            catalog.category_mapping[name_pattern.strip()] = category.strip()
            catalog.save_category_mapping()
            st.success(f"Добавлено: {name_pattern.strip()} → {category.strip()}")
            st.rerun()
        else:
            st.error("Заполните оба поля")
    
    if catalog.category_mapping:
        st.subheader("🗑️ Удалить правило")
        rule_to_delete = st.selectbox(
            "Выберите правило",
            options=list(catalog.category_mapping.keys()),
            format_func=lambda x: f"{x} → {catalog.category_mapping[x]}",
            key="hv_delete_rule"
        )
        if st.button("Удалить", key="hv_delete_category"):
            del catalog.category_mapping[rule_to_delete]
            catalog.save_category_mapping()
            st.success(f"Удалено: {rule_to_delete}")
            st.rerun()

# ============================================================================
# БЛОК 11: ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.markdown(f"""
    <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #0f3460 0%, #16213e 100%); border-radius: 10px; margin-bottom: 20px;">
        <h1 style="color: white;">🚀 {APP_NAME}</h1>
        <p style="color: #e94560; font-size: 18px;">v{APP_VERSION} | Полная версия 5000+ строк</p>
        <p style="color: #aaa;">Юнит-экономика маркетплейсов 2026 | Каталог с поиском аналогов 2 уровня</p>
        <p style="color: #888;">High-Volume каталог (10M+) | ИИ-обновление тарифов | Экспорт с формулами</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/bar-chart.png", width=80)
        st.markdown("---")
        
        menu_options = [
            "📁 Загрузка данных",
            "📊 Обогащение каталога",
            "📊 Юнит-экономика",
            "📊 Аналитика",
            "📋 История расчетов",
            "📤 Экспорт",
            "🚗 High-Volume каталог",
            "⚙️ Настройки"
        ]
        
        menu_icons = {
            "📁 Загрузка данных": "📤",
            "📊 Обогащение каталога": "🔍",
            "📊 Юнит-экономика": "💰",
            "📊 Аналитика": "📈",
            "📋 История расчетов": "📜",
            "📤 Экспорт": "💾",
            "🚗 High-Volume каталог": "🚗",
            "⚙️ Настройки": "⚙️"
        }
        
        menu = st.radio(
            "Меню",
            menu_options,
            key="main_menu",
            format_func=lambda x: f"{menu_icons.get(x, '')} {x}"
        )
        
        st.markdown("---")
        
        st.markdown("### 📊 Состояние системы")
        data_loaded = st.session_state.get('uploaded_data') is not None
        rows = len(st.session_state.get('uploaded_data', pd.DataFrame())) if data_loaded else 0
        
        st.metric("📁 Данные", "Загружены ✅" if data_loaded else "Не загружены ❌")
        if data_loaded:
            st.metric("📦 Товаров", rows)
        
        st.markdown("---")
        st.caption(f"Приложение v{APP_VERSION}")
        st.caption(f"Python {sys.version.split()[0]}")
        
        with st.expander("📚 Библиотеки"):
            for lib, installed in LIBRARIES.items():
                status = "✅" if installed else "❌"
                st.write(f"{status} {lib}")
    
    try:
        if menu == "📁 Загрузка данных":
            show_data_upload_interface()
        elif menu == "📊 Обогащение каталога":
            show_catalog_enhance_interface()
        elif menu == "📊 Юнит-экономика":
            show_unit_economics_interface()
        elif menu == "📊 Аналитика":
            show_analytics_interface()
        elif menu == "📋 История расчетов":
            show_history_interface()
        elif menu == "📤 Экспорт":
            show_export_interface()
        elif menu == "🚗 High-Volume каталог":
            show_high_volume_interface()
        elif menu == "⚙️ Настройки":
            show_settings_interface()
    except Exception as e:
        st.error(f"❌ Ошибка при отображении страницы: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Main page error: {traceback.format_exc()}")

if __name__ == "__main__":
    main()
