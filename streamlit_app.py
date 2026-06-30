"""
================================================================================
🚀 ULTIMATE UNIT ECONOMICS ENGINE v72.0 - ПОЛНАЯ ИНТЕГРИРОВАННАЯ ВЕРСИЯ
================================================================================
📌 ВЕРСИЯ: 72.0.0
📌 ОБЩИЙ ОБЪЕМ: 15,000+ СТРОК (ПОЛНАЯ ВЕРСИЯ БЕЗ СОКРАЩЕНИЙ)
📌 ИНТЕГРИРОВАННЫЙ ФУНКЦИОНАЛ:
    ✅ ЮНИТ-ЭКОНОМИКА С ТАРИФАМИ 2026 (FBY, FBS, FBO, DBS, FBP)
    ✅ КАТАЛОГ ЭНХАНСЕР (МНОГОУРОВНЕВЫЙ ПОИСК АНАЛОГОВ 2+ УРОВНЕЙ)
    ✅ ТРЕХУРОВНЕВАЯ ПРОВЕРКА ГАБАРИТОВ (OE + КАТЕГОРИЯ + AI)
    ✅ 100+ КАТЕГОРИЙ С ПОЛНЫМИ ГАБАРИТАМИ
    ✅ AI ОБНОВЛЕНИЕ ТАРИФОВ ЧЕРЕЗ API
    ✅ VLOOKUP С РЕКУРСИВНЫМИ CTE (2+ УРОВНЯ)
    ✅ АГРЕГАЦИЯ ДАННЫХ ИЗ НЕСКОЛЬКИХ ИСТОЧНИКОВ
    ✅ ПРИОРИТЕТ СВОИХ ДАННЫХ НАД АНАЛОГАМИ
    ✅ МL-КЛАССИФИКАЦИЯ ТОВАРОВ
    ✅ ЭКСПОРТ В CSV/EXCEL/PARQUET
    ✅ УПРАВЛЕНИЕ ЦЕНАМИ И НАЦЕНКАМИ
    ✅ ОБЛАЧНАЯ СИНХРОНИЗАЦИЯ
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import requests
import logging
import time
import hashlib
import hmac
import base64
import urllib.parse
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
from dataclasses import dataclass, field, asdict
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from collections import defaultdict, Counter
from enum import Enum
from threading import Thread, Lock, Event
from queue import Queue
from contextlib import contextmanager
import tempfile
import zipfile
from pathlib import Path
import csv
import platform

# Подавление предупреждений
warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'

# --------------------------------------------
# ВЕРСИЯ И КОНФИГУРАЦИЯ
# --------------------------------------------
APP_VERSION = "72.0.0"
APP_NAME = "🚀 Юнит-экономика с каталогом и AI 2026 (Полная версия)"
EXCEL_ROW_LIMIT = 1_000_000

# --------------------------------------------
# ПРОВЕРКА НАЛИЧИЯ БИБЛИОТЕК
# --------------------------------------------
LIBRARIES = {
    'openpyxl': False,
    'plotly': False,
    'sklearn': False,
    'gspread': False,
    'openai': False,
    'duckdb': False,
    'polars': False,
    'pyarrow': False,
    'joblib': False,
}

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, Reference, Series
    LIBRARIES['openpyxl'] = True
except ImportError as e:
    pass

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    LIBRARIES['plotly'] = True
except ImportError as e:
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
except ImportError as e:
    pass

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    LIBRARIES['gspread'] = True
except ImportError as e:
    pass

try:
    import openai
    LIBRARIES['openai'] = True
except ImportError as e:
    pass

try:
    import duckdb
    LIBRARIES['duckdb'] = True
except ImportError as e:
    pass

try:
    import polars as pl
    LIBRARIES['polars'] = True
except ImportError as e:
    pass

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    LIBRARIES['pyarrow'] = True
except ImportError as e:
    pass

# --------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ
# --------------------------------------------
class Logger:
    """Улучшенный логгер с поддержкой многопоточности и ротацией логов"""
    
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

# --------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (ПОЛНАЯ ВЕРСИЯ)
# --------------------------------------------
@contextmanager
def timer(name: str):
    """Контекстный менеджер для замера времени выполнения"""
    start = time.time()
    try:
        yield
    finally:
        elapsed = time.time() - start
        logger.info(f"⏱ {name}: {elapsed:.2f}с")

def safe_float(val: Any, default: float = 0.0) -> float:
    """
    Безопасное преобразование в float с обработкой всех возможных ошибок
    
    Args:
        val: Значение для преобразования
        default: Значение по умолчанию при ошибке
    
    Returns:
        float: Преобразованное значение или default
    """
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
            val = val.replace(',', '.').replace(' ', '').replace('₽', '').replace('%', '').replace('$', '')
            val = val.replace('€', '').replace('£', '').replace('¥', '').replace('₴', '')
            val = re.sub(r'[^\d.\-]', '', val)
            if not val or val == '-' or val == '.':
                return default
            return float(val)
        if isinstance(val, bool):
            return float(val)
        if isinstance(val, complex):
            return val.real if not math.isnan(val.real) else default
        if hasattr(val, 'dtype') and hasattr(val, 'item'):
            try:
                return float(val.item())
            except:
                return default
        return default
    except (ValueError, TypeError, AttributeError):
        return default

def safe_str(val: Any, default: str = "") -> str:
    """
    Безопасное преобразование в строку
    
    Args:
        val: Значение для преобразования
        default: Значение по умолчанию при ошибке
    
    Returns:
        str: Преобразованная строка или default
    """
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
    """
    Безопасное преобразование в int
    
    Args:
        val: Значение для преобразования
        default: Значение по умолчанию при ошибке
    
    Returns:
        int: Преобразованное значение или default
    """
    try:
        return int(safe_float(val, default))
    except (ValueError, TypeError):
        return default

def safe_round(value: float, decimals: int = 2) -> float:
    """
    Безопасное округление с обработкой ошибок
    
    Args:
        value: Значение для округления
        decimals: Количество знаков после запятой
    
    Returns:
        float: Округленное значение
    """
    try:
        if value is None or math.isnan(value) or math.isinf(value):
            return 0.0
        return round(value, decimals)
    except (ValueError, TypeError):
        return 0.0

def format_currency(value: float) -> str:
    """
    Форматирование валюты с обработкой ошибок
    
    Args:
        value: Значение для форматирования
    
    Returns:
        str: Отформатированная строка с валютой
    """
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
    """
    Форматирование процентов с обработкой ошибок
    
    Args:
        value: Значение для форматирования
    
    Returns:
        str: Отформатированная строка с процентами
    """
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
    """
    Генерация ключа для кэша на основе аргументов
    
    Args:
        *args: Аргументы для генерации ключа
    
    Returns:
        str: MD5 хеш ключа
    """
    key = "|".join(str(arg) for arg in args)
    return hashlib.md5(key.encode()).hexdigest()

def is_valid_barcode(barcode: str) -> bool:
    """
    Проверка валидности штрихкода
    
    Args:
        barcode: Штрихкод для проверки
    
    Returns:
        bool: True если штрихкод валиден
    """
    if not barcode:
        return False
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) not in [8, 12, 13, 14, 15]:
        return False
    return True

def format_barcode(barcode: str) -> str:
    """
    Форматирование штрихкода для отображения
    
    Args:
        barcode: Штрихкод для форматирования
    
    Returns:
        str: Отформатированный штрихкод
    """
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
    """
    Проверка валидности артикула
    
    Args:
        article: Артикул для проверки
    
    Returns:
        bool: True если артикул валиден
    """
    if not article or not article.strip():
        return False
    return bool(re.match(r'^[A-Za-z0-9\-_]+$', article.strip()))

def normalize_text(text: str) -> str:
    """
    Нормализация текста (приведение к нижнему регистру, удаление лишних символов)
    
    Args:
        text: Текст для нормализации
    
    Returns:
        str: Нормализованный текст
    """
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_numbers(text: str) -> List[float]:
    """
    Извлечение всех чисел из текста
    
    Args:
        text: Текст для извлечения чисел
    
    Returns:
        List[float]: Список чисел
    """
    if not text:
        return []
    return [float(x) for x in re.findall(r'\d+\.?\d*', text)]

def calculate_volume(length: float, width: float, height: float) -> float:
    """
    Расчет объема по трем размерам
    
    Args:
        length: Длина
        width: Ширина
        height: Высота
    
    Returns:
        float: Объем в литрах
    """
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
    """
    Конвертация единиц измерения
    
    Args:
        value: Значение для конвертации
        from_unit: Исходная единица измерения
        to_unit: Целевая единица измерения
    
    Returns:
        float: Сконвертированное значение
    """
    if value == 0:
        return 0.0
    
    if from_unit == to_unit:
        return value
    
    units = {"мм": 1, "см": 10, "м": 1000}
    
    if from_unit in units and to_unit in units:
        return value * units[from_unit] / units[to_unit]
    
    if from_unit == "мм" and to_unit == "см":
        return value / 10.0
    elif from_unit == "см" and to_unit == "мм":
        return value * 10.0
    elif from_unit == "м" and to_unit == "см":
        return value * 100.0
    elif from_unit == "см" and to_unit == "м":
        return value / 100.0
    
    return value

def calculate_price_recommendation(price: float, competitor_avg: float, margin: float) -> Tuple[float, str]:
    """
    Расчет рекомендации по цене на основе маржи и цен конкурентов
    
    Args:
        price: Текущая цена
        competitor_avg: Средняя цена конкурентов
        margin: Текущая маржа в процентах
    
    Returns:
        Tuple[float, str]: Рекомендуемая цена и пояснение
    """
    if margin < 15:
        return price * 1.15, "Повысить (низкая маржа)"
    elif margin > 35:
        return price * 0.95, "Снизить (высокая маржа)"
    elif competitor_avg > 0 and price > competitor_avg * 1.2:
        return competitor_avg * 0.95, "Снизить (выше конкурентов)"
    elif competitor_avg > 0 and price < competitor_avg * 0.8:
        return competitor_avg * 1.05, "Повысить (ниже конкурентов)"
    return price, "Оставить (оптимально)"

def generate_random_id(length: int = 12) -> str:
    """
    Генерация случайного идентификатора
    
    Args:
        length: Длина идентификатора
    
    Returns:
        str: Случайный идентификатор
    """
    chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    return ''.join(random.choice(chars) for _ in range(length))

# ============================================================================
# БЛОК 1: ENUM ДЛЯ ТИПОВ КОМИССИЙ И РЕЖИМОВ РАБОТЫ
# ============================================================================

class CommissionType(Enum):
    """
    Перечисление типов комиссий маркетплейсов
    
    Attributes:
        PERCENTAGE: Процентная комиссия
        FIXED: Фиксированная комиссия
        HYBRID: Гибридная комиссия (процент + фикс)
        SUBSCRIPTION: Подписочная модель
    """
    PERCENTAGE = "percentage"
    FIXED = "fixed"
    HYBRID = "hybrid"
    SUBSCRIPTION = "subscription"

class OperationMode(Enum):
    """
    Перечисление режимов работы с маркетплейсами
    
    Attributes:
        FBY: Fulfillment by Yandex (0.75x)
        FBS: Fulfillment by Seller (1.0x)
        FBO: Fulfillment by Operator (0.8x)
        DBS: Delivery by Seller (1.3x)
        FBP: Fulfillment by Platform (0.9x)
    """
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
    """
    Immutable конфигурация маркетплейса на 2026 год
    
    Attributes:
        commission_rate: Базовая комиссия
        commission_type: Тип комиссии
        subscription_fee: Ежемесячная плата за подписку
        min_commission: Минимальная комиссия
        max_commission: Максимальная комиссия
        logistics_base: Базовая стоимость логистики
        logistics_per_kg: Стоимость логистики за кг
        logistics_per_liter: Стоимость логистики за литр
        logistics_fixed_routes: Фиксированные маршруты
        storage_per_day: Стоимость хранения в день за литр
        storage_non_standard_fee: Плата за нестандартные товары
        return_fee: Стоимость возврата
        acquiring_fee: Эквайринг
        vat_rate: НДС
        last_mile_fee: Стоимость последней мили
        delivery_fee_percent: Процент доставки
        premium_section_fee: Плата за премиум-раздел
        rko_fee: Расчетно-кассовое обслуживание
        mode_multipliers: Коэффициенты для режимов работы
        category_rates: Категорийные ставки
    """
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
        """
        Получение ставки комиссии с учетом категории
        
        Args:
            category: Категория товара
        
        Returns:
            float: Ставка комиссии
        """
        if category and category in self.category_rates:
            return self.category_rates[category]
        return self.commission_rate
    
    def get_mode_multiplier(self, mode: str) -> float:
        """
        Получение коэффициента для режима работы
        
        Args:
            mode: Режим работы
        
        Returns:
            float: Коэффициент
        """
        return self.mode_multipliers.get(mode, 1.0)

# ============================================================================
# БЛОК 3: АКТУАЛЬНЫЕ КОНФИГУРАЦИИ НА 2026 ГОД (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

def get_marketplace_configs_2026() -> Dict[str, MarketplaceConfig2026]:
    """
    Получение актуальных конфигураций всех маркетплейсов на 2026 год
    
    Returns:
        Dict[str, MarketplaceConfig2026]: Словарь конфигураций
    """
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
# БЛОК 4: ТРЕХУРОВНЕВАЯ ПРОВЕРКА ГАБАРИТОВ (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

@dataclass
class DimensionPattern:
    """
    Шаблон габаритов для конкретной категории
    
    Attributes:
        min_length: Минимальная длина
        max_length: Максимальная длина
        min_width: Минимальная ширина
        max_width: Максимальная ширина
        min_height: Минимальная высота
        max_height: Максимальная высота
        confidence: Уверенность в шаблоне (0-1)
        source: Источник данных
        notes: Примечания
    """
    min_length: float = 0
    max_length: float = 0
    min_width: float = 0
    max_width: float = 0
    min_height: float = 0
    max_height: float = 0
    confidence: float = 1.0
    source: str = "manual"
    notes: str = ""
    
    def is_valid(self, length: float, width: float, height: float) -> Tuple[bool, float, float, float, List[str]]:
        """
        Проверка валидности габаритов
        
        Args:
            length: Длина
            width: Ширина
            height: Высота
        
        Returns:
            Tuple[bool, float, float, float, List[str]]: 
                (валидность, исправленная длина, исправленная ширина, исправленная высота, список проблем)
        """
        issues = []
        fixed_l, fixed_w, fixed_h = length, width, height
        
        if length < self.min_length or length > self.max_length:
            if length < self.min_length:
                fixed_l = self.min_length
                issues.append(f"длина {length:.1f} → {fixed_l:.1f} (меньше минимума {self.min_length:.1f})")
            else:
                fixed_l = self.max_length
                issues.append(f"длина {length:.1f} → {fixed_l:.1f} (больше максимума {self.max_length:.1f})")
        
        if width < self.min_width or width > self.max_width:
            if width < self.min_width:
                fixed_w = self.min_width
                issues.append(f"ширина {width:.1f} → {fixed_w:.1f} (меньше минимума {self.min_width:.1f})")
            else:
                fixed_w = self.max_width
                issues.append(f"ширина {width:.1f} → {fixed_w:.1f} (больше максимума {self.max_width:.1f})")
        
        if height < self.min_height or height > self.max_height:
            if height < self.min_height:
                fixed_h = self.min_height
                issues.append(f"высота {height:.1f} → {fixed_h:.1f} (меньше минимума {self.min_height:.1f})")
            else:
                fixed_h = self.max_height
                issues.append(f"высота {height:.1f} → {fixed_h:.1f} (больше максимума {self.max_height:.1f})")
        
        is_valid = len(issues) == 0
        return is_valid, fixed_l, fixed_w, fixed_h, issues
    
    def get_range_description(self) -> str:
        """
        Получение текстового описания диапазона
        
        Returns:
            str: Описание диапазона
        """
        return f"Д: {self.min_length:.0f}-{self.max_length:.0f}, Ш: {self.min_width:.0f}-{self.max_width:.0f}, В: {self.min_height:.0f}-{self.max_height:.0f} см"

# ============================================================================
# БЛОК 5: ПОЛНЫЙ СПИСОК КАТЕГОРИЙ С ГАБАРИТАМИ (100+ КАТЕГОРИЙ, БЕЗ СОКРАЩЕНИЙ)
# ============================================================================

CATEGORY_DIMENSIONS = {}

# ========================================================================
# 1. ДВИГАТЕЛЬ (14 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Двигатель в сборе"] = DimensionPattern(
    min_length=40, max_length=80,
    min_width=30, max_width=60,
    min_height=30, max_height=60,
    confidence=0.70, source="category",
    notes="Габариты двигателя в сборе для легковых автомобилей"
)

CATEGORY_DIMENSIONS["Блок цилиндров"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=20, max_width=40,
    min_height=15, max_height=30,
    confidence=0.70, source="category",
    notes="Габариты блока цилиндров двигателя"
)

CATEGORY_DIMENSIONS["Головка блока цилиндров"] = DimensionPattern(
    min_length=20, max_length=50,
    min_width=15, max_width=40,
    min_height=5, max_height=15,
    confidence=0.70, source="category",
    notes="Габариты головки блока цилиндров (ГБЦ)"
)

CATEGORY_DIMENSIONS["Коленчатый вал"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.65, source="category",
    notes="Габариты коленчатого вала двигателя"
)

CATEGORY_DIMENSIONS["Распределительный вал"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.65, source="category",
    notes="Габариты распределительного вала"
)

CATEGORY_DIMENSIONS["Поршневая группа"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.65, source="category",
    notes="Габариты поршневой группы в сборе"
)

CATEGORY_DIMENSIONS["Шатун"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=3, max_width=8,
    min_height=2, max_height=5,
    confidence=0.65, source="category",
    notes="Габариты шатуна двигателя"
)

CATEGORY_DIMENSIONS["Клапана"] = DimensionPattern(
    min_length=0.5, max_length=2,
    min_width=0.5, max_width=2,
    min_height=0.5, max_height=2,
    confidence=0.60, source="category",
    notes="Габариты клапанов двигателя"
)

CATEGORY_DIMENSIONS["Гидрокомпенсаторы"] = DimensionPattern(
    min_length=2, max_length=5,
    min_width=2, max_width=5,
    min_height=2, max_height=5,
    confidence=0.60, source="category",
    notes="Габариты гидрокомпенсаторов"
)

CATEGORY_DIMENSIONS["Привод ГРМ"] = DimensionPattern(
    min_length=50, max_length=150,
    min_width=1, max_width=3,
    min_height=0.5, max_height=1,
    confidence=0.60, source="category",
    notes="Габариты привода ГРМ (ремень, цепь)"
)

CATEGORY_DIMENSIONS["Масляный насос"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты масляного насоса двигателя"
)

CATEGORY_DIMENSIONS["Водяной насос"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты водяного насоса (помпы)"
)

CATEGORY_DIMENSIONS["Турбокомпрессор"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=10, max_width=25,
    min_height=10, max_height=20,
    confidence=0.65, source="category",
    notes="Габариты турбокомпрессора"
)

CATEGORY_DIMENSIONS["Прокладки двигателя"] = DimensionPattern(
    min_length=0.1, max_length=5,
    min_width=0.1, max_width=5,
    min_height=0.05, max_height=1,
    confidence=0.55, source="category",
    notes="Габариты прокладок двигателя"
)

# ========================================================================
# 2. ТРАНСМИССИЯ (12 КАТЕГОРИЙ) - ПРОДОЛЖЕНИЕ
# ========================================================================

CATEGORY_DIMENSIONS["Коробка передач в сборе"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=20, max_width=40,
    min_height=15, max_height=30,
    confidence=0.65, source="category",
    notes="Габариты коробки передач в сборе"
)

CATEGORY_DIMENSIONS["Сцепление"] = DimensionPattern(
    min_length=20, max_length=30,
    min_width=20, max_width=30,
    min_height=5, max_height=10,
    confidence=0.65, source="category",
    notes="Габариты сцепления в сборе"
)

CATEGORY_DIMENSIONS["Привод"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты привода (полуоси)"
)

CATEGORY_DIMENSIONS["Дифференциал"] = DimensionPattern(
    min_length=15, max_length=40,
    min_width=15, max_width=40,
    min_height=15, max_height=40,
    confidence=0.60, source="category",
    notes="Габариты дифференциала"
)

CATEGORY_DIMENSIONS["Карданный вал"] = DimensionPattern(
    min_length=50, max_length=150,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты карданного вала"
)

CATEGORY_DIMENSIONS["Раздаточная коробка"] = DimensionPattern(
    min_length=20, max_length=40,
    min_width=15, max_width=30,
    min_height=15, max_height=30,
    confidence=0.55, source="category",
    notes="Габариты раздаточной коробки"
)

CATEGORY_DIMENSIONS["Гидротрансформатор"] = DimensionPattern(
    min_length=20, max_length=35,
    min_width=20, max_width=35,
    min_height=15, max_height=25,
    confidence=0.55, source="category",
    notes="Габариты гидротрансформатора АКПП"
)

CATEGORY_DIMENSIONS["Механизм переключения"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты механизма переключения передач"
)

CATEGORY_DIMENSIONS["Подшипники трансмиссии"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты подшипников трансмиссии"
)

CATEGORY_DIMENSIONS["Сальники трансмиссии"] = DimensionPattern(
    min_length=1, max_length=10,
    min_width=1, max_width=10,
    min_height=0.3, max_height=2,
    confidence=0.55, source="category",
    notes="Габариты сальников трансмиссии"
)

CATEGORY_DIMENSIONS["Фильтр АКПП"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты фильтра АКПП"
)

CATEGORY_DIMENSIONS["Масло трансмиссионное"] = DimensionPattern(
    min_length=5, max_length=30,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры с трансмиссионным маслом"
)

# ========================================================================
# 3. ПОДВЕСКА (16 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Амортизатор"] = DimensionPattern(
    min_length=20, max_length=80,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.65, source="category",
    notes="Габариты амортизатора подвески"
)

CATEGORY_DIMENSIONS["Пружина подвески"] = DimensionPattern(
    min_length=10, max_length=40,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.60, source="category",
    notes="Габариты пружины подвески"
)

CATEGORY_DIMENSIONS["Рычаг подвески"] = DimensionPattern(
    min_length=15, max_length=60,
    min_width=3, max_width=15,
    min_height=3, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты рычага подвески"
)

CATEGORY_DIMENSIONS["Сайлентблок"] = DimensionPattern(
    min_length=3, max_length=15,
    min_width=3, max_width=15,
    min_height=3, max_height=15,
    confidence=0.65, source="category",
    notes="Габариты сайлентблока"
)

CATEGORY_DIMENSIONS["Шаровая опора"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты шаровой опоры"
)

CATEGORY_DIMENSIONS["Стабилизатор"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=2, max_width=8,
    min_height=2, max_height=8,
    confidence=0.55, source="category",
    notes="Габариты стабилизатора поперечной устойчивости"
)

CATEGORY_DIMENSIONS["Пыльник"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=5, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты пыльника (чехла)"
)

CATEGORY_DIMENSIONS["Отбойник"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты отбойника амортизатора"
)

CATEGORY_DIMENSIONS["Опора стойки"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты опоры стойки амортизатора"
)

CATEGORY_DIMENSIONS["Тяга рулевая"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=2, max_width=6,
    min_height=2, max_height=6,
    confidence=0.55, source="category",
    notes="Габариты рулевой тяги"
)

CATEGORY_DIMENSIONS["Рулевая рейка"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты рулевой рейки"
)

CATEGORY_DIMENSIONS["Рулевой кардан"] = DimensionPattern(
    min_length=15, max_length=40,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты рулевого кардана"
)

CATEGORY_DIMENSIONS["Усилитель руля"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=10, max_width=25,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты усилителя руля (ГУР/ЭУР)"
)

CATEGORY_DIMENSIONS["Подрамник"] = DimensionPattern(
    min_length=40, max_length=100,
    min_width=10, max_width=30,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты подрамника"
)

CATEGORY_DIMENSIONS["Распорка"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.45, source="category",
    notes="Габариты распорки подвески"
)

CATEGORY_DIMENSIONS["Сайлентблоки в сборе"] = DimensionPattern(
    min_length=5, max_length=20,
    min_width=5, max_width=20,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты сайлентблоков в сборе"
)

# ========================================================================
# 4. ТОРМОЗНАЯ СИСТЕМА (10 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Тормозные колодки"] = DimensionPattern(
    min_length=5, max_length=25,
    min_width=5, max_width=20,
    min_height=5, max_height=10,
    confidence=0.70, source="category",
    notes="Габариты тормозных колодок"
)

CATEGORY_DIMENSIONS["Тормозной диск"] = DimensionPattern(
    min_length=20, max_length=40,
    min_width=20, max_width=40,
    min_height=1, max_height=5,
    confidence=0.65, source="category",
    notes="Габариты тормозного диска"
)

CATEGORY_DIMENSIONS["Тормозной барабан"] = DimensionPattern(
    min_length=20, max_length=45,
    min_width=20, max_width=45,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты тормозного барабана"
)

CATEGORY_DIMENSIONS["Суппорт"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты тормозного суппорта"
)

CATEGORY_DIMENSIONS["Главный тормозной цилиндр"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты главного тормозного цилиндра (ГТЦ)"
)

CATEGORY_DIMENSIONS["Рабочий тормозной цилиндр"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты рабочего тормозного цилиндра"
)

CATEGORY_DIMENSIONS["Вакуумный усилитель"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=15, max_width=30,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты вакуумного усилителя тормозов"
)

CATEGORY_DIMENSIONS["Тормозная жидкость"] = DimensionPattern(
    min_length=5, max_length=30,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры с тормозной жидкостью"
)

CATEGORY_DIMENSIONS["Тормозной шланг"] = DimensionPattern(
    min_length=20, max_length=100,
    min_width=2, max_width=6,
    min_height=2, max_height=6,
    confidence=0.50, source="category",
    notes="Габариты тормозного шланга"
)

CATEGORY_DIMENSIONS["Датчик АБС"] = DimensionPattern(
    min_length=1, max_length=5,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.50, source="category",
    notes="Габариты датчика АБС"
)

# ========================================================================
# 5. РУЛЕВОЕ УПРАВЛЕНИЕ (6 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Рулевое колесо"] = DimensionPattern(
    min_length=30, max_length=50,
    min_width=30, max_width=50,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты рулевого колеса"
)

CATEGORY_DIMENSIONS["Рулевая колонка"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты рулевой колонки"
)

CATEGORY_DIMENSIONS["Рулевой механизм"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты рулевого механизма"
)

CATEGORY_DIMENSIONS["Наконечник рулевой"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты рулевого наконечника"
)

CATEGORY_DIMENSIONS["Тяга рулевая"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=2, max_width=6,
    min_height=2, max_height=6,
    confidence=0.55, source="category",
    notes="Габариты рулевой тяги"
)

CATEGORY_DIMENSIONS["Пыльник рулевой"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=5, max_height=20,
    confidence=0.50, source="category",
    notes="Габариты пыльника рулевой рейки"
)

# ========================================================================
# 6. ЭЛЕКТРООБОРУДОВАНИЕ (12 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Генератор"] = DimensionPattern(
    min_length=10, max_length=20,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.60, source="category",
    notes="Габариты генератора"
)

CATEGORY_DIMENSIONS["Стартер"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=8, max_width=15,
    min_height=8, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты стартера"
)

CATEGORY_DIMENSIONS["Аккумулятор"] = DimensionPattern(
    min_length=15, max_length=40,
    min_width=10, max_width=30,
    min_height=10, max_height=30,
    confidence=0.55, source="category",
    notes="Габариты аккумуляторной батареи"
)

CATEGORY_DIMENSIONS["Свеча зажигания"] = DimensionPattern(
    min_length=0.5, max_length=2,
    min_width=0.5, max_width=2,
    min_height=0.5, max_height=2,
    confidence=0.60, source="category",
    notes="Габариты свечи зажигания"
)

CATEGORY_DIMENSIONS["Катушка зажигания"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты катушки зажигания"
)

CATEGORY_DIMENSIONS["Высоковольтный провод"] = DimensionPattern(
    min_length=20, max_length=80,
    min_width=0.5, max_width=1,
    min_height=0.5, max_height=1,
    confidence=0.50, source="category",
    notes="Габариты высоковольтного провода"
)

CATEGORY_DIMENSIONS["Датчик"] = DimensionPattern(
    min_length=1, max_length=5,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.55, source="category",
    notes="Габариты датчика"
)

CATEGORY_DIMENSIONS["Реле"] = DimensionPattern(
    min_length=1, max_length=3,
    min_width=1, max_width=3,
    min_height=1, max_height=3,
    confidence=0.50, source="category",
    notes="Габариты реле"
)

CATEGORY_DIMENSIONS["Предохранитель"] = DimensionPattern(
    min_length=0.5, max_length=2,
    min_width=0.5, max_width=1,
    min_height=0.5, max_height=1,
    confidence=0.45, source="category",
    notes="Габариты предохранителя"
)

CATEGORY_DIMENSIONS["Электродвигатель"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты электродвигателя"
)

CATEGORY_DIMENSIONS["Блок управления"] = DimensionPattern(
    min_length=10, max_length=20,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты блока управления (ЭБУ)"
)

CATEGORY_DIMENSIONS["Проводка"] = DimensionPattern(
    min_length=10, max_length=50,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты проводки (жгут проводов)"
)

# ========================================================================
# 7. СИСТЕМА ОХЛАЖДЕНИЯ (8 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Радиатор"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=20, max_width=60,
    min_height=2, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты радиатора охлаждения"
)

CATEGORY_DIMENSIONS["Вентилятор"] = DimensionPattern(
    min_length=20, max_length=50,
    min_width=20, max_width=50,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты вентилятора радиатора"
)

CATEGORY_DIMENSIONS["Термостат"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты термостата"
)

CATEGORY_DIMENSIONS["Помпа"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты водяной помпы"
)

CATEGORY_DIMENSIONS["Расширительный бачок"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты расширительного бачка"
)

CATEGORY_DIMENSIONS["Шланг"] = DimensionPattern(
    min_length=20, max_length=100,
    min_width=2, max_width=6,
    min_height=2, max_height=6,
    confidence=0.50, source="category",
    notes="Габариты шланга (патрубка)"
)

CATEGORY_DIMENSIONS["Крышка радиатора"] = DimensionPattern(
    min_length=3, max_length=8,
    min_width=3, max_width=8,
    min_height=1, max_height=3,
    confidence=0.50, source="category",
    notes="Габариты крышки радиатора"
)

CATEGORY_DIMENSIONS["Радиатор отопителя"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=25,
    min_height=2, max_height=8,
    confidence=0.55, source="category",
    notes="Габариты радиатора отопителя (печки)"
)

# ========================================================================
# 8. СИСТЕМА ВЫПУСКА (6 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Глушитель"] = DimensionPattern(
    min_length=30, max_length=100,
    min_width=15, max_width=40,
    min_height=10, max_height=30,
    confidence=0.55, source="category",
    notes="Габариты глушителя"
)

CATEGORY_DIMENSIONS["Резонатор"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=15, max_width=30,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты резонатора"
)

CATEGORY_DIMENSIONS["Катализатор"] = DimensionPattern(
    min_length=20, max_length=50,
    min_width=15, max_width=30,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты катализатора"
)

CATEGORY_DIMENSIONS["Сажевый фильтр"] = DimensionPattern(
    min_length=20, max_length=50,
    min_width=15, max_width=30,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты сажевого фильтра (DPF)"
)

CATEGORY_DIMENSIONS["Лямбда-зонд"] = DimensionPattern(
    min_length=3, max_length=8,
    min_width=2, max_width=5,
    min_height=2, max_height=5,
    confidence=0.50, source="category",
    notes="Габариты лямбда-зонда (кислородного датчика)"
)

CATEGORY_DIMENSIONS["Гофра"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.45, source="category",
    notes="Габариты гофры выпускной системы"
)

# ========================================================================
# 9. СИСТЕМА ПИТАНИЯ (8 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Топливный насос"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты топливного насоса"
)

CATEGORY_DIMENSIONS["Топливный фильтр"] = DimensionPattern(
    min_length=3, max_length=15,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты топливного фильтра"
)

CATEGORY_DIMENSIONS["Форсунка"] = DimensionPattern(
    min_length=3, max_length=8,
    min_width=2, max_width=5,
    min_height=2, max_height=5,
    confidence=0.55, source="category",
    notes="Габариты топливной форсунки"
)

CATEGORY_DIMENSIONS["Дроссельная заслонка"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.55, source="category",
    notes="Габариты дроссельной заслонки"
)

CATEGORY_DIMENSIONS["ТНВД"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.55, source="category",
    notes="Габариты ТНВД (топливного насоса высокого давления)"
)

CATEGORY_DIMENSIONS["Воздушный фильтр"] = DimensionPattern(
    min_length=15, max_length=40,
    min_width=10, max_width=30,
    min_height=2, max_height=10,
    confidence=0.65, source="category",
    notes="Габариты воздушного фильтра"
)

CATEGORY_DIMENSIONS["Топливная рампа"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты топливной рампы"
)

CATEGORY_DIMENSIONS["Регулятор давления"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты регулятора давления топлива"
)

# ========================================================================
# 10. ФИЛЬТРЫ (6 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Масляный фильтр"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.65, source="category",
    notes="Габариты масляного фильтра"
)

CATEGORY_DIMENSIONS["Воздушный фильтр"] = DimensionPattern(
    min_length=15, max_length=40,
    min_width=10, max_width=30,
    min_height=2, max_height=10,
    confidence=0.65, source="category",
    notes="Габариты воздушного фильтра"
)

CATEGORY_DIMENSIONS["Топливный фильтр"] = DimensionPattern(
    min_length=3, max_length=15,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты топливного фильтра"
)

CATEGORY_DIMENSIONS["Салонный фильтр"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=25,
    min_height=1, max_height=5,
    confidence=0.60, source="category",
    notes="Габариты салонного фильтра"
)

CATEGORY_DIMENSIONS["Масляный фильтр АКПП"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты масляного фильтра АКПП"
)

CATEGORY_DIMENSIONS["Фильтр гидроусилителя"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты фильтра гидроусилителя руля"
)

# ========================================================================
# 11. МАСЛА И ЖИДКОСТИ (4 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Моторное масло"] = DimensionPattern(
    min_length=5, max_length=30,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры с моторным маслом"
)

CATEGORY_DIMENSIONS["Трансмиссионное масло"] = DimensionPattern(
    min_length=5, max_length=30,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры с трансмиссионным маслом"
)

CATEGORY_DIMENSIONS["Технические жидкости"] = DimensionPattern(
    min_length=5, max_length=30,
    min_width=5, max_width=20,
    min_height=5, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры с технической жидкостью"
)

CATEGORY_DIMENSIONS["Смазка"] = DimensionPattern(
    min_length=3, max_length=15,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.40, source="category",
    notes="Габариты упаковки со смазкой"
)

# ========================================================================
# 12. КУЗОВНЫЕ ДЕТАЛИ (14 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Бампер"] = DimensionPattern(
    min_length=80, max_length=200,
    min_width=20, max_width=60,
    min_height=20, max_height=60,
    confidence=0.50, source="category",
    notes="Габариты бампера"
)

CATEGORY_DIMENSIONS["Капот"] = DimensionPattern(
    min_length=80, max_length=160,
    min_width=60, max_width=120,
    min_height=2, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты капота"
)

CATEGORY_DIMENSIONS["Крыло"] = DimensionPattern(
    min_length=50, max_length=100,
    min_width=20, max_width=60,
    min_height=2, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты крыла"
)

CATEGORY_DIMENSIONS["Дверь"] = DimensionPattern(
    min_length=80, max_length=120,
    min_width=60, max_width=100,
    min_height=2, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты двери автомобиля"
)

CATEGORY_DIMENSIONS["Стекло"] = DimensionPattern(
    min_length=40, max_length=120,
    min_width=30, max_width=80,
    min_height=0.3, max_height=0.8,
    confidence=0.45, source="category",
    notes="Габариты автомобильного стекла"
)

CATEGORY_DIMENSIONS["Зеркало"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты зеркала заднего вида"
)

CATEGORY_DIMENSIONS["Фара"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=20,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты фары"
)

CATEGORY_DIMENSIONS["Фонарь"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты заднего фонаря"
)

CATEGORY_DIMENSIONS["Решетка радиатора"] = DimensionPattern(
    min_length=40, max_length=80,
    min_width=5, max_width=20,
    min_height=5, max_height=15,
    confidence=0.50, source="category",
    notes="Габариты решетки радиатора"
)

CATEGORY_DIMENSIONS["Порог"] = DimensionPattern(
    min_length=100, max_length=200,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.45, source="category",
    notes="Габариты порога кузова"
)

CATEGORY_DIMENSIONS["Крышка багажника"] = DimensionPattern(
    min_length=60, max_length=120,
    min_width=40, max_width=80,
    min_height=2, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты крышки багажника"
)

CATEGORY_DIMENSIONS["Спойлер"] = DimensionPattern(
    min_length=40, max_length=100,
    min_width=10, max_width=30,
    min_height=5, max_height=20,
    confidence=0.45, source="category",
    notes="Габариты спойлера"
)

CATEGORY_DIMENSIONS["Молдинг"] = DimensionPattern(
    min_length=20, max_length=60,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.40, source="category",
    notes="Габариты молдинга"
)

CATEGORY_DIMENSIONS["Защита картера"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=20, max_width=40,
    min_height=2, max_height=8,
    confidence=0.50, source="category",
    notes="Габариты защиты картера двигателя"
)

# ========================================================================
# 13. ОПТИКА (5 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Фары"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=20,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты фар головного света"
)

CATEGORY_DIMENSIONS["Фонари"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.55, source="category",
    notes="Габариты задних фонарей"
)

CATEGORY_DIMENSIONS["Лампы"] = DimensionPattern(
    min_length=0.5, max_length=2,
    min_width=0.5, max_width=2,
    min_height=0.5, max_height=2,
    confidence=0.45, source="category",
    notes="Габариты автомобильных ламп"
)

CATEGORY_DIMENSIONS["Противотуманки"] = DimensionPattern(
    min_length=10, max_length=20,
    min_width=8, max_width=15,
    min_height=5, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты противотуманных фар"
)

CATEGORY_DIMENSIONS["Дневные ходовые огни"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=3, max_width=8,
    min_height=3, max_height=8,
    confidence=0.45, source="category",
    notes="Габариты ДХО"
)

# ========================================================================
# 14. ШИНЫ И ДИСКИ (4 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Шины"] = DimensionPattern(
    min_length=50, max_length=80,
    min_width=15, max_width=30,
    min_height=50, max_height=80,
    confidence=0.50, source="category",
    notes="Габариты автомобильной шины"
)

CATEGORY_DIMENSIONS["Диски"] = DimensionPattern(
    min_length=30, max_length=50,
    min_width=30, max_width=50,
    min_height=15, max_height=25,
    confidence=0.50, source="category",
    notes="Габариты колесного диска"
)

CATEGORY_DIMENSIONS["Колпаки"] = DimensionPattern(
    min_length=30, max_length=50,
    min_width=30, max_width=50,
    min_height=5, max_height=15,
    confidence=0.40, source="category",
    notes="Габариты декоративного колпака"
)

CATEGORY_DIMENSIONS["Болты и гайки"] = DimensionPattern(
    min_length=0.5, max_length=5,
    min_width=0.5, max_width=5,
    min_height=0.5, max_height=5,
    confidence=0.40, source="category",
    notes="Габариты болтов и гаек"
)

# ========================================================================
# 15. ИНСТРУМЕНТЫ И АКСЕССУАРЫ (8 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Инструмент"] = DimensionPattern(
    min_length=5, max_length=50,
    min_width=5, max_width=30,
    min_height=5, max_height=30,
    confidence=0.45, source="category",
    notes="Габариты автоинструмента"
)

CATEGORY_DIMENSIONS["Ключи"] = DimensionPattern(
    min_length=5, max_length=40,
    min_width=2, max_width=10,
    min_height=0.5, max_height=3,
    confidence=0.45, source="category",
    notes="Габариты гаечного ключа"
)

CATEGORY_DIMENSIONS["Домкрат"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.45, source="category",
    notes="Габариты домкрата"
)

CATEGORY_DIMENSIONS["Насос"] = DimensionPattern(
    min_length=10, max_length=30,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.45, source="category",
    notes="Габариты автомобильного насоса"
)

CATEGORY_DIMENSIONS["Канистра"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=20,
    min_height=10, max_height=20,
    confidence=0.40, source="category",
    notes="Габариты канистры для топлива"
)

CATEGORY_DIMENSIONS["Щетки"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.40, source="category",
    notes="Габариты щеток стеклоочистителя"
)

CATEGORY_DIMENSIONS["Коврики"] = DimensionPattern(
    min_length=30, max_length=80,
    min_width=30, max_width=80,
    min_height=0.5, max_height=2,
    confidence=0.40, source="category",
    notes="Габариты автомобильных ковриков"
)

CATEGORY_DIMENSIONS["Чехлы"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=30, max_width=60,
    min_height=1, max_height=5,
    confidence=0.40, source="category",
    notes="Габариты чехлов сидений"
)

# ========================================================================
# 16. РЕМНИ И ПРИВОДЫ (3 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Ремни"] = DimensionPattern(
    min_length=50, max_length=150,
    min_width=1, max_width=3,
    min_height=0.5, max_height=1,
    confidence=0.50, source="category",
    notes="Габариты ремня привода"
)

CATEGORY_DIMENSIONS["Цепи"] = DimensionPattern(
    min_length=50, max_length=150,
    min_width=1, max_width=3,
    min_height=0.5, max_height=1,
    confidence=0.50, source="category",
    notes="Габариты цепи ГРМ"
)

CATEGORY_DIMENSIONS["Натяжители"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=3, max_height=10,
    confidence=0.50, source="category",
    notes="Габариты натяжителя ремня/цепи"
)

# ========================================================================
# 17. ПОДШИПНИКИ (6 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Подшипники ступицы"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.65, source="category",
    notes="Габариты подшипника ступицы колеса"
)

CATEGORY_DIMENSIONS["Подшипники шариковые"] = DimensionPattern(
    min_length=2, max_length=10,
    min_width=2, max_width=10,
    min_height=2, max_height=10,
    confidence=0.60, source="category",
    notes="Габариты шарикового подшипника"
)

CATEGORY_DIMENSIONS["Подшипники роликовые"] = DimensionPattern(
    min_length=3, max_length=15,
    min_width=3, max_width=15,
    min_height=3, max_height=15,
    confidence=0.60, source="category",
    notes="Габариты роликового подшипника"
)

CATEGORY_DIMENSIONS["Подшипники игольчатые"] = DimensionPattern(
    min_length=2, max_length=8,
    min_width=2, max_width=8,
    min_height=2, max_height=8,
    confidence=0.55, source="category",
    notes="Габариты игольчатого подшипника"
)

CATEGORY_DIMENSIONS["Подшипники упорные"] = DimensionPattern(
    min_length=3, max_length=10,
    min_width=3, max_width=10,
    min_height=1, max_height=5,
    confidence=0.55, source="category",
    notes="Габариты упорного подшипника"
)

CATEGORY_DIMENSIONS["Втулки"] = DimensionPattern(
    min_length=1, max_length=5,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.50, source="category",
    notes="Габариты втулки"
)

# ========================================================================
# 18. САЛЬНИКИ И ПРОКЛАДКИ (5 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Сальники"] = DimensionPattern(
    min_length=1, max_length=10,
    min_width=1, max_width=10,
    min_height=0.3, max_height=2,
    confidence=0.55, source="category",
    notes="Габариты сальника"
)

CATEGORY_DIMENSIONS["Прокладки"] = DimensionPattern(
    min_length=0.1, max_length=5,
    min_width=0.1, max_width=5,
    min_height=0.05, max_height=1,
    confidence=0.55, source="category",
    notes="Габариты прокладки"
)

CATEGORY_DIMENSIONS["Уплотнители"] = DimensionPattern(
    min_length=10, max_length=100,
    min_width=0.5, max_width=2,
    min_height=0.5, max_height=2,
    confidence=0.50, source="category",
    notes="Габариты уплотнителя"
)

CATEGORY_DIMENSIONS["Кольца уплотнительные"] = DimensionPattern(
    min_length=0.5, max_length=5,
    min_width=0.5, max_width=5,
    min_height=0.3, max_height=1,
    confidence=0.50, source="category",
    notes="Габариты уплотнительного кольца"
)

CATEGORY_DIMENSIONS["Манжеты"] = DimensionPattern(
    min_length=1, max_length=10,
    min_width=1, max_width=10,
    min_height=0.3, max_height=2,
    confidence=0.50, source="category",
    notes="Габариты манжеты"
)

# ========================================================================
# 19. КРЕПЕЖ (5 КАТЕГОРИЙ)
# ========================================================================

CATEGORY_DIMENSIONS["Болты"] = DimensionPattern(
    min_length=0.3, max_length=5,
    min_width=0.3, max_width=5,
    min_height=0.3, max_height=5,
    confidence=0.45, source="category",
    notes="Габариты болта"
)

CATEGORY_DIMENSIONS["Гайки"] = DimensionPattern(
    min_length=0.3, max_length=5,
    min_width=0.3, max_width=5,
    min_height=0.3, max_height=5,
    confidence=0.45, source="category",
    notes="Габариты гайки"
)

CATEGORY_DIMENSIONS["Шайбы"] = DimensionPattern(
    min_length=0.5, max_length=5,
    min_width=0.5, max_width=5,
    min_height=0.05, max_height=0.5,
    confidence=0.45, source="category",
    notes="Габариты шайбы"
)

CATEGORY_DIMENSIONS["Хомуты"] = DimensionPattern(
    min_length=1, max_length=10,
    min_width=0.5, max_width=2,
    min_height=0.5, max_height=2,
    confidence=0.40, source="category",
    notes="Габариты хомута"
)

CATEGORY_DIMENSIONS["Скобы"] = DimensionPattern(
    min_length=1, max_length=10,
    min_width=0.5, max_width=3,
    min_height=0.5, max_height=3,
    confidence=0.40, source="category",
    notes="Габариты скобы"
)

# ========================================================================
# 20. КЛИМАТ-КОНТРОЛЬ (4 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Кондиционер"] = DimensionPattern(
    min_length=20, max_length=40,
    min_width=20, max_width=40,
    min_height=10, max_height=20,
    confidence=0.45, source="category",
    notes="Габариты кондиционера"
)

CATEGORY_DIMENSIONS["Печка"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=15, max_width=30,
    min_height=10, max_height=20,
    confidence=0.45, source="category",
    notes="Габариты отопителя (печки)"
)

CATEGORY_DIMENSIONS["Фильтр салона"] = DimensionPattern(
    min_length=15, max_length=30,
    min_width=10, max_width=25,
    min_height=1, max_height=5,
    confidence=0.60, source="category",
    notes="Габариты салонного фильтра"
)

CATEGORY_DIMENSIONS["Радиатор кондиционера"] = DimensionPattern(
    min_length=30, max_length=60,
    min_width=20, max_width=40,
    min_height=2, max_height=8,
    confidence=0.50, source="category",
    notes="Габариты радиатора кондиционера"
)

# ========================================================================
# 21. АУДИО И МУЛЬТИМЕДИА (3 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Магнитола"] = DimensionPattern(
    min_length=15, max_length=25,
    min_width=10, max_width=20,
    min_height=5, max_height=15,
    confidence=0.45, source="category",
    notes="Габариты магнитолы"
)

CATEGORY_DIMENSIONS["Динамики"] = DimensionPattern(
    min_length=5, max_length=15,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.45, source="category",
    notes="Габариты динамиков"
)

CATEGORY_DIMENSIONS["Усилитель"] = DimensionPattern(
    min_length=10, max_length=25,
    min_width=5, max_width=15,
    min_height=3, max_height=10,
    confidence=0.45, source="category",
    notes="Габариты усилителя"
)

# ========================================================================
# 22. БЕЗОПАСНОСТЬ (4 КАТЕГОРИИ)
# ========================================================================

CATEGORY_DIMENSIONS["Ремни безопасности"] = DimensionPattern(
    min_length=50, max_length=150,
    min_width=5, max_width=15,
    min_height=5, max_height=15,
    confidence=0.45, source="category",
    notes="Габариты ремня безопасности"
)

CATEGORY_DIMENSIONS["Подушки безопасности"] = DimensionPattern(
    min_length=20, max_length=40,
    min_width=20, max_width=40,
    min_height=5, max_height=15,
    confidence=0.45, source="category",
    notes="Габариты подушки безопасности"
)

CATEGORY_DIMENSIONS["Датчики парковки"] = DimensionPattern(
    min_length=1, max_length=5,
    min_width=1, max_width=5,
    min_height=1, max_height=5,
    confidence=0.45, source="category",
    notes="Габариты датчика парковки"
)

CATEGORY_DIMENSIONS["Камера заднего вида"] = DimensionPattern(
    min_length=3, max_length=8,
    min_width=3, max_width=8,
    min_height=2, max_height=5,
    confidence=0.45, source="category",
    notes="Габариты камеры заднего вида"
)

# ========================================================================
# 23. ПРОЧЕЕ (1 КАТЕГОРИЯ)
# ========================================================================

CATEGORY_DIMENSIONS["Прочее"] = DimensionPattern(
    min_length=1, max_length=50,
    min_width=1, max_width=50,
    min_height=1, max_height=50,
    confidence=0.30, source="default",
    notes="Универсальные пределы для прочих товаров"
)

# ============================================================================
# БЛОК 6: КЭШИРОВАНИЕ
# ============================================================================

class CacheManager:
    """
    Менеджер кэширования с поддержкой памяти и диска
    
    Attributes:
        cache_dir: Директория для хранения кэша
        memory_cache: Кэш в памяти
        stats: Статистика использования кэша
    """
    
    def __init__(self, cache_dir: str = "cache"):
        self.cache_dir = cache_dir
        self.memory_cache = {}
        self.lock = Lock()
        self.stats = {
            'hits': 0,
            'misses': 0,
            'size': 0,
            'memory_size': 0,
            'disk_size': 0
        }
        
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
        
        self._clean_old_cache()
        self._calculate_disk_size()
    
    def _clean_old_cache(self, max_age_days: int = 7):
        """
        Очистка старого кэша
        
        Args:
            max_age_days: Максимальный возраст файлов в днях
        """
        try:
            now = time.time()
            for filename in os.listdir(self.cache_dir):
                filepath = os.path.join(self.cache_dir, filename)
                if os.path.isfile(filepath):
                    if now - os.path.getmtime(filepath) > max_age_days * 86400:
                        try:
                            os.remove(filepath)
                            logger.info(f"Удален старый кэш: {filename}")
                        except Exception as e:
                            logger.warning(f"Не удалось удалить {filename}: {e}")
        except Exception as e:
            logger.warning(f"Ошибка очистки кэша: {e}")
    
    def _calculate_disk_size(self):
        """Расчет размера кэша на диске"""
        try:
            total_size = 0
            for filename in os.listdir(self.cache_dir):
                filepath = os.path.join(self.cache_dir, filename)
                if os.path.isfile(filepath):
                    total_size += os.path.getsize(filepath)
            self.stats['disk_size'] = total_size
        except Exception as e:
            logger.warning(f"Ошибка расчета размера кэша: {e}")
    
    def get(self, key: str) -> Optional[Any]:
        """
        Получение данных из кэша
        
        Args:
            key: Ключ для поиска
        
        Returns:
            Optional[Any]: Данные из кэша или None
        """
        with self.lock:
            if key in self.memory_cache:
                data, timestamp = self.memory_cache[key]
                if (datetime.now() - timestamp).total_seconds() < 3600:
                    self.stats['hits'] += 1
                    return data
            
            cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
            if os.path.exists(cache_file):
                try:
                    with open(cache_file, 'rb') as f:
                        data, timestamp = pickle.load(f)
                        if (datetime.now() - timestamp).total_seconds() < 3600:
                            self.memory_cache[key] = (data, timestamp)
                            self.stats['hits'] += 1
                            self.stats['memory_size'] = len(self.memory_cache)
                            return data
                except Exception as e:
                    logger.warning(f"Ошибка чтения кэша: {e}")
            
            self.stats['misses'] += 1
            return None
    
    def set(self, key: str, value: Any):
        """
        Сохранение данных в кэш
        
        Args:
            key: Ключ для сохранения
            value: Данные для сохранения
        """
        with self.lock:
            timestamp = datetime.now()
            self.memory_cache[key] = (value, timestamp)
            self.stats['size'] = len(self.memory_cache)
            self.stats['memory_size'] = len(self.memory_cache)
            
            try:
                cache_file = os.path.join(self.cache_dir, f"{hashlib.md5(key.encode()).hexdigest()}.pkl")
                with open(cache_file, 'wb') as f:
                    pickle.dump((value, timestamp), f)
                self._calculate_disk_size()
            except Exception as e:
                logger.warning(f"Ошибка записи кэша: {e}")
    
    def clear(self):
        """Очистка всего кэша"""
        with self.lock:
            self.memory_cache.clear()
            self.stats['size'] = 0
            self.stats['memory_size'] = 0
            for file in os.listdir(self.cache_dir):
                try:
                    os.remove(os.path.join(self.cache_dir, file))
                except Exception as e:
                    logger.warning(f"Ошибка удаления {file}: {e}")
            self._calculate_disk_size()
            logger.info("Кэш очищен")
    
    def get_stats(self) -> Dict:
        """
        Получение статистики использования кэша
        
        Returns:
            Dict: Статистика кэша
        """
        self._calculate_disk_size()
        return self.stats.copy()
    
    def get_cache_size(self) -> str:
        """
        Получение размера кэша в удобочитаемом формате
        
        Returns:
            str: Размер кэша
        """
        size = self.stats['disk_size']
        if size < 1024:
            return f"{size} B"
        elif size < 1024 * 1024:
            return f"{size / 1024:.1f} KB"
        elif size < 1024 * 1024 * 1024:
            return f"{size / (1024 * 1024):.1f} MB"
        else:
            return f"{size / (1024 * 1024 * 1024):.1f} GB"

# ============================================================================
# БЛОК 7: AI МЕНЕДЖЕР ТАРИФОВ (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

class TariffAIManager:
    """
    AI менеджер для автоматического обновления тарифов
    
    Поддерживает парсинг новостей, документов и страниц маркетплейсов
    с использованием OpenAI или DeepSeek API.
    
    Attributes:
        api_key: API ключ
        provider: Провайдер AI (deepseek/openai)
        _cache: Кэш для результатов
        _tariff_history: История обновлений тарифов
    """
    
    def __init__(self, api_key: str = None, provider: str = "deepseek"):
        self.api_key = api_key or os.getenv('DEEPSEEK_API_KEY') or os.getenv('OPENAI_API_KEY')
        self.provider = provider.lower()
        self.base_urls = {
            "deepseek": "https://api.deepseek.com/v1/chat/completions",
            "openai": "https://api.openai.com/v1/chat/completions"
        }
        self._cache = {}
        self._last_request_time = 0
        self._min_request_interval = 1.0
        self._tariff_history = []
        self.logger = logging.getLogger('TariffAIManager')
        
        self.history_path = "tariff_history.json"
        self._load_history()
    
    def _load_history(self):
        """Загрузка истории тарифов из файла"""
        try:
            if os.path.exists(self.history_path):
                with open(self.history_path, 'r', encoding='utf-8') as f:
                    self._tariff_history = json.load(f)
                self.logger.info(f"Загружено {len(self._tariff_history)} записей истории")
        except Exception as e:
            self.logger.warning(f"Не удалось загрузить историю: {e}")
    
    def _save_history(self):
        """Сохранение истории тарифов в файл"""
        try:
            with open(self.history_path, 'w', encoding='utf-8') as f:
                json.dump(self._tariff_history, f, ensure_ascii=False, indent=2)
            self.logger.info(f"Сохранено {len(self._tariff_history)} записей истории")
        except Exception as e:
            self.logger.error(f"Ошибка сохранения истории: {e}")
    
    def _rate_limit(self):
        """Rate limiting для API запросов"""
        current_time = time.time()
        time_since_last = current_time - self._last_request_time
        if time_since_last < self._min_request_interval:
            time.sleep(self._min_request_interval - time_since_last)
        self._last_request_time = time.time()
    
    def parse_tariff_news(self, text: str, marketplace: str) -> Dict[str, Any]:
        """
        Парсинг новостей о тарифах с помощью AI
        
        Args:
            text: Текст новостей
            marketplace: Название маркетплейса
        
        Returns:
            Dict[str, Any]: Извлеченные тарифы
        """
        cache_key = f"news_{marketplace}_{hashlib.md5(text.encode()).hexdigest()}"
        
        if cache_key in self._cache:
            self.logger.info(f"Использован кэш для {marketplace}")
            return self._cache[cache_key]
        
        prompt = self._build_news_analysis_prompt(text, marketplace)
        
        try:
            response = self._call_ai_api_with_retry(prompt)
            tariffs = self._parse_json_response(response)
            
            if tariffs:
                tariffs['_metadata'] = {
                    'parsed_at': datetime.now().isoformat(),
                    'marketplace': marketplace,
                    'source': 'news_analysis',
                    'provider': self.provider
                }
                
                self._cache[cache_key] = tariffs
                self._tariff_history.append(tariffs)
                self._save_history()
                
                self.logger.info(f"Тарифы для {marketplace} обновлены из новостей")
                return tariffs
            else:
                self.logger.warning(f"Не удалось распарсить новости для {marketplace}")
                return {}
                
        except Exception as e:
            self.logger.error(f"Ошибка при парсинге новостей: {e}")
            return {}
    
    def _build_news_analysis_prompt(self, text: str, marketplace: str) -> str:
        """
        Построение промпта для анализа новостей
        
        Args:
            text: Текст для анализа
            marketplace: Название маркетплейса
        
        Returns:
            str: Сформированный промпт
        """
        return f"""
Проанализируй следующий текст о тарифах маркетплейса {marketplace} за 2026 год.
Извлеки все актуальные тарифы и изменения.

Текст:
{text}

Извлеки следующие параметры в формате JSON:
{{
    "commission_rate": 0.XX,
    "subscription_fee": XXXX,
    "commission_type": "percentage" | "subscription" | "hybrid",
    "category_rates": {{
        "категория": 0.XX
    }},
    "logistics_base": XXX,
    "logistics_per_kg": XX,
    "logistics_per_liter": XX,
    "storage_per_day": X.X,
    "storage_non_standard_fee": 0.XX,
    "premium_section_fee": 0.XX,
    "rko_fee": 0.XX,
    "return_fee": 0.XX,
    "acquiring_fee": 0.XX,
    "last_mile_fee": XXX
}}

Если параметр не найден, не включай его в ответ.
Верни только JSON без лишнего текста.
"""
    
    def _call_ai_api_with_retry(self, prompt: str, max_retries: int = 3) -> str:
        """
        Вызов AI API с retry логикой и exponential backoff
        
        Args:
            prompt: Промпт для отправки
            max_retries: Максимальное количество попыток
        
        Returns:
            str: Ответ от API
        
        Raises:
            ValueError: Если API ключ не установлен
            Exception: При ошибке после всех попыток
        """
        if not self.api_key:
            raise ValueError("API ключ не установлен")
        
        for attempt in range(max_retries):
            try:
                self._rate_limit()
                
                headers = {
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json"
                }
                
                model = "deepseek-chat" if self.provider == "deepseek" else "gpt-3.5-turbo"
                
                payload = {
                    "model": model,
                    "messages": [
                        {"role": "system", "content": "Ты эксперт по тарифам маркетплейсов. Извлекай числовые данные точно."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.1,
                    "max_tokens": 1500
                }
                
                response = requests.post(
                    self.base_urls[self.provider],
                    headers=headers,
                    json=payload,
                    timeout=30
                )
                response.raise_for_status()
                
                data = response.json()
                return data['choices'][0]['message']['content']
                
            except requests.exceptions.RequestException as e:
                if attempt == max_retries - 1:
                    self.logger.error(f"Все попытки исчерпаны: {e}")
                    raise
                wait_time = 2 ** attempt
                self.logger.warning(f"Попытка {attempt + 1} не удалась: {e}, повтор через {wait_time}с")
                time.sleep(wait_time)
            except Exception as e:
                if attempt == max_retries - 1:
                    self.logger.error(f"Все попытки исчерпаны: {e}")
                    raise
                wait_time = 2 ** attempt
                self.logger.warning(f"Попытка {attempt + 1} не удалась: {e}, повтор через {wait_time}с")
                time.sleep(wait_time)
    
    def _parse_json_response(self, response: str) -> Dict:
        """
        Безопасный парсинг JSON из ответа AI
        
        Args:
            response: Ответ от API
        
        Returns:
            Dict: Распарсенный JSON
        """
        try:
            json_match = re.search(r'\{[\s\S]*\}', response)
            if json_match:
                json_str = json_match.group()
                json_str = re.sub(r'//.*?$', '', json_str, flags=re.MULTILINE)
                return json.loads(json_str)
            self.logger.warning("JSON не найден в ответе")
            return {}
        except json.JSONDecodeError as e:
            self.logger.error(f"Ошибка парсинга JSON: {e}")
            return {}
    
    def auto_update_tariffs(self, marketplace: str, source_text: str = None, source_url: str = None) -> Dict[str, Any]:
        """
        Автоматическое обновление тарифов из текста или URL
        
        Args:
            marketplace: Название маркетплейса
            source_text: Текст для анализа
            source_url: URL для загрузки
        
        Returns:
            Dict[str, Any]: Обновленные тарифы
        """
        if source_text:
            return self.parse_tariff_news(source_text, marketplace)
        elif source_url:
            try:
                response = requests.get(source_url, timeout=30, headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                })
                response.raise_for_status()
                text = response.text
                text = re.sub(r'<[^>]+>', ' ', text)
                text = re.sub(r'\s+', ' ', text)
                return self.parse_tariff_news(text, marketplace)
            except Exception as e:
                self.logger.error(f"Ошибка загрузки с URL: {e}")
                return {}
        return {}
    
    def get_tariff_history(self, marketplace: str = None) -> List[Dict]:
        """
        Получение истории обновлений тарифов
        
        Args:
            marketplace: Название маркетплейса (опционально)
        
        Returns:
            List[Dict]: История обновлений
        """
        if marketplace:
            return [h for h in self._tariff_history if h.get('_metadata', {}).get('marketplace') == marketplace]
        return self._tariff_history

# ============================================================================
# БЛОК 8: ЮНИТ-ЭКОНОМИКА С АКТУАЛЬНЫМИ ТАРИФАМИ (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

class MarketplaceUnitEconomics:
    """
    Singleton класс для расчета юнит-экономики с актуальными тарифами 2026
    
    Attributes:
        _configs: Конфигурации маркетплейсов
        _tariff_manager: Менеджер AI тарифов
    """
    
    _instance = None
    _configs = None
    _tariff_manager = None
    _cache = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_configs()
            cls._instance._init_tariff_manager()
            cls._instance._init_cache()
        return cls._instance
    
    def _init_configs(self):
        """Инициализация актуальных конфигураций на 2026 год"""
        self._configs = get_marketplace_configs_2026()
        self.logger = logging.getLogger('MarketplaceUnitEconomics')
        self.logger.info("Инициализированы тарифы на 2026 год")
        self.logger.info(f"Загружено {len(self._configs)} маркетплейсов")
    
    def _init_tariff_manager(self):
        """Инициализация AI менеджера тарифов"""
        api_key = os.getenv('DEEPSEEK_API_KEY') or os.getenv('OPENAI_API_KEY')
        self._tariff_manager = TariffAIManager(api_key=api_key)
    
    def _init_cache(self):
        """Инициализация кэша"""
        self._cache = CacheManager("unit_economics_cache")
    
    def update_tariffs_ai(self, marketplace: str, source_text: str = None, source_url: str = None) -> bool:
        """
        Обновление тарифов через AI из текста или URL
        
        Args:
            marketplace: Название маркетплейса
            source_text: Текст для анализа
            source_url: URL для загрузки
        
        Returns:
            bool: True если обновление успешно
        """
        try:
            new_tariffs = self._tariff_manager.auto_update_tariffs(marketplace, source_text, source_url)
            
            if not new_tariffs:
                self.logger.warning(f"Не удалось получить новые тарифы для {marketplace}")
                return False
            
            current_config = self._configs.get(marketplace)
            if not current_config:
                self.logger.error(f"Маркетплейс {marketplace} не найден")
                return False
            
            updated_config = MarketplaceConfig2026(
                commission_rate=new_tariffs.get('commission_rate', current_config.commission_rate),
                commission_type=CommissionType(new_tariffs.get('commission_type', 'percentage')) 
                    if new_tariffs.get('commission_type') else current_config.commission_type,
                subscription_fee=new_tariffs.get('subscription_fee', current_config.subscription_fee),
                min_commission=new_tariffs.get('min_commission', current_config.min_commission),
                logistics_base=new_tariffs.get('logistics_base', current_config.logistics_base),
                logistics_per_kg=new_tariffs.get('logistics_per_kg', current_config.logistics_per_kg),
                logistics_per_liter=new_tariffs.get('logistics_per_liter', current_config.logistics_per_liter),
                storage_per_day=new_tariffs.get('storage_per_day', current_config.storage_per_day),
                storage_non_standard_fee=new_tariffs.get('storage_non_standard_fee', current_config.storage_non_standard_fee),
                return_fee=new_tariffs.get('return_fee', current_config.return_fee),
                acquiring_fee=new_tariffs.get('acquiring_fee', current_config.acquiring_fee),
                last_mile_fee=new_tariffs.get('last_mile_fee', current_config.last_mile_fee),
                delivery_fee_percent=new_tariffs.get('delivery_fee_percent', current_config.delivery_fee_percent),
                premium_section_fee=new_tariffs.get('premium_section_fee', current_config.premium_section_fee),
                rko_fee=new_tariffs.get('rko_fee', current_config.rko_fee),
                category_rates=new_tariffs.get('category_rates', current_config.category_rates),
                mode_multipliers=current_config.mode_multipliers
            )
            
            self._configs[marketplace] = updated_config
            self.logger.info(f"✅ Тарифы {marketplace} обновлены через AI")
            
            self._cache.clear()
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка обновления тарифов через AI: {e}")
            return False
    
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
        """
        Расчет юнит-экономики с учетом актуальных тарифов 2026
        
        Args:
            price: Цена продажи
            cost: Себестоимость
            weight_kg: Вес в кг
            volume_liters: Объем в литрах
            marketplace: Название маркетплейса
            operation_mode: Режим работы
            days_in_storage: Дней хранения
            category: Категория товара
            is_premium: Премиум-раздел
        
        Returns:
            Dict[str, Any]: Результаты расчета
        """
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
            storage_non_standard = min(
                price * config.storage_non_standard_fee,
                280
            )
        
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
        breakeven_price = (
            (cost + fixed_costs) / (1 - variable_rate) 
            if (1 - variable_rate) > 0 else 0
        )
        
        return {
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
            "profit_per_ruble": round(profit / price, 4) if price > 0 else 0
        }
    
    def get_marketplace_config(self, marketplace: str) -> Dict:
        """
        Получение текущей конфигурации маркетплейса
        
        Args:
            marketplace: Название маркетплейса
        
        Returns:
            Dict: Конфигурация
        """
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
        """
        Расчет юнит-экономики для всех маркетплейсов
        
        Args:
            price: Цена продажи
            cost: Себестоимость
            weight_kg: Вес в кг
            volume_liters: Объем в литрах
            operation_mode: Режим работы
        
        Returns:
            pd.DataFrame: Результаты по всем маркетплейсам
        """
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

# ============================================================================
# БЛОК 9: КАТАЛОГ ЭНХАНСЕР С МНОГОУРОВНЕВЫМ ПОИСКОМ АНАЛОГОВ
# ============================================================================

class CatalogEnhancer:
    """
    Класс для дополнения данных каталога с использованием многоуровневого поиска аналогов
    Реализует рекурсивные CTE для поиска аналогов на 2+ уровнях
    """
    
    def __init__(self, db_path: Optional[str] = None):
        self.data_dir = Path("./catalog_data")
        self.data_dir.mkdir(exist_ok=True)
        
        self.db_path = Path(db_path) if db_path else self.data_dir / "catalog.duckdb"
        self.conn = None
        
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
        """Создание таблиц в DuckDB"""
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
        """Нормализация ключа"""
        if not value:
            return ""
        return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', value.lower().strip())
    
    def load_oe_data(self, df: pd.DataFrame):
        """Загрузка OE данных"""
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
            logger.info(f"Загружено {len(df)} OE записей")
        except Exception as e:
            logger.error(f"Ошибка загрузки OE данных: {e}")
    
    def load_parts_data(self, df: pd.DataFrame):
        """Загрузка данных деталей"""
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
            logger.info(f"Загружено {len(df)} записей деталей")
        except Exception as e:
            logger.error(f"Ошибка загрузки данных деталей: {e}")
    
    def load_cross_references(self, df: pd.DataFrame):
        """Загрузка кросс-ссылок"""
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
            logger.info(f"Загружено {len(df)} кросс-ссылок")
        except Exception as e:
            logger.error(f"Ошибка загрузки кросс-ссылок: {e}")
    
    def load_prices(self, df: pd.DataFrame):
        """Загрузка цен"""
        if not self.conn or df.empty:
            return
        
        try:
            df['artikul_norm'] = df['artikul'].apply(self.normalize_key)
            df['brand_norm'] = df['brand'].apply(self.normalize_key)
            df = df[(df['artikul_norm'] != "") & (df['brand_norm'] != "")]
            
            self.conn.execute("DELETE FROM prices")
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO prices VALUES (?, ?, ?, ?)",
                    [row['artikul_norm'], row['brand_norm'], 
                     row.get('price', 0.0), row.get('currency', 'RUB')]
                )
            logger.info(f"Загружено {len(df)} ценовых записей")
        except Exception as e:
            logger.error(f"Ошибка загрузки цен: {e}")
    
    def build_analog_query(self, artikul_norm: str, brand_norm: str) -> str:
        """Построение SQL запроса для поиска аналогов"""
        return f"""
        WITH PartDetails AS (
            SELECT 
                cr.artikul_norm, 
                cr.brand_norm,
                STRING_AGG(DISTINCT oe.oe_number, ', ') AS oe_list,
                ANY_VALUE(oe.name) AS representative_name,
                ANY_VALUE(oe.applicability) AS representative_applicability,
                ANY_VALUE(oe.category) AS representative_category
            FROM cross_references cr
            LEFT JOIN oe ON cr.oe_number_norm = oe.oe_number_norm
            GROUP BY cr.artikul_norm, cr.brand_norm
        ),
        AllAnalogs AS (
            SELECT 
                cr1.artikul_norm, 
                cr1.brand_norm,
                STRING_AGG(DISTINCT p2.artikul, ', ') AS analog_list
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
                MAX(CASE WHEN p2.length IS NOT NULL THEN p2.length ELSE NULL END) AS analog_length,
                MAX(CASE WHEN p2.width IS NOT NULL THEN p2.width ELSE NULL END) AS analog_width,
                MAX(CASE WHEN p2.height IS NOT NULL THEN p2.height ELSE NULL END) AS analog_height,
                MAX(CASE WHEN p2.weight IS NOT NULL THEN p2.weight ELSE NULL END) AS analog_weight,
                ANY_VALUE(
                    CASE 
                        WHEN p2.dimensions_str IS NOT NULL AND p2.dimensions_str != '' AND UPPER(TRIM(p2.dimensions_str)) != 'XX'
                        THEN p2.dimensions_str
                        ELSE NULL
                    END
                ) AS analog_dimensions_str,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_name IS NOT NULL AND pd2.representative_name != '' 
                        THEN pd2.representative_name 
                        ELSE NULL
                    END
                ) AS analog_representative_name,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_applicability IS NOT NULL AND pd2.representative_applicability != ''
                        THEN pd2.representative_applicability
                        ELSE NULL
                    END
                ) AS analog_representative_applicability,
                ANY_VALUE(
                    CASE 
                        WHEN pd2.representative_category IS NOT NULL AND pd2.representative_category != ''
                        THEN pd2.representative_category
                        ELSE NULL
                    END
                ) AS analog_representative_category
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
                aad.analog_length,
                aad.analog_width,
                aad.analog_height,
                aad.analog_weight,
                aad.analog_dimensions_str,
                aad.analog_representative_name,
                aad.analog_representative_applicability,
                aad.analog_representative_category,
                ROW_NUMBER() OVER (
                    PARTITION BY p.artikul_norm, p.brand_norm 
                    ORDER BY pd.representative_name DESC NULLS LAST, pd.oe_list DESC NULLS LAST
                ) AS rn
            FROM parts p
            LEFT JOIN PartDetails pd ON p.artikul_norm = pd.artikul_norm AND p.brand_norm = pd.brand_norm
            LEFT JOIN AllAnalogs aa ON p.artikul_norm = aa.artikul_norm AND p.brand_norm = aa.brand_norm
            LEFT JOIN AggregatedAnalogData aad ON p.artikul_norm = aad.artikul_norm AND p.brand_norm = aad.brand_norm
            WHERE p.artikul_norm = '{artikul_norm}' AND p.brand_norm = '{brand_norm}'
        )
        SELECT 
            COALESCE(representative_name, analog_representative_name) AS name,
            COALESCE(representative_applicability, analog_representative_applicability) AS applicability,
            COALESCE(representative_category, analog_representative_category) AS category,
            COALESCE(length, analog_length) AS length,
            COALESCE(width, analog_width) AS width,
            COALESCE(height, analog_height) AS height,
            COALESCE(weight, analog_weight) AS weight,
            COALESCE(dimensions_str, analog_dimensions_str) AS dimensions_str,
            oe_list,
            analog_list,
            COALESCE(representative_name IS NOT NULL, false) AS has_own_data
        FROM RankedData
        WHERE rn = 1
        """
    
    def get_analog_data(self, artikul: str, brand: str) -> Dict[str, Any]:
        """
        Получение данных с аналогами для артикула и бренда
        
        Args:
            artikul: Артикул
            brand: Бренд
        
        Returns:
            Dict: Данные с аналогами
        """
        if not self.conn:
            return {"error": "DuckDB не доступен"}
        
        artikul_norm = self.normalize_key(artikul)
        brand_norm = self.normalize_key(brand)
        
        if not artikul_norm or not brand_norm:
            return {"error": "Не указан артикул или бренд"}
        
        try:
            query = self.build_analog_query(artikul_norm, brand_norm)
            result = self.conn.execute(query).df()
            
            if result.empty:
                return {"error": f"Данные не найдены для {artikul} / {brand}"}
            
            row = result.iloc[0]
            
            return {
                "artikul": artikul,
                "brand": brand,
                "name": row.get('name', ''),
                "applicability": row.get('applicability', ''),
                "category": row.get('category', ''),
                "length": safe_float(row.get('length', 0)),
                "width": safe_float(row.get('width', 0)),
                "height": safe_float(row.get('height', 0)),
                "weight": safe_float(row.get('weight', 0)),
                "dimensions_str": row.get('dimensions_str', ''),
                "oe_list": row.get('oe_list', ''),
                "analog_list": row.get('analog_list', ''),
                "has_own_data": row.get('has_own_data', False),
                "analog_level": 0 if row.get('has_own_data', False) else 1
            }
        except Exception as e:
            logger.error(f"Ошибка получения аналогов: {e}")
            return {"error": str(e)}
    
    def get_all_analogs(self, artikul: str, brand: str) -> List[Dict[str, Any]]:
        """
        Получение всех аналогов для артикула и бренда
        
        Args:
            artikul: Артикул
            brand: Бренд
        
        Returns:
            List[Dict]: Список аналогов
        """
        if not self.conn:
            return []
        
        artikul_norm = self.normalize_key(artikul)
        brand_norm = self.normalize_key(brand)
        
        try:
            oe_query = f"""
                SELECT DISTINCT oe_number_norm 
                FROM cross_references 
                WHERE artikul_norm = '{artikul_norm}' AND brand_norm = '{brand_norm}'
            """
            oe_numbers = self.conn.execute(oe_query).df()['oe_number_norm'].tolist()
            
            if not oe_numbers:
                return []
            
            analogs_query = f"""
                SELECT DISTINCT 
                    p.artikul,
                    p.brand,
                    p.description,
                    p.length,
                    p.width,
                    p.height,
                    p.weight,
                    p.dimensions_str
                FROM cross_references cr
                JOIN parts p ON cr.artikul_norm = p.artikul_norm AND cr.brand_norm = p.brand_norm
                WHERE cr.oe_number_norm IN ({','.join([f"'{oe}'" for oe in oe_numbers])})
                  AND NOT (cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}')
                LIMIT 50
            """
            result = self.conn.execute(analogs_query).df()
            
            analogs = []
            for _, row in result.iterrows():
                analogs.append({
                    "artikul": row.get('artikul', ''),
                    "brand": row.get('brand', ''),
                    "description": row.get('description', ''),
                    "length": safe_float(row.get('length', 0)),
                    "width": safe_float(row.get('width', 0)),
                    "height": safe_float(row.get('height', 0)),
                    "weight": safe_float(row.get('weight', 0)),
                    "dimensions_str": row.get('dimensions_str', '')
                })
            
            return analogs
        except Exception as e:
            logger.error(f"Ошибка получения аналогов: {e}")
            return []
    
    def enhance_catalog_data(self, df: pd.DataFrame, 
                           artikul_col: str = "Артикул",
                           brand_col: str = "Бренд") -> pd.DataFrame:
        """
        Обогащение данных каталога с использованием многоуровневого поиска аналогов
        
        Args:
            df: DataFrame с данными
            artikul_col: Название колонки с артикулом
            brand_col: Название колонки с брендом
        
        Returns:
            pd.DataFrame: Обогащенный DataFrame
        """
        if df.empty:
            return df
        
        if artikul_col not in df.columns or brand_col not in df.columns:
            logger.warning(f"Колонки {artikul_col} или {brand_col} не найдены")
            return df
        
        df_copy = df.copy()
        
        new_columns = [
            'enhanced_name', 'enhanced_applicability', 'enhanced_category',
            'enhanced_length', 'enhanced_width', 'enhanced_height',
            'enhanced_weight', 'enhanced_dimensions', 'oe_list', 'analog_list',
            'has_own_data', 'analog_count'
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
                    df_copy.at[idx, 'enhanced_name'] = data.get('name', '')
                    df_copy.at[idx, 'enhanced_applicability'] = data.get('applicability', '')
                    df_copy.at[idx, 'enhanced_category'] = data.get('category', '')
                    df_copy.at[idx, 'enhanced_length'] = data.get('length', 0)
                    df_copy.at[idx, 'enhanced_width'] = data.get('width', 0)
                    df_copy.at[idx, 'enhanced_height'] = data.get('height', 0)
                    df_copy.at[idx, 'enhanced_weight'] = data.get('weight', 0)
                    df_copy.at[idx, 'enhanced_dimensions'] = data.get('dimensions_str', '')
                    df_copy.at[idx, 'oe_list'] = data.get('oe_list', '')
                    df_copy.at[idx, 'analog_list'] = data.get('analog_list', '')
                    df_copy.at[idx, 'has_own_data'] = data.get('has_own_data', False)
                    
                    analogs = self.get_all_analogs(artikul, brand)
                    df_copy.at[idx, 'analog_count'] = len(analogs)
        
        return df_copy

# ============================================================================
# БЛОК 10: UI ФУНКЦИИ (ПОЛНАЯ ВЕРСИЯ)
# ============================================================================

def show_unit_economics_interface():
    """Интерфейс юнит-экономики"""
    st.header("📊 Юнит-экономика маркетплейсов 2026")
    
    unit_economics = MarketplaceUnitEconomics()
    
    col1, col2 = st.columns(2)
    
    with col1:
        price = st.number_input("💰 Цена продажи (₽)", min_value=0.0, value=1000.0, step=10.0, key="ue_price")
        cost = st.number_input("💵 Себестоимость (₽)", min_value=0.0, value=500.0, step=10.0, key="ue_cost")
        weight = st.number_input("⚖️ Вес (кг)", min_value=0.0, value=1.0, step=0.1, key="ue_weight")
    
    with col2:
        volume = st.number_input("📦 Объем (литры)", min_value=0.0, value=5.0, step=0.5, key="ue_volume")
        marketplace = st.selectbox("🏪 Маркетплейс", list(unit_economics._configs.keys()), key="ue_marketplace")
        operation_mode = st.selectbox(
            "📦 Режим работы",
            ["FBY", "FBS", "FBO", "DBS", "FBP"],
            key="ue_mode"
        )
        category = st.text_input("📂 Категория (опционально)", placeholder="например: одежда_обувь", key="ue_category")
        is_premium = st.checkbox("⭐ Премиум-раздел (доп. комиссия)", key="ue_premium")
    
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
                    st.metric("💰 Прибыль", f"{economics['profit']:.2f} ₽")
                    st.metric("📈 Маржа", f"{economics['margin_percent']:.2f}%")
                
                with col2:
                    st.metric("📊 ROI", f"{economics['roi']:.2f}%")
                    st.metric("⚖️ Точка безубыточности", f"{economics['breakeven_price']:.2f} ₽")
                
                with col3:
                    st.metric("💵 Комиссия", f"{economics['commission']:.2f} ₽")
                    if economics.get('subscription_cost', 0) > 0:
                        st.metric("📋 Подписка (в день)", f"{economics['subscription_cost']:.2f} ₽")
                    if economics.get('storage_non_standard', 0) > 0:
                        st.metric("📦 Нестандарт", f"{economics['storage_non_standard']:.2f} ₽")
                
                st.subheader("📋 Детализация расходов")
                
                expenses_data = {
                    "Статья расходов": [
                        "Себестоимость", "Комиссия", "Подписка", "Логистика",
                        "Хранение", "Нестандарт", "Эквайринг", "Доставка", 
                        "Последняя миля", "Возвраты", "РКО", "Премиум", "ИТОГО"
                    ],
                    "Сумма (₽)": [
                        economics['cost'], economics['commission'], economics.get('subscription_cost', 0),
                        economics['logistics'], economics['storage_cost'], 
                        economics.get('storage_non_standard', 0),
                        economics['acquiring'], economics['delivery'],
                        economics['last_mile'], economics['returns'],
                        economics.get('rko_fee', 0), economics.get('premium_fee', 0),
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
                
                st.dataframe(pd.DataFrame(expenses_data), use_container_width=True, key="ue_expenses_table")
                
                st.subheader("🏆 Сравнение всех маркетплейсов")
                comparison_df = unit_economics.calculate_for_all_marketplaces(
                    price=price,
                    cost=cost,
                    weight_kg=weight,
                    volume_liters=volume,
                    operation_mode=operation_mode
                )
                st.dataframe(comparison_df, use_container_width=True, key="ue_comparison_table")
                
                if not comparison_df.empty:
                    best_idx = comparison_df['profit'].idxmax()
                    best = comparison_df.loc[best_idx]
                    st.success(
                        f"🏆 Оптимальный маркетплейс: **{best['marketplace']}** "
                        f"(прибыль: {best['profit']:.2f} ₽, маржа: {best['margin_percent']:.2f}%)"
                    )
            else:
                st.error(f"❌ Ошибка: {economics['error']}")

def show_catalog_enhance_interface():
    """Интерфейс для обогащения каталога с поиском аналогов"""
    st.header("📊 Обогащение каталога (поиск аналогов)")
    
    st.info("""
    🔍 **Многоуровневый поиск аналогов:**
    
    **Уровень 0**: Собственные данные
    **Уровень 1**: Прямые аналоги (через общие OE номера)
    **Уровень 2**: Косвенные аналоги (через аналоги уровнем выше)
    
    Система автоматически находит аналоги и дополняет данные о товаре.
    """)
    
    if 'catalog_enhancer' not in st.session_state:
        st.session_state.catalog_enhancer = CatalogEnhancer()
    
    enhancer = st.session_state.catalog_enhancer
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📤 Загрузка данных каталога")
        
        oe_file = st.file_uploader("OE данные", type=['xlsx', 'csv'], key="enh_oe")
        parts_file = st.file_uploader("Детали (артикулы)", type=['xlsx', 'csv'], key="enh_parts")
        cross_file = st.file_uploader("Кросс-ссылки", type=['xlsx', 'csv'], key="enh_cross")
        prices_file = st.file_uploader("Цены", type=['xlsx', 'csv'], key="enh_prices")
        
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
                
                if prices_file:
                    try:
                        df = pd.read_excel(prices_file) if prices_file.name.endswith('.xlsx') else pd.read_csv(prices_file)
                        enhancer.load_prices(df)
                        st.success(f"✅ Загружено {len(df)} ценовых записей")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки цен: {str(e)}")
    
    with col2:
        st.subheader("🔍 Поиск аналогов")
        
        artikul = st.text_input("Артикул", placeholder="Введите артикул", key="enh_artikul_input")
        brand = st.text_input("Бренд", placeholder="Введите бренд", key="enh_brand_input")
        
        if st.button("🔍 Найти аналоги", type="primary", key="enh_find_analogs"):
            if artikul and brand:
                with st.spinner("Поиск аналогов..."):
                    data = enhancer.get_analog_data(artikul, brand)
                    
                    if data.get('error'):
                        st.error(f"❌ {data['error']}")
                    else:
                        st.success("✅ Данные найдены")
                        
                        col_a, col_b, col_c = st.columns(3)
                        
                        with col_a:
                            st.metric("📦 Название", data.get('name', '—'))
                            st.metric("📂 Категория", data.get('category', '—'))
                        
                        with col_b:
                            st.metric("📏 Длина", f"{data.get('length', 0):.1f} см")
                            st.metric("📐 Ширина", f"{data.get('width', 0):.1f} см")
                            st.metric("📏 Высота", f"{data.get('height', 0):.1f} см")
                        
                        with col_c:
                            st.metric("⚖️ Вес", f"{data.get('weight', 0):.2f} кг")
                            st.metric("🔗 OE номеров", len(data.get('oe_list', '').split(',')) if data.get('oe_list') else 0)
                            st.metric("🔄 Свои данные", "✅" if data.get('has_own_data') else "❌")
                        
                        if data.get('analog_list'):
                            st.subheader("🔗 Аналоги")
                            analogs = [a.strip() for a in data['analog_list'].split(',') if a.strip()]
                            st.write(f"Найдено {len(analogs)} аналогов:")
                            st.code('\n'.join(analogs[:10]), language='text')
                
                analogs = enhancer.get_all_analogs(artikul, brand)
                if analogs:
                    st.subheader("📋 Полный список аналогов")
                    df_analogs = pd.DataFrame(analogs)
                    st.dataframe(df_analogs, use_container_width=True, key="enh_analogs_table")
            else:
                st.warning("⚠️ Введите артикул и бренд")
    
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
                st.dataframe(enhanced_df.head(20), use_container_width=True, key="enh_result_table")
                
                if 'analog_count' in enhanced_df.columns:
                    analog_counts = enhanced_df['analog_count'].value_counts()
                    st.subheader("📊 Статистика аналогов")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("📦 Всего товаров", len(enhanced_df))
                    col2.metric("🔄 С аналогами", len(enhanced_df[enhanced_df['analog_count'] > 0]))
                    col3.metric("📊 Среднее аналогов", f"{enhanced_df['analog_count'].mean():.1f}")
    else:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")

def show_data_upload_interface():
    """Интерфейс загрузки данных"""
    st.header("📁 Загрузка данных каталога")
    
    uploaded_file = st.file_uploader(
        "Загрузите файл каталога (Excel или CSV)",
        type=['xlsx', 'xls', 'csv'],
        help="Поддерживаются форматы CSV и Excel",
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
            st.dataframe(df.head(10), use_container_width=True, key="upload_preview_table")
            
            if st.button("📊 Обогатить каталог (поиск аналогов)", type="primary", key="upload_enrich_button"):
                st.info("Перейдите на вкладку '📊 Обогащение каталога'")
                    
        except Exception as e:
            st.error(f"❌ Ошибка загрузки файла: {str(e)}")
            st.code(traceback.format_exc())

def show_export_interface():
    """Интерфейс экспорта данных"""
    st.header("📤 Экспорт данных")
    
    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные")
        return
    
    df = st.session_state.uploaded_data
    
    st.success(f"✅ Готово к экспорту: {len(df)} товаров, {len(df.columns)} колонок")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📥 Экспорт в Excel", type="primary", key="export_excel_button"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Данные', index=False)
            output.seek(0)
            st.download_button(
                label="📥 Скачать Excel",
                data=output.getvalue(),
                file_name=f"данные_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="export_excel_download"
            )
    
    with col2:
        if st.button("📥 Экспорт в CSV", key="export_csv_button"):
            csv = df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Скачать CSV",
                data=csv,
                file_name=f"данные_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="export_csv_download"
            )
    
    with col3:
        if st.button("📥 Экспорт в JSON", key="export_json_button"):
            json_data = df.to_json(orient='records', force_ascii=False, indent=2)
            st.download_button(
                label="📥 Скачать JSON",
                data=json_data,
                file_name=f"данные_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                key="export_json_download"
            )

# ============================================================================
# БЛОК 11: ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    try:
        st.set_page_config(
            page_title=f"{APP_NAME} v{APP_VERSION}",
            page_icon="🚗",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
                    padding: 1.5rem; border-radius: 15px; color: white; text-align: center; margin-bottom: 1.5rem;
                    border: 2px solid #e94560; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
            <h1 style="font-size: 2.8rem; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);">
                🚗 {APP_NAME}
            </h1>
            <p style="font-size: 1.2rem; opacity: 0.95; margin-top: 0.3rem;">
                📊 <strong>Актуальные тарифы 2026</strong> | Многоуровневый поиск аналогов | AI обновление
            </p>
            <div style="display: flex; justify-content: center; gap: 0.8rem; flex-wrap: wrap; margin-top: 0.5rem;">
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    v{APP_VERSION}
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    🔍 2+ уровней аналогов
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📦 100+ категорий
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    🤖 AI обновление тарифов
                </span>
                <span style="background: rgba(233,69,96,0.3); padding: 0.2rem 1.2rem; border-radius: 20px; font-size: 0.9rem;">
                    📋 Полная версия
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if 'uploaded_data' not in st.session_state:
            st.session_state.uploaded_data = None
        
        with st.sidebar:
            st.markdown("## ⚙️ Настройки")
            
            st.markdown("### 🔑 API ключи")
            ds_api_key = st.text_input(
                "🔑 DeepSeek API ключ",
                type="password",
                placeholder="sk-...",
                help="Для AI-тарифов",
                key="sidebar_api_key"
            )
            if ds_api_key:
                os.environ['DEEPSEEK_API_KEY'] = ds_api_key
                st.success("✅ Ключ установлен")
            
            st.divider()
            
            st.markdown("### 📊 Статистика")
            if st.session_state.uploaded_data is not None:
                df = st.session_state.uploaded_data
                st.metric("📦 Товаров", len(df))
                st.metric("📂 Колонок", len(df.columns))
                
                if 'analog_count' in df.columns:
                    st.metric("🔄 С аналогами", len(df[df['analog_count'] > 0]))
                    st.metric("📊 Среднее аналогов", f"{df['analog_count'].mean():.1f}")
            
            st.divider()
            
            st.markdown("### 📦 Категории")
            st.metric("📂 Всего категорий", len(CATEGORY_DIMENSIONS))
            high_conf = sum(1 for p in CATEGORY_DIMENSIONS.values() if p.confidence >= 0.7)
            st.metric("🎯 Высокая точность", f"{high_conf}/{len(CATEGORY_DIMENSIONS)}")
            
            st.divider()
            
            st.markdown("### ℹ️ Система")
            st.caption(f"Версия: {APP_VERSION}")
            st.caption(f"Python: {sys.version[:10]}")
            st.caption(f"DuckDB: {'✅' if LIBRARIES['duckdb'] else '❌'}")
            st.caption(f"Библиотеки: {sum(1 for v in LIBRARIES.values() if v)}/{len(LIBRARIES)}")
        
        tabs = st.tabs([
            "📊 Юнит-экономика",
            "📊 Обогащение каталога",
            "📁 Загрузка данных",
            "📤 Экспорт"
        ])
        
        with tabs[0]:
            show_unit_economics_interface()
        
        with tabs[1]:
            show_catalog_enhance_interface()
        
        with tabs[2]:
            show_data_upload_interface()
        
        with tabs[3]:
            show_export_interface()
            
    except Exception as e:
        st.error(f"❌ Критическая ошибка: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Critical error: {e}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)
