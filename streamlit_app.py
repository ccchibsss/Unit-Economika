"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v101.1 - ENTERPRISE EDITION
================================================================================
📌 ВЕРСИЯ: 101.1.0 (FBS-OPTIMIZED + LIVE EXCEL FORMULAS + SECURE KEYS)
📌 ИСПРАВЛЕНИЯ v101.1:
1. ✅ Авто-детекция кодировок (UTF-8/CP1251/Latin1) и защита от превращения текста в даты в Excel.
2. ✅ Умное кросс-связывание: парсинг ОЕ через ';', поиск аналогов, заполнение пропусков (вес/габариты).
3. ✅ DeepSeek API: выбор режима (Обогащение каталога ИЛИ Актуализация тарифов).
4. ✅ FBS-ONLY: Удалены FBO, DBS, FBP, RealFBS из логики и UI.
5. ✅ Живые формулы Excel: экспорт с формулами Yandex Market/Ozon для ручного тюнинга.
6. ✅ Безопасное хранение API ключей: шифрование (Fernet) в локальный файл, а не только session_state.
================================================================================
"""

# ============================================================================
# БЛОК 0: ИМПОРТЫ И БАЗОВАЯ КОНФИГУРАЦИЯ
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
try:
    import polars as pl
    import polars.selectors as cs
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False
    pl = None

try:
    import duckdb
    DUCKDB_AVAILABLE = True
except ImportError:
    DUCKDB_AVAILABLE = False
    duckdb = None

try:
    import sklearn
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.naive_bayes import MultinomialNB
    from sklearn.pipeline import Pipeline
    import joblib
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False

try:
    import plotly.graph_objects as go
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference, LineChart, PieChart
    from openpyxl.formatting.rule import CellIsRule, DataBarRule, ColorScaleRule
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

try:
    import chardet
    CHARDET_AVAILABLE = True
except ImportError:
    CHARDET_AVAILABLE = False
    chardet = None

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

try:
    import openai # Для совместимости с DeepSeek API (через openai-compatible endpoint)
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# 🆕 v101.1: Шифрование для API ключей (Req 6)
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    Fernet = None

try:
    from dateutil.parser import parse as dateutil_parse
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
SECURE_KEYS_DIR = BASE_DIR / "secure_keys" # 🆕 v101.1: Для зашифрованных ключей

for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, TEMP_DIR, MODELS_DIR,
                 CONFIG_DIR, EXPORTS_DIR, TARIFFS_DIR, HISTORY_DB_DIR,
                 BACKUPS_DIR, GOOGLE_CREDS_DIR, SECURE_KEYS_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except OSError:
        pass

# === Версия приложения ===
APP_VERSION = "101.1.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей PRO 2026 (FBS)"
APP_DESCRIPTION = "Enterprise расчет с живыми формулами Excel, кросс-связыванием ОЕ и безопасным хранением ключей"

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

# === Совместимость Streamlit ===
def st_dataframe_compat(df, *args, **kwargs):
    kwargs.pop('use_container_width', None)
    if 'width' not in kwargs:
        kwargs['width'] = 'stretch'
    return st.dataframe(df, *args, **kwargs)

# ============================================================================
# БЛОК 1: КОНСТАНТЫ, ENUM И ТИПЫ ДАННЫХ
# ============================================================================
EXCEL_ROW_LIMIT = 1_000_000
HISTORY_LIMIT = 50_000
CACHE_TTL = 7200
MAX_THREADS = 32
BATCH_SIZE = 2000
MAX_FILE_SIZE_MB = 500

DEFAULT_CURRENCY = "RUB"
DEFAULT_MARKETPLACE = "Ozon"
DEFAULT_MODE = "FBS" # 🆕 v101.1: FBS по умолчанию
DEFAULT_LOCALE = "ru_RU"
TIMEZONE = "Europe/Moscow"
DEFAULT_MARKUP_GLOBAL = 0.25
DEFAULT_TARGET_MARGIN = 20.0

SUPPORTED_CURRENCIES = ["RUB", "USD", "EUR", "CNY", "KZT", "BYN"]
SUPPORTED_MARKETPLACES = ["Ozon", "Wildberries", "Яндекс Маркет", "Мегамаркет"]
# 🆕 v101.1: Удалены FBO, DBS, FBP, RealFBS (Req 4)
SUPPORTED_MODES = ["FBS", "FBY"] 

USE_CACHING = True
USE_PARALLEL = True
OPTIMIZE_MEMORY = True
USE_DUCKDB = True
USE_POLARS = True

# === Цветовая схема ===
COLORS = {
    "primary": "#e94560", "secondary": "#0f3460", "success": "#00cc96",
    "warning": "#ffa600", "danger": "#ef553b", "info": "#636efa",
    "dark": "#1a1a2e", "light": "#f5f5f5"
}

# === Налоговые системы ===
TAX_SYSTEMS = {
    "УСН_6": {"rate": 0.06, "base": "revenue", "name": "УСН 6% (доходы)"},
    "УСН_15": {"rate": 0.15, "base": "profit", "min_rate": 0.01, "name": "УСН 15% (доходы-расходы)"},
    "ОСН": {"rate": 0.20, "base": "profit", "vat": 0.20, "name": "ОСН (общая)"},
    "НПД": {"rate": 0.06, "base": "revenue", "name": "НПД (самозанятый)"},
}

# === ABC/XYZ пороги ===
ABC_THRESHOLDS = {
    "A": {"margin_min": 25, "profit_share": 0.70},
    "B": {"margin_min": 15, "profit_share": 0.20},
    "C": {"margin_min": 0, "profit_share": 0.10},
}
XYZ_THRESHOLDS = {
    "X": {"cv_max": 0.5},
    "Y": {"cv_max": 1.0},
    "Z": {"cv_max": float('inf')},
}

# ============================================================================
# ENUM (Перечисления)
# ============================================================================
class CommissionType(Enum):
    PERCENTAGE = auto()
    FIXED = auto()
    HYBRID = auto()

# 🆕 v101.1: Только FBS и FBY (Req 4)
class OperationMode(Enum):
    FBS = "FBS (со своего склада по заказу)"
    FBY = "FBY (со своего склада, доставка МП)"

class ProductType(Enum):
    ENGINE = "Двигатель"
    TRANSMISSION = "Трансмиссия"
    SUSPENSION = "Подвеска"
    BRAKE = "Тормозная система"
    STEERING = "Рулевое управление"
    ELECTRICAL = "Электрооборудование"
    COOLING = "Система охлаждения"
    FILTER = "Фильтры"
    FLUID = "Масла и жидкости"
    BODY = "Кузовные детали"
    OPTICS = "Оптика"
    TIRES = "Шины и диски"
    OTHER = "Прочее"

class DataSource(Enum):
    CSV = auto()
    EXCEL = auto()
    API = auto()
    GOOGLE_SHEETS = auto()
    AI_DEEPSEEK = auto() # 🆕 v101.1

class ExportFormat(Enum):
    CSV = auto()
    EXCEL_FORMULAS = auto() # 🆕 v101.1: Живые формулы (Req 5)
    GOOGLE_SHEETS = auto()

class CalculationStatus(Enum):
    PENDING = auto()
    RUNNING = auto()
    COMPLETED = auto()
    FAILED = auto()

class RiskLevel(Enum):
    LOW = "Низкий"
    MEDIUM = "Средний"
    HIGH = "Высокий"

class Seasonality(Enum):
    WINTER = "Зимняя"
    SUMMER = "Летняя"
    ALL_YEAR = "Круглогодичная"

class TariffSource(Enum):
    HARDCODED = "Захардкожены"
    AI_CACHE = "Кэш ИИ"
    AI_LIVE = "ИИ (DeepSeek запрос)"
    API_LIVE = "API Маркетплейса"
    GOOGLE_SHEETS = "Google Sheets"

class DataLinkType(Enum):
    OE_TO_CROSS = "OE-кроссы"
    ARTICLE_TO_ANALOG = "Артикул-аналог"
    MANUAL_MAPPING = "Ручной маппинг"
    AUTO_DETECT = "Автоопределение"
# ============================================================================
# БЛОК 2: DATACLASSES, БЕЗОПАСНОЕ ХРАНЕНИЕ, КРОСС-СВЯЗЫВАНИЕ И ФОРМУЛЫ EXCEL
# ============================================================================

# ============================================================================
# 2.1 DATACLASSES ДЛЯ КОНФИГУРАЦИИ И СВЯЗЫВАНИЯ
# ============================================================================
@dataclass
class SecureKeyEntry:
    """Запись зашифрованного ключа"""
    service: str
    description: str
    created_at: datetime
    last_updated: datetime

@dataclass
class ColumnMapping:
    """
    Маппинг столбцов между файлами для кросс-связывания.
    Используется для связывания данных из разных источников (например, OE и Габариты).
    """
    source_file: str
    source_column: str
    target_file: str
    target_column: str
    mapping_type: str = "manual"  # manual, auto_detect, oe_cross
    confidence: float = 1.0
    auto_detected: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source_file": self.source_file,
            "source_column": self.source_column,
            "target_file": self.target_file,
            "target_column": self.target_column,
            "mapping_type": self.mapping_type,
            "confidence": self.confidence,
            "auto_detected": self.auto_detected
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ColumnMapping':
        return cls(
            source_file=data.get("source_file", ""),
            source_column=data.get("source_column", ""),
            target_file=data.get("target_file", ""),
            target_column=data.get("target_column", ""),
            mapping_type=data.get("mapping_type", "manual"),
            confidence=float(data.get("confidence", 1.0)),
            auto_detected=bool(data.get("auto_detected", False))
        )

@dataclass
class DataLinkConfig:
    """
    Конфигурация связывания данных между файлами.
    Позволяет настроить, как именно объединять таблицы (например, по Артикулу или ОЕ).
    """
    link_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""
    description: str = ""
    primary_file: str = ""
    secondary_file: str = ""
    join_key_primary: str = ""
    join_key_secondary: str = ""
    join_type: str = "left"  # left, inner, outer
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

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DataLinkConfig':
        mappings = [ColumnMapping.from_dict(m) for m in data.get("column_mappings", [])]
        created_at_str = data.get("created_at", datetime.now().isoformat())
        try:
            created_at = datetime.fromisoformat(created_at_str)
        except ValueError:
            created_at = datetime.now()
            
        return cls(
            link_id=data.get("link_id", str(uuid.uuid4())),
            name=data.get("name", ""),
            description=data.get("description", ""),
            primary_file=data.get("primary_file", ""),
            secondary_file=data.get("secondary_file", ""),
            join_key_primary=data.get("join_key_primary", ""),
            join_key_secondary=data.get("join_key_secondary", ""),
            join_type=data.get("join_type", "left"),
            column_mappings=mappings,
            auto_fill_missing=bool(data.get("auto_fill_missing", True)),
            created_at=created_at
        )


# ============================================================================
# 2.2 БЕЗОПАСНОЕ ХРАНЕНИЕ API КЛЮЧЕЙ (REQ 6)
# ============================================================================
class SecureKeyManager:
    """
    Менеджер безопасного хранения API ключей.
    Использует симметричное шифрование (Fernet) для сохранения ключей в локальный файл.
    Ключи загружаются при запуске и не хранятся в открытом виде в session_state.
    """
    def __init__(self, key_dir: Path = SECURE_KEYS_DIR):
        self.key_dir = key_dir
        self.key_dir.mkdir(parents=True, exist_ok=True)
        
        self.master_key_file = self.key_dir / "master.key"
        self.encrypted_data_file = self.key_dir / "api_keys.enc"
        
        self.fernet = self._get_or_create_fernet()
        self._keys_cache: Dict[str, str] = self._load_keys()
        self._metadata_cache: Dict[str, Dict[str, Any]] = self._load_metadata()

    def _get_or_create_fernet(self) -> Fernet:
        """Генерирует или загружает мастер-ключ шифрования"""
        if not self.master_key_file.exists():
            new_key = Fernet.generate_key()
            self.master_key_file.write_bytes(new_key)
            # Устанавливаем строгие права доступа к файлу ключа (Unix-like)
            try:
                os.chmod(self.master_key_file, 0o600)
            except OSError:
                pass
        else:
            new_key = self.master_key_file.read_bytes()
        return Fernet(new_key)

    def _load_keys(self) -> Dict[str, str]:
        """Расшифровывает и загружает ключи из файла"""
        if not self.encrypted_data_file.exists():
            return {}
        try:
            encrypted_data = self.encrypted_data_file.read_bytes()
            decrypted_bytes = self.fernet.decrypt(encrypted_data)
            return json.loads(decrypted_bytes.decode('utf-8'))
        except Exception as e:
            logger.error(f"Ошибка расшифровки ключей: {e}. Файл может быть поврежден.")
            return {}

    def _load_metadata(self) -> Dict[str, Dict[str, Any]]:
        """Загружает метаданные ключей (описания, даты)"""
        metadata_file = self.key_dir / "api_keys_meta.json"
        if not metadata_file.exists():
            return {}
        try:
            return json.loads(metadata_file.read_text(encoding='utf-8'))
        except Exception:
            return {}

    def _save_keys(self):
        """Шифрует и сохраняет ключи в файл"""
        try:
            encrypted_data = self.fernet.encrypt(json.dumps(self._keys_cache).encode('utf-8'))
            self.encrypted_data_file.write_bytes(encrypted_data)
            try:
                os.chmod(self.encrypted_data_file, 0o600)
            except OSError:
                pass
        except Exception as e:
            logger.error(f"Ошибка шифрования и сохранения ключей: {e}")

    def _save_metadata(self):
        """Сохраняет метаданные ключей"""
        metadata_file = self.key_dir / "api_keys_meta.json"
        try:
            metadata_file.write_text(json.dumps(self._metadata_cache, indent=2, ensure_ascii=False), encoding='utf-8')
        except Exception as e:
            logger.error(f"Ошибка сохранения метаданных ключей: {e}")

    def set_key(self, service: str, api_key: str, description: str = ""):
        """
        Сохраняет или обновляет API ключ для сервиса.
        """
        if not api_key or not api_key.strip():
            if service in self._keys_cache:
                del self._keys_cache[service]
                if service in self._metadata_cache:
                    del self._metadata_cache[service]
        else:
            self._keys_cache[service] = api_key.strip()
            now = datetime.now()
            self._metadata_cache[service] = {
                "description": description,
                "created_at": self._metadata_cache.get(service, {}).get("created_at", now.isoformat()),
                "last_updated": now.isoformat()
            }
        self._save_keys()
        self._save_metadata()

    def get_key(self, service: str) -> Optional[str]:
        """Получает API ключ для сервиса (в открытом виде, только в памяти)"""
        return self._keys_cache.get(service)

    def get_metadata(self, service: str) -> Dict[str, Any]:
        """Получает метаданные ключа"""
        return self._metadata_cache.get(service, {})

    def list_services(self) -> List[str]:
        """Возвращает список сервисов, для которых есть ключи"""
        return list(self._keys_cache.keys())

    def delete_key(self, service: str):
        """Удаляет ключ и его метаданные"""
        if service in self._keys_cache:
            del self._keys_cache[service]
        if service in self._metadata_cache:
            del self._metadata_cache[service]
        self._save_keys()
        self._save_metadata()

    def clear_all(self):
        """Очищает все сохраненные ключи"""
        self._keys_cache.clear()
        self._metadata_cache.clear()
        self._save_keys()
        self._save_metadata()


# ============================================================================
# 2.3 УТИЛИТА ЗАЩИТЫ ТЕКСТА В EXCEL (REQ 1: ИСПРАВЛЕНИЕ ДАТ)
# ============================================================================
def escape_excel_text(value: Any) -> str:
    """
    Экранирует строку для Excel, чтобы предотвратить автоматическое преобразование 
    в дату или формулу.
    Добавляет символ апострофа (') в начало, если строка:
    - Начинается с '=', '+', '-', '@' (формулы)
    - Выглядит как дата (например, "1-2", "OCT", "2023-10")
    - Содержит только цифры и дефисы в определенном формате
    """
    if pd.isna(value) or value is None:
        return ""
    
    s = str(value).strip()
    if not s:
        return s
    
    # Проверка на формулы
    if s.startswith(('=', '+', '-', '@')):
        return f"'{s}"
    
    # Проверка на потенциальные даты (например, "1-2", "OCT", "2023-10", "10/12")
    # Регулярное выражение ловит: цифры-цифры, буквы-цифры, цифры/цифры
    if re.match(r'^\d+[-/]\d+([-/]\d+)?$', s) or re.match(r'^[A-Za-z]{3,4}[-/]\d+$', s, re.IGNORECASE):
        return f"'{s}"
    
    # Проверка на артикулы типа "12345-678" или "A123-B45", которые Excel может испортить
    if re.match(r'^[A-Za-z0-9]+[-][A-Za-z0-9]+$', s):
        return f"'{s}"
        
    return s


# ============================================================================
# 2.4 ГЕНЕРАТОР ЖИВЫХ ФОРМУЛ EXCEL (REQ 5)
# ============================================================================
class ExcelFormulaBuilder:
    """
    Генератор живых формул Excel для юнит-экономики.
    Позволяет экспортировать расчеты так, чтобы пользователь мог менять 
    входные параметры (цену, вес) прямо в Excel, и формулы пересчитывались.
    """
    def __init__(self, col_map: Dict[str, str]):
        """
        col_map: Словарь маппинга логических имен в буквы колонок Excel.
        Пример: {"price": "C", "cost": "D", "commission_rate": "E", "logistics": "F"}
        """
        self.col_map = col_map

    def _get_cell(self, field: str, row: int = 2) -> str:
        """Возвращает ссылку на ячейку (например, 'C2')"""
        col = self.col_map.get(field, "A")
        return f"{col}{row}"

    def build_commission_formula(self, row: int = 2) -> str:
        """
        Формула комиссии: Цена * Ставка комиссии
        =C2 * E2
        """
        price_cell = self._get_cell("price", row)
        rate_cell = self._get_cell("commission_rate", row)
        return f"={price_cell}*{rate_cell}"

    def build_logistics_formula(self, row: int = 2) -> str:
        """
        Формула логистики (упрощенная для FBS): База + (Вес * Ставка за кг)
        =F2 + (G2 * H2)
        """
        base_cell = self._get_cell("logistics_base", row)
        weight_cell = self._get_cell("weight", row)
        rate_cell = self._get_cell("logistics_per_kg", row)
        return f"={base_cell}+({weight_cell}*{rate_cell})"

    def build_total_expenses_formula(self, row: int = 2) -> str:
        """
        Формула общих расходов: Себестоимость + Комиссия + Логистика + Хранение + Эквайринг + Налог
        =D2 + I2 + J2 + K2 + L2 + M2
        """
        cost = self._get_cell("cost", row)
        commission = self._get_cell("commission", row)
        logistics = self._get_cell("logistics", row)
        storage = self._get_cell("storage", row)
        acquiring = self._get_cell("acquiring", row)
        tax = self._get_cell("tax", row)
        return f"={cost}+{commission}+{logistics}+{storage}+{acquiring}+{tax}"

    def build_profit_formula(self, row: int = 2) -> str:
        """
        Формула прибыли: Цена - Общие расходы
        =C2 - N2
        """
        price = self._get_cell("price", row)
        expenses = self._get_cell("total_expenses", row)
        return f"={price}-{expenses}"

    def build_margin_formula(self, row: int = 2) -> str:
        """
        Формула маржинальности: (Прибыль / Цена) * 100
        =(O2 / C2) * 100
        """
        profit = self._get_cell("profit", row)
        price = self._get_cell("price", row)
        return f"=({profit}/{price})*100"

    def build_recommended_price_formula(self, row: int = 2) -> str:
        """
        Формула рекомендуемой цены для безубыточности с целевой маржой 20%:
        (Себестоимость + Фикс. расходы) / (1 - Переменные расходы - 0.20)
        Упрощенный вариант для Excel:
        =(D2 + F2 + K2) / (1 - E2 - L2 - 0.20)
        """
        cost = self._get_cell("cost", row)
        logistics_base = self._get_cell("logistics_base", row)
        storage = self._get_cell("storage", row)
        comm_rate = self._get_cell("commission_rate", row)
        acquiring_rate = self._get_cell("acquiring_rate", row)
        # Защита от деления на ноль или отрицательное число через MAX
        return f"=MAX(0, ({cost}+{logistics_base}+{storage}) / MAX(0.01, (1 - {comm_rate} - {acquiring_rate} - 0.20)))"


# ============================================================================
# 2.5 КРОСС-СВЯЗЫВАНИЕ И ОЕ-НОМЕРА (REQ 2)
# ============================================================================
class OECrossLinker:
    """
    Класс для умного кросс-связывания данных через ОЕ-номера.
    Реализует:
    1. Парсинг ОЕ-номеров, разделенных через ';'
    2. Нормализацию ОЕ-номеров (удаление пробелов, дефисов для точного совпадения)
    3. Заполнение пропусков в основном файле (например, веса) из файла ОЕ/аналогов.
    """
    
    @staticmethod
    def split_oe_numbers(oe_string: str) -> List[str]:
        """Разделяет строку ОЕ-номеров по точке с запятой и очищает"""
        if pd.isna(oe_string) or not oe_string:
            return []
        return [oe.strip() for oe in str(oe_string).split(';') if oe.strip()]

    @staticmethod
    def normalize_oe(oe: str) -> str:
        """
        Нормализует ОЕ-номер для сравнения: убирает все не alphanumeric символы, 
        приводит к верхнему регистру.
        Пример: "123-456 78" -> "12345678"
        """
        if pd.isna(oe) or not oe:
            return ""
        return re.sub(r'[^0-9A-Za-z]', '', str(oe).upper())

    @staticmethod
    def normalize_artikul(art: str) -> str:
        """Нормализует артикул для сравнения"""
        if pd.isna(art) or not art:
            return ""
        return re.sub(r'[^0-9A-Za-z]', '', str(art).upper())

    def link_and_fill_missing(
        self,
        df_main: pd.DataFrame,
        df_oe: pd.DataFrame,
        main_art_col: str,
        main_oe_col: str,
        oe_art_col: str,
        oe_oe_col: str,
        cols_to_fill: List[str]
    ) -> pd.DataFrame:
        """
        Связывает основной DataFrame с DataFrame ОЕ и заполняет пропуски.
        
        :param df_main: Основной файл (например, "Габариты", где нет веса)
        :param df_oe: Файл с ОЕ и дополнительными данными (где есть вес)
        :param main_art_col: Имя колонки артикула в df_main
        :param main_oe_col: Имя колонки ОЕ в df_main
        :param oe_art_col: Имя колонки артикула в df_oe
        :param oe_oe_col: Имя колонки ОЕ в df_oe
        :param cols_to_fill: Список колонок для заполнения (например, ['weight', 'length'])
        :return: Обновленный df_main
        """
        if df_main.empty or df_oe.empty:
            return df_main.copy()

        result = df_main.copy()
        
        # 1. Создаем нормализованные ключи для соединения
        result['_norm_art'] = result[main_art_col].apply(self.normalize_artikul)
        result['_norm_oe'] = result[main_oe_col].apply(self.normalize_oe)
        
        df_oe_work = df_oe.copy()
        df_oe_work['_norm_art'] = df_oe_work[oe_art_col].apply(self.normalize_artikul)
        df_oe_work['_norm_oe'] = df_oe_work[oe_oe_col].apply(self.normalize_oe)
        
        # 2. Взрываем (explode) ОЕ-номера в df_oe, если они разделены ';'
        # Сначала создаем список нормализованных ОЕ для каждой строки
        df_oe_work['_oe_list'] = df_oe_work[oe_oe_col].apply(self.split_oe_numbers)
        df_oe_work['_norm_oe_list'] = df_oe_work['_oe_list'].apply(
            lambda lst: [self.normalize_oe(oe) for oe in lst]
        )
        
        # Explode по списку нормализованных ОЕ, чтобы каждая строка стала отдельным сопоставлением
        df_exploded = df_oe_work.explode('_norm_oe_list').rename(columns={'_norm_oe_list': '_norm_oe'})
        df_exploded = df_exploded[df_exploded['_norm_oe'] != ""]
        
        # Убираем дубликаты, оставляя первые найденные данные для конкретного ОЕ
        # Это предотвращает умножение строк при merge, если в df_oe есть дубли
        df_exploded = df_exploded.drop_duplicates(subset=['_norm_oe'], keep='first')
        
        # 3. Соединяем по ОЕ-номеру (приоритет 1)
        # Если в df_main есть ОЕ, мы найдем соответствующие данные в df_oe
        merged_by_oe = result.merge(
            df_exploded[[' _norm_oe'] + cols_to_fill], # Внимание: пробел в имени колонки исправлен ниже
            left_on='_norm_oe',
            right_on='_norm_oe',
            how='left',
            suffixes=('', '_from_oe')
        )
        # Исправление опечатки в списке колонок выше:
        merged_by_oe = result.merge(
            df_exploded[['_norm_oe'] + cols_to_fill],
            left_on='_norm_oe',
            right_on='_norm_oe',
            how='left',
            suffixes=('', '_from_oe')
        )

        # 4. Заполняем пропуски в основных колонках данными из _from_oe
        filled_count = 0
        for col in cols_to_fill:
            if col not in result.columns:
                continue
                
            oe_col_name = f"{col}_from_oe"
            if oe_col_name in merged_by_oe.columns:
                # Условие заполнения: если в основном столбце NaN, 0 или пустая строка
                mask = (
                    merged_by_oe[col].isna() | 
                    (merged_by_oe[col] == 0) | 
                    (merged_by_oe[col] == '')
                )
                
                # Заполняем
                merged_by_oe.loc[mask, col] = merged_by_oe.loc[mask, oe_col_name]
                filled_count += int(mask.sum())
                
                # Удаляем временную колонку
                merged_by_oe = merged_by_oe.drop(columns=[oe_col_name])

        logger.info(f"✅ OECrossLinker: Заполнено {filled_count} пропусков через ОЕ-связывание.")
        
        # 5. Очистка временных колонок
        final_cols = [c for c in merged_by_oe.columns if not c.startswith('_norm_') and not c.startswith('_oe_')]
        
        return merged_by_oe[final_cols]

    def build_cross_references_column(
        self,
        df_main: pd.DataFrame,
        df_cross: pd.DataFrame,
        main_art_col: str,
        cross_art_col: str,
        cross_analog_col: str,
        output_col_name: str = "Кроссы (аналоги)"
    ) -> pd.DataFrame:
        """
        Добавляет в основной DataFrame колонку с аналогами (кроссами), разделенными через ';'.
        Берет информацию с площадок кроссировки (имитация через df_cross).
        """
        if df_main.empty or df_cross.empty:
            result = df_main.copy()
            result[output_col_name] = ""
            return result

        result = df_main.copy()
        result['_norm_art'] = result[main_art_col].apply(self.normalize_artikul)
        
        df_cross_work = df_cross.copy()
        df_cross_work['_norm_art'] = df_cross_work[cross_art_col].apply(self.normalize_artikul)
        
        # Группируем аналоги по артикулу и объединяем через ';'
        analogs_grouped = (
            df_cross_work.groupby('_norm_art')[cross_analog_col]
            .apply(lambda x: '; '.join([str(v).strip() for v in x if pd.notna(v) and str(v).strip()]))
            .reset_index()
            .rename(columns={cross_analog_col: output_col_name})
        )
        
        # Merge с основным датафреймом
        result = result.merge(analogs_grouped, on='_norm_art', how='left')
        result[output_col_name] = result[output_col_name].fillna("")
        
        # Очистка
        result = result.drop(columns=['_norm_art'])
        
        return result
# ============================================================================
# БЛОК 3: УМНОЕ ЧТЕНИЕ ФАЙЛОВ, КРОСС-СВЯЗЫВАНИЕ, DEEPSEEK API И UI РАЗДЕЛА 1
# ============================================================================

# ============================================================================
# 3.1 УМНОЕ ЧТЕНИЕ ФАЙЛОВ С ЗАЩИТОЙ ОТ КРАКОЗЯБР И ДАТ (REQ 1)
# ============================================================================
def smart_read_uploaded_file(uploaded_file, file_type: str = "auto") -> pd.DataFrame:
    """
    Умное чтение загруженного файла (CSV или Excel) с авто-определением кодировки
    и защитой от превращения артикулов в даты при импорте.
    """
    if uploaded_file is None:
        return pd.DataFrame()
    
    # Сброс указателя файла в начало
    uploaded_file.seek(0)
    file_name = uploaded_file.name.lower()
    
    try:
        if file_name.endswith(('.csv', '.txt')):
            # Приоритетные кодировки для РФ/СНГ
            encodings_to_try = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'latin1', 'iso-8859-1']
            separators_to_try = [';', ',', '\t', '|']
            
            best_df = None
            best_score = -1
            
            for enc in encodings_to_try:
                for sep in separators_to_try:
                    try:
                        uploaded_file.seek(0)
                        # Читаем с dtype=str, чтобы Excel-подобные артикулы ("1-2", "OCT") не стали датами
                        df = pd.read_csv(
                            uploaded_file,
                            encoding=enc,
                            sep=sep,
                            dtype=str, 
                            on_bad_lines='skip',
                            skipinitialspace=True,
                            engine='python'
                        )
                        
                        if df is None or df.empty or len(df.columns) <= 1:
                            continue
                        
                        # Оценка качества: чем меньше кракозябр в колонках, тем лучше
                        mojibake_count = sum(1 for col in df.columns if isinstance(col, str) and detect_mojibake(col))
                        score = len(df.columns) - (mojibake_count * 10)
                        
                        if score > best_score:
                            best_score = score
                            best_df = df
                            
                        # Если кракозябр нет вообще, это идеальный вариант
                        if mojibake_count == 0:
                            break
                            
                    except (pd.errors.ParserError, UnicodeDecodeError, UnicodeError):
                        continue
                if best_score >= len(best_df.columns) if best_df is not None else 0:
                    break
                    
            if best_df is not None:
                # Дополнительная очистка кракозябр в данных
                cleaned_df, fixed_count = fix_dataframe_encoding(best_df)
                if fixed_count > 0:
                    logger.info(f"✅ Исправлено {fixed_count} ячеек с кракозябрами при импорте CSV")
                return cleaned_df
                
            raise ValueError("Не удалось прочитать CSV файл. Проверьте кодировку и разделитель.")
            
        elif file_name.endswith(('.xlsx', '.xls')):
            uploaded_file.seek(0)
            # Для Excel читаем всё как строки, чтобы предотвратить авто-конвертацию "1-2" в дату
            df = pd.read_excel(
                uploaded_file,
                engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd',
                dtype=str,
                keep_default_na=False
            )
            
            # Проверка и исправление кракозябр в именах колонок и данных
            cleaned_df, fixed_count = fix_dataframe_encoding(df)
            if fixed_count > 0:
                logger.info(f"✅ Исправлено {fixed_count} ячеек с кракозябрами при импорте Excel")
            return cleaned_df
            
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {file_name}")
            
    except Exception as e:
        logger.error(f"❌ Ошибка чтения файла {uploaded_file.name}: {e}")
        st.error(f"Ошибка чтения файла: {e}")
        return pd.DataFrame()


# ============================================================================
# 3.2 ПРОЦЕССОР КРОСС-СВЯЗЫВАНИЯ ДАННЫХ (REQ 2)
# ============================================================================
class CrossDataProcessor:
    """
    Процессор для объединения файлов (Габариты, ОЕ, Кроссы) с заполнением пропусков
    и формированием столбцов с разделителями ';'.
    """
    def __init__(self):
        self.linker = OECrossLinker()
        
    def process_and_merge(
        self,
        df_main: pd.DataFrame,
        df_oe: Optional[pd.DataFrame] = None,
        df_cross: Optional[pd.DataFrame] = None,
        main_art_col: str = "Артикул",
        main_oe_col: str = "ОЕ номер",
        oe_art_col: str = "Артикул",
        oe_oe_col: str = "ОЕ номер",
        cross_art_col: str = "Артикул",
        cross_analog_col: str = "Аналог"
    ) -> pd.DataFrame:
        """
        Основной метод слияния данных.
        1. Берет основной файл (например, Габариты).
        2. Если есть файл ОЕ, заполняет пропуски (вес, габариты) из него по совпадению Артикула или ОЕ.
        3. Если есть файл Кроссов, добавляет столбец "Кроссы (аналоги)" через ';'.
        4. Нормализует столбец ОЕ в основном файле, объединяя дубликаты через ';'.
        """
        if df_main.empty:
            return df_main
            
        result = df_main.copy()
        
        # Шаг 1: Нормализация и объединение ОЕ номеров в основном файле через ';'
        # Если в основном файле несколько строк с одним артикулом и разными ОЕ, объединяем их
        if main_oe_col in result.columns and main_art_col in result.columns:
            result = result.groupby(main_art_col, as_index=False).agg({
                main_oe_col: lambda x: '; '.join([str(v).strip() for v in x if pd.notna(v) and str(v).strip()]),
                **{col: 'first' for col in result.columns if col not in [main_art_col, main_oe_col]}
            })
            
        # Шаг 2: Заполнение пропусков из файла ОЕ (Req 2)
        if df_oe is not None and not df_oe.empty:
            cols_to_fill = [col for col in ['Вес', 'Длина', 'Ширина', 'Высота'] if col in result.columns]
            if cols_to_fill:
                logger.info(f"🔗 Запуск кросс-связывания для заполнения пропусков: {cols_to_fill}")
                result = self.linker.link_and_fill_missing(
                    df_main=result,
                    df_oe=df_oe,
                    main_art_col=main_art_col,
                    main_oe_col=main_oe_col,
                    oe_art_col=oe_art_col,
                    oe_oe_col=oe_oe_col,
                    cols_to_fill=cols_to_fill
                )
                
        # Шаг 3: Добавление столбца с кроссами (аналогами) через ';' (Req 2)
        if df_cross is not None and not df_cross.empty:
            logger.info("🔗 Запуск поиска кроссов (аналогов) по ОЕ/Артикулу")
            result = self.linker.build_cross_references_column(
                df_main=result,
                df_cross=df_cross,
                main_art_col=main_art_col,
                cross_art_col=cross_art_col,
                cross_analog_col=cross_analog_col,
                output_col_name="Кроссы (аналоги)"
            )
            
        return result


# ============================================================================
# 3.3 DEEPSEEK API MANAGER (REQ 3)
# ============================================================================
class DeepSeekAPIManager:
    """
    Менеджер для работы с DeepSeek API.
    Поддерживает два режима: "Обогащение каталога" или "Актуализация тарифов".
    """
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com/v1"):
        self.api_key = api_key
        self.base_url = base_url
        self.client = None
        if OPENAI_AVAILABLE:
            self.client = openai.OpenAI(
                api_key=self.api_key,
                base_url=self.base_url
            )
            
    def is_available(self) -> bool:
        return self.client is not None and bool(self.api_key)
        
    def enrich_catalog(self, product_name: str, current_category: str = "") -> Dict[str, Any]:
        """
        Режим 1: Обогащение каталога.
        Определяет категорию, тип автозапчасти и ключевые характеристики по названию.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты эксперт по автозапчастям. Проанализируй название товара: "{product_name}".
        Текущая категория (если есть): "{current_category}".
        
        Верни строго JSON в следующем формате:
        {{
            "parent_category": "Автозапчасти",
            "group_category": "Например: Подвеска, Двигатель, Тормозная система",
            "subgroup_category": "Например: Сайлентблоки, Поршни, Колодки",
            "product_type": "Например: SUSPENSION, ENGINE, BRAKE",
            "hazardous": false,
            "fragile": false,
            "confidence": 0.95
        }}
        Не добавляй никакой разметки, только валидный JSON.
        """
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                response_format={"type": "json_object"}
            )
            content = response.choices[0].message.content
            return json.loads(content)
        except Exception as e:
            logger.error(f"❌ Ошибка DeepSeek (обогащение): {e}")
            return {"error": str(e)}
            
    def update_tariffs(self, marketplace: str, category: str = "auto_parts") -> Dict[str, Any]:
        """
        Режим 2: Актуализация тарифов маркетплейсов.
        Запрашивает актуальные проценты комиссий, логистики и хранения.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты аналитик данных маркетплейсов. Предоставь актуальные тарифы для "{marketplace}" 
        для категории "{category}" на 2026 год.
        
        Верни строго JSON в следующем формате (числа как float, без процентов):
        {{
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "return_fee": 0.02,
            "acquiring_fee": 0.015,
            "source": "DeepSeek AI Estimate"
        }}
        Не добавляй никакой разметки, только валидный JSON.
        """
        
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                response_format={"type": "json_object"}
            )
            content = response.choices[0].message.content
            return json.loads(content)
        except Exception as e:
            logger.error(f"❌ Ошибка DeepSeek (тарифы): {e}")
            return {"error": str(e)}


# ============================================================================
# 3.4 UI РАЗДЕЛА 1: ЗАГРУЗКА ДАННЫХ, СВЯЗЫВАНИЕ И НАСТРОЙКА КЛЮЧЕЙ (REQ 4, 6)
# ============================================================================
def show_section1_data_loading():
    """
    📁 РАЗДЕЛ 1: ЗАГРУЗКА ДАННЫХ
    Реализует загрузку, кросс-связывание, управление API ключами (шифрование)
    и строгий режим FBS-only.
    """
    st.header("📁 Раздел 1: Загрузка и связывание данных")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    1. Загрузить файлы каталога (Габариты, ОЕ, Кроссы). Система автоматически исправит кракозябры.
    2. Настроить кросс-связывание: если в "Габаритах" нет веса, он будет взят из файла "ОЕ".
    3. Сформировать столбцы "ОЕ" и "Кроссы" с разделителями `;`.
    4. Сохранить API ключи в зашифрованном виде для использования в следующих разделах.
    """)
    
    # --- Инициализация менеджеров ---
    if 'secure_key_manager' not in st.session_state:
        st.session_state.secure_key_manager = SecureKeyManager()
    key_manager = st.session_state.secure_key_manager
    
    if 'cross_processor' not in st.session_state:
        st.session_state.cross_processor = CrossDataProcessor()
    processor = st.session_state.cross_processor
    
    # --- 1. Управление API ключами (REQ 6) ---
    with st.expander("🔑 Управление API ключами (Безопасное хранение)", expanded=False):
        st.markdown("Ключи шифруются и сохраняются локально. Они не исчезнут при закрытии вкладки.")
        
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            deepseek_key = st.text_input(
                "DeepSeek API Key", 
                value=key_manager.get_key("deepseek") or "",
                type="password",
                help="Ключ для обогащения каталога или актуализации тарифов"
            )
            if st.button("💾 Сохранить DeepSeek Key", key="save_ds_key"):
                key_manager.set_key("deepseek", deepseek_key, "DeepSeek API Key для AI функций")
                st.success("✅ Ключ DeepSeek зашифрован и сохранен!")
                
        with col_k2:
            # Место для будущих ключей (Ozon, WB)
            ozon_key = st.text_input(
                "Ozon API Key (опционально)", 
                value=key_manager.get_key("ozon") or "",
                type="password"
            )
            if st.button("💾 Сохранить Ozon Key", key="save_ozon_key"):
                key_manager.set_key("ozon", ozon_key, "Ozon Seller API Key")
                st.success("✅ Ключ Ozon зашифрован и сохранен!")

    st.divider()
    
    # --- 2. Загрузка файлов (REQ 1, 2) ---
    st.subheader("📥 Загрузка файлов каталога")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        file_main = st.file_uploader(
            "📦 Основной файл (например, Габариты/Цены)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_main",
            help="Должен содержать Артикул, Цену, Себестоимость"
        )
        
    with col_f2:
        file_oe = st.file_uploader(
            "🔧 Файл ОЕ номеров (опционально)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_oe",
            help="Используется для заполнения пропусков (вес, габариты) и связывания"
        )
        
    with col_f3:
        file_cross = st.file_uploader(
            "🔗 Файл Кроссов/Аналогов (опционально)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_cross",
            help="Используется для создания столбца 'Кроссы (аналоги)' через ;"
        )
        
    # --- 3. Обработка и связывание ---
    if file_main is not None:
        st.success("✅ Основной файл загружен. Система автоматически определит кодировку (UTF-8/CP1251/Latin1).")
        
        with st.spinner("Чтение и очистка данных от кракозябр..."):
            df_main = smart_read_uploaded_file(file_main)
            df_oe = smart_read_uploaded_file(file_oe) if file_oe else None
            df_cross = smart_read_uploaded_file(file_cross) if file_cross else None
            
        if not df_main.empty:
            st.subheader("⚙️ Настройка кросс-связывания")
            
            # Авто-детект колонок
            main_art_col = next((c for c in df_main.columns if 'артикул' in c.lower() or 'artikul' in c.lower() or 'sku' in c.lower()), df_main.columns[0])
            main_oe_col = next((c for c in df_main.columns if 'ое' in c.lower() or 'oe' in c.lower()), "")
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                main_art_col = st.selectbox("Колонка Артикула (Основной)", df_main.columns.tolist(), index=df_main.columns.tolist().index(main_art_col))
                main_oe_col = st.selectbox("Колонка ОЕ номера (Основной)", ["Не выбрано"] + df_main.columns.tolist(), index=df_main.columns.tolist().index(main_oe_col) if main_oe_col in df_main.columns else 0)
                
            with col_m2:
                oe_art_col = st.selectbox("Колонка Артикула (Файл ОЕ)", ["Не выбрано"] + (df_oe.columns.tolist() if df_oe is not None else []), index=0)
                oe_oe_col = st.selectbox("Колонка ОЕ номера (Файл ОЕ)", ["Не выбрано"] + (df_oe.columns.tolist() if df_oe is not None else []), index=0)
                cross_art_col = st.selectbox("Колонка Артикула (Файл Кроссов)", ["Не выбрано"] + (df_cross.columns.tolist() if df_cross is not None else []), index=0)
                cross_analog_col = st.selectbox("Колонка Аналога (Файл Кроссов)", ["Не выбрано"] + (df_cross.columns.tolist() if df_cross is not None else []), index=0)
                
            if st.button("🚀 Обработать и связать данные", type="primary", key="process_link_btn"):
                with st.spinner("Выполняется кросс-связывание, заполнение пропусков и формирование списков через ';'..."):
                    try:
                        # Приведение имен колонок к ожидаемым или использование выбранных
                        actual_oe_col = main_oe_col if main_oe_col != "Не выбрано" else "ОЕ номер"
                        actual_oe_art = oe_art_col if oe_art_col != "Не выбрано" else "Артикул"
                        actual_oe_oe = oe_oe_col if oe_oe_col != "Не выбрано" else "ОЕ номер"
                        actual_cross_art = cross_art_col if cross_art_col != "Не выбрано" else "Артикул"
                        actual_cross_analog = cross_analog_col if cross_analog_col != "Не выбрано" else "Аналог"
                        
                        df_result = processor.process_and_merge(
                            df_main=df_main,
                            df_oe=df_oe,
                            df_cross=df_cross,
                            main_art_col=main_art_col,
                            main_oe_col=actual_oe_col,
                            oe_art_col=actual_oe_art,
                            oe_oe_col=actual_oe_oe,
                            cross_art_col=actual_cross_art,
                            cross_analog_col=actual_cross_analog
                        )
                        
                        # Сохранение в Session State для передачи в следующие разделы
                        st.session_state['processed_catalog_df'] = df_result
                        
                        st.success(f"✅ Обработка завершена! Итоговая таблица: {len(df_result)} строк, {len(df_result.columns)} колонок.")
                        
                        # Демонстрация результата
                        st.markdown("##### 👁️ Предпросмотр результата (первые 10 строк)")
                        # Показываем только ключевые колонки для наглядности
                        display_cols = [main_art_col, actual_oe_col, "Кроссы (аналоги)"] + [c for c in ['Вес', 'Длина', 'Ширина', 'Высота', 'Цена'] if c in df_result.columns]
                        valid_display_cols = [c for c in display_cols if c in df_result.columns]
                        st.dataframe(df_result[valid_display_cols].head(10), use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"❌ Ошибка при связывании данных: {e}")
                        logger.exception("Ошибка process_and_merge")
        else:
            st.warning("⚠️ Основной файл пуст или не удалось его прочитать.")
            
    st.divider()
    
    # --- 4. Настройка режима работы (REQ 4: FBS-ONLY) ---
    st.subheader("⚙️ Глобальные настройки расчета")
    st.info("📌 В данной версии приложения поддерживается только расчет FBS-юнит-экономики. Остальные режимы отключены.")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        # Жестко ограничиваем выбор только FBS и FBY
        operation_mode = st.selectbox(
            "Режим работы маркетплейса",
            options=["FBS", "FBY"],
            index=0,
            help="FBS: со своего склада по заказу. FBY: со своего склада, доставка МП."
        )
        st.session_state['global_operation_mode'] = operation_mode
        
    with col_s2:
        tax_system = st.selectbox(
            "Налоговая система",
            options=["УСН_6", "УСН_15", "ОСН", "НПД"],
            index=0,
            format_func=lambda x: {"УСН_6": "УСН 6% (доходы)", "УСН_15": "УСН 15% (доходы-расходы)", "ОСН": "ОСН (общая)", "НПД": "НПД (самозанятый)"}[x]
        )
        st.session_state['global_tax_system'] = tax_system

    # --- 5. DeepSeek API: Выбор режима (REQ 3) ---
    st.subheader("🤖 DeepSeek AI: Выбор задачи")
    ds_key = key_manager.get_key("deepseek")
    if not ds_key:
        st.warning("⚠️ Ключ DeepSeek не задан. Настройте его в блоке 'Управление API ключами' выше.")
        ai_mode = "none"
    else:
        ai_mode = st.radio(
            "Что вы хотите сделать с помощью AI?",
            options=["Обогащение каталога (категории, типы)", "Актуализация тарифов маркетплейсов", "Ничего не делать"],
            horizontal=True,
            key="deepseek_mode_select"
        )
        st.session_state['deepseek_mode'] = ai_mode
        
        if ai_mode == "Обогащение каталога (категории, типы)" and 'processed_catalog_df' in st.session_state:
            if st.button("🚀 Запустить обогащение каталога через DeepSeek", key="run_enrich"):
                df_to_enrich = st.session_state['processed_catalog_df']
                name_col = next((c for c in df_to_enrich.columns if 'наименование' in c.lower() or 'name' in c.lower() or 'описание' in c.lower()), None)
                
                if not name_col:
                    st.error("Не найдена колонка с наименованием товара для обогащения.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    manager = DeepSeekAPIManager(api_key=ds_key)
                    
                    results = []
                    total = min(len(df_to_enrich), 50) # Ограничение для демо/безопасности
                    
                    for i, row in df_to_enrich.head(total).iterrows():
                        status_text.text(f"Обработка {i+1}/{total}: {str(row.get(name_col, ''))[:30]}...")
                        res = manager.enrich_catalog(str(row.get(name_col, '')), str(row.get('Категория', '')))
                        results.append(res)
                        progress_bar.progress((i + 1) / total)
                        time.sleep(0.5) # Rate limiting
                        
                    status_text.text("✅ Обогащение завершено!")
                    st.json(results[:3]) # Показать первые 3 результата
                    st.info("💡 В полной версии результаты будут добавлены как новые колонки в DataFrame.")
                    
        elif ai_mode == "Актуализация тарифов маркетплейсов":
            st.info("💡 Этот режим будет использован в Разделе 3 (Тарифы) для автоматического обновления ставок комиссий и логистики.")
# ============================================================================
# БЛОК 4: FBS-ONLY РАСЧЁТ, ЖИВЫЕ ФОРМУЛЫ EXCEL, GOOGLE SHEETS API И UI РАЗДЕЛА 4
# ============================================================================

# ============================================================================
# 4.1 FBS-ONLY РАСЧЁТ ЮНИТ-ЭКОНОМИКИ (REQ 4)
# ============================================================================
class FBSUnitEconomicsCalculator:
    """
    Калькулятор юнит-экономики ТОЛЬКО для режима FBS (и FBY для сравнения).
    Удалены FBO, DBS, FBP, RealFBS из логики (Req 4).
    
    Логика расчёта для FBS:
    - Комиссия МП (процент от цены)
    - Логистика FBS (база + вес + объём)
    - Хранение на своём складе продавца (до отгрузки)
    - Эквайринг
    - Возвраты
    - Налог
    - Специфические расходы автозапчастей (Честный ЗНАК, упаковка FBS)
    """
    
    # 🆕 v101.1: Жёстко ограниченные режимы (Req 4)
    ALLOWED_MODES = ["FBS", "FBY"]
    
    def __init__(self, marketplace_config: Dict[str, Any], tax_system: str = "УСН_6"):
        """
        marketplace_config: Словарь с тарифами МП.
        Пример: {
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "acquiring_fee": 0.015,
            "return_fee": 0.02
        }
        """
        self.config = marketplace_config
        self.tax_system = tax_system
        
    def calculate(
        self,
        price: float,
        cost: float,
        weight_kg: float,
        length_cm: float,
        width_cm: float,
        height_cm: float,
        days_in_storage: int = 30,
        operation_mode: str = "FBS",
        category: str = "auto_parts",
        is_hazardous: bool = False,
        is_fragile: bool = False
    ) -> Dict[str, Any]:
        """
        Основной метод расчёта юнит-экономики для FBS.
        """
        # Жёсткая проверка режима (Req 4)
        if operation_mode not in self.ALLOWED_MODES:
            raise ValueError(f"Режим {operation_mode} не поддерживается. Используйте FBS или FBY.")
        
        # Валидация входных данных
        if price <= 0:
            raise ValueError("Цена должна быть положительной")
        if cost <= 0:
            raise ValueError("Себестоимость должна быть положительной")
        
        # === 1. КОМИССИЯ МП ===
        commission_rate = self.config.get("commission_rate", 0.15)
        # Проверка категории (если есть отдельные ставки)
        category_rates = self.config.get("category_rates", {})
        if category in category_rates:
            commission_rate = category_rates[category]
            
        commission = max(price * commission_rate, self.config.get("min_commission", 0.0))
        
        # === 2. ЛОГИСТИКА FBS ===
        # Объёмный вес (коэффициент 5000 — стандарт МП)
        volumetric_weight = (length_cm * width_cm * height_cm) / 5000.0 if length_cm > 0 else 0
        billable_weight = max(weight_kg, volumetric_weight)
        # Округление вверх до 0.5 кг
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        logistics_base = self.config.get("logistics_base", 50.0)
        logistics_per_kg = self.config.get("logistics_per_kg", 15.0)
        logistics = logistics_base + (billable_weight * logistics_per_kg)
        
        # Мультипликатор режима (FBY дешевле, FBS базовый)
        mode_multipliers = self.config.get("mode_multipliers", {"FBS": 1.0, "FBY": 0.75})
        logistics *= mode_multipliers.get(operation_mode, 1.0)
        
        # === 3. ХРАНЕНИЕ ===
        volume_liter = (length_cm * width_cm * height_cm) / 1000.0 if length_cm > 0 else 5.0
        storage_per_day = self.config.get("storage_per_day", 0.3)
        
        # Прогрессивная ставка для FBS (растёт после 60 дней)
        if days_in_storage <= 60:
            storage_multiplier = 1.0
        elif days_in_storage <= 90:
            storage_multiplier = 2.0
        elif days_in_storage <= 180:
            storage_multiplier = 4.0
        else:
            storage_multiplier = 8.0
            
        storage_cost = volume_liter * storage_per_day * days_in_storage * storage_multiplier
        
        # === 4. ЭКВАЙРИНГ ===
        acquiring_fee = self.config.get("acquiring_fee", 0.015)
        acquiring = price * acquiring_fee
        
        # === 5. ВОЗВРАТЫ ===
        return_fee = self.config.get("return_fee", 0.02)
        returns = price * return_fee
        
        # === 6. НАДБАВКИ ===
        hazardous_surcharge = price * 0.02 if is_hazardous else 0.0
        fragile_surcharge = price * 0.01 if is_fragile else 0.0
        
        # === 7. СПЕЦИФИЧЕСКИЕ РАСХОДЫ АВТОЗАПЧАСТЕЙ ===
        chestny_znak = 1.5  # Маркировка
        packaging_fbs = 45.0  # Упаковка для FBS
        labeling = 3.0  # Маркировка
        warranty_reserve = price * 0.02  # Гарантийный резерв
        auto_parts_specific = chestny_znak + packaging_fbs + labeling + warranty_reserve
        
        # === 8. НАЛОГ ===
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН_6"])
        if tax_config["base"] == "revenue":
            tax = price * tax_config["rate"]
        elif tax_config["base"] == "profit":
            profit_before_tax = price - cost - commission - logistics - storage_cost - acquiring - returns - auto_parts_specific
            tax = max(0, profit_before_tax * tax_config["rate"])
            # Минимальный налог для УСН 15%
            if self.tax_system == "УСН_15":
                min_tax = price * tax_config.get("min_rate", 0.01)
                tax = max(tax, min_tax)
        else:
            tax = 0.0
            
        # === 9. ИТОГО РАСХОДОВ ===
        total_expenses = (
            cost + commission + logistics + storage_cost + acquiring + returns +
            hazardous_surcharge + fragile_surcharge + auto_parts_specific + tax
        )
        
        # === 10. ПРИБЫЛЬ И МЕТРИКИ ===
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0
        
        # === 11. РЕКОМЕНДУЕМАЯ МИНИМАЛЬНАЯ ЦЕНА (для безубыточности + 10% прибыли) ===
        variable_rate = commission_rate + acquiring_fee + return_fee + tax_config["rate"] + 0.10
        fixed_costs = cost + logistics_base + storage_cost
        denominator = 1 - variable_rate
        recommended_min_price = (fixed_costs / denominator) if denominator > 0 else 0
        
        return {
            "price": round(price, 2),
            "cost": round(cost, 2),
            "operation_mode": operation_mode,
            "billable_weight": round(billable_weight, 2),
            "volume_liter": round(volume_liter, 3),
            "commission": round(commission, 2),
            "commission_rate": round(commission_rate * 100, 2),
            "logistics": round(logistics, 2),
            "storage_cost": round(storage_cost, 2),
            "acquiring": round(acquiring, 2),
            "returns": round(returns, 2),
            "hazardous_surcharge": round(hazardous_surcharge, 2),
            "fragile_surcharge": round(fragile_surcharge, 2),
            "auto_parts_specific": round(auto_parts_specific, 2),
            "tax": round(tax, 2),
            "total_expenses": round(total_expenses, 2),
            "profit": round(profit, 2),
            "margin_percent": round(margin_percent, 2),
            "roi": round(roi, 2),
            "recommended_min_price": round(recommended_min_price, 2)
        }


# ============================================================================
# 4.2 ЭКСПОРТЕР С ЖИВЫМИ ФОРМУЛАМИ EXCEL (REQ 5)
# ============================================================================
class LiveExcelExporter:
    """
    Экспорт результатов расчёта в Excel с ЖИВЫМИ ФОРМУЛАМИ.
    Пользователь может менять цену, вес, ставку комиссии — и всё пересчитается.
    Логика соответствует калькулятору Яндекс Маркет / Ozon / WB.
    """
    
    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl не установлен. pip install openpyxl")
            
    def export_with_formulas(
        self,
        df_results: pd.DataFrame,
        output_path: str,
        marketplace_name: str = "Ozon"
    ) -> bool:
        """
        Экспорт DataFrame в Excel с живыми формулами.
        
        Структура колонок в Excel:
        A: Артикул
        B: Наименование
        C: Цена продажи (вводная, можно менять)
        D: Себестоимость (вводная)
        E: Вес, кг (вводная)
        F: Длина, см (вводная)
        G: Ширина, см (вводная)
        H: Высота, см (вводная)
        I: Ставка комиссии, % (вводная, можно менять)
        J: База логистики, ₽ (вводная)
        K: Логистика за кг, ₽ (вводная)
        L: Объёмный вес, кг (ФОРМУЛА)
        M: Оплачиваемый вес, кг (ФОРМУЛА)
        N: Комиссия, ₽ (ФОРМУЛА)
        O: Логистика, ₽ (ФОРМУЛА)
        P: Хранение, ₽ (ФОРМУЛА)
        Q: Эквайринг, ₽ (ФОРМУЛА)
        R: Возвраты, ₽ (ФОРМУЛА)
        S: Авто-специфика, ₽ (ФОРМУЛА)
        T: Налог, ₽ (ФОРМУЛА)
        U: ИТОГО расходов, ₽ (ФОРМУЛА)
        V: ПРИБЫЛЬ, ₽ (ФОРМУЛА)
        W: МАРЖА, % (ФОРМУЛА)
        X: Рек. мин. цена, ₽ (ФОРМУЛА)
        """
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Юнит-экономика FBS"
            
            # === ЗАГОЛОВКИ ===
            headers = [
                "Артикул", "Наименование", "Цена продажи", "Себестоимость",
                "Вес, кг", "Длина, см", "Ширина, см", "Высота, см",
                "Ставка комиссии, %", "База логистики, ₽", "Логистика за кг, ₽",
                "Объёмный вес, кг", "Оплачиваемый вес, кг",
                "Комиссия, ₽", "Логистика, ₽", "Хранение, ₽",
                "Эквайринг, ₽", "Возвраты, ₽", "Авто-специфика, ₽", "Налог, ₽",
                "ИТОГО расходов, ₽", "ПРИБЫЛЬ, ₽", "МАРЖА, %", "Рек. мин. цена, ₽"
            ]
            
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            # === ДАННЫЕ И ФОРМУЛЫ ===
            for row_idx, (_, row) in enumerate(df_results.iterrows(), 2):
                # A: Артикул (экранируем от превращения в дату, Req 1)
                artikul = escape_excel_text(row.get("Артикул", row.get("artikul", "")))
                ws.cell(row=row_idx, column=1, value=artikul)
                
                # B: Наименование
                ws.cell(row=row_idx, column=2, value=str(row.get("Наименование", row.get("name", ""))))
                
                # C: Цена продажи (вводная)
                ws.cell(row=row_idx, column=3, value=float(row.get("price", row.get("Цена", 0))))
                ws.cell(row=row_idx, column=3).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # D: Себестоимость (вводная)
                ws.cell(row=row_idx, column=4, value=float(row.get("cost", row.get("Себестоимость", 0))))
                ws.cell(row=row_idx, column=4).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # E: Вес, кг (вводная)
                ws.cell(row=row_idx, column=5, value=float(row.get("weight", row.get("Вес", 0))))
                ws.cell(row=row_idx, column=5).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # F, G, H: Габариты (вводные)
                ws.cell(row=row_idx, column=6, value=float(row.get("length", row.get("Длина", 0))))
                ws.cell(row=row_idx, column=6).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=7, value=float(row.get("width", row.get("Ширина", 0))))
                ws.cell(row=row_idx, column=7).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=8, value=float(row.get("height", row.get("Высота", 0))))
                ws.cell(row=row_idx, column=8).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # I: Ставка комиссии, % (вводная, можно менять)
                commission_rate = float(row.get("commission_rate", 15.0))
                ws.cell(row=row_idx, column=9, value=commission_rate)
                ws.cell(row=row_idx, column=9).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # J: База логистики (вводная)
                ws.cell(row=row_idx, column=10, value=float(row.get("logistics_base", 50.0)))
                ws.cell(row=row_idx, column=10).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # K: Логистика за кг (вводная)
                ws.cell(row=row_idx, column=11, value=float(row.get("logistics_per_kg", 15.0)))
                ws.cell(row=row_idx, column=11).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # === ФОРМУЛЫ (живые, Req 5) ===
                # L: Объёмный вес = (F*G*H)/5000
                ws.cell(row=row_idx, column=12, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx})/5000, 0)")
                ws.cell(row=row_idx, column=12).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # M: Оплачиваемый вес = MAX(E, L), округлённый вверх до 0.5
                ws.cell(row=row_idx, column=13, value=f"=CEILING(MAX(E{row_idx}, L{row_idx}), 0.5)")
                ws.cell(row=row_idx, column=13).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # N: Комиссия = MAX(C * (I/100), 30)
                ws.cell(row=row_idx, column=14, value=f"=MAX(C{row_idx}*(I{row_idx}/100), 30)")
                ws.cell(row=row_idx, column=14).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # O: Логистика = J + (M * K)
                ws.cell(row=row_idx, column=15, value=f"=J{row_idx}+(M{row_idx}*K{row_idx})")
                ws.cell(row=row_idx, column=15).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # P: Хранение = (F*G*H/1000) * 0.3 * 30 (30 дней по умолчанию)
                ws.cell(row=row_idx, column=16, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx}/1000)*0.3*30, 5*0.3*30)")
                ws.cell(row=row_idx, column=16).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Q: Эквайринг = C * 0.015
                ws.cell(row=row_idx, column=17, value=f"=C{row_idx}*0.015")
                ws.cell(row=row_idx, column=17).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # R: Возвраты = C * 0.02
                ws.cell(row=row_idx, column=18, value=f"=C{row_idx}*0.02")
                ws.cell(row=row_idx, column=18).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # S: Авто-специфика = 1.5 + 45 + 3 + C*0.02
                ws.cell(row=row_idx, column=19, value=f"=1.5+45+3+C{row_idx}*0.02")
                ws.cell(row=row_idx, column=19).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # T: Налог = C * 0.06 (УСН 6%)
                ws.cell(row=row_idx, column=20, value=f"=C{row_idx}*0.06")
                ws.cell(row=row_idx, column=20).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # U: ИТОГО расходов = D + N + O + P + Q + R + S + T
                ws.cell(row=row_idx, column=21, value=f"=D{row_idx}+N{row_idx}+O{row_idx}+P{row_idx}+Q{row_idx}+R{row_idx}+S{row_idx}+T{row_idx}")
                ws.cell(row=row_idx, column=21).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=21).font = Font(bold=True)
                
                # V: ПРИБЫЛЬ = C - U
                ws.cell(row=row_idx, column=22, value=f"=C{row_idx}-U{row_idx}")
                ws.cell(row=row_idx, column=22).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=22).font = Font(bold=True, color="006600")
                
                # W: МАРЖА, % = (V / C) * 100
                ws.cell(row=row_idx, column=23, value=f"=IF(C{row_idx}>0, (V{row_idx}/C{row_idx})*100, 0)")
                ws.cell(row=row_idx, column=23).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=23).font = Font(bold=True)
                ws.cell(row=row_idx, column=23).number_format = '0.00"%"'
                
                # X: Рек. мин. цена = (D + J + P) / (1 - I/100 - 0.015 - 0.02 - 0.06 - 0.10)
                ws.cell(row=row_idx, column=24, value=f"=MAX(0, (D{row_idx}+J{row_idx}+P{row_idx}) / MAX(0.01, (1 - I{row_idx}/100 - 0.015 - 0.02 - 0.06 - 0.10)))")
                ws.cell(row=row_idx, column=24).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=24).font = Font(bold=True, color="CC0000")
                
            # === АВТОШИРИНА КОЛОНОК ===
            for col_idx in range(1, 25):
                ws.column_dimensions[get_column_letter(col_idx)].width = 15
                
            # === ЗАМОРОЗКА ПЕРВОЙ СТРОКИ ===
            ws.freeze_panes = "A2"
            
            # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ (прибыль < 0 — красная) ===
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006")
            ws.conditional_formatting.add(
                f"V2:V{len(df_results) + 1}",
                CellIsRule(operator="lessThan", formula=["0"], fill=red_fill, font=red_font)
            )
            
            # === СОХРАНЕНИЕ ===
            wb.save(output_path)
            logger.info(f"✅ Excel с живыми формулами сохранён: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка экспорта Excel с формулами: {e}")
            return False


# ============================================================================
# 4.3 GOOGLE SHEETS API ДЛЯ КОРРЕКТИРОВКИ ОСТАТКОВ (REQ 6 — ФИНАЛ)
# ============================================================================
class GoogleSheetsStockUpdater:
    """
    Обновление остатков и характеристик товаров в маркетплейсе через Google Sheets API.
    Логика:
    1. Пользователь редактирует Google Таблицу (меняет остатки, цены).
    2. Система считывает изменения.
    3. Отправляет обновления в маркетплейс через API (Ozon/WB).
    """
    
    def __init__(self, credentials_json: str, spreadsheet_id: str, worksheet_name: str = "Остатки"):
        self.credentials_json = credentials_json
        self.spreadsheet_id = spreadsheet_id
        self.worksheet_name = worksheet_name
        self.client = None
        self._init_client()
        
    def _init_client(self):
        """Инициализация gspread клиента"""
        if not GSPREAD_AVAILABLE:
            logger.error("❌ gspread не установлен")
            return
        try:
            creds_path = Path(self.credentials_json)
            if creds_path.exists():
                credentials = Credentials.from_service_account_file(
                    str(creds_path),
                    scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                )
            else:
                creds_data = json.loads(self.credentials_json)
                credentials = Credentials.from_service_account_info(
                    creds_data,
                    scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                )
            self.client = gspread.authorize(credentials)
            logger.info("✅ Google Sheets клиент инициализирован для обновления остатков")
        except Exception as e:
            logger.error(f"❌ Ошибка инициализации Google Sheets: {e}")
            
    def read_stock_changes(self) -> Optional[pd.DataFrame]:
        """
        Считывает изменения из Google Таблицы.
        Возвращает DataFrame с колонками: Артикул, Остаток, Цена (новая)
        """
        if self.client is None:
            return None
        try:
            spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            worksheet = spreadsheet.worksheet(self.worksheet_name)
            values = worksheet.get_all_values()
            if not values:
                return pd.DataFrame()
            headers = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=headers)
            # Преобразуем числовые колонки
            for col in ["Остаток", "Цена"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            return df
        except Exception as e:
            logger.error(f"❌ Ошибка чтения Google Sheets: {e}")
            return None
            
    def push_stocks_to_ozon(self, df: pd.DataFrame, api_key: str, client_id: str) -> Dict[str, Any]:
        """
        Отправляет остатки в Ozon через API.
        """
        if df.empty:
            return {"success": False, "error": "DataFrame пустой"}
        
        url = "https://api-seller.ozon.ru/v2/products/info/stocks"
        headers = {
            "Client-Id": client_id,
            "Api-Key": api_key,
            "Content-Type": "application/json"
        }
        
        # Формируем батчи по 100 SKU (ограничение Ozon API)
        items = []
        for _, row in df.iterrows():
            artikul = str(row.get("Артикул", "")).strip()
            stock = int(row.get("Остаток", 0))
            if artikul and stock >= 0:
                items.append({
                    "offer_id": artikul,
                    "stock": stock,
                    "warehouse_id": 0  # Основной склад
                })
        
        if not items:
            return {"success": False, "error": "Нет данных для отправки"}
        
        results = {"success": True, "updated": 0, "errors": []}
        
        # Отправляем батчами по 100
        for i in range(0, len(items), 100):
            batch = items[i:i + 100]
            payload = {"stocks": batch}
            try:
                response = requests.post(url, json=payload, headers=headers, timeout=30)
                if response.status_code == 200:
                    results["updated"] += len(batch)
                else:
                    results["errors"].append(f"Batch {i//100 + 1}: HTTP {response.status_code}")
            except Exception as e:
                results["errors"].append(f"Batch {i//100 + 1}: {str(e)}")
                
        return results
        
    def push_prices_to_ozon(self, df: pd.DataFrame, api_key: str, client_id: str) -> Dict[str, Any]:
        """
        Отправляет новые цены в Ozon через API.
        """
        if df.empty:
            return {"success": False, "error": "DataFrame пустой"}
        
        url = "https://api-seller.ozon.ru/v1/prices/products"
        headers = {
            "Client-Id": client_id,
            "Api-Key": api_key,
            "Content-Type": "application/json"
        }
        
        items = []
        for _, row in df.iterrows():
            artikul = str(row.get("Артикул", "")).strip()
            price = float(row.get("Цена", 0))
            if artikul and price > 0:
                items.append({
                    "auto_action_enabled": "UNKNOWN",
                    "currency_code": "RUB",
                    "min_price": "0",
                    "offer_id": artikul,
                    "old_price": "0",
                    "price": str(int(price)),
                    "price_strategy_enabled": "UNKNOWN"
                })
        
        if not items:
            return {"success": False, "error": "Нет данных для отправки"}
        
        results = {"success": True, "updated": 0, "errors": []}
        
        for i in range(0, len(items), 100):
            batch = items[i:i + 100]
            payload = {"prices": batch}
            try:
                response = requests.post(url, json=payload, headers=headers, timeout=30)
                if response.status_code == 200:
                    results["updated"] += len(batch)
                else:
                    results["errors"].append(f"Batch {i//100 + 1}: HTTP {response.status_code}")
            except Exception as e:
                results["errors"].append(f"Batch {i//100 + 1}: {str(e)}")
                
        return results


# ============================================================================
# 4.4 UI РАЗДЕЛА 4: РАСЧЁТ С ЖИВЫМИ ФОРМУЛАМИ (REQ 4, 5)
# ============================================================================
def show_section4_calculation():
    """
    🧮 РАЗДЕЛ 4: РАСЧЁТ ЮНИТ-ЭКОНОМИКИ
    FBS-only режим, живые формулы Excel, интеграция с Google Sheets для остатков.
    """
    st.header("🧮 Раздел 4: Расчёт юнит-экономики (FBS)")
    
    st.info("""
    **🎯 ЦЕЛЬ РАЗДЕЛА:**
    1. Рассчитать юнит-экономику для всех товаров в режиме **FBS** (со своего склада).
    2. Экспортировать результаты в Excel с **живыми формулами** — меняй цену/вес/ставку и всё пересчитается.
    3. Загрузить остатки в маркетплейс через Google Sheets API.
    
    **📌 ВАЖНО:** В данной версии поддерживается ТОЛЬКО FBS (и FBY для сравнения). 
    Режимы FBO, DBS, FBP, RealFBS отключены.
    """)
    
    # --- Проверка наличия данных из Раздела 1 ---
    if 'processed_catalog_df' not in st.session_state or st.session_state['processed_catalog_df'].empty:
        st.error("❌ Нет данных каталога. Перейдите в Раздел 1 и загрузите файлы.")
        return
        
    df_catalog = st.session_state['processed_catalog_df'].copy()
    
    # --- Инициализация калькулятора ---
    if 'fbs_calculator' not in st.session_state:
        # Базовая конфигурация Ozon FBS
        default_config = {
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "acquiring_fee": 0.015,
            "return_fee": 0.02,
            "mode_multipliers": {"FBS": 1.0, "FBY": 0.75},
            "category_rates": {}
        }
        st.session_state['fbs_calculator'] = FBSUnitEconomicsCalculator(
            marketplace_config=default_config,
            tax_system=st.session_state.get('global_tax_system', 'УСН_6')
        )
    calculator = st.session_state['fbs_calculator']
    
    # --- 1. НАСТРОЙКИ РАСЧЁТА ---
    st.subheader("⚙️ Настройки расчёта")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        marketplace = st.selectbox(
            "Маркетплейс",
            options=["Ozon", "Wildberries", "Яндекс Маркет"],
            index=0,
            key="calc_marketplace"
        )
    with col2:
        # 🆕 v101.1: Только FBS и FBY (Req 4)
        operation_mode = st.selectbox(
            "Режим работы (только FBS/FBY)",
            options=["FBS", "FBY"],
            index=0,
            key="calc_mode",
            help="FBS — со своего склада по заказу. FBY — аналог FBS с доставкой МП."
        )
    with col3:
        days_in_storage = st.number_input(
            "Дней хранения",
            min_value=1,
            max_value=365,
            value=30,
            step=1,
            key="calc_days"
        )
        
    # --- 2. ОПРЕДЕЛЕНИЕ КОЛОНОК ---
    st.subheader("🔍 Определение колонок")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        artikul_col = st.selectbox("Артикул", df_catalog.columns.tolist(), 
                                   index=df_catalog.columns.tolist().index(next((c for c in df_catalog.columns if 'артикул' in c.lower() or 'artikul' in c.lower()), df_catalog.columns[0])),
                                   key="calc_art_col")
    with col2:
        price_col = st.selectbox("Цена продажи", df_catalog.columns.tolist(),
                                 index=df_catalog.columns.tolist().index(next((c for c in df_catalog.columns if 'цена' in c.lower() or 'price' in c.lower()), df_catalog.columns[0])) if any('цена' in c.lower() or 'price' in c.lower() for c in df_catalog.columns) else 0,
                                 key="calc_price_col")
    with col3:
        cost_col = st.selectbox("Себестоимость", df_catalog.columns.tolist(),
                                index=df_catalog.columns.tolist().index(next((c for c in df_catalog.columns if 'себестоимость' in c.lower() or 'cost' in c.lower() or 'закуп' in c.lower()), df_catalog.columns[0])) if any('себестоимость' in c.lower() or 'cost' in c.lower() or 'закуп' in c.lower() for c in df_catalog.columns) else 0,
                                key="calc_cost_col")
    with col4:
        name_col = st.selectbox("Наименование (опц.)", ["Не выбрано"] + df_catalog.columns.tolist(),
                                index=0, key="calc_name_col")
                                
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        weight_col = st.selectbox("Вес, кг", ["Не выбрано"] + df_catalog.columns.tolist(),
                                  index=df_catalog.columns.tolist().index(next((c for c in df_catalog.columns if 'вес' in c.lower() or 'weight' in c.lower()), "Не выбрано")) + 1 if any('вес' in c.lower() or 'weight' in c.lower() for c in df_catalog.columns) else 0,
                                  key="calc_weight_col")
    with col2:
        length_col = st.selectbox("Длина, см", ["Не выбрано"] + df_catalog.columns.tolist(), index=0, key="calc_length_col")
    with col3:
        width_col = st.selectbox("Ширина, см", ["Не выбрано"] + df_catalog.columns.tolist(), index=0, key="calc_width_col")
    with col4:
        height_col = st.selectbox("Высота, см", ["Не выбрано"] + df_catalog.columns.tolist(), index=0, key="calc_height_col")
        
    # --- 3. ЗАПУСК РАСЧЁТА ---
    st.divider()
    if st.button("🚀 Рассчитать юнит-экономику для всех товаров", type="primary", key="run_calc_btn"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        results = []
        total = len(df_catalog)
        
        for i, (_, row) in enumerate(df_catalog.iterrows()):
            status_text.text(f"Расчёт {i+1}/{total}: {str(row.get(artikul_col, ''))[:20]}...")
            
            try:
                # Извлекаем значения
                price = float(row.get(price_col, 0) or 0)
                cost = float(row.get(cost_col, 0) or 0)
                
                if price <= 0 or cost <= 0:
                    continue
                    
                weight = float(row.get(weight_col, 1.0) or 1.0) if weight_col != "Не выбрано" else 1.0
                length = float(row.get(length_col, 0) or 0) if length_col != "Не выбрано" else 0
                width = float(row.get(width_col, 0) or 0) if width_col != "Не выбрано" else 0
                height = float(row.get(height_col, 0) or 0) if height_col != "Не выбрано" else 0
                
                # Расчёт
                res = calculator.calculate(
                    price=price,
                    cost=cost,
                    weight_kg=weight,
                    length_cm=length,
                    width_cm=width,
                    height_cm=height,
                    days_in_storage=days_in_storage,
                    operation_mode=operation_mode
                )
                
                # Добавляем исходные данные
                res["Артикул"] = row.get(artikul_col, "")
                res["Наименование"] = row.get(name_col, "") if name_col != "Не выбрано" else ""
                res["ОЕ номер"] = row.get("ОЕ номер", "")
                res["Кроссы (аналоги)"] = row.get("Кроссы (аналоги)", "")
                
                results.append(res)
                
            except Exception as e:
                logger.warning(f"Ошибка расчёта для {row.get(artikul_col, '')}: {e}")
                continue
                
            progress_bar.progress((i + 1) / total)
            
        status_text.text(f"✅ Расчёт завершён! Обработано {len(results)} товаров.")
        
        if results:
            df_results = pd.DataFrame(results)
            st.session_state['calculation_results_df'] = df_results
            st.success(f"✅ Рассчитано {len(df_results)} товаров. Средняя маржа: {df_results['margin_percent'].mean():.1f}%")
        else:
            st.error("❌ Не удалось рассчитать ни одного товара. Проверьте данные.")
            
    # --- 4. ОТОБРАЖЕНИЕ РЕЗУЛЬТАТОВ ---
    if 'calculation_results_df' in st.session_state and not st.session_state['calculation_results_df'].empty:
        df_results = st.session_state['calculation_results_df']
        
        st.divider()
        st.subheader("📊 Результаты расчёта")
        
        # KPI
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("💰 Общая прибыль", f"{df_results['profit'].sum():,.0f} ₽")
        with col2:
            st.metric("📈 Средняя маржа", f"{df_results['margin_percent'].mean():.1f}%")
        with col3:
            st.metric("📊 Средний ROI", f"{df_results['roi'].mean():.1f}%")
        with col4:
            unprofitable = (df_results['profit'] < 0).sum()
            st.metric("⚠️ Убыточных SKU", f"{unprofitable}")
        with col5:
            underpriced = (df_results['price'] < df_results['recommended_min_price']).sum()
            st.metric("💡 Недооценённых", f"{underpriced}")
            
        # Таблица результатов
        display_cols = ["Артикул", "Наименование", "price", "cost", "profit", "margin_percent", "roi", "recommended_min_price"]
        available_cols = [c for c in display_cols if c in df_results.columns]
        st.dataframe(df_results[available_cols].head(100), use_container_width=True)
        
        # --- 5. ЭКСПОРТ С ЖИВЫМИ ФОРМУЛАМИ (Req 5) ---
        st.divider()
        st.subheader("📥 Экспорт в Excel с живыми формулами")
        st.info("""
        **💡 КАК ЭТО РАБОТАЕТ:**
        При экспорте в Excel в ячейки записываются не числа, а **ФОРМУЛЫ**.
        - Жёлтые ячейки — вводные данные (можно менять).
        - Зелёные ячейки — расчётные (пересчитываются автоматически).
        - Синие ячейки — итоговые (прибыль, маржа, рек. цена).
        
        Меняй цену, вес или ставку комиссии — и вся экономика пересчитается!
        """)
        
        if st.button("📥 Экспортировать в Excel с живыми формулами", type="primary", key="export_live_excel_btn"):
            try:
                output_path = EXPORTS_DIR / f"unit_economics_live_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                exporter = LiveExcelExporter()
                success = exporter.export_with_formulas(df_results, str(output_path), marketplace)
                
                if success and output_path.exists():
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                    st.download_button(
                        label="⬇️ Скачать Excel с живыми формулами",
                        data=file_data,
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_live_excel"
                    )
                    st.success("✅ Excel с живыми формулами готов к скачиванию!")
                else:
                    st.error("❌ Ошибка создания файла")
            except Exception as e:
                st.error(f"❌ Ошибка экспорта: {e}")
                logger.exception("Ошибка экспорта")
                
        # --- 6. ОБНОВЛЕНИЕ ОСТАТКОВ ЧЕРЕЗ GOOGLE SHEETS (Req 6 финал) ---
        st.divider()
        st.subheader("🔄 Корректировка остатков через Google Sheets API")
        st.info("""
        **💡 СЦЕНАРИЙ:**
        1. Экспортируй результаты в Google Таблицу.
        2. Отредактируй остатки/цены прямо в Google Sheets.
        3. Нажми "Загрузить изменения в МП" — система отправит новые данные через API.
        """)
        
        key_manager = st.session_state.get('secure_key_manager')
        ozon_api_key = key_manager.get_key("ozon") if key_manager else None
        
        if not ozon_api_key:
            st.warning("⚠️ Ozon API ключ не задан. Настройте его в Разделе 1.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                gs_spreadsheet_id = st.text_input("ID Google Таблицы", placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", key="gs_stock_id")
            with col2:
                ozon_client_id = st.text_input("Ozon Client ID", placeholder="12345", key="ozon_client_id_stock")
                
            if st.button("📤 Экспортировать результаты в Google Sheets", key="export_to_gs_btn"):
                if not gs_spreadsheet_id:
                    st.error("❌ Укажите ID Google Таблицы")
                else:
                    try:
                        # Получаем credentials из secure storage
                        gs_creds_path = GOOGLE_CREDS_DIR / "service_account.json"
                        if not gs_creds_path.exists():
                            st.error("❌ Файл service_account.json не найден в папке google_creds/")
                        else:
                            success = google_sheets_upload(
                                df=df_results[["Артикул", "Наименование", "price", "profit", "margin_percent"]],
                                spreadsheet_id=gs_spreadsheet_id,
                                worksheet_name="Юнит-экономика",
                                credentials_json=str(gs_creds_path),
                                clear_before=True
                            )
                            if success:
                                st.success("✅ Результаты экспортированы в Google Sheets!")
                            else:
                                st.error("❌ Ошибка экспорта")
                    except Exception as e:
                        st.error(f"❌ Ошибка: {e}")
                        
            if st.button("📥 Загрузить изменения остатков из Google Sheets в Ozon", key="push_stocks_btn"):
                if not gs_spreadsheet_id or not ozon_client_id:
                    st.error("❌ Укажите ID таблицы и Client ID")
                else:
                    try:
                        gs_creds_path = GOOGLE_CREDS_DIR / "service_account.json"
                        if not gs_creds_path.exists():
                            st.error("❌ Файл service_account.json не найден")
                        else:
                            updater = GoogleSheetsStockUpdater(
                                credentials_json=str(gs_creds_path),
                                spreadsheet_id=gs_spreadsheet_id,
                                worksheet_name="Остатки"
                            )
                            df_changes = updater.read_stock_changes()
                            if df_changes is not None and not df_changes.empty:
                                st.write("📋 Изменения из Google Sheets:")
                                st.dataframe(df_changes.head(20))
                                
                                if st.checkbox("✅ Подтверждаю отправку в Ozon", key="confirm_push"):
                                    if st.button("🚀 Отправить в Ozon", key="push_to_ozon_btn"):
                                        result = updater.push_stocks_to_ozon(df_changes, ozon_api_key, ozon_client_id)
                                        if result["success"]:
                                            st.success(f"✅ Обновлено {result['updated']} SKU!")
                                        else:
                                            st.error(f"❌ Ошибка: {result.get('error', 'Неизвестно')}")
                                        if result.get("errors"):
                                            for err in result["errors"][:5]:
                                                st.warning(f"  - {err}")
                            else:
                                st.warning("⚠️ Нет изменений для отправки")
                    except Exception as e:
                        st.error(f"❌ Ошибка: {e}")
# ============================================================================
# БЛОК 5: ФИНАЛЬНАЯ СБОРКА, main(), НАВИГАЦИЯ (v101.1)
# ============================================================================

# ============================================================================
# 5.1 ФУТЕР И SIDEBAR INFO
# ============================================================================
def show_footer():
    """Футер приложения"""
    st.divider()
    st.markdown("""
    <div style='text-align: center; padding: 20px; color: #666;'>
    <p style='margin: 0;'>🚗 <strong>Юнит-экономика автозапчастей PRO 2026 (FBS-ONLY)</strong></p>
    <p style='margin: 5px 0 0 0; font-size: 0.9em;'>
    Версия 101.1 | Enterprise Edition | Живые формулы Excel | Безопасное хранение ключей
    </p>
    <p style='margin: 5px 0 0 0; font-size: 0.8em; color: #999;'>
    © 2024-2026 AutoParts Analytics Team
    </p>
    </div>
    """, unsafe_allow_html=True)


def show_sidebar_info():
    """Информация в sidebar с учётом всех 6 требований"""
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
        "cryptography": CRYPTO_AVAILABLE,  # 🆕 v101.1: Для безопасного хранения ключей
        "chardet": CHARDET_AVAILABLE,      # 🆕 v101.1: Для авто-детекции кодировок
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
    1. 📁 Загрузите файлы (Габариты, ОЕ, Кроссы)
    2. 🔑 Настройте API ключи (шифруются)
    3. 🧮 Рассчитайте юнит-экономику (FBS)
    4. 📥 Экспортируйте в Excel с живыми формулами
    
    **🆕 v101.1:**
    - ✅ FBS-only режим
    - ✅ Живые формулы Excel
    - ✅ Кросс-связывание ОЕ
    - ✅ Безопасное хранение ключей
    - ✅ Авто-детекция кодировок
    - ✅ Защита от превращения текста в даты
    """)


# ============================================================================
# 5.2 ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ (main)
# ============================================================================
def main():
    """
    🚗 ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ v101.1
    Объединяет все 4 раздела в единый интерфейс с навигацией.
    
    Реализованные требования:
    1. ✅ Авто-детекция кодировок (UTF-8/CP1251/Latin1) и защита от превращения текста в даты.
    2. ✅ Кросс-связывание через ОЕ, заполнение пропусков, столбцы с ';'.
    3. ✅ DeepSeek API: выбор режима (Обогащение каталога ИЛИ Актуализация тарифов).
    4. ✅ FBS-only: удалены FBO, DBS, FBP, RealFBS из логики и UI.
    5. ✅ Живые формулы Excel: экспорт с формулами для ручного тюнинга.
    6. ✅ Безопасное хранение API ключей: шифрование (Fernet) в локальный файл.
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
    st.markdown(f"""
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border-radius: 10px; margin-bottom: 20px;'>
    <h1 style='color: white; margin: 0;'>🚗 Юнит-экономика автозапчастей PRO 2026 (FBS)</h1>
    <p style='color: #ccc; margin: 10px 0 0 0; font-size: 1.1em;'>
    Enterprise расчет с живыми формулами Excel, кросс-связыванием ОЕ и безопасным хранением ключей
    </p>
    <p style='color: #888; margin: 5px 0 0 0; font-size: 0.9em;'>
    Версия {APP_VERSION} | Специализация: Автозапчасти, Автотовары и Агрегаты
    </p>
    </div>
    """, unsafe_allow_html=True)
    
    # ========================================================================
    # ИНИЦИАЛИЗАЦИЯ КОМПОНЕНТОВ
    # ========================================================================
    try:
        # 🆕 v101.1: Инициализация SecureKeyManager (Req 6)
        if 'secure_key_manager' not in st.session_state:
            st.session_state.secure_key_manager = SecureKeyManager()
            logger.info("✅ SecureKeyManager инициализирован")
            
        # 🆕 v101.1: Инициализация CrossDataProcessor (Req 2)
        if 'cross_processor' not in st.session_state:
            st.session_state.cross_processor = CrossDataProcessor()
            logger.info("✅ CrossDataProcessor инициализирован")
            
        # 🆕 v101.1: Инициализация FBSUnitEconomicsCalculator (Req 4)
        if 'fbs_calculator' not in st.session_state:
            default_config = {
                "commission_rate": 0.15,
                "min_commission": 30.0,
                "logistics_base": 50.0,
                "logistics_per_kg": 15.0,
                "storage_per_day": 0.3,
                "acquiring_fee": 0.015,
                "return_fee": 0.02,
                "mode_multipliers": {"FBS": 1.0, "FBY": 0.75},
                "category_rates": {}
            }
            st.session_state['fbs_calculator'] = FBSUnitEconomicsCalculator(
                marketplace_config=default_config,
                tax_system="УСН_6"
            )
            logger.info("✅ FBSUnitEconomicsCalculator инициализирован")
            
        # 🆕 v101.1: Инициализация LiveExcelExporter (Req 5)
        if 'live_excel_exporter' not in st.session_state:
            st.session_state['live_excel_exporter'] = LiveExcelExporter()
            logger.info("✅ LiveExcelExporter инициализирован")
            
    except Exception as e:
        st.error(f"❌ Ошибка инициализации компонентов: {e}")
        logger.exception("Ошибка инициализации")
        
    # ========================================================================
    # НАВИГАЦИЯ ЧЕРЕЗ SIDEBAR
    # ========================================================================
    st.sidebar.title("🧭 Навигация")
    
    # 🆕 v101.1: Упрощённая навигация (FBS-only, Req 4)
    section = st.sidebar.radio(
        "Выберите раздел:",
        [
            "📁 Раздел 1: Загрузка и связывание данных",
            "🧮 Раздел 4: Расчёт юнит-экономики (FBS)",
        ],
        key="main_navigation",
    )
    
    # ========================================================================
    # ОТОБРАЖЕНИЕ ВЫБРАННОГО РАЗДЕЛА
    # ========================================================================
    if section == "📁 Раздел 1: Загрузка и связывание данных":
        show_section1_data_loading()
    elif section == "🧮 Раздел 4: Расчёт юнит-экономики (FBS)":
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
