#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v23.4 — API-COMPLIANT & OPTIMIZED
============================================================================
Исправления и улучшения (v23.4):
1. URL для расчёта тарифов: /v2/tariffs/calculate
2. Получение реального categoryId через /v2/categories/tree (итеративный поиск)
3. Пакетная отправка офферов (макс. 200 за запрос)
4. Rate limiter: 1 запрос в секунду
5. В запрос тарифов передаются реальные параметры товаров
6. Добавлены параметры transitWarehouseType и orderCargoType
7. Использование Enum для sellingProgram
8. Обработка ошибок с логированием ответов
9. Улучшена структура конфигурации
10. ФИКС: Удален псевдовекторизованный Decimal, заменен на нативный NumPy round.
11. ФИКС: Математика анализа чувствительности переписана через реальный пересчет P&L.
12. ФИКС: Убран ProcessPoolExecutor для Pandas (оверхед сериализации), оставлен ThreadPool.
13. ФИКС: Колонки 'source' и 'scheme' больше не удаляются из финального датафрейма (аудит).
14. ФИКС: Безопасная работа с временными файлами и кодировками.
15. ФИКС v23.4: Исправлено хеширование dict в st.cache_data.
16. ФИКС v23.4: Google Sheets авторизация через память (без записи JSON на диск).
17. ФИКС v23.4: Порог многопоточности поднят до 50 000 строк.
18. ФИКС v23.4: Корректная обработка NaN в артикулах и жесткий clip(lower=1) для quantity.
"""
import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
import logging
import warnings
from datetime import datetime, timedelta
from enum import Enum
from typing import (
    Dict, List, Optional, Tuple, Any, Union, TypedDict, Final, Callable
)
from dataclasses import dataclass
from functools import wraps
from collections import OrderedDict
import time
from contextlib import contextmanager
import json
import os
import sys
import re
import tempfile

# Дополнительные импорты для улучшений
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

try:
    from concurrent.futures import ThreadPoolExecutor
    import multiprocessing as mp
    CONCURRENT_AVAILABLE = True
except ImportError:
    CONCURRENT_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side,
        ColorScaleRule, DataBarRule, IconSetRule,
        CellIsRule, FormulaRule
    )
    from openpyxl.formatting.rule import ConditionalFormattingValueObject
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, Reference, PieChart, LineChart
    from openpyxl.chart.label import DataLabelList
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Подавляем только конкретные ожидаемые предупреждения, а не все подряд
warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
warnings.filterwarnings('ignore', message='.*openpyxl.*')

# ----------------------------------------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ (с безопасным путем и ротацией)
# ----------------------------------------------------------------------------
def setup_logging():
    """Настройка логирования с ротацией файлов в безопасной директории."""
    logger = logging.getLogger('YandexMarketUE')
    logger.setLevel(logging.INFO)
    if logger.handlers:
        logger.handlers.clear()
    
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    
    try:
        from logging.handlers import RotatingFileHandler
        # Используем временную директорию ОС или текущую, чтобы избежать ошибок прав доступа
        log_dir = tempfile.gettempdir()
        log_file = os.path.join(log_dir, 'yandex_market_ue.log')
        
        fh = RotatingFileHandler(
            log_file,
            maxBytes=10*1024*1024,
            backupCount=5,
            encoding='utf-8'
        )
        fh.setLevel(logging.INFO)
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except Exception as e:
        # Фоллбэк, если даже временная директория недоступна (редкий случай)
        logger.error(f"Не удалось настроить файловый логгер: {e}")
    
    return logger

logger = setup_logging()

# ----------------------------------------------------------------------------
# КОНФИГУРАЦИЯ ИЗ YAML (опционально)
# ----------------------------------------------------------------------------
DEFAULT_CONFIG = {
    'app': {
        'version': '23.4.0',
        'cache_ttl': 3600,
        'lru_cache_size': 128,
        'max_retries': 3,
        'request_timeout': 15
    },
    'api': {
        'base_url': 'https://api.partner.market.yandex.ru',
        'tariffs_endpoint': '/v2/tariffs/calculate',
        'categories_endpoint': '/v2/categories/tree',
        'recommendations_endpoint': '/v2/businesses/{businessId}/offers/recommendations',
        'max_offers_per_request': 200,
        'rate_limit_per_second': 1
    },
    'tax_systems': {
        'usn_6': {'label': 'УСН 6% (доходы)', 'rate': 0.06, 'base': 'revenue'},
        'usn_15': {'label': 'УСН 15% (доходы-расходы)', 'rate': 0.15, 'base': 'profit'},
        'osn': {'label': 'ОСН (общая с НДС 20%)', 'rate': 0.20, 'base': 'profit_vat'},
        'ausn_8': {'label': 'АУСН 8% (доходы)', 'rate': 0.08, 'base': 'revenue'}
    },
    'excel': {
        'styles': {
            'header_bg': '1F4E79',
            'header_fg': 'FFFFFF',
            'zebra_1': 'F2F2F2',
            'zebra_2': 'FFFFFF'
        }
    }
}

def load_config(config_path: str = 'config.yaml') -> dict:
    """Загрузка конфигурации из YAML."""
    if not YAML_AVAILABLE:
        logger.warning('PyYAML не установлен, используются настройки по умолчанию')
        return DEFAULT_CONFIG
    if not os.path.exists(config_path):
        logger.info(f'Файл {config_path} не найден, используются настройки по умолчанию')
        return DEFAULT_CONFIG
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            logger.info('Конфигурация загружена из YAML')
            return config
    except Exception as e:
        logger.error(f'Ошибка загрузки конфигурации: {e}')
        return DEFAULT_CONFIG

CONFIG = load_config()

APP_VERSION: Final[str] = CONFIG['app']['version']
APP_NAME: Final[str] = "Yandex Market Unit Economics PRO"
CACHE_TTL: Final[int] = CONFIG['app']['cache_ttl']
LRU_CACHE_SIZE: Final[int] = CONFIG['app']['lru_cache_size']
MAX_RETRIES: Final[int] = CONFIG['app']['max_retries']
REQUEST_TIMEOUT: Final[int] = CONFIG['app']['request_timeout']
API_BASE_URL: Final[str] = CONFIG['api']['base_url']
TARIFFS_ENDPOINT: Final[str] = CONFIG['api']['tariffs_endpoint']
CATEGORIES_ENDPOINT: Final[str] = CONFIG['api']['categories_endpoint']
RECOMMENDATIONS_ENDPOINT: Final[str] = CONFIG['api']['recommendations_endpoint']
MAX_OFFERS_PER_REQUEST: Final[int] = CONFIG['api']['max_offers_per_request']
RATE_LIMIT_PER_SECOND: Final[int] = CONFIG['api']['rate_limit_per_second']

# ============================================================================
# ENUM для программ продаж
# ============================================================================
class SellingProgram(str, Enum):
    FBS = "FBS"
    FBY = "FBY"
    DBS = "DBS"
    EXPRESS = "EXPRESS"
    LAAS = "LAAS"

# ============================================================================
# ТИПИЗИРОВАННЫЕ КОНФИГУРАЦИИ
# ============================================================================
class TariffDict(TypedDict, total=False):
    category: str
    commission_rate: float
    min_commission: float
    sorting_cost: float
    delivery_rate: float
    delivery_min: float
    delivery_max: float
    acquiring_transfer_rate: float
    acquiring_sku_cost: float
    return_rate: float
    return_processing: float
    storage_fee_per_day: float
    special_tariff_rate: float
    source: str
    scheme: str

class PaymentRate(TypedDict):
    frequency: str
    rate: float

# ============================================================================
# УТИЛИТЫ (БЛОК 0)
# ============================================================================
class NumericUtils:
    """Утилиты для точных и БЫСТРЫХ денежных расчётов (нативная векторизация NumPy)."""
    
    @staticmethod
    def money_round(values: Union[pd.Series, np.ndarray]) -> Union[pd.Series, np.ndarray]:
        """Быстрое округление до 2 знаков без создания объектов Decimal."""
        if isinstance(values, pd.Series):
            return values.fillna(0.0).round(2)
        return np.round(np.nan_to_num(values, nan=0.0, posinf=0.0, neginf=0.0), 2)
    
    @staticmethod
    def percent_round(values: Union[pd.Series, np.ndarray]) -> Union[pd.Series, np.ndarray]:
        """Быстрое округление процентов до 2 знаков."""
        if isinstance(values, pd.Series):
            return values.fillna(0.0).round(2)
        return np.round(np.nan_to_num(values, nan=0.0, posinf=0.0, neginf=0.0), 2)
    
    @staticmethod
    def safe_divide(
        numerator: Union[pd.Series, np.ndarray],
        denominator: Union[pd.Series, np.ndarray],
        default: float = 0.0
    ) -> Union[pd.Series, np.ndarray]:
        """Безопасное деление без предупреждений и деления на ноль."""
        if isinstance(numerator, pd.Series):
            return np.where(denominator != 0, numerator / denominator, default)
        
        # Для numpy массивов
        with np.errstate(divide='ignore', invalid='ignore'):
            result = np.divide(numerator, denominator)
            result[~np.isfinite(result)] = default
            return result

class DtypeOptimizer:
    """Оптимизация типов данных для экономии памяти."""
    
    INT_COLS = {'daily_sales', 'stock_depth_days', 'quantity_per_order'}
    FLOAT_COLS = {
        'selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm',
        'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost',
        'marketing_budget_per_unit', 'warehouse_cost'
    }
    
    @classmethod
    def optimize(cls, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        for col in ['artikul', 'category', 'abc_category', 'xyz_category', 
                     'profitability_status', 'source', 'scheme']:
            if col in df.columns:
                df[col] = df[col].astype('category')
        for col in cls.INT_COLS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                df[col] = df[col].astype(np.int32)
        for col in cls.FLOAT_COLS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                df[col] = df[col].astype(np.float32)
        if 'is_special_tariff' in df.columns:
            df['is_special_tariff'] = df['is_special_tariff'].astype(bool)
        return df

class StringUtils:
    """Утилиты для работы со строками."""
    
    @staticmethod
    def fix_double_utf8(text: str) -> str:
        """Безопасное исправление двойной кодировки (mojibake)."""
        if not isinstance(text, str) or not text:
            return text
        
        # Проверяем наличие типичных артефактов двойной кодировки UTF-8 -> Windows-1251
        if 'Ð' in text or 'Ã' in text or 'â€™' in text:
            try:
                # Пытаемся декодировать как если бы UTF-8 был прочитан как cp1251
                fixed = text.encode('cp1251').decode('utf-8')
                # Дополнительная проверка, чтобы не испортить валидные строки
                if fixed and len(fixed) > 0:
                    return fixed
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
            
            try:
                fixed = text.encode('latin1').decode('utf-8')
                if fixed and len(fixed) > 0:
                    return fixed
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
                
        return text
    
    @staticmethod
    def make_hash(obj: Any) -> str:
        try:
            if isinstance(obj, pd.DataFrame):
                return hashlib.sha256(
                    pd.util.hash_pandas_object(obj, index=True).values.tobytes()
                ).hexdigest()[:16]
            return hashlib.sha256(str(obj).encode('utf-8')).hexdigest()[:16]
        except Exception:
            return hashlib.sha256(b"hash_fallback").hexdigest()[:16]

class LRUCache:
    """LRU-кэш с TTL для тарифов."""
    
    def __init__(self, max_size: int = LRU_CACHE_SIZE, ttl: int = CACHE_TTL):
        self.max_size = max_size
        self.ttl = ttl
        self._cache: OrderedDict[str, Tuple[Any, float]] = OrderedDict()
    
    def get(self, key: str) -> Optional[Any]:
        if key not in self._cache:
            return None
        value, timestamp = self._cache[key]
        if time.time() - timestamp > self.ttl:
            del self._cache[key]
            return None
        self._cache.move_to_end(key)
        return value
    
    def set(self, key: str, value: Any) -> None:
        if key in self._cache:
            self._cache.move_to_end(key)
        self._cache[key] = (value, time.time())
        if len(self._cache) > self.max_size:
            self._cache.popitem(last=False)
    
    def clear(self) -> None:
        self._cache.clear()

# ============================================================================
# МОНИТОРИНГ ПРОИЗВОДИТЕЛЬНОСТИ
# ============================================================================
class PerformanceMonitor:
    def __init__(self):
        self.metrics = {
            'load_time': [],
            'calc_time': [],
            'export_time': [],
            'api_time': [],
            'memory_usage': []
        }
        self._start_time = None
        self._start_memory = None
    
    @contextmanager
    def measure(self, metric_name: str):
        start = time.time()
        start_memory = self._get_memory()
        try:
            yield
        finally:
            end = time.time()
            end_memory = self._get_memory()
            duration = end - start
            memory_delta = end_memory - start_memory if start_memory is not None else 0
            if metric_name not in self.metrics:
                self.metrics[metric_name] = []
            self.metrics[metric_name].append({
                'duration': duration,
                'memory_mb': memory_delta
            })
            logger.info(f"{metric_name} выполнен за {duration:.2f}с, память: {memory_delta:.2f} МБ")
    
    def _get_memory(self) -> Optional[float]:
        if PSUTIL_AVAILABLE:
            try:
                process = psutil.Process()
                return process.memory_info().rss / 1024 / 1024
            except Exception:
                return None
        return None
    
    def report(self) -> Dict[str, Any]:
        report = {}
        for key, values in self.metrics.items():
            if values:
                durations = [v['duration'] for v in values if 'duration' in v]
                memories = [v['memory_mb'] for v in values if 'memory_mb' in v and v['memory_mb'] is not None]
                report[key] = {
                    'count': len(values),
                    'avg_duration': float(np.mean(durations)) if durations else 0.0,
                    'total_duration': float(np.sum(durations)) if durations else 0.0,
                    'avg_memory': float(np.mean(memories)) if memories else 0.0,
                    'max_memory': float(np.max(memories)) if memories else 0.0
                }
        return report
    
    def display_report(self):
        report = self.report()
        if not report:
            return
        st.subheader("📊 Отчёт по производительности")
        for key, data in report.items():
            st.write(f"**{key}**: {data['count']} операций, "
                     f"среднее {data['avg_duration']:.2f}с, "
                     f"память {data['avg_memory']:.2f} МБ")

# ============================================================================
# ОПТИМИЗАЦИЯ ПАМЯТИ
# ============================================================================
class MemoryOptimizer:
    @staticmethod
    def downcast_floats(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        float_cols = df.select_dtypes(include=['float64']).columns
        for col in float_cols:
            col_min = df[col].min()
            col_max = df[col].max()
            if pd.isna(col_min) or pd.isna(col_max):
                continue
            if col_min > np.finfo(np.float16).min and col_max < np.finfo(np.float16).max:
                df[col] = df[col].astype(np.float16)
            elif col_min > np.finfo(np.float32).min and col_max < np.finfo(np.float32).max:
                df[col] = df[col].astype(np.float32)
        return df
    
    @staticmethod
    def downcast_integers(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        int_cols = df.select_dtypes(include=['int64']).columns
        for col in int_cols:
            col_min = df[col].min()
            col_max = df[col].max()
            if pd.isna(col_min) or pd.isna(col_max):
                continue
            if col_min >= 0:
                if col_max <= np.iinfo(np.uint8).max:
                    df[col] = df[col].astype(np.uint8)
                elif col_max <= np.iinfo(np.uint16).max:
                    df[col] = df[col].astype(np.uint16)
                elif col_max <= np.iinfo(np.uint32).max:
                    df[col] = df[col].astype(np.uint32)
            else:
                if col_min >= np.iinfo(np.int8).min and col_max <= np.iinfo(np.int8).max:
                    df[col] = df[col].astype(np.int8)
                elif col_min >= np.iinfo(np.int16).min and col_max <= np.iinfo(np.int16).max:
                    df[col] = df[col].astype(np.int16)
                elif col_min >= np.iinfo(np.int32).min and col_max <= np.iinfo(np.int32).max:
                    df[col] = df[col].astype(np.int32)
        return df
    
    @classmethod
    def optimize_all(cls, df: pd.DataFrame) -> pd.DataFrame:
        df = cls.downcast_floats(df)
        df = cls.downcast_integers(df)
        return df

# ============================================================================
# ПАРАЛЛЕЛЬНАЯ ОБРАБОТКА (Безопасная для Pandas)
# ============================================================================
class ParallelProcessor:
    @staticmethod
    def chunk_dataframe(df: pd.DataFrame, n_chunks: int) -> List[pd.DataFrame]:
        if n_chunks <= 1:
            return [df]
        chunk_size = max(1, len(df) // n_chunks)
        return [df.iloc[i:i+chunk_size].copy() for i in range(0, len(df), chunk_size)]
    
    @classmethod
    def process_in_parallel(
        cls,
        df: pd.DataFrame,
        func: Callable[[pd.DataFrame], pd.DataFrame],
        n_workers: Optional[int] = None,
        chunk_size: Optional[int] = None
    ) -> pd.DataFrame:
        """
        Используем ТОЛЬКО ThreadPoolExecutor. 
        ProcessPoolExecutor для Pandas вызывает огромный оверхед на pickle-сериализацию 
        при передаче данных между процессами, что делает его медленнее последовательного 
        выполнения для чанков < 50 000 строк. NumPy освобождает GIL, поэтому потоки эффективны.
        """
        if not CONCURRENT_AVAILABLE:
            logger.warning("Модуль concurrent.futures не доступен, выполняем последовательно")
            return func(df)
            
        if n_workers is None:
            n_workers = max(1, mp.cpu_count() - 1)
            
        if chunk_size is not None:
            chunks = [df.iloc[i:i+chunk_size].copy() for i in range(0, len(df), chunk_size)]
        else:
            chunks = cls.chunk_dataframe(df, n_workers)
            
        if len(chunks) <= 1:
            return func(df)
            
        logger.info(f"Потоковая обработка: {len(chunks)} чанков, {n_workers} потоков")
        
        # Используем ThreadPoolExecutor, так как он безопасен и эффективен для C-расширений Pandas/NumPy
        with ThreadPoolExecutor(max_workers=n_workers) as executor:
            results = list(executor.map(func, chunks))
            
        return pd.concat(results, ignore_index=True)

# ============================================================================
# МОДЕЛИ (БЛОК 1)
# ============================================================================
@dataclass(frozen=True)
class TaxConfig:
    label: str
    rate: float
    base: str
    min_rate: float

class TaxSystem(Enum):
    USN_6 = TaxConfig("УСН 6% (доходы)", 0.06, "revenue", 0.0)
    USN_15 = TaxConfig("УСН 15% (доходы-расходы)", 0.15, "profit", 0.01)
    OSN = TaxConfig("ОСН (общая с НДС 20%)", 0.20, "profit_vat", 0.0)
    AUSN_8 = TaxConfig("АУСН 8% (доходы)", 0.08, "revenue", 0.0)
    
    @classmethod
    def by_label(cls, label: str) -> TaxConfig:
        for item in cls:
            if item.value.label == label:
                return item.value
        return cls.USN_6.value

@dataclass
class Tariff:
    category: str
    commission_rate: float = 0.15
    min_commission: float = 0.0
    sorting_cost: float = 45.0
    delivery_rate: float = 0.045
    delivery_min: float = 60.0
    delivery_max: float = 500.0
    acquiring_transfer_rate: float = 0.016
    acquiring_sku_cost: float = 0.12
    return_rate: float = 0.05
    return_processing: float = 15.0
    storage_fee_per_day: float = 0.50
    special_tariff_rate: float = 0.15
    source: str = "Базовый фоллбэк"
    scheme: str = "FBS"
    
    def __post_init__(self):
        self.category = str(self.category).lower().strip()
        self.commission_rate = max(0.0, float(self.commission_rate))
        self.min_commission = max(0.0, float(self.min_commission))
        self.sorting_cost = max(0.0, float(self.sorting_cost))
        self.delivery_rate = max(0.0, float(self.delivery_rate))
        self.delivery_min = max(0.0, float(self.delivery_min))
        self.delivery_max = max(float(self.delivery_min), float(self.delivery_max))
        self.acquiring_transfer_rate = max(0.0, float(self.acquiring_transfer_rate))
        self.acquiring_sku_cost = max(0.0, float(self.acquiring_sku_cost))
        self.return_rate = max(0.0, float(self.return_rate))
        self.return_processing = max(0.0, float(self.return_processing))
        self.storage_fee_per_day = max(0.0, float(self.storage_fee_per_day))
        self.special_tariff_rate = max(0.0, float(self.special_tariff_rate))
    
    def to_dict(self) -> TariffDict:
        return {
            'category': self.category,
            'commission_rate': self.commission_rate,
            'min_commission': self.min_commission,
            'sorting_cost': self.sorting_cost,
            'delivery_rate': self.delivery_rate,
            'delivery_min': self.delivery_min,
            'delivery_max': self.delivery_max,
            'acquiring_transfer_rate': self.acquiring_transfer_rate,
            'acquiring_sku_cost': self.acquiring_sku_cost,
            'return_rate': self.return_rate,
            'return_processing': self.return_processing,
            'storage_fee_per_day': self.storage_fee_per_day,
            'special_tariff_rate': self.special_tariff_rate,
            'source': self.source,
            'scheme': self.scheme
        }
    
    @classmethod
    def default(cls, category: str = "default") -> 'Tariff':
        return cls(category=category)

# ============================================================================
# API КЛИЕНТ (БЛОК 2)
# ============================================================================
class RateLimiter:
    def __init__(self, max_calls: int = RATE_LIMIT_PER_SECOND, period: float = 1.0):
        self.max_calls = max_calls
        self.period = period
        self.calls: List[float] = []
    
    def wait_if_needed(self) -> None:
        now = time.time()
        self.calls = [t for t in self.calls if now - t < self.period]
        if len(self.calls) >= self.max_calls:
            sleep_time = self.period - (now - self.calls[0])
            if sleep_time > 0:
                time.sleep(sleep_time)
        self.calls.append(time.time())

class APIClient:
    """HTTP-клиент с retry, backoff и rate limiting (1 запрос/сек)."""
    
    def __init__(
        self,
        base_url: str,
        api_key: str,
        max_retries: int = MAX_RETRIES,
        timeout: int = REQUEST_TIMEOUT
    ):
        self.base_url = base_url.rstrip('/')
        self.api_key = api_key
        self.timeout = timeout
        self.rate_limiter = RateLimiter(max_calls=RATE_LIMIT_PER_SECOND, period=1.0)
        self.max_retries = max_retries
        
        self.session = requests.Session()
        retry_strategy = Retry(
            total=max_retries,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS", "TRACE"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
    
    @contextmanager
    def _request_context(self):
        self.rate_limiter.wait_if_needed()
        try:
            yield
        finally:
            pass
    
    def _request(
        self,
        method: str,
        endpoint: str,
        **kwargs
    ) -> Dict[str, Any]:
        if not self.api_key:
            return {}
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        with self._request_context():
            for attempt in range(self.max_retries + 1):
                try:
                    resp = self.session.request(
                        method,
                        url,
                        timeout=self.timeout,
                        **kwargs
                    )
                    resp.raise_for_status()
                    return resp.json()
                except requests.exceptions.Timeout:
                    wait_time = 2 ** attempt
                    logger.warning(f"Таймаут {self.timeout}s, попытка {attempt+1}/{self.max_retries+1}, ждём {wait_time}s")
                    if attempt < self.max_retries:
                        time.sleep(wait_time)
                    else:
                        logger.error(f"Таймаут после {self.max_retries+1} попыток: {url}")
                except requests.exceptions.HTTPError as e:
                    if e.response.status_code in [429, 500, 502, 503, 504]:
                        wait_time = 2 ** attempt
                        logger.warning(f"HTTP {e.response.status_code}, попытка {attempt+1}, ждём {wait_time}s")
                        if attempt < self.max_retries:
                            time.sleep(wait_time)
                        else:
                            logger.error(f"HTTP ошибка после {self.max_retries+1} попыток: {e}")
                    else:
                        logger.warning(f"HTTP {e.response.status_code}: {e.response.text[:200]}")
                        logger.debug(f"Response body: {e.response.text}")
                        break
                except Exception as e:
                    logger.warning(f"Ошибка запроса {url}: {e}")
                    break
            return {}

class YandexMarketAPI(APIClient):
    """API-клиент Яндекс Маркета с поддержкой категорий и пакетной отправки."""
    
    def __init__(self, api_key: str, business_id: Optional[str] = None):
        super().__init__(API_BASE_URL, api_key)
        self.business_id = business_id
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if business_id:
            self.headers["X-Business-Id"] = business_id
        self._category_cache: Dict[str, int] = {}
    
    def get_campaigns(self) -> List[Dict[str, Any]]:
        data = self._request("GET", "/campaigns", headers=self.headers)
        return data.get("campaigns", [])
    
    def get_categories_tree(self) -> List[Dict[str, Any]]:
        """Получение дерева категорий."""
        data = self._request("GET", CATEGORIES_ENDPOINT, headers=self.headers)
        return data.get("result", {}).get("categories", [])
    
    def get_category_id_by_name(self, category_name: str) -> Optional[int]:
        """Поиск categoryId по названию категории с кэшированием (итеративный BFS для безопасности)."""
        if category_name in self._category_cache:
            return self._category_cache[category_name]
        
        categories = self.get_categories_tree()
        
        # Итеративный поиск в ширину (BFS) вместо рекурсии, чтобы избежать RecursionError
        queue = list(categories)
        found_id = None
        
        while queue:
            cat = queue.pop(0)
            if cat.get('name', '').lower() == category_name.lower():
                found_id = cat.get('id')
                break
            if 'children' in cat and cat['children']:
                queue.extend(cat['children'])
        
        if found_id is not None:
            self._category_cache[category_name] = found_id
            
        return found_id
    
    def calculate_tariffs_batch(
        self,
        offers: List[Dict[str, Any]],
        campaign_id: Optional[int] = None,
        selling_program: SellingProgram = SellingProgram.FBS,
        transit_warehouse_type: Optional[str] = None,
        order_cargo_type: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Отправка офферов пачками (макс. MAX_OFFERS_PER_REQUEST).
        Возвращает список результатов для всех офферов.
        """
        if not offers:
            return []
        
        all_results = []
        for i in range(0, len(offers), MAX_OFFERS_PER_REQUEST):
            chunk = offers[i:i+MAX_OFFERS_PER_REQUEST]
            payload = {
                "parameters": {
                    "sellingProgram": selling_program.value,
                    "frequency": "WEEKLY",
                    "paymentDelayWeeks": 4,
                    "currency": "RUR"
                },
                "offers": chunk
            }
            if campaign_id:
                payload["parameters"]["campaignId"] = campaign_id
                del payload["parameters"]["sellingProgram"]
            
            if transit_warehouse_type:
                payload["parameters"]["transitWarehouseType"] = transit_warehouse_type
            if order_cargo_type:
                payload["parameters"]["orderCargoType"] = order_cargo_type
            
            data = self._request(
                "POST",
                TARIFFS_ENDPOINT,
                headers=self.headers,
                json=payload
            )
            chunk_results = data.get("result", {}).get("offers", [])
            all_results.extend(chunk_results)
        return all_results
    
    def calculate_tariffs(
        self,
        offers: List[Dict[str, Any]],
        campaign_id: Optional[int] = None,
        selling_program: str = "FBS",
        transit_warehouse_type: Optional[str] = None,
        order_cargo_type: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        try:
            prog = SellingProgram(selling_program.upper())
        except ValueError:
            prog = SellingProgram.FBS
        return self.calculate_tariffs_batch(
            offers, campaign_id, prog, transit_warehouse_type, order_cargo_type
        )
    
    def get_price_recommendations(
        self,
        offers: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """Получение рекомендаций по ценам (метод v2)."""
        if not self.business_id:
            logger.warning("Business ID не указан, рекомендации недоступны")
            return []
        endpoint = RECOMMENDATIONS_ENDPOINT.format(businessId=self.business_id)
        payload = {"offers": offers}
        data = self._request(
            "POST",
            endpoint,
            headers=self.headers,
            json=payload
        )
        return data.get("result", {}).get("offers", [])
    
    def calculate_tariffs_with_retry(
        self,
        offers: List[Dict[str, Any]],
        campaign_id: Optional[int] = None,
        selling_program: str = "FBS",
        fallback_tariff: Optional[Tariff] = None,
        transit_warehouse_type: Optional[str] = None,
        order_cargo_type: Optional[str] = None
    ) -> Optional[Tariff]:
        try:
            results = self.calculate_tariffs(
                offers, campaign_id, selling_program,
                transit_warehouse_type, order_cargo_type
            )
            if results and len(results) > 0:
                return self._parse_tariff(results[0])
        except Exception as e:
            logger.error(f"Ошибка при расчёте тарифов через API: {e}")
        return fallback_tariff or Tariff.default()
    
    def _parse_tariff(self, tariff_data: Dict) -> Tariff:
        return Tariff(
            category=tariff_data.get('category', 'default'),
            commission_rate=tariff_data.get('commission', 0.15) / 100,
            delivery_rate=tariff_data.get('delivery', 0.05),
            source="API Яндекс Маркета",
            scheme=tariff_data.get('sellingProgram', 'FBS')
        )

# ============================================================================
# МЕНЕДЖЕР ТАРИФОВ (БЛОК 3)
# ============================================================================
class HybridTariffManager:
    """Управление тарифами с LRU-кэшем."""
    
    def __init__(self):
        if 'tariffs' not in st.session_state:
            st.session_state.tariffs = {}
        self._cache = LRUCache()
    
    @property
    def tariffs(self) -> Dict[str, Tariff]:
        return st.session_state.tariffs
    
    def load_tariffs_from_file(self, df: pd.DataFrame) -> int:
        required_cols = {'category', 'commission_rate'}
        if not required_cols.issubset(df.columns):
            raise ValueError(
                f"Файл тарифов должен содержать: {required_cols}"
            )
        loaded = 0
        for _, row in df.iterrows():
            cat = str(row.get('category', '')).lower().strip()
            if not cat or cat == 'nan':
                continue
            self.tariffs[cat] = Tariff(
                category=cat,
                commission_rate=float(row.get('commission_rate', 0.15)),
                min_commission=float(row.get('min_commission', 0)),
                sorting_cost=float(row.get('sorting_cost', 45)),
                delivery_rate=float(row.get('delivery_rate', 0.045)),
                delivery_min=float(row.get('delivery_min', 60)),
                delivery_max=float(row.get('delivery_max', 500)),
                acquiring_transfer_rate=float(row.get('acquiring_transfer_rate', 0.016)),
                acquiring_sku_cost=float(row.get('acquiring_sku_cost', 0.12)),
                return_rate=float(row.get('return_rate', 0.05)),
                return_processing=float(row.get('return_processing', 15)),
                storage_fee_per_day=float(row.get('storage_fee_per_day', 0.5)),
                special_tariff_rate=float(row.get('special_tariff_rate', 0.15)),
                source="Загружено пользователем",
                scheme=str(row.get('scheme', 'FBS'))
            )
            loaded += 1
        logger.info(f"Загружено {loaded} тарифов")
        return loaded
    
    def get_best_tariff(
        self,
        category_name: str,
        scheme: str,
        ym_api: Optional[YandexMarketAPI] = None,
        use_api: bool = True,
        offer_params: Optional[Dict[str, Any]] = None
    ) -> Tariff:
        cat_clean = category_name.lower().strip()
        cache_key = f"{cat_clean}_{scheme}"
        cached = self._cache.get(cache_key)
        if cached:
            return cached
        
        if use_api and ym_api and ym_api.api_key and offer_params:
            try:
                offer = {
                    "price": offer_params.get('selling_price', 1000),
                    "weight": offer_params.get('weight_kg', 1.0),
                    "length": offer_params.get('length_cm', 10),
                    "width": offer_params.get('width_cm', 10),
                    "height": offer_params.get('height_cm', 10),
                    "quantity": offer_params.get('quantity_per_order', 1)
                }
                if 'category_id' in offer_params and offer_params['category_id']:
                    offer["categoryId"] = offer_params['category_id']
                else:
                    cat_id = ym_api.get_category_id_by_name(cat_clean)
                    if cat_id:
                        offer["categoryId"] = cat_id
                    else:
                        offer["categoryId"] = 0
                        logger.warning(f"Не найден categoryId для {cat_clean}, используется 0")
                
                transit_type = offer_params.get('transit_warehouse_type')
                cargo_type = offer_params.get('order_cargo_type')
                
                tariff = ym_api.calculate_tariffs_with_retry(
                    [offer],
                    selling_program=scheme,
                    transit_warehouse_type=transit_type,
                    order_cargo_type=cargo_type
                )
                if tariff:
                    tariff.category = cat_clean
                    self._cache.set(cache_key, tariff)
                    return tariff
            except Exception as e:
                logger.warning(f"API ЯМ сбой для {cat_clean}: {e}")
        
        if cat_clean in self.tariffs:
            t = self.tariffs[cat_clean]
            t.scheme = scheme
            self._cache.set(cache_key, t)
            return t
        
        logger.warning(
            f"Тариф для '{cat_clean}' не найден. Применён фоллбэк 15%."
        )
        t = Tariff(
            category=cat_clean,
            commission_rate=0.15,
            source="⚠️ БАЗОВЫЙ ФОЛЛБЭК",
            scheme=scheme
        )
        self._cache.set(cache_key, t)
        return t
    
    def get_tariffs_vectorized(
        self,
        df: pd.DataFrame,
        scheme: str,
        ym_api: Optional[YandexMarketAPI] = None,
        use_api: bool = True
    ) -> pd.DataFrame:
        unique_cats = df['category'].cat.categories if hasattr(df['category'], 'cat') else df['category'].unique()
        
        tariff_map = {}
        for cat in unique_cats:
            sample = df[df['category'] == cat].iloc[0] if not df[df['category'] == cat].empty else {}
            offer_params = {
                'selling_price': sample.get('selling_price', 1000),
                'weight_kg': sample.get('weight_kg', 1.0),
                'length_cm': sample.get('length_cm', 10),
                'width_cm': sample.get('width_cm', 10),
                'height_cm': sample.get('height_cm', 10),
                'quantity_per_order': sample.get('quantity_per_order', 1),
                'category_id': sample.get('category_id', None),
                'transit_warehouse_type': sample.get('transit_warehouse_type', None),
                'order_cargo_type': sample.get('order_cargo_type', None)
            }
            tariff_map[cat] = self.get_best_tariff(cat, scheme, ym_api, use_api, offer_params)
        
        tariff_df = pd.DataFrame([
            {'category': cat, **t.to_dict()}
            for cat, t in tariff_map.items()
        ])
        return tariff_df

# ============================================================================
# ВАЛИДАТОР ДАННЫХ (БЛОК 4)
# ============================================================================
class DataValidator:
    REQUIRED_COLS: Final[Tuple[str, ...]] = (
        'artikul', 'category', 'selling_price', 'cogs'
    )
    NUMERIC_COLS: Final[Tuple[str, ...]] = (
        'selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm',
        'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost',
        'marketing_budget_per_unit', 'stock_depth_days', 'quantity_per_order',
        'daily_sales'
    )
    
    @classmethod
    def validate(cls, df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
        errors = []
        if df.empty:
            return df, ["DataFrame пустой"]
        df_validated = df.copy()
        missing = [c for c in cls.REQUIRED_COLS if c not in df_validated.columns]
        if missing:
            errors.append(f"Отсутствуют обязательные колонки: {missing}")
        for col in cls.NUMERIC_COLS:
            if col in df_validated.columns:
                df_validated[col] = pd.to_numeric(
                    df_validated[col],
                    errors='coerce'
                ).fillna(0).clip(lower=0)
        if 'selling_price' in df_validated.columns:
            zero_prices = (df_validated['selling_price'] == 0).sum()
            if zero_prices > 0:
                errors.append(
                    f"selling_price: {zero_prices} SKU с нулевой ценой"
                )
        if 'quantity_per_order' in df_validated.columns:
            # ФИКС: Жесткая гарантия, что количество в заказе не может быть < 1
            df_validated['quantity_per_order'] = pd.to_numeric(
                df_validated['quantity_per_order'], errors='coerce'
            ).fillna(1).clip(lower=1).astype(int)
        return df_validated, errors

# ============================================================================
# ФИНАНСОВЫЙ ДВИЖОК (БЛОК 5)
# ============================================================================
class FinancialEngine:
    PICK_PACK_COST: Final[float] = 35.0
    SPECIAL_PRICE_THRESHOLD: Final[float] = 300.0
    SPECIAL_VOLUME_THRESHOLD: Final[float] = 5.0
    
    @staticmethod
    def calculate_billable_weight(df: pd.DataFrame) -> np.ndarray:
        vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
        billable = np.maximum(df['weight_kg'].values, vol_weight.values)
        return np.ceil(billable * 2) / 2
    
    @staticmethod
    def calculate_middle_mile(billable_weight: np.ndarray) -> np.ndarray:
        return np.select(
            [billable_weight <= 4, billable_weight <= 10],
            [100, 300],
            default=600
        )
    
    @classmethod
    def calculate_all(
        cls,
        df: pd.DataFrame,
        tax_config: TaxConfig,
        scheme: str,
        payment_rate: float,
        tariffs_map: Dict[str, Dict]
    ) -> pd.DataFrame:
        df = df.copy()
        for col in ['artikul', 'category']:
            if col in df.columns:
                df[col] = df[col].astype(str).apply(StringUtils.fix_double_utf8)
        defaults = {
            'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0,
            'length_cm': 0.0, 'width_cm': 0.0, 'height_cm': 0.0,
            'packaging_cost': 0.0, 'marketing_budget_per_unit': 0.0,
            'daily_sales': 0.0, 'stock_depth_days': 0.0,
            'first_mile_cost': 0.0, 'warehouse_cost': 0.0,
            'volume_liters': 0.0, 'quantity_per_order': 1.0
        }
        for col, default in defaults.items():
            if col not in df.columns:
                df[col] = default
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
        
        tariff_df = pd.DataFrame.from_dict(tariffs_map, orient='index').reset_index()
        tariff_df.columns = ['category'] + list(tariff_df.columns[1:])
        df = df.merge(tariff_df, on='category', how='left')
        default_tariff = Tariff.default().to_dict()
        for col in default_tariff.keys():
            if col in df.columns:
                df[col] = df[col].fillna(default_tariff[col])
        
        df['billable_weight'] = cls.calculate_billable_weight(df)
        df['is_special_tariff'] = (
            (df['selling_price'] <= cls.SPECIAL_PRICE_THRESHOLD) &
            (df['volume_liters'] <= cls.SPECIAL_VOLUME_THRESHOLD)
        )
        df['commission'] = np.where(
            df['is_special_tariff'],
            df['selling_price'] * df['special_tariff_rate'],
            np.maximum(
                df['selling_price'] * df['commission_rate'],
                df['min_commission']
            )
        )
        raw_delivery = df['selling_price'] * df['delivery_rate']
        df['delivery_to_customer'] = np.where(
            df['is_special_tariff'],
            0.0,
            np.clip(raw_delivery, df['delivery_min'], df['delivery_max'])
        )
        df['middle_mile_cost'] = np.where(
            df['is_special_tariff'],
            0.0,
            cls.calculate_middle_mile(df['billable_weight'].values)
        )
        df['sorting_cost'] = np.where(
            df['is_special_tariff'],
            0.0,
            np.where(scheme == 'FBS', df['sorting_cost'], 0.0)
        )
        df['acquiring_sku_cost'] = df['acquiring_sku_cost'] / df['quantity_per_order']
        df['acquiring_transfer_cost'] = df['selling_price'] * payment_rate
        df['acquiring_cost'] = df['acquiring_sku_cost'] + df['acquiring_transfer_cost']
        df['return_processing_cost'] = np.where(
            df['is_special_tariff'], 0.0, df['return_processing']
        )
        df['return_delivery_cost'] = np.where(
            df['is_special_tariff'],
            0.0,
            df['middle_mile_cost'] * df['return_rate']
        )
        df['return_cost'] = df['return_processing_cost'] + df['return_delivery_cost']
        df['pick_pack_cost'] = cls.PICK_PACK_COST
        df['warehouse_cost'] = np.where(
            df['warehouse_cost'] == 0,
            (df['stock_depth_days'] * df['daily_sales']) * df['storage_fee_per_day'],
            df['warehouse_cost']
        )
        df['fixed_operational_costs'] = (
            df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
            df['packaging_cost'] + df['return_cost'] +
            df['marketing_budget_per_unit'] + df['warehouse_cost']
        )
        df['marketplace_fees'] = (
            df['commission'] + df['delivery_to_customer'] +
            df['middle_mile_cost'] + df['sorting_cost'] + df['acquiring_cost']
        )
        df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']
        if tax_config.base == "revenue":
            df['tax_cost'] = df['selling_price'] * tax_config.rate
        elif tax_config.base == "profit_vat":
            revenue_without_vat = df['selling_price'] / 1.20
            pre_tax_profit = revenue_without_vat - df['pre_tax_expenses']
            df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_config.rate
        else:
            pre_tax_profit = df['selling_price'] - df['pre_tax_expenses']
            df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_config.rate
        df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
        df['gross_profit'] = df['selling_price'] - df['total_expenses']
        df['margin_percent'] = np.where(
            df['selling_price'] > 0,
            (df['gross_profit'] / df['selling_price']) * 100,
            0.0
        )
        var_fees = np.where(
            df['is_special_tariff'],
            df['special_tariff_rate'] + payment_rate + (
                tax_config.rate if tax_config.base == "revenue" else 0
            ),
            df['commission_rate'] + df['delivery_rate'] + payment_rate + (
                tax_config.rate if tax_config.base == "revenue" else 0
            )
        )
        denom = 1.0 - var_fees
        denom_valid = denom > 0.01
        fixed_no_return = (
            df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
            df['packaging_cost'] + df['marketing_budget_per_unit'] +
            df['warehouse_cost']
        )
        df['rec_price_min'] = np.where(
            denom_valid,
            NumericUtils.safe_divide(fixed_no_return.values, denom.values),
            np.nan
        )
        df['rec_price_15'] = np.where(
            denom_valid,
            NumericUtils.safe_divide(
                fixed_no_return.values,
                denom.values - 0.15
            ),
            np.nan
        )
        df['rec_price_25'] = np.where(
            denom_valid,
            NumericUtils.safe_divide(
                fixed_no_return.values,
                denom.values - 0.25
            ),
            np.nan
        )
        df['variable_costs'] = (
            df['commission'] + df['delivery_to_customer'] +
            df['middle_mile_cost'] + df['sorting_cost'] +
            df['acquiring_cost'] + df['return_cost']
        )
        df['fixed_costs'] = fixed_no_return + df['return_cost']
        df['contribution_margin'] = df['selling_price'] - df['variable_costs']
        df['roi_percent'] = np.where(
            df['cogs'] > 0,
            (df['gross_profit'] / df['cogs']) * 100,
            0.0
        )
        df['break_even_units'] = NumericUtils.safe_divide(
            df['fixed_costs'].values,
            df['contribution_margin'].values
        )
        df['abc_category'] = np.select(
            [df['daily_sales'] >= 10, df['daily_sales'] >= 3],
            ['A', 'B'],
            default='C'
        )
        df['xyz_category'] = np.select(
            [df['margin_percent'] >= 20, df['margin_percent'] >= 10],
            ['X', 'Y'],
            default='Z'
        )
        df['abc_xyz'] = df['abc_category'] + df['xyz_category']
        df['profitability_status'] = np.where(
            df['gross_profit'] > 0,
            np.where(
                df['margin_percent'] >= 20,
                'Высокомаржинальный',
                'Низкомаржинальный'
            ),
            'Убыточный'
        )
        money_cols = [
            'commission', 'delivery_to_customer', 'middle_mile_cost',
            'sorting_cost', 'acquiring_cost', 'return_cost',
            'gross_profit', 'total_expenses', 'rec_price_min',
            'rec_price_15', 'rec_price_25', 'tax_cost'
        ]
        for col in money_cols:
            if col in df.columns:
                df[col] = NumericUtils.money_round(df[col])
        pct_cols = ['margin_percent', 'roi_percent']
        for col in pct_cols:
            if col in df.columns:
                df[col] = NumericUtils.percent_round(df[col])
        
        # ВАЖНО: Мы сохраняем 'source' и 'scheme' для аудита, удаляем только чисто технические расчетные поля,
        # которые не несут смысловой нагрузки для конечного пользователя, но оставляем ключевые.
        tech_cols_to_drop = [
            'is_special_tariff', 'billable_weight'
        ]
        df = df.drop(columns=[c for c in tech_cols_to_drop if c in df.columns])
        return df

# ============================================================================
# КЭШИРОВАННЫЙ РАСЧЁТ
# ============================================================================
# ФИКС: Добавлен hash_funcs для корректного хеширования словаря tariffs_map
@st.cache_data(ttl=CACHE_TTL, show_spinner=False, hash_funcs={dict: lambda d: str(sorted(d.items()))})
def run_calculations_cached(
    df_hash: str,
    df: pd.DataFrame,
    tax_label: str,
    scheme_label: str,
    payment_frequency: str,
    tariffs_map: Dict[str, Dict]
) -> pd.DataFrame:
    if df.empty:
        return df
    tax_config = TaxSystem.by_label(tax_label)
    scheme = scheme_label.split(" ")[0]
    payment_rates = {
        "Ежемесячно (1.0%)": 0.01,
        "Раз в 2 недели (1.3%)": 0.013,
        "Еженедельно, 4 нед. (1.6%)": 0.016,
        "Ежедневно (3.3%)": 0.033
    }
    payment_rate = payment_rates.get(payment_frequency, 0.016)
    result = FinancialEngine.calculate_all(
        df, tax_config, scheme, payment_rate, tariffs_map
    )
    result = DtypeOptimizer.optimize(result)
    # ФИКС: Удален дублирующий вызов MemoryOptimizer.optimize_all(result)
    return result

# ============================================================================
# ЭКСПОРТ EXCEL — ULTRA DESIGN v3
# ============================================================================
class UltimateExcelExporter:
    C = {
        "navy":       "1F4E79",
        "blue":       "2E75B6",
        "light_blue": "BDD7EE",
        "sky":        "DEEAF1",
        "green":      "375623",
        "lime":       "E2EFDA",
        "yellow":     "FFF2CC",
        "orange":     "FCE4D6",
        "red":        "FF0000",
        "crimson":    "C00000",
        "white":      "FFFFFF",
        "gray":       "F2F2F2",
        "dark_gray":  "595959",
        "mid_gray":   "A6A6A6",
        "profit_hi":  "375623",
        "profit_lo":  "7F6000",
        "loss":       "C00000",
        "profit_hi_bg": "E2EFDA",
        "profit_lo_bg": "FFEB9C",
        "loss_bg":    "FFC7CE",
        "abc_a":      "1F4E79",
        "abc_b":      "2E75B6",
        "abc_c":      "BDD7EE",
    }

    @classmethod
    def _fill(cls, hex_color: str) -> PatternFill:
        return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    @classmethod
    def _font(
        cls,
        bold: bool = False,
        color: str = "000000",
        size: int = 10,
        italic: bool = False,
        name: str = "Calibri",
    ) -> Font:
        return Font(bold=bold, color=color, size=size, italic=italic, name=name)

    @classmethod
    def _align(cls, h: str = "left", v: str = "center", wrap: bool = False) -> Alignment:
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    @classmethod
    def _style_header_row(cls, ws, row: int, col_start: int, col_end: int, bg: str = "navy") -> None:
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = cls._fill(cls.C[bg])
            cell.font = cls._font(bold=True, color="FFFFFF", size=10)
            cell.alignment = cls._align(h="center", wrap=True)

    @classmethod
    def _apply_zebra(cls, ws, row_start: int, row_end: int, col_start: int, col_end: int) -> None:
        for r in range(row_start, row_end + 1):
            bg = cls.C["gray"] if r % 2 == 0 else cls.C["white"]
            for c in range(col_start, col_end + 1):
                cell = ws.cell(row=r, column=c)
                if cell.fill.patternType in (None, "none", ""):
                    cell.fill = cls._fill(bg)

    FMT_MONEY   = '#,##0.00 ₽'
    FMT_PCT     = '0.00%'
    FMT_PCT1    = '0.0%'
    FMT_INT     = '#,##0'
    FMT_DATE    = 'DD.MM.YYYY HH:MM'

    @classmethod
    def export_max_info(
        cls,
        df: pd.DataFrame,
        tax_label: str,
        scheme_label: str,
        payment_frequency: str,
        include_advanced: bool = True
    ) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty:
            return b""

        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
            from openpyxl.utils import get_column_letter
            from openpyxl.worksheet.table import Table, TableStyleInfo
            from openpyxl.chart import PieChart, Reference, LineChart
            from openpyxl.chart.label import DataLabelList
        except ImportError as e:
            logger.error(f"openpyxl import error: {e}")
            return b""

        def thin_border(left=True, right=True, top=True, bottom=True) -> Border:
            s = Side(style="thin", color="D9D9D9")
            return Border(left=s if left else None, right=s if right else None,
                          top=s if top else None, bottom=s if bottom else None)

        def thick_border_bottom() -> Border:
            s_thin = Side(style="thin", color="D9D9D9")
            s_med = Side(style="medium", color=cls.C["navy"])
            return Border(left=s_thin, right=s_thin, top=s_thin, bottom=s_med)

        df = df.copy()
        for col in df.select_dtypes(include=['category']).columns:
            df[col] = df[col].astype(str)

        expected_cols = [
            'artikul', 'category', 'selling_price', 'cogs',
            'commission', 'delivery_to_customer', 'middle_mile_cost',
            'sorting_cost', 'acquiring_cost', 'return_cost',
            'pick_pack_cost', 'packaging_cost', 'first_mile_cost',
            'marketing_budget_per_unit', 'warehouse_cost',
            'tax_cost', 'total_expenses', 'gross_profit',
            'margin_percent', 'roi_percent', 'break_even_units',
            'rec_price_min', 'rec_price_15', 'rec_price_25',
            'daily_sales', 'abc_category', 'xyz_category',
            'abc_xyz', 'profitability_status', 'source', 'scheme'
        ]
        for col in expected_cols:
            if col not in df.columns:
                if col in ('artikul', 'category', 'source', 'scheme'):
                    df[col] = "—"
                elif col in ('margin_percent', 'roi_percent'):
                    df[col] = 0.0
                else:
                    df[col] = 0.0

        wb = Workbook()

        # ① ДАШБОРД
        ws_dash = wb.active
        ws_dash.title = "📊 Дашборд"
        ws_dash.sheet_view.showGridLines = False
        ws_dash.column_dimensions["A"].width = 2
        ws_dash.column_dimensions["Q"].hidden = True
        ws_dash.column_dimensions["R"].hidden = True

        ws_dash.row_dimensions[1].height = 8
        ws_dash.row_dimensions[2].height = 40
        ws_dash.row_dimensions[3].height = 14

        ws_dash.merge_cells("B2:N2")
        title_cell = ws_dash["B2"]
        title_cell.value = "📊 UNIT-ECONOMICS · ЯНДЕКС МАРКЕТ"
        title_cell.font = Font(name="Calibri", bold=True, size=20, color=cls.C["white"])
        title_cell.fill = cls._fill(cls.C["navy"])
        title_cell.alignment = cls._align(h="center")

        ws_dash.merge_cells("B3:N3")
        sub_cell = ws_dash["B3"]
        sub_cell.value = f"{scheme_label} · {tax_label} · {payment_frequency} · Сформировано: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
        sub_cell.font = Font(name="Calibri", italic=True, size=10, color=cls.C["white"])
        sub_cell.fill = cls._fill(cls.C["blue"])
        sub_cell.alignment = cls._align(h="center")

        ws_dash.row_dimensions[4].height = 10

        total_sku = len(df)
        avg_margin = df["margin_percent"].mean() if "margin_percent" in df.columns else 0
        profitable_cnt = int((df["gross_profit"] > 0).sum()) if "gross_profit" in df.columns else 0
        loss_cnt = int((df["gross_profit"] < 0).sum()) if "gross_profit" in df.columns else 0
        hi_margin_cnt = int((df["profitability_status"] == "Высокомаржинальный").sum()) if "profitability_status" in df.columns else 0
        avg_roi = df["roi_percent"].mean() if "roi_percent" in df.columns else 0

        kpi_cards = [
            ("Всего SKU", total_sku, "#,##0", cls.C["navy"], cls.C["sky"], "🏷️"),
            ("Средняя маржа", avg_margin / 100, "0.0%", cls.C["green"], cls.C["lime"], "📈"),
            ("Прибыльных", profitable_cnt, "#,##0", cls.C["green"], cls.C["lime"], "✅"),
            ("Убыточных", loss_cnt, "#,##0", cls.C["crimson"], cls.C["orange"], "❌"),
            ("Высокомаржинальных", hi_margin_cnt, "#,##0", cls.C["navy"], cls.C["sky"], "⭐"),
            ("Средний ROI", avg_roi / 100, "0.0%", cls.C["blue"], cls.C["light_blue"], "💰"),
        ]

        card_cols = [("B", "C"), ("D", "E"), ("F", "G"), ("H", "I"), ("J", "K"), ("L", "M")]

        for idx, ((label, value, fmt, fg, bg, icon), (start_col, end_col)) in enumerate(zip(kpi_cards, card_cols)):
            sc = ord(start_col) - ord("A") + 1
            ec = ord(end_col) - ord("A") + 1

            for r in range(5, 10):
                for c in range(sc, ec + 1):
                    cell = ws_dash.cell(row=r, column=c)
                    cell.fill = cls._fill(bg)
                    cell.border = thin_border()

            ws_dash.merge_cells(start_row=5, start_column=sc, end_row=5, end_column=ec)
            lbl = ws_dash.cell(row=5, column=sc, value=f"{icon} {label}")
            lbl.font = Font(name="Calibri", bold=True, size=9, color=fg)
            lbl.fill = cls._fill(bg)
            lbl.alignment = cls._align(h="center")

            ws_dash.merge_cells(start_row=6, start_column=sc, end_row=8, end_column=ec)
            val_cell = ws_dash.cell(row=6, column=sc, value=value)
            val_cell.font = Font(name="Calibri", bold=True, size=22, color=fg)
            val_cell.fill = cls._fill(bg)
            val_cell.number_format = fmt
            val_cell.alignment = cls._align(h="center", v="center")

            ws_dash.row_dimensions[5].height = 18
            ws_dash.row_dimensions[6].height = 38
            ws_dash.row_dimensions[7].height = 38
            ws_dash.row_dimensions[8].height = 18

        ws_dash.row_dimensions[10].height = 10

        status_header_row = 11
        ws_dash.merge_cells(f"B{status_header_row}:F{status_header_row}")
        h = ws_dash.cell(row=status_header_row, column=2, value="Структура портфеля по маржинальности")
        h.font = cls._font(bold=True, color="FFFFFF", size=11)
        h.fill = cls._fill(cls.C["navy"])
        h.alignment = cls._align(h="center")

        if "profitability_status" in df.columns and "gross_profit" in df.columns:
            status_summary = (
                df.groupby("profitability_status", observed=True)
                  .agg(SKU=("artikul", "count"), Прибыль=("gross_profit", "sum"), Маржа_avg=("margin_percent", "mean"))
                  .reset_index().rename(columns={"profitability_status": "Статус"})
            )

            col_hdrs = ["Статус", "SKU, шт.", "Сум. прибыль, ₽", "Ср. маржа, %"]
            STATUS_COLORS = {
                "Высокомаржинальный": (cls.C["profit_hi"], cls.C["profit_hi_bg"]),
                "Низкомаржинальный": (cls.C["profit_lo"], cls.C["profit_lo_bg"]),
                "Убыточный": (cls.C["loss"], cls.C["loss_bg"]),
            }

            for ci, hdr in enumerate(col_hdrs, 2):
                cell = ws_dash.cell(row=status_header_row + 1, column=ci, value=hdr)
                cell.font = cls._font(bold=True, color="FFFFFF", size=9)
                cell.fill = cls._fill(cls.C["blue"])
                cell.alignment = cls._align(h="center")
                cell.border = thin_border()

            for ri, row_data in enumerate(status_summary.itertuples(), 2):
                fg, bg = STATUS_COLORS.get(row_data.Статус, ("000000", "FFFFFF"))
                row_idx = status_header_row + ri
                values = [row_data.Статус, row_data.SKU, row_data.Прибыль, row_data.Маржа_avg / 100 if row_data.Маржа_avg else 0]
                fmts = [None, "#,##0", "#,##0.00 ₽", "0.0%"]

                for ci, (val, fmt2) in enumerate(zip(values, fmts), 2):
                    cell = ws_dash.cell(row=row_idx, column=ci, value=val)
                    cell.fill = cls._fill(bg)
                    cell.font = Font(name="Calibri", bold=(ci == 2), color=fg, size=10)
                    cell.alignment = cls._align(h="center" if ci > 2 else "left")
                    cell.border = thin_border()
                    if fmt2:
                        cell.number_format = fmt2

        if "profitability_status" in df.columns:
            try:
                pc = PieChart()
                pc.title = "Структура по статусам"
                pc.width = 14
                pc.height = 10
                status_counts = df["profitability_status"].value_counts()
                hidden_start = 20
                ws_dash.cell(row=hidden_start, column=17, value="Статус")
                ws_dash.cell(row=hidden_start, column=18, value="Кол-во")
                for i, (st_name, cnt) in enumerate(status_counts.items(), 1):
                    ws_dash.cell(row=hidden_start + i, column=17, value=st_name)
                    ws_dash.cell(row=hidden_start + i, column=18, value=int(cnt))

                data_ref = Reference(ws_dash, min_col=18, min_row=hidden_start, max_row=hidden_start + len(status_counts))
                labels_ref = Reference(ws_dash, min_col=17, min_row=hidden_start + 1, max_row=hidden_start + len(status_counts))
                pc.add_data(data_ref, titles_from_data=True)
                pc.set_categories(labels_ref)
                dl = DataLabelList()
                dl.showPercent = True
                dl.showVal = False
                pc.dataLabels = dl
                ws_dash.add_chart(pc, "H11")
            except Exception as chart_err:
                logger.warning(f"PieChart: {chart_err}")

        # ② ДЕТАЛЬНЫЙ РАСЧЁТ
        ws_det = wb.create_sheet("📋 Детальный расчёт")
        ws_det.sheet_view.showGridLines = False
        ws_det.freeze_panes = "C2"

        COLUMN_GROUPS: List[Tuple[str, List[str], str]] = [
            ("🏷️ Товар", ["artikul", "category"], cls.C["navy"]),
            ("💰 Цены", ["selling_price", "cogs"], cls.C["blue"]),
            ("📦 Затраты МП", ["commission", "delivery_to_customer", "middle_mile_cost", "sorting_cost", "acquiring_cost"], "366092"),
            ("🔄 Возвраты", ["return_cost"], "7F6000"),
            ("🏭 Опер. затраты", ["pick_pack_cost", "packaging_cost", "first_mile_cost", "marketing_budget_per_unit", "warehouse_cost"], "375623"),
            ("💸 Налоги", ["tax_cost", "total_expenses"], "595959"),
            ("📈 Финрезультат", ["gross_profit", "margin_percent", "roi_percent", "break_even_units"], cls.C["green"]),
            ("🎯 Рекомендации", ["rec_price_min", "rec_price_15", "rec_price_25"], "833C00"),
            ("📊 ABC/XYZ", ["daily_sales", "abc_category", "xyz_category", "abc_xyz", "profitability_status"], "1F4E79"),
            ("🔍 Аудит", ["source", "scheme"], cls.C["mid_gray"]),
        ]

        COL_DISPLAY_NAMES: Dict[str, str] = {
            "artikul": "Артикул", "category": "Категория", "selling_price": "Цена продажи", "cogs": "Себестоимость",
            "commission": "Комиссия МП", "delivery_to_customer": "Доставка покупателю", "middle_mile_cost": "Магистраль",
            "sorting_cost": "Сортировка", "acquiring_cost": "Эквайринг", "return_cost": "Затраты на возврат",
            "pick_pack_cost": "Пик/Пак", "packaging_cost": "Упаковка", "first_mile_cost": "Первая миля",
            "marketing_budget_per_unit": "Маркетинг / ед.", "warehouse_cost": "Хранение", "tax_cost": "Налоги",
            "total_expenses": "Итого затраты", "gross_profit": "Прибыль", "margin_percent": "Маржа, %",
            "roi_percent": "ROI, %", "break_even_units": "Точка безубыт., шт.", "rec_price_min": "Цена 0% маржа",
            "rec_price_15": "Цена 15% маржа", "rec_price_25": "Цена 25% маржа", "daily_sales": "Продажи / день",
            "abc_category": "ABC", "xyz_category": "XYZ", "abc_xyz": "ABC·XYZ", "profitability_status": "Статус",
            "source": "Источник тарифа", "scheme": "Схема"
        }

        COL_FORMATS: Dict[str, str] = {
            "selling_price": cls.FMT_MONEY, "cogs": cls.FMT_MONEY, "commission": cls.FMT_MONEY,
            "delivery_to_customer": cls.FMT_MONEY, "middle_mile_cost": cls.FMT_MONEY, "sorting_cost": cls.FMT_MONEY,
            "acquiring_cost": cls.FMT_MONEY, "return_cost": cls.FMT_MONEY, "pick_pack_cost": cls.FMT_MONEY,
            "packaging_cost": cls.FMT_MONEY, "first_mile_cost": cls.FMT_MONEY, "marketing_budget_per_unit": cls.FMT_MONEY,
            "warehouse_cost": cls.FMT_MONEY, "tax_cost": cls.FMT_MONEY, "total_expenses": cls.FMT_MONEY,
            "gross_profit": cls.FMT_MONEY, "margin_percent": "0.00", "roi_percent": "0.00",
            "break_even_units": "#,##0.0", "rec_price_min": cls.FMT_MONEY, "rec_price_15": cls.FMT_MONEY,
            "rec_price_25": cls.FMT_MONEY, "daily_sales": "#,##0",
        }

        ordered_cols: List[str] = []
        for _, cols, _ in COLUMN_GROUPS:
            for c in cols:
                if c in df.columns and c not in ordered_cols:
                    ordered_cols.append(c)
        for c in df.columns:
            if c not in ordered_cols:
                ordered_cols.append(c)

        df_export = df[ordered_cols].copy()

        group_row = 1
        current_col = 1
        for group_name, group_cols, group_color in COLUMN_GROUPS:
            valid_cols = [c for c in group_cols if c in df_export.columns]
            if not valid_cols:
                continue
            span = len(valid_cols)
            end_col = current_col + span - 1

            if span > 1:
                ws_det.merge_cells(start_row=group_row, start_column=current_col, end_row=group_row, end_column=end_col)
            cell = ws_det.cell(row=group_row, column=current_col, value=group_name)
            cell.fill = cls._fill(group_color)
            cell.font = cls._font(bold=True, color="FFFFFF", size=9)
            cell.alignment = cls._align(h="center")
            cell.border = thick_border_bottom()
            current_col = end_col + 1

        ws_det.row_dimensions[group_row].height = 22

        header_row = 2
        for ci, col in enumerate(df_export.columns, 1):
            cell = ws_det.cell(row=header_row, column=ci, value=COL_DISPLAY_NAMES.get(col, col))
            cell.fill = cls._fill(cls.C["navy"])
            cell.font = cls._font(bold=True, color="FFFFFF", size=9)
            cell.alignment = cls._align(h="center", wrap=True)
            cell.border = thin_border()
        ws_det.row_dimensions[header_row].height = 32

        STATUS_BG_MAP = {
            "Высокомаржинальный": cls.C["profit_hi_bg"], "Низкомаржинальный": cls.C["profit_lo_bg"], "Убыточный": cls.C["loss_bg"],
        }
        STATUS_FG_MAP = {
            "Высокомаржинальный": cls.C["profit_hi"], "Низкомаржинальный": cls.C["profit_lo"], "Убыточный": cls.C["loss"],
        }
        ABC_BG_MAP = {
            "A": cls.C["abc_a"], "B": cls.C["abc_b"], "C": cls.C["abc_c"],
            "AX": cls.C["profit_hi_bg"], "AY": cls.C["lime"], "AZ": cls.C["yellow"],
            "BX": cls.C["sky"], "BY": cls.C["sky"], "BZ": cls.C["yellow"],
            "CX": cls.C["gray"], "CY": cls.C["orange"], "CZ": cls.C["loss_bg"],
        }

        data_start_row = 3
        n_rows = len(df_export)
        n_cols = len(df_export.columns)

        margin_col_idx = col_name_list.index("margin_percent") + 1 if "margin_percent" in (col_name_list := list(df_export.columns)) else None
        profit_col_idx = col_name_list.index("gross_profit") + 1 if "gross_profit" in col_name_list else None

        for ri, (_, row_data) in enumerate(df_export.iterrows()):
            r_idx = data_start_row + ri
            zebra = cls.C["gray"] if ri % 2 == 0 else cls.C["white"]
            status = str(row_data.get("profitability_status", ""))
            row_bg = STATUS_BG_MAP.get(status, zebra)

            for ci, col in enumerate(col_name_list, 1):
                val = row_data[col]
                cell = ws_det.cell(row=r_idx, column=ci, value=val)
                cell.border = thin_border()
                cell.alignment = cls._align(h="right" if col in COL_FORMATS else "left", v="center")

                if col in COL_FORMATS:
                    cell.fill = cls._fill(row_bg)
                    cell.number_format = COL_FORMATS[col]
                    if col == "gross_profit" and isinstance(val, (int, float)) and val < 0:
                        cell.font = cls._font(bold=True, color=cls.C["crimson"])
                    else:
                        cell.font = cls._font(size=10)
                elif col == "profitability_status":
                    cell.fill = cls._fill(STATUS_BG_MAP.get(status, cls.C["white"]))
                    cell.font = Font(name="Calibri", bold=True, size=9, color=STATUS_FG_MAP.get(status, "000000"))
                    cell.alignment = cls._align(h="center")
                elif col in ("abc_category", "abc_xyz"):
                    cell.fill = cls._fill(ABC_BG_MAP.get(str(val), cls.C["white"]))
                    cell.font = cls._font(bold=True, color="FFFFFF" if str(val) in ("A", "B") else "000000", size=10)
                    cell.alignment = cls._align(h="center")
                else:
                    cell.fill = cls._fill(zebra)
                    cell.font = cls._font(size=10)

        if margin_col_idx:
            m_col_letter = get_column_letter(margin_col_idx)
            m_range = f"{m_col_letter}{data_start_row}:{m_col_letter}{data_start_row + n_rows - 1}"
            try:
                db_rule = DataBarRule(start_type="num", start_value=-50, end_type="num", end_value=50, color=cls.C["blue"])
                ws_det.conditional_formatting.add(m_range, db_rule)
            except Exception:
                pass

        if profit_col_idx:
            p_col_letter = get_column_letter(profit_col_idx)
            p_range = f"{p_col_letter}{data_start_row}:{p_col_letter}{data_start_row + n_rows - 1}"
            try:
                cs_rule = ColorScaleRule(start_type="min", start_color="FFC7CE", mid_type="num", mid_value=0, mid_color="FFEB9C", end_type="max", end_color="C6EFCE")
                ws_det.conditional_formatting.add(p_range, cs_rule)
            except Exception:
                pass

        MIN_WIDTH, MAX_WIDTH = 8, 35
        for ci, col in enumerate(col_name_list, 1):
            col_letter = get_column_letter(ci)
            sample_vals = df_export[col].astype(str).str.len()
            best_len = max(len(COL_DISPLAY_NAMES.get(col, col)), sample_vals.quantile(0.95) if not sample_vals.empty else 0)
            ws_det.column_dimensions[col_letter].width = min(max(float(best_len) + 2, MIN_WIDTH), MAX_WIDTH)

        table_ref = f"A{header_row}:{get_column_letter(n_cols)}{data_start_row + n_rows - 1}"
        try:
            tab = Table(displayName="UnitEconomics", ref=table_ref)
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False)
            ws_det.add_table(tab)
        except Exception as te:
            logger.warning(f"Table add error: {te}")
        ws_det.row_dimensions[1].height = 6

        # ③ ЛИСТ РЕКОМЕНДАЦИЙ
        ws_rec = wb.create_sheet("💡 Рекомендации")
        ws_rec.sheet_view.showGridLines = False
        ws_rec.merge_cells("A1:G1")
        h = ws_rec["A1"]
        h.value = "💡 ЦЕНОВЫЕ РЕКОМЕНДАЦИИ И СТАТУС SKU"
        h.font = cls._font(bold=True, color="FFFFFF", size=14)
        h.fill = cls._fill(cls.C["navy"])
        h.alignment = cls._align(h="center")
        ws_rec.row_dimensions[1].height = 32

        rec_cols = ["artikul", "category", "selling_price", "cogs", "gross_profit", "margin_percent", "rec_price_min", "rec_price_15", "rec_price_25", "profitability_status"]
        rec_cols = [c for c in rec_cols if c in df.columns]
        df_rec = df[rec_cols].copy()

        REC_HEADER_NAMES = {
            "artikul": "Артикул", "category": "Категория", "selling_price": "Цена продажи", "cogs": "Себестоимость",
            "gross_profit": "Прибыль / ед.", "margin_percent": "Маржа, %", "rec_price_min": "▶ Цена 0%",
            "rec_price_15": "▶ Цена 15%", "rec_price_25": "▶ Цена 25%", "profitability_status": "Статус",
        }
        REC_FMTS = {
            "selling_price": cls.FMT_MONEY, "cogs": cls.FMT_MONEY, "gross_profit": cls.FMT_MONEY,
            "margin_percent": "0.00", "rec_price_min": cls.FMT_MONEY, "rec_price_15": cls.FMT_MONEY, "rec_price_25": cls.FMT_MONEY,
        }

        hdr_row = 2
        for ci, col in enumerate(rec_cols, 1):
            cell = ws_rec.cell(row=hdr_row, column=ci, value=REC_HEADER_NAMES.get(col, col))
            cell.fill = cls._fill(cls.C["blue"])
            cell.font = cls._font(bold=True, color="FFFFFF", size=9)
            cell.alignment = cls._align(h="center", wrap=True)
            cell.border = thin_border()
        ws_rec.row_dimensions[hdr_row].height = 30

        for ri, (_, row_data) in enumerate(df_rec.iterrows()):
            r_idx = hdr_row + 1 + ri
            status = str(row_data.get("profitability_status", ""))
            zebra = cls.C["gray"] if ri % 2 == 0 else cls.C["white"]

            for ci, col in enumerate(rec_cols, 1):
                val = row_data[col]
                cell = ws_rec.cell(row=r_idx, column=ci, value=val)
                cell.border = thin_border()
                cell.alignment = cls._align(h="right" if col in REC_FMTS else "center" if col == "profitability_status" else "left")
                if col in REC_FMTS:
                    cell.number_format = REC_FMTS[col]
                if col == "profitability_status":
                    cell.fill = cls._fill(STATUS_BG_MAP.get(status, cls.C["white"]))
                    cell.font = Font(name="Calibri", bold=True, size=9, color=STATUS_FG_MAP.get(status, "000000"))
                elif col in ("rec_price_min", "rec_price_15", "rec_price_25"):
                    cell.fill = cls._fill(cls.C["sky"])
                    cell.font = cls._font(bold=True, color=cls.C["navy"])
                else:
                    cell.fill = cls._fill(zebra)
                    cell.font = cls._font(size=10)

        for ci, col in enumerate(rec_cols, 1):
            cl = get_column_letter(ci)
            ws_rec.column_dimensions[cl].width = min(max(len(REC_HEADER_NAMES.get(col, col)) + 2, df_rec[col].astype(str).str.len().quantile(0.9) + 2 if not df_rec.empty else 10), 30)
        ws_rec.freeze_panes = "C3"

        # ④ ABC/XYZ МАТРИЦА
        ws_abc = wb.create_sheet("🔢 ABC·XYZ матрица")
        ws_abc.sheet_view.showGridLines = False
        ws_abc.merge_cells("A1:D1")
        t = ws_abc["A1"]
        t.value = "🔢 ABC·XYZ МАТРИЦА ПОРТФЕЛЯ"
        t.font = cls._font(bold=True, color="FFFFFF", size=14)
        t.fill = cls._fill(cls.C["navy"])
        t.alignment = cls._align(h="center")
        ws_abc.row_dimensions[1].height = 32

        if "abc_xyz" in df.columns:
            matrix_data = (
                df.groupby("abc_xyz", observed=True)
                  .agg(SKU=("artikul", "count"), Прибыль=("gross_profit", "sum") if "gross_profit" in df.columns else ("artikul", "count"), Маржа=("margin_percent", "mean") if "margin_percent" in df.columns else ("artikul", "count"))
                  .reset_index().sort_values("abc_xyz")
            )
            mat_hdrs = ["Сегмент", "SKU, шт.", "Сум. прибыль, ₽", "Ср. маржа, %"]
            for ci, hdr in enumerate(mat_hdrs, 1):
                c = ws_abc.cell(row=2, column=ci, value=hdr)
                c.fill = cls._fill(cls.C["blue"])
                c.font = cls._font(bold=True, color="FFFFFF", size=10)
                c.alignment = cls._align(h="center")
                c.border = thin_border()
            ws_abc.row_dimensions[2].height = 26

            for ri, row_data in enumerate(matrix_data.itertuples(), 3):
                seg = row_data.abc_xyz
                seg_bg = ABC_BG_MAP.get(seg, cls.C["white"])
                cells_data = [
                    (seg, None, "center"),
                    (row_data.SKU, "#,##0", "center"),
                    (row_data.Прибыль, cls.FMT_MONEY, "right"),
                    (row_data.Маржа / 100 if row_data.Маржа else 0, "0.0%", "right"),
                ]
                for ci, (val, fmt2, h_align) in enumerate(cells_data, 1):
                    c = ws_abc.cell(row=ri, column=ci, value=val)
                    c.fill = cls._fill(seg_bg)
                    c.font = cls._font(bold=(ci == 1), size=10, color=cls.C["navy"] if ci == 1 else "000000")
                    c.alignment = cls._align(h=h_align)
                    c.border = thin_border()
                    if fmt2:
                        c.number_format = fmt2
            for col_letter, width in zip("ABCD", [14, 12, 22, 16]):
                ws_abc.column_dimensions[col_letter].width = width

        # ⑤ ПАРАМЕТРЫ РАСЧЁТА
        ws_par = wb.create_sheet("⚙️ Параметры")
        ws_par.sheet_view.showGridLines = False
        ws_par.merge_cells("A1:C1")
        p_title = ws_par["A1"]
        p_title.value = "⚙️ ПАРАМЕТРЫ РАСЧЁТА"
        p_title.font = cls._font(bold=True, color="FFFFFF", size=13)
        p_title.fill = cls._fill(cls.C["navy"])
        p_title.alignment = cls._align(h="center")
        ws_par.row_dimensions[1].height = 28

        params_list = [
            ("Версия приложения", APP_VERSION),
            ("Дата формирования", datetime.now().strftime("%d.%m.%Y %H:%M:%S")),
            ("Система налогообложения", tax_label),
            ("Схема работы", scheme_label),
            ("Частота выплат", payment_frequency),
            ("Всего SKU", len(df)),
            ("Прибыльных SKU", int((df["gross_profit"] > 0).sum()) if "gross_profit" in df.columns else "—"),
            ("Убыточных SKU", int((df["gross_profit"] < 0).sum()) if "gross_profit" in df.columns else "—"),
            ("Средняя маржа", f"{df['margin_percent'].mean():.2f}%" if "margin_percent" in df.columns else "—"),
        ]
        for ri, (param, val) in enumerate(params_list, 2):
            bg = cls.C["sky"] if ri % 2 == 0 else cls.C["white"]
            p_cell = ws_par.cell(row=ri, column=1, value=param)
            p_cell.font = cls._font(bold=True, size=10)
            p_cell.fill = cls._fill(bg)
            p_cell.alignment = cls._align(h="left")
            p_cell.border = thin_border()
            v_cell = ws_par.cell(row=ri, column=2, value=val)
            v_cell.font = cls._font(size=10)
            v_cell.fill = cls._fill(bg)
            v_cell.alignment = cls._align(h="left")
            v_cell.border = thin_border()
        ws_par.column_dimensions["A"].width = 30
        ws_par.column_dimensions["B"].width = 40

        # ⑥ ДОПОЛНИТЕЛЬНЫЕ ЛИСТЫ
        if include_advanced:
            cls.add_sensitivity_analysis(wb, df)
            cls.add_forecast(wb, df)

        sheet_order = ["📊 Дашборд", "📋 Детальный расчёт", "💡 Рекомендации", "🔢 ABC·XYZ матрица", "⚙️ Параметры"]
        if include_advanced:
            sheet_order.extend(["🎯 Анализ чувствительности", "📈 Прогноз"])

        for i, sheet_name in enumerate(sheet_order):
            if sheet_name in wb.sheetnames:
                wb.move_sheet(sheet_name, offset=-wb.index(wb[sheet_name]) + i)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    @classmethod
    def add_sensitivity_analysis(cls, wb: Workbook, df: pd.DataFrame):
        """
        ИСПРАВЛЕННАЯ ВЕРСИЯ: Пересчитывает P&L математически корректно, 
        а не линейно масштабирует проценты, что давало ложные результаты.
        """
        ws = wb.create_sheet("🎯 Анализ чувствительности")
        ws.sheet_view.showGridLines = False
        
        ws.merge_cells("A1:E1")
        title = ws["A1"]
        title.value = "🎯 АНАЛИЗ ЧУВСТВИТЕЛЬНОСТИ (реальный пересчет P&L)"
        title.font = cls._font(bold=True, color="FFFFFF", size=14)
        title.fill = cls._fill(cls.C["navy"])
        title.alignment = cls._align(h="center")
        ws.row_dimensions[1].height = 32
        
        headers = ["Изменение цены, %", "Изменение комиссии, п.п.", "Средняя маржа, %", "Кол-во прибыльных", "Кол-во убыточных"]
        for ci, hdr in enumerate(headers, 1):
            cell = ws.cell(row=2, column=ci, value=hdr)
            cell.fill = cls._fill(cls.C["blue"])
            cell.font = cls._font(bold=True, color="FFFFFF", size=10)
            cell.alignment = cls._align(h="center")
            cell.border = thin_border()
        ws.row_dimensions[2].height = 26
        
        price_changes = [-20, -10, 0, 10, 20]
        commission_changes = [-5, -2, 0, 2, 5]
        
        # Базовые постоянные издержки для пересчета
        base_fixed = (df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + 
                      df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost'])
        acq_rate = df['acquiring_transfer_rate'].values if 'acquiring_transfer_rate' in df.columns else np.full(len(df), 0.016)
        acq_sku = df['acquiring_sku_cost'].values if 'acquiring_sku_cost' in df.columns else np.full(len(df), 0.12)
        
        row = 3
        for pc in price_changes:
            for cc in commission_changes:
                # 1. Новая цена
                new_price = df['selling_price'].values * (1 + pc / 100)
                
                # 2. Пересчет комиссии (зависит от цены) + дельта комиссии в п.п.
                new_comm_rate = np.maximum(0, df['commission_rate'].values + (cc / 100))
                new_commission = np.maximum(new_price * new_comm_rate, df['min_commission'].values)
                
                # 3. Пересчет эквайринга (зависит от цены)
                new_acquiring = acq_sku + (new_price * acq_rate)
                
                # 4. Новые полные затраты
                new_total_expenses = (base_fixed.values + new_commission + df['delivery_to_customer'].values + 
                                      df['middle_mile_cost'].values + df['sorting_cost'].values + 
                                      new_acquiring + df['return_cost'].values)
                
                # 5. Новая прибыль и маржа
                new_gross_profit = new_price - new_total_expenses
                new_margin = np.where(new_price > 0, (new_gross_profit / new_price) * 100, 0.0)
                
                avg_margin = float(np.mean(new_margin))
                profit_cnt = int(np.sum(new_gross_profit > 0))
                loss_cnt = int(np.sum(new_gross_profit <= 0))
                
                ws.cell(row=row, column=1, value=pc)
                ws.cell(row=row, column=2, value=cc)
                ws.cell(row=row, column=3, value=avg_margin / 100)
                ws.cell(row=row, column=4, value=profit_cnt)
                ws.cell(row=row, column=5, value=loss_cnt)
                
                ws.cell(row=row, column=3).number_format = "0.0%"
                ws.cell(row=row, column=4).number_format = "#,##0"
                ws.cell(row=row, column=5).number_format = "#,##0"
                
                if avg_margin > 20:
                    bg = cls.C["profit_hi_bg"]
                elif avg_margin > 0:
                    bg = cls.C["profit_lo_bg"]
                else:
                    bg = cls.C["loss_bg"]
                for c in range(1, 6):
                    ws.cell(row=row, column=c).fill = cls._fill(bg)
                    ws.cell(row=row, column=c).border = thin_border()
                
                row += 1
        
        for col_letter, width in zip("ABCDE", [18, 18, 18, 18, 18]):
            ws.column_dimensions[col_letter].width = width
        ws.freeze_panes = "A3"
    
    @classmethod
    def add_forecast(cls, wb: Workbook, df: pd.DataFrame):
        ws = wb.create_sheet("📈 Прогноз")
        ws.sheet_view.showGridLines = False
        
        ws.merge_cells("A1:M1")
        title = ws["A1"]
        title.value = "📈 ПРОГНОЗ ПРИБЫЛИ НА 12 МЕСЯЦЕВ"
        title.font = cls._font(bold=True, color="FFFFFF", size=14)
        title.fill = cls._fill(cls.C["navy"])
        title.alignment = cls._align(h="center")
        ws.row_dimensions[1].height = 32
        
        months = [f"Месяц {i}" for i in range(1, 13)]
        for ci, m in enumerate(months, 2):
            cell = ws.cell(row=2, column=ci, value=m)
            cell.fill = cls._fill(cls.C["blue"])
            cell.font = cls._font(bold=True, color="FFFFFF", size=10)
            cell.alignment = cls._align(h="center")
            cell.border = thin_border()
        ws.row_dimensions[2].height = 26
        
        indicators = [
            ("Прогноз роста продаж, %", [1.0, 1.05, 1.10, 1.08, 1.12, 1.15, 1.10, 1.05, 1.08, 1.12, 1.15, 1.20]),
            ("Средняя маржа, %", None),
            ("Прогноз прибыли, ₽", None)
        ]
        
        base_profit = float(df['gross_profit'].mean()) if 'gross_profit' in df.columns else 0.0
        base_margin = float(df['margin_percent'].mean()) if 'margin_percent' in df.columns else 0.0
        
        row = 3
        for ind, values in indicators:
            cell = ws.cell(row=row, column=1, value=ind)
            cell.font = cls._font(bold=True, size=10)
            cell.alignment = cls._align(h="left")
            cell.border = thin_border()
            
            if values is not None:
                for ci, val in enumerate(values, 2):
                    cell = ws.cell(row=row, column=ci, value=val)
                    cell.number_format = "0.0%"
                    cell.fill = cls._fill(cls.C["gray"] if ci % 2 == 0 else cls.C["white"])
                    cell.border = thin_border()
                    cell.alignment = cls._align(h="center")
            elif ind == "Средняя маржа, %":
                for ci in range(2, 14):
                    val = (base_margin / 100) * (1 + 0.01 * (ci-1))
                    cell = ws.cell(row=row, column=ci, value=val)
                    cell.number_format = "0.0%"
                    cell.fill = cls._fill(cls.C["gray"] if ci % 2 == 0 else cls.C["white"])
                    cell.border = thin_border()
                    cell.alignment = cls._align(h="center")
            else:
                growth_row = 3
                for ci in range(2, 14):
                    growth = ws.cell(row=growth_row, column=ci).value or 1.0
                    margin_factor = 1 + 0.01 * (ci-1)
                    val = base_profit * growth * margin_factor
                    cell = ws.cell(row=row, column=ci, value=val)
                    cell.number_format = cls.FMT_MONEY
                    cell.fill = cls._fill(cls.C["gray"] if ci % 2 == 0 else cls.C["white"])
                    cell.border = thin_border()
                    cell.alignment = cls._align(h="center")
            
            ws.row_dimensions[row].height = 22
            row += 1
        
        try:
            chart = LineChart()
            chart.title = "Прогноз прибыли"
            chart.style = 12
            chart.width = 20
            chart.height = 12
            
            data = Reference(ws, min_col=2, min_row=row-1, max_row=row-1, max_col=13)
            chart.add_data(data, titles_from_data=True)
            cats = Reference(ws, min_col=2, min_row=2, max_row=2, max_col=13)
            chart.set_categories(cats)
            ws.add_chart(chart, "A15")
        except Exception as e:
            logger.warning(f"Ошибка создания диаграммы прогноза: {e}")
        
        for col in "ABCDEFGHIJKLM":
            ws.column_dimensions[col].width = 14

# ============================================================================
# НОРМАЛИЗАТОР ДАННЫХ (БЛОК 7)
# ============================================================================
class UniversalDataNormalizer:
    COLUMN_MAPPING: Final[Dict[str, List[str]]] = {
        'artikul': ['artikul', 'артикул', 'sku', 'offer_id', 'id', 'код'],
        'category': ['category', 'категория', 'группа', 'предмет', 'тип'],
        'selling_price': ['selling_price', 'цена продажи', 'цена', 'price', 'стоимость'],
        'cogs': ['cogs', 'себестоимость', 'закупка', 'cost', 'закупочная'],
        'daily_sales': ['daily_sales', 'заказы, шт.', 'продажи, шт.', 'quantity', 'продажи'],
        'weight_kg': ['weight_kg', 'вес', 'weight', 'вес кг', 'вес, кг'],
        'length_cm': ['length_cm', 'длина', 'length', 'длина, см'],
        'width_cm': ['width_cm', 'ширина', 'width', 'ширина, см'],
        'height_cm': ['height_cm', 'высота', 'height', 'высота, см'],
        'volume_liters': ['volume_liters', 'объем', 'volume', 'объем, л'],
        'packaging_cost': ['packaging_cost', 'упаковка', 'стоимость упаковки'],
        'first_mile_cost': ['first_mile_cost', 'магистраль', 'первая миля'],
        'marketing_budget_per_unit': ['marketing_budget_per_unit', 'реклама', 'дрр', 'маркетинг'],
        'stock_depth_days': ['stock_depth_days', 'глубина запаса', 'дни запаса'],
        'quantity_per_order': ['quantity_per_order', 'количество в заказе', 'шт в заказе'],
        'warehouse_cost': ['warehouse_cost', 'стоимость хранения', 'хранение'],
    }
    NUMERIC_COLS: Final[Tuple[str, ...]] = (
        'selling_price', 'cogs', 'daily_sales', 'weight_kg',
        'length_cm', 'width_cm', 'height_cm', 'volume_liters',
        'packaging_cost', 'first_mile_cost', 'marketing_budget_per_unit',
        'stock_depth_days', 'quantity_per_order', 'warehouse_cost'
    )
    
    @classmethod
    def normalize_dataframe(cls, raw_df: pd.DataFrame) -> pd.DataFrame:
        if raw_df.empty:
            return raw_df
        df = raw_df.copy()
        df.columns = [str(col).strip().lower() for col in df.columns]
        final_data = {}
        for target_col, synonyms in cls.COLUMN_MAPPING.items():
            found = False
            for syn in synonyms:
                if syn in df.columns:
                    final_data[target_col] = df[syn]
                    found = True
                    break
            if not found:
                if target_col == 'artikul':
                    final_data[target_col] = [f"SKU_{i+1}" for i in range(len(df))]
                elif target_col == 'category':
                    final_data[target_col] = "не указано"
                elif target_col == 'quantity_per_order':
                    final_data[target_col] = 1.0
                else:
                    final_data[target_col] = 0.0
        norm_df = pd.DataFrame(final_data)
        for col in cls.NUMERIC_COLS:
            if col in norm_df.columns:
                norm_df[col] = pd.to_numeric(
                    norm_df[col].astype(str).str.replace(r'[\s,;%₽]', '', regex=True),
                    errors='coerce'
                ).fillna(0.0).abs()
        
        # ФИКС: Корректная обработка NaN в артикулах до конвертации в строку
        norm_df['artikul'] = norm_df['artikul'].fillna('UNKNOWN').astype(str).str.strip()
        norm_df['artikul'] = norm_df['artikul'].replace('nan', 'UNKNOWN')
        norm_df['category'] = norm_df['category'].fillna('не указано').astype(str).str.strip().str.lower()
        
        return norm_df.drop_duplicates(subset=['artikul'], keep='first')
    
    @classmethod
    def load_file(cls, file_buffer: io.BytesIO, file_name: str) -> pd.DataFrame:
        try:
            if file_name.endswith('.csv'):
                return pd.read_csv(file_buffer, sep=None, engine='python', encoding='utf-8', encoding_errors='replace')
            elif file_name.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file_buffer)
            else:
                raise ValueError("Неподдерживаемый формат файла")
        except UnicodeDecodeError:
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, sep=None, engine='python', encoding='cp1251', encoding_errors='replace')

# ============================================================================
# ИНТЕГРАЦИЯ С ЯНДЕКС.ДИРЕКТ (заготовка)
# ============================================================================
class YandexDirectIntegration:
    def __init__(self, api_token: str):
        self.api_token = api_token
        self.base_url = "https://api.direct.yandex.com/v5"
    
    def get_campaign_stats(self, date_from: str, date_to: str) -> pd.DataFrame:
        logger.warning("Метод get_campaign_stats не реализован")
        return pd.DataFrame()
    
    def get_drr_by_category(self, category: str) -> float:
        logger.warning("Метод get_drr_by_category не реализован")
        return 0.0

# ============================================================================
# ЭКСПОРТ В GOOGLE SHEETS (ИСПРАВЛЕННАЯ ВЕРСИЯ)
# ============================================================================
class GoogleSheetsExporter:
    def __init__(self, credentials_info: dict):
        """Инициализация через словарь данных, а не путь к файлу, для безопасности."""
        self.credentials_info = credentials_info
        self.service = None
        try:
            from googleapiclient.discovery import build
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_info(
                self.credentials_info,
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            self.service = build("sheets", "v4", credentials=creds)
            logger.info("Google Sheets сервис инициализирован из памяти")
        except Exception as e:
            logger.error(f"Ошибка инициализации Google Sheets: {e}")
    
    def export(self, df: pd.DataFrame, spreadsheet_id: str, sheet_name: str = "Unit Economics") -> None:
        if self.service is None:
            logger.error("Сервис Google Sheets не доступен")
            return
        try:
            body = {
                'values': [df.columns.tolist()] + df.values.tolist()
            }
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1",
                valueInputOption="RAW",
                body=body
            ).execute()
            logger.info(f"Данные экспортированы в Google Sheets: {spreadsheet_id}")
        except Exception as e:
            logger.error(f"Ошибка экспорта в Google Sheets: {e}")

# ============================================================================
# ОРКЕСТРАТОР ПАЙПЛАЙНА (БЛОК 8)
# ============================================================================
import contextlib

@dataclass
class DataPipeline:
    tax_label: str = "УСН 6% (доходы)"
    scheme_label: str = "FBS (склад продавца)"
    payment_frequency: str = "Еженедельно, 4 нед. (1.6%)"
    use_api: bool = True
    parallel: bool = False
    chunk_size: Optional[int] = None
    
    def process(
        self,
        raw_df: pd.DataFrame,
        tariff_manager: HybridTariffManager,
        ym_api: Optional[YandexMarketAPI] = None,
        perf_monitor: Optional[PerformanceMonitor] = None
    ) -> pd.DataFrame:
        with (perf_monitor.measure('load_time') if perf_monitor else contextlib.nullcontext()):
            norm_df = UniversalDataNormalizer.normalize_dataframe(raw_df)
        validated_df, errors = DataValidator.validate(norm_df)
        if errors:
            logger.warning(f"Ошибки валидации: {errors}")
        scheme = self.scheme_label.split(" ")[0]
        with (perf_monitor.measure('api_time') if perf_monitor else contextlib.nullcontext()):
            tariff_df = tariff_manager.get_tariffs_vectorized(
                validated_df,
                scheme=scheme,
                ym_api=ym_api,
                use_api=self.use_api
            )
            tariffs_map = tariff_df.set_index('category').to_dict(orient='index')
        current_hash = StringUtils.make_hash(validated_df)
        def calc_func(df_chunk):
            return run_calculations_cached(
                df_hash=current_hash,
                df=df_chunk,
                tax_label=self.tax_label,
                scheme_label=self.scheme_label,
                payment_frequency=self.payment_frequency,
                tariffs_map=tariffs_map
            )
        with (perf_monitor.measure('calc_time') if perf_monitor else contextlib.nullcontext()):
            # ФИКС: Порог многопоточности поднят до 50 000 строк для избежания оверхеда
            if self.parallel and CONCURRENT_AVAILABLE and len(validated_df) > 50000:
                calc_df = ParallelProcessor.process_in_parallel(
                    validated_df,
                    calc_func,
                    chunk_size=self.chunk_size
                )
            else:
                calc_df = calc_func(validated_df)
        return calc_df

# ============================================================================
# STREAMLIT UI (БЛОК 9)
# ============================================================================
def init_session_state():
    defaults = {
        'main_df': pd.DataFrame(),
        'calc_df': pd.DataFrame(),
        'tariffs': {},
        'last_hash': '',
        'api_key': '',
        'business_id': '',
        'perf_monitor': PerformanceMonitor()
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def render_sidebar() -> DataPipeline:
    with st.sidebar:
        st.header("⚙️ Настройки")
        api_key = st.text_input(
            "API Key Яндекс Маркета",
            value=st.session_state.api_key,
            type="password"
        )
        business_id = st.text_input(
            "Business ID",
            value=st.session_state.business_id
        )
        st.session_state.api_key = api_key
        st.session_state.business_id = business_id
        
        tax_options = [tax.value.label for tax in TaxSystem]
        tax_label = st.selectbox(
            "Система налогообложения",
            tax_options,
            index=0
        )
        scheme_options = [
            "FBS (склад продавца)",
            "FBY (склад Маркета)",
            "Экспресс",
            "DBS (доставка продавца)"
        ]
        scheme_label = st.selectbox(
            "Схема работы",
            scheme_options,
            index=0
        )
        payment_options = [
            "Ежемесячно (1.0%)",
            "Раз в 2 недели (1.3%)",
            "Еженедельно, 4 нед. (1.6%)",
            "Ежедневно (3.3%)"
        ]
        payment_frequency = st.selectbox(
            "Частота выплат",
            payment_options,
            index=2
        )
        use_api = st.checkbox(
            "Использовать API ЯМ для тарифов",
            value=True
        )
        st.subheader("⚡ Производительность")
        parallel = st.checkbox("Параллельная обработка (только для >50k строк)", value=False)
        chunk_size = st.number_input("Размер чанка (строк)", min_value=100, max_value=100000, value=5000, step=100) if parallel else None
        if st.button("Показать отчёт по производительности"):
            st.session_state.perf_monitor.display_report()
    return DataPipeline(
        tax_label=tax_label,
        scheme_label=scheme_label,
        payment_frequency=payment_frequency,
        use_api=use_api,
        parallel=parallel,
        chunk_size=chunk_size
    )

def render_upload_section() -> Tuple[Optional[pd.DataFrame], HybridTariffManager]:
    col1, col2 = st.columns(2)
    tariff_manager = HybridTariffManager()
    main_df = None
    with col1:
        st.subheader("1. Загрузка данных товаров")
        uploaded_file = st.file_uploader(
            "Загрузите Excel или CSV",
            type=['xlsx', 'xls', 'csv']
        )
        if uploaded_file is not None:
            try:
                raw_df = UniversalDataNormalizer.load_file(
                    io.BytesIO(uploaded_file.read()),
                    uploaded_file.name
                )
                norm_df = UniversalDataNormalizer.normalize_dataframe(raw_df)
                validated_df, errors = DataValidator.validate(norm_df)
                if errors:
                    for err in errors:
                        st.warning(err)
                main_df = validated_df
                st.session_state.main_df = validated_df
                st.success(f"Загружено {len(validated_df)} уникальных SKU")
            except Exception as e:
                st.error(f"Ошибка чтения файла: {e}")
    with col2:
        st.subheader("2. Загрузка справочника тарифов (опционально)")
        tariff_file = st.file_uploader(
            "Тарифы (Excel/CSV)",
            type=['xlsx', 'xls', 'csv'],
            key="tariff_uploader"
        )
        if tariff_file is not None:
            try:
                t_raw = UniversalDataNormalizer.load_file(
                    io.BytesIO(tariff_file.read()),
                    tariff_file.name
                )
                t_raw.columns = [str(c).strip().lower() for c in t_raw.columns]
                tariff_manager.load_tariffs_from_file(t_raw)
                st.success(f"Загружено {len(tariff_manager.tariffs)} тарифов")
            except Exception as e:
                st.error(f"Ошибка тарифов: {e}")
    return main_df, tariff_manager

def render_results(df_calc: pd.DataFrame) -> None:
    st.subheader("3. Результаты расчёта")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Всего SKU", len(df_calc))
    c2.metric("Средняя маржа, %", f"{df_calc['margin_percent'].mean():.1f}%")
    c3.metric("Убыточных SKU", len(df_calc[df_calc['gross_profit'] < 0]))
    c4.metric("Высокомаржинальных", len(df_calc[df_calc['profitability_status'] == 'Высокомаржинальный']))
    status_filter = st.multiselect(
        "Фильтр по статусу",
        options=df_calc['profitability_status'].unique(),
        default=df_calc['profitability_status'].unique()
    )
    filtered_df = df_calc[df_calc['profitability_status'].isin(status_filter)]
    display_cols = [
        'artikul', 'category', 'selling_price', 'cogs',
        'gross_profit', 'margin_percent', 'roi_percent',
        'profitability_status', 'source', 'scheme'
    ]
    st.dataframe(filtered_df[display_cols], use_container_width=True, height=400)
    return filtered_df

def render_export_section(filtered_df: pd.DataFrame, pipeline: DataPipeline) -> None:
    st.subheader("4. Экспорт")
    include_advanced = st.checkbox("Включить расширенные листы (анализ чувствительности, прогноз)", value=True)
    excel_data = UltimateExcelExporter.export_max_info(
        df=filtered_df,
        tax_label=pipeline.tax_label,
        scheme_label=pipeline.scheme_label,
        payment_frequency=pipeline.payment_frequency,
        include_advanced=include_advanced
    )
    if excel_data:
        st.download_button(
            label="📥 Скачать улучшенный Excel (Дашборд, Светофор, Freeze, Аналитика)",
            data=excel_data,
            file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Не удалось сформировать Excel. Проверьте установку openpyxl.")

def render_google_sheets_export(df: pd.DataFrame) -> None:
    st.subheader("5. Экспорт в Google Sheets")
    with st.expander("🔐 Настройки доступа к Google Sheets"):
        sheets_url = st.text_input(
            "Ссылка на Google Sheets (открыта на редактирование по ссылке)",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit"
        )
        credentials_file = st.file_uploader(
            "JSON-ключ сервисного аккаунта",
            type=['json'],
            help="Скачайте ключ из Google Cloud Console → Сервисные аккаунты"
        )
        sheet_name = st.text_input("Название листа", value="Unit Economics")
        if st.button("📤 Экспортировать в Google Sheets") and sheets_url and credentials_file:
            try:
                match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheets_url)
                if not match:
                    st.error("Не удалось распознать ID таблицы. Проверьте ссылку.")
                    return
                spreadsheet_id = match.group(1)
                
                # ФИКС: Чтение JSON напрямую в память, без создания временных файлов на диске
                credentials_info = json.loads(credentials_file.getvalue().decode('utf-8'))
                
                exporter = GoogleSheetsExporter(credentials_info)
                exporter.export(df, spreadsheet_id, sheet_name)
                st.success(f"✅ Данные успешно экспортированы в таблицу: {sheets_url}")
            except json.JSONDecodeError:
                st.error("❌ Ошибка: Загруженный файл не является валидным JSON.")
            except Exception as e:
                st.error(f"❌ Ошибка экспорта: {e}")

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================
def main():
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="📈",
        layout="wide"
    )
    init_session_state()
    st.title(f"📊 {APP_NAME} v{APP_VERSION}")
    st.markdown(
        "Монолитный калькулятор юнит-экономики "
        "с векторизованными вычислениями, профессиональным экспортом "
        "и улучшенной производительностью."
    )
    pipeline = render_sidebar()
    main_df, tariff_manager = render_upload_section()
    st.markdown("---")
    if main_df is not None and not main_df.empty:
        if st.button("🚀 Рассчитать юнит-экономику", type="primary"):
            with st.spinner("Выполняется векторизованный расчёт..."):
                ym_api = YandexMarketAPI(
                    api_key=st.session_state.api_key,
                    business_id=st.session_state.business_id
                ) if st.session_state.api_key else None
                perf_monitor = st.session_state.perf_monitor
                calc_df = pipeline.process(main_df, tariff_manager, ym_api, perf_monitor)
                st.session_state.calc_df = calc_df
                st.session_state.last_hash = StringUtils.make_hash(main_df)
                st.success("Расчёт завершён успешно!")
                perf_monitor.display_report()
        if not st.session_state.calc_df.empty:
            filtered_df = render_results(st.session_state.calc_df)
            render_export_section(filtered_df, pipeline)
            render_google_sheets_export(filtered_df)

if __name__ == "__main__":
    main()
