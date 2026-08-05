#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v24.0 — МАСШТАБИРУЕМАЯ ВЕРСИЯ
============================================================================
КЛЮЧЕВЫЕ УЛУЧШЕНИЯ ДЛЯ 300K+ SKU:
1. Экспорт через XlsxWriter с режимом constant_memory (минимальное потребление памяти)
2. Полноценное форматирование Excel: дашборд, KPI, таблицы, условное форматирование, диаграммы
3. Кэширование тарифов по категориям (один запрос на категорию вместо каждого SKU)
4. Параллельные API-запросы с учётом rate limit (ThreadPoolExecutor)
5. Оптимизация типов данных при загрузке (dtype для экономии памяти)
6. Пакетная обработка по чанкам при расчёте
7. Предложение CSV для больших объёмов (>100k строк)
8. Исправлена математика анализа чувствительности (реальный P&L)
9. Сохранение колонок source и scheme для аудита
============================================================================
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
import contextlib  # ДОБАВЛЕНО: исправление ошибки отсутствия импорта

# Дополнительные импорты
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
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import multiprocessing as mp
    CONCURRENT_AVAILABLE = True
except ImportError:
    CONCURRENT_AVAILABLE = False

# Заменяем openpyxl на xlsxwriter
try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

# Подавляем предупреждения
warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)

# ----------------------------------------------------------------------------
# НАСТРОЙКА ЛОГИРОВАНИЯ
# ----------------------------------------------------------------------------
def setup_logging():
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
    except Exception:
        pass
    return logger

logger = setup_logging()

# ----------------------------------------------------------------------------
# КОНФИГУРАЦИЯ
# ----------------------------------------------------------------------------
DEFAULT_CONFIG = {
    'app': {
        'version': '24.0.0',
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
        'rate_limit_per_second': 1,
        'parallel_api_workers': 5
    },
    'tax_systems': {
        'usn_6': {'label': 'УСН 6% (доходы)', 'rate': 0.06, 'base': 'revenue'},
        'usn_15': {'label': 'УСН 15% (доходы-расходы)', 'rate': 0.15, 'base': 'profit'},
        'osn': {'label': 'ОСН (общая с НДС 20%)', 'rate': 0.20, 'base': 'profit_vat'},
        'ausn_8': {'label': 'АУСН 8% (доходы)', 'rate': 0.08, 'base': 'revenue'}
    },
    'excel': {
        'max_rows_for_excel': 100000,
        'constant_memory': True
    }
}

def load_config(config_path: str = 'config.yaml') -> dict:
    if not YAML_AVAILABLE or not os.path.exists(config_path):
        return DEFAULT_CONFIG
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception:
        return DEFAULT_CONFIG

CONFIG = load_config()

APP_VERSION = CONFIG['app']['version']
APP_NAME = "Yandex Market Unit Economics PRO"
CACHE_TTL = CONFIG['app']['cache_ttl']
LRU_CACHE_SIZE = CONFIG['app']['lru_cache_size']
MAX_RETRIES = CONFIG['app']['max_retries']
REQUEST_TIMEOUT = CONFIG['app']['request_timeout']
API_BASE_URL = CONFIG['api']['base_url']
TARIFFS_ENDPOINT = CONFIG['api']['tariffs_endpoint']
CATEGORIES_ENDPOINT = CONFIG['api']['categories_endpoint']
RECOMMENDATIONS_ENDPOINT = CONFIG['api']['recommendations_endpoint']
MAX_OFFERS_PER_REQUEST = CONFIG['api']['max_offers_per_request']
RATE_LIMIT_PER_SECOND = CONFIG['api']['rate_limit_per_second']
PARALLEL_API_WORKERS = CONFIG['api']['parallel_api_workers']
MAX_ROWS_FOR_EXCEL = CONFIG['excel']['max_rows_for_excel']

# ============================================================================
# ENUM и типы
# ============================================================================
class SellingProgram(str, Enum):
    FBS = "FBS"
    FBY = "FBY"
    DBS = "DBS"
    EXPRESS = "EXPRESS"
    LAAS = "LAAS"

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

# ============================================================================
# УТИЛИТЫ
# ============================================================================
class NumericUtils:
    @staticmethod
    def money_round(values):
        if isinstance(values, pd.Series):
            return values.fillna(0.0).round(2)
        return np.round(np.nan_to_num(values, nan=0.0, posinf=0.0, neginf=0.0), 2)
    
    @staticmethod
    def percent_round(values):
        if isinstance(values, pd.Series):
            return values.fillna(0.0).round(2)
        return np.round(np.nan_to_num(values, nan=0.0, posinf=0.0, neginf=0.0), 2)
    
    @staticmethod
    def safe_divide(num, den, default=0.0):
        if isinstance(num, pd.Series):
            return np.where(den != 0, num / den, default)
        with np.errstate(divide='ignore', invalid='ignore'):
            result = np.divide(num, den)
            result[~np.isfinite(result)] = default
            return result

class DtypeOptimizer:
    INT_COLS = {'daily_sales', 'stock_depth_days', 'quantity_per_order'}
    FLOAT_COLS = {
        'selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm',
        'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost',
        'marketing_budget_per_unit', 'warehouse_cost'
    }
    
    @classmethod
    def optimize(cls, df):
        df = df.copy()
        for col in ['artikul', 'category', 'abc_category', 'xyz_category', 
                     'profitability_status', 'source', 'scheme']:
            if col in df.columns:
                df[col] = df[col].astype('category')
        for col in cls.INT_COLS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(np.int32)
        for col in cls.FLOAT_COLS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(np.float32)
        if 'is_special_tariff' in df.columns:
            df['is_special_tariff'] = df['is_special_tariff'].astype(bool)
        return df

class StringUtils:
    @staticmethod
    def fix_double_utf8(text):
        if not isinstance(text, str) or not text:
            return text
        if 'Ð' in text or 'Ã' in text or 'â€™' in text:
            try:
                return text.encode('cp1251').decode('utf-8')
            except:
                try:
                    return text.encode('latin1').decode('utf-8')
                except:
                    return text
        return text
    
    @staticmethod
    def make_hash(obj):
        try:
            if isinstance(obj, pd.DataFrame):
                return hashlib.sha256(
                    pd.util.hash_pandas_object(obj, index=True).values.tobytes()
                ).hexdigest()[:16]
            return hashlib.sha256(str(obj).encode('utf-8')).hexdigest()[:16]
        except:
            return hashlib.sha256(b"hash_fallback").hexdigest()[:16]

class LRUCache:
    def __init__(self, max_size=LRU_CACHE_SIZE, ttl=CACHE_TTL):
        self.max_size = max_size
        self.ttl = ttl
        self._cache = OrderedDict()
    
    def get(self, key):
        if key not in self._cache:
            return None
        value, timestamp = self._cache[key]
        if time.time() - timestamp > self.ttl:
            del self._cache[key]
            return None
        self._cache.move_to_end(key)
        return value
    
    def set(self, key, value):
        if key in self._cache:
            self._cache.move_to_end(key)
        self._cache[key] = (value, time.time())
        if len(self._cache) > self.max_size:
            self._cache.popitem(last=False)

# ============================================================================
# МОНИТОРИНГ
# ============================================================================
class PerformanceMonitor:
    def __init__(self):
        self.metrics = {}
    
    @contextmanager
    def measure(self, metric_name):
        start = time.time()
        start_mem = self._get_memory()
        try:
            yield
        finally:
            end = time.time()
            end_mem = self._get_memory()
            if metric_name not in self.metrics:
                self.metrics[metric_name] = []
            self.metrics[metric_name].append({
                'duration': end - start,
                'memory_mb': (end_mem - start_mem) if start_mem is not None else 0
            })
            logger.info(f"{metric_name} выполнен за {end-start:.2f}с")
    
    def _get_memory(self):
        if PSUTIL_AVAILABLE:
            try:
                return psutil.Process().memory_info().rss / 1024 / 1024
            except:
                return None
        return None
    
    def report(self):
        report = {}
        for key, vals in self.metrics.items():
            durations = [v['duration'] for v in vals]
            memories = [v['memory_mb'] for v in vals if v['memory_mb'] is not None]
            report[key] = {
                'count': len(vals),
                'avg_duration': np.mean(durations) if durations else 0,
                'total_duration': np.sum(durations),
                'avg_memory': np.mean(memories) if memories else 0,
                'max_memory': np.max(memories) if memories else 0
            }
        return report
    
    def display_report(self):
        report = self.report()
        if not report:
            return
        st.subheader("📊 Отчёт по производительности")
        for key, data in report.items():
            st.write(f"**{key}**: {data['count']} опер., среднее {data['avg_duration']:.2f}с, память {data['avg_memory']:.2f} МБ")

# ============================================================================
# МОДЕЛИ
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
    def by_label(cls, label):
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
    
    def to_dict(self):
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
    def default(cls, category="default"):
        return cls(category=category)

# ============================================================================
# API КЛИЕНТ С ПАРАЛЛЕЛЬНЫМИ ЗАПРОСАМИ
# ============================================================================
class RateLimiter:
    def __init__(self, max_calls=RATE_LIMIT_PER_SECOND, period=1.0):
        self.max_calls = max_calls
        self.period = period
        self.calls = []
        self._lock = None
        try:
            import threading
            self._lock = threading.Lock()
        except:
            pass
    
    def wait_if_needed(self):
        if self._lock:
            with self._lock:
                now = time.time()
                self.calls = [t for t in self.calls if now - t < self.period]
                if len(self.calls) >= self.max_calls:
                    sleep_time = self.period - (now - self.calls[0])
                    if sleep_time > 0:
                        time.sleep(sleep_time)
                self.calls.append(time.time())
        else:
            now = time.time()
            self.calls = [t for t in self.calls if now - t < self.period]
            if len(self.calls) >= self.max_calls:
                sleep_time = self.period - (now - self.calls[0])
                if sleep_time > 0:
                    time.sleep(sleep_time)
            self.calls.append(time.time())

class APIClient:
    def __init__(self, base_url, api_key, max_retries=MAX_RETRIES, timeout=REQUEST_TIMEOUT):
        self.base_url = base_url.rstrip('/')
        self.api_key = api_key
        self.timeout = timeout
        self.max_retries = max_retries
        self.rate_limiter = RateLimiter()
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
    
    def _request(self, method, endpoint, **kwargs):
        if not self.api_key:
            return {}
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        self.rate_limiter.wait_if_needed()
        for attempt in range(self.max_retries + 1):
            try:
                resp = self.session.request(
                    method, url,
                    timeout=self.timeout,
                    **kwargs
                )
                resp.raise_for_status()
                return resp.json()
            except requests.exceptions.Timeout:
                wait_time = 2 ** attempt
                logger.warning(f"Таймаут, попытка {attempt+1}, ждём {wait_time}s")
                if attempt < self.max_retries:
                    time.sleep(wait_time)
                else:
                    logger.error(f"Таймаут после {self.max_retries+1} попыток")
            except requests.exceptions.HTTPError as e:
                if e.response.status_code in [429, 500, 502, 503, 504]:
                    wait_time = 2 ** attempt
                    logger.warning(f"HTTP {e.response.status_code}, попытка {attempt+1}")
                    if attempt < self.max_retries:
                        time.sleep(wait_time)
                    else:
                        logger.error(f"HTTP ошибка после попыток: {e}")
                else:
                    logger.warning(f"HTTP {e.response.status_code}: {e.response.text[:200]}")
                    break
            except Exception as e:
                logger.warning(f"Ошибка запроса: {e}")
                break
        return {}
    
    def close(self):
        """Закрывает сессию requests"""
        if hasattr(self, 'session'):
            self.session.close()

class YandexMarketAPI(APIClient):
    def __init__(self, api_key, business_id=None):
        super().__init__(API_BASE_URL, api_key)
        self.business_id = business_id
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if business_id:
            self.headers["X-Business-Id"] = business_id
        self._category_cache = {}
    
    def get_categories_tree(self):
        data = self._request("GET", CATEGORIES_ENDPOINT, headers=self.headers)
        return data.get("result", {}).get("categories", [])
    
    def get_category_id_by_name(self, category_name):
        if category_name in self._category_cache:
            return self._category_cache[category_name]
        categories = self.get_categories_tree()
        queue = list(categories)
        found = None
        while queue:
            cat = queue.pop(0)
            if cat.get('name', '').lower() == category_name.lower():
                found = cat.get('id')
                break
            if 'children' in cat and cat['children']:
                queue.extend(cat['children'])
        if found is not None:
            self._category_cache[category_name] = found
        return found
    
    def calculate_tariffs_batch(self, offers, campaign_id=None, selling_program=SellingProgram.FBS,
                                 transit_warehouse_type=None, order_cargo_type=None):
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
            data = self._request("POST", TARIFFS_ENDPOINT, headers=self.headers, json=payload)
            chunk_results = data.get("result", {}).get("offers", [])
            all_results.extend(chunk_results)
        return all_results
    
    def calculate_tariffs(self, offers, campaign_id=None, selling_program="FBS",
                          transit_warehouse_type=None, order_cargo_type=None):
        try:
            prog = SellingProgram(selling_program.upper())
        except ValueError:
            prog = SellingProgram.FBS
        return self.calculate_tariffs_batch(
            offers, campaign_id, prog, transit_warehouse_type, order_cargo_type
        )
    
    @staticmethod
    def _parse_tariff(tariff_data):
        return Tariff(
            category=tariff_data.get('category', 'default'),
            commission_rate=tariff_data.get('commission', 0.15) / 100,
            delivery_rate=tariff_data.get('delivery', 0.05),
            source="API Яндекс Маркета",
            scheme=tariff_data.get('sellingProgram', 'FBS')
        )

    def fetch_tariff_for_offer(self, offer_params, category_name, scheme):
        try:
            result = self.calculate_tariffs(
                [offer_params],
                selling_program=scheme,
                transit_warehouse_type=offer_params.get('transit_warehouse_type'),
                order_cargo_type=offer_params.get('order_cargo_type')
            )
            if result and len(result) > 0:
                tariff = self._parse_tariff(result[0])
                tariff.category = category_name
                return tariff
        except Exception as e:
            logger.warning(f"API ошибка для {category_name}: {e}")
        return None

# ============================================================================
# МЕНЕДЖЕР ТАРИФОВ С ГРУППИРОВКОЙ
# ============================================================================
class HybridTariffManager:
    def __init__(self):
        if 'tariffs' not in st.session_state:
            st.session_state.tariffs = {}
        self._cache = LRUCache()
        self._api = None
    
    @property
    def tariffs(self):
        return st.session_state.tariffs
    
    def load_tariffs_from_file(self, df):
        required = {'category', 'commission_rate'}
        if not required.issubset(df.columns):
            raise ValueError(f"Файл тарифов должен содержать: {required}")
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
    
    def get_tariffs_vectorized(self, df, scheme, ym_api=None, use_api=True):
        unique_cats = df['category'].cat.categories if hasattr(df['category'], 'cat') else df['category'].unique()
        tariff_map = {}
        
        for cat in unique_cats:
            if cat in self.tariffs:
                t = self.tariffs[cat]
                t.scheme = scheme
                tariff_map[cat] = t
        
        if use_api and ym_api and ym_api.api_key:
            missing_cats = [cat for cat in unique_cats if cat not in tariff_map]
            if missing_cats:
                logger.info(f"Запрос тарифов для {len(missing_cats)} категорий через API")
                offers_by_cat = {}
                for cat in missing_cats:
                    sample = df[df['category'] == cat].iloc[0] if not df[df['category'] == cat].empty else {}
                    offer = {
                        "price": float(sample.get('selling_price', 1000)),
                        "weight": float(sample.get('weight_kg', 1.0)),
                        "length": float(sample.get('length_cm', 10)),
                        "width": float(sample.get('width_cm', 10)),
                        "height": float(sample.get('height_cm', 10)),
                        "quantity": int(sample.get('quantity_per_order', 1))
                    }
                    cat_id = ym_api.get_category_id_by_name(cat)
                    if cat_id:
                        offer["categoryId"] = cat_id
                    else:
                        offer["categoryId"] = 0
                    offers_by_cat[cat] = offer
                
                def fetch_for_category(cat):
                    t = ym_api.fetch_tariff_for_offer(offers_by_cat[cat], cat, scheme)
                    if t:
                        return cat, t
                    return cat, None
                
                with ThreadPoolExecutor(max_workers=PARALLEL_API_WORKERS) as executor:
                    future_to_cat = {executor.submit(fetch_for_category, cat): cat for cat in missing_cats}
                    for future in as_completed(future_to_cat):
                        cat, tariff = future.result()
                        if tariff:
                            tariff_map[cat] = tariff
                            self.tariffs[cat] = tariff
                            logger.info(f"Получен тариф для {cat}")
                        else:
                            logger.warning(f"Не удалось получить тариф для {cat}, используется базовый")
                            tariff_map[cat] = Tariff.default(cat)
                            tariff_map[cat].scheme = scheme
        
        for cat in unique_cats:
            if cat not in tariff_map:
                tariff_map[cat] = Tariff.default(cat)
                tariff_map[cat].scheme = scheme
        
        tariff_df = pd.DataFrame([
            {'category': cat, **t.to_dict()}
            for cat, t in tariff_map.items()
        ])
        return tariff_df

# ============================================================================
# ВАЛИДАТОР
# ============================================================================
class DataValidator:
    REQUIRED_COLS = ('artikul', 'category', 'selling_price', 'cogs')
    NUMERIC_COLS = (
        'selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm',
        'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost',
        'marketing_budget_per_unit', 'stock_depth_days', 'quantity_per_order',
        'daily_sales'
    )
    
    @classmethod
    def validate(cls, df):
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
                errors.append(f"selling_price: {zero_prices} SKU с нулевой ценой")
        if 'quantity_per_order' in df_validated.columns:
            df_validated['quantity_per_order'] = df_validated['quantity_per_order'].replace(0, 1)
        return df_validated, errors

# ============================================================================
# ФИНАНСОВЫЙ ДВИЖОК
# ============================================================================
class FinancialEngine:
    PICK_PACK_COST = 35.0
    SPECIAL_PRICE_THRESHOLD = 300.0
    SPECIAL_VOLUME_THRESHOLD = 5.0
    
    @staticmethod
    def calculate_billable_weight(df):
        vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
        billable = np.maximum(df['weight_kg'].values, vol_weight.values)
        return np.ceil(billable * 2) / 2
    
    @staticmethod
    def calculate_middle_mile(billable_weight):
        return np.select(
            [billable_weight <= 4, billable_weight <= 10],
            [100, 300],
            default=600
        )
    
    @classmethod
    def calculate_all(cls, df, tax_config, scheme, payment_rate, tariffs_map):
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
        
        tech_cols_to_drop = ['is_special_tariff', 'billable_weight']
        df = df.drop(columns=[c for c in tech_cols_to_drop if c in df.columns])
        return df

# ============================================================================
# КЭШИРОВАННЫЙ РАСЧЁТ
# ============================================================================
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
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
    return result

# ============================================================================
# ЭКСПОРТ ЧЕРЕЗ XLSXWRITER С ПОЛНЫМ ФОРМАТИРОВАНИЕМ
# ============================================================================
class UltimateExcelExporter:
    """
    Полноценный экспорт в Excel с использованием xlsxwriter.
    Включает дашборд с KPI, детальный лист, рекомендации, ABC/XYZ, анализ чувствительности, прогноз.
    Поддерживает режим constant_memory для больших данных.
    """
    @staticmethod
    def export_to_excel(df, tax_label, scheme_label, payment_frequency, include_advanced=True):
        if not XLSXWRITER_AVAILABLE or df.empty:
            return b""
        
        if len(df) > MAX_ROWS_FOR_EXCEL:
            logger.warning(f"Слишком много строк ({len(df)}) для Excel, генерируем только сводку")
            return UltimateExcelExporter._export_summary_only(df, tax_label, scheme_label, payment_frequency)
        
        return UltimateExcelExporter._export_full_excel(df, tax_label, scheme_label, payment_frequency, include_advanced)
    
    @staticmethod
    def _export_full_excel(df, tax_label, scheme_label, payment_frequency, include_advanced):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'constant_memory': True, 'in_memory': False})
        
        # Определяем форматы
        money_fmt = workbook.add_format({'num_format': '#,##0.00 ₽'})
        pct_fmt = workbook.add_format({'num_format': '0.00%'})
        pct_fmt1 = workbook.add_format({'num_format': '0.0%'})
        int_fmt = workbook.add_format({'num_format': '#,##0'})
        bold = workbook.add_format({'bold': True})
        header_fmt = workbook.add_format({
            'bold': True,
            'bg_color': '#1F4E79',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#DEEAF1'})
        zebra_light = workbook.add_format({'bg_color': '#F2F2F2'})
        zebra_dark = workbook.add_format({'bg_color': '#FFFFFF'})
        green_fill = workbook.add_format({'bg_color': '#C6EFCE'})
        red_fill = workbook.add_format({'bg_color': '#FFC7CE'})
        yellow_fill = workbook.add_format({'bg_color': '#FFEB9C'})
        blue_fill = workbook.add_format({'bg_color': '#DEEAF1'})
        navy_font = workbook.add_format({'bold': True, 'color': '#1F4E79'})
        white_font = workbook.add_format({'color': 'white'})
        
        # 1. ДАШБОРД
        ws_dash = workbook.add_worksheet("📊 Дашборд")
        ws_dash.set_column('A:A', 2)
        ws_dash.set_column('B:M', 12)
        ws_dash.set_row(0, 30)
        ws_dash.set_row(1, 40)
        ws_dash.set_row(2, 20)
        
        # Заголовок
        title_format = workbook.add_format({'bold': True, 'size': 20, 'bg_color': '#1F4E79', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
        ws_dash.merge_range('B2:N2', '📊 UNIT-ECONOMICS · ЯНДЕКС МАРКЕТ', title_format)
        sub_format = workbook.add_format({'italic': True, 'size': 10, 'bg_color': '#2E75B6', 'font_color': 'white', 'align': 'center'})
        ws_dash.merge_range('B3:N3', f"{scheme_label} · {tax_label} · {payment_frequency} · {datetime.now().strftime('%d.%m.%Y %H:%M')}", sub_format)
        
        # KPI карточки
        total_sku = len(df)
        avg_margin = df["margin_percent"].mean() if "margin_percent" in df.columns else 0
        profitable = int((df["gross_profit"] > 0).sum()) if "gross_profit" in df.columns else 0
        loss = int((df["gross_profit"] < 0).sum()) if "gross_profit" in df.columns else 0
        hi_margin = int((df["profitability_status"] == "Высокомаржинальный").sum()) if "profitability_status" in df.columns else 0
        avg_roi = df["roi_percent"].mean() if "roi_percent" in df.columns else 0
        
        kpi_data = [
            ("🏷️ Всего SKU", total_sku, int_fmt, '#1F4E79', '#DEEAF1'),
            ("📈 Средняя маржа", avg_margin/100, pct_fmt, '#375623', '#E2EFDA'),
            ("✅ Прибыльных", profitable, int_fmt, '#375623', '#E2EFDA'),
            ("❌ Убыточных", loss, int_fmt, '#C00000', '#FFC7CE'),
            ("⭐ Высокомаржинальных", hi_margin, int_fmt, '#1F4E79', '#DEEAF1'),
            ("💰 Средний ROI", avg_roi/100, pct_fmt, '#2E75B6', '#BDD7EE')
        ]
        for i, (label, val, fmt, fg, bg) in enumerate(kpi_data):
            row = 5 + (i // 3)*4
            col = 2 + (i % 3)*2
            # Фон для карточки
            card_format = workbook.add_format({'bg_color': bg, 'border': 1})
            ws_dash.merge_range(row, col, row+3, col+1, '', card_format)
            # Метка
            label_format = workbook.add_format({'bold': True, 'color': fg, 'bg_color': bg, 'align': 'center'})
            ws_dash.merge_range(row, col, row, col+1, label, label_format)
            # Значение
            val_format = workbook.add_format({'bold': True, 'size': 18, 'color': fg, 'bg_color': bg, 'align': 'center', 'valign': 'vcenter'})
            ws_dash.merge_range(row+1, col, row+3, col+1, val, val_format)
            if fmt:
                # Нельзя применить формат к merged range напрямую, поэтому пишем значение с форматом в первую ячейку
                ws_dash.write(row+1, col, val, fmt)
                # Остальные ячейки оставляем пустыми, но они уже имеют фон
        ws_dash.set_row(5, 22)
        ws_dash.set_row(6, 40)
        ws_dash.set_row(7, 40)
        ws_dash.set_row(8, 22)
        ws_dash.set_row(9, 5)
        
        # Таблица статусов
        if "profitability_status" in df.columns:
            status_summary = df.groupby('profitability_status', observed=True).agg(
                SKU=('artikul', 'count'),
                Прибыль=('gross_profit', 'sum'),
                Маржа_avg=('margin_percent', 'mean')
            ).reset_index()
            status_summary.columns = ['Статус', 'SKU', 'Прибыль', 'Маржа_avg']
            row_start = 12
            ws_dash.merge_range(row_start, 2, row_start, 5, 'Структура портфеля по маржинальности', 
                                workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white', 'align': 'center'}))
            row_start += 1
            headers = ['Статус', 'SKU, шт.', 'Сум. прибыль, ₽', 'Ср. маржа, %']
            for c, h in enumerate(headers):
                ws_dash.write(row_start, c+2, h, header_fmt)
            status_colors = {
                'Высокомаржинальный': ('#375623', '#E2EFDA'),
                'Низкомаржинальный': ('#7F6000', '#FFEB9C'),
                'Убыточный': ('#C00000', '#FFC7CE')
            }
            for r_idx, row in enumerate(status_summary.itertuples(), 1):
                fg, bg = status_colors.get(row.Статус, ('#000000', '#FFFFFF'))
                fmt_status = workbook.add_format({'bold': True, 'color': fg, 'bg_color': bg})
                ws_dash.write(row_start + r_idx, 2, row.Статус, fmt_status)
                ws_dash.write(row_start + r_idx, 3, row.SKU, int_fmt)
                ws_dash.write(row_start + r_idx, 4, row.Прибыль, money_fmt)
                ws_dash.write(row_start + r_idx, 5, row.Маржа_avg/100, pct_fmt)
        
        # 2. ДЕТАЛЬНЫЙ ЛИСТ
        ws_det = workbook.add_worksheet("📋 Детальный расчёт")
        ws_det.freeze_panes(1, 2)
        # Определяем колонки для вывода
        cols_to_export = [
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
        cols_to_export = [c for c in cols_to_export if c in df.columns]
        df_export = df[cols_to_export].copy()
        
        # Заголовки
        for c, col in enumerate(cols_to_export):
            ws_det.write(0, c, col, header_fmt)
        # Запись данных с форматированием
        for r_idx, row in enumerate(df_export.itertuples(index=False), 1):
            # Зебра
            bg = zebra_light if r_idx % 2 == 0 else zebra_dark
            for c_idx, val in enumerate(row):
                col_name = cols_to_export[c_idx]
                if col_name in ['selling_price', 'cogs', 'commission', 'delivery_to_customer', 
                                'middle_mile_cost', 'sorting_cost', 'acquiring_cost', 'return_cost',
                                'pick_pack_cost', 'packaging_cost', 'first_mile_cost',
                                'marketing_budget_per_unit', 'warehouse_cost',
                                'tax_cost', 'total_expenses', 'gross_profit',
                                'rec_price_min', 'rec_price_15', 'rec_price_25']:
                    fmt = money_fmt
                elif col_name in ['margin_percent', 'roi_percent']:
                    fmt = pct_fmt
                elif col_name in ['break_even_units']:
                    fmt = int_fmt
                else:
                    fmt = None
                # Особое форматирование для прибыли/убытка и статусов
                if col_name == 'gross_profit' and isinstance(val, (int, float)) and val < 0:
                    cell_format = workbook.add_format({'num_format': '#,##0.00 ₽', 'bg_color': '#FFC7CE', 'font_color': '#C00000'})
                elif col_name == 'profitability_status':
                    if val == 'Высокомаржинальный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'color': '#375623'})
                    elif val == 'Низкомаржинальный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'color': '#7F6000'})
                    elif val == 'Убыточный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#FFC7CE', 'color': '#C00000'})
                    else:
                        cell_format = bg
                elif col_name in ['abc_category', 'abc_xyz']:
                    if val in ['A', 'AX', 'AY', 'AZ']:
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'color': 'white'})
                    elif val in ['B', 'BX', 'BY', 'BZ']:
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#2E75B6', 'color': 'white'})
                    else:
                        cell_format = bg
                else:
                    cell_format = bg
                if fmt and not col_name in ['gross_profit', 'profitability_status', 'abc_category', 'abc_xyz']:
                    # Создаём формат с числовым форматом и фоном
                    cell_format = workbook.add_format({'num_format': fmt.get_num_format(), 'bg_color': bg.bg_color if hasattr(bg, 'bg_color') else '#FFFFFF'})
                ws_det.write(r_idx, c_idx, val, cell_format)
        
        # Автоширина колонок (приблизительная)
        for c, col in enumerate(cols_to_export):
            max_len = max(len(col), df_export[col].astype(str).str.len().max() if not df_export.empty else 0)
            ws_det.set_column(c, c, min(max(max_len+2, 8), 30))
        
        # 3. РЕКОМЕНДАЦИИ
        ws_rec = workbook.add_worksheet("💡 Рекомендации")
        rec_cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent', 
                    'rec_price_min', 'rec_price_15', 'rec_price_25', 'profitability_status']
        rec_cols = [c for c in rec_cols if c in df.columns]
        df_rec = df[rec_cols].copy()
        for c, col in enumerate(rec_cols):
            ws_rec.write(0, c, col, header_fmt)
        for r_idx, row in enumerate(df_rec.itertuples(index=False), 1):
            bg = zebra_light if r_idx % 2 == 0 else zebra_dark
            for c_idx, val in enumerate(row):
                col_name = rec_cols[c_idx]
                if col_name in ['selling_price', 'cogs', 'gross_profit', 'rec_price_min', 'rec_price_15', 'rec_price_25']:
                    fmt = money_fmt
                elif col_name == 'margin_percent':
                    fmt = pct_fmt
                else:
                    fmt = None
                if col_name == 'profitability_status':
                    if val == 'Высокомаржинальный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'color': '#375623'})
                    elif val == 'Низкомаржинальный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'color': '#7F6000'})
                    elif val == 'Убыточный':
                        cell_format = workbook.add_format({'bold': True, 'bg_color': '#FFC7CE', 'color': '#C00000'})
                    else:
                        cell_format = bg
                else:
                    cell_format = bg
                if fmt and col_name not in ['profitability_status']:
                    cell_format = workbook.add_format({'num_format': fmt.get_num_format(), 'bg_color': bg.bg_color if hasattr(bg, 'bg_color') else '#FFFFFF'})
                ws_rec.write(r_idx, c_idx, val, cell_format)
        # Автоширина
        for c, col in enumerate(rec_cols):
            max_len = max(len(col), df_rec[col].astype(str).str.len().max() if not df_rec.empty else 0)
            ws_rec.set_column(c, c, min(max(max_len+2, 8), 30))
        ws_rec.freeze_panes(1, 2)
        
        # 4. ABC/XYZ матрица
        ws_abc = workbook.add_worksheet("🔢 ABC·XYZ матрица")
        if 'abc_xyz' in df.columns:
            matrix = df.groupby('abc_xyz', observed=True).agg(
                SKU=('artikul', 'count'),
                Прибыль=('gross_profit', 'sum'),
                Маржа=('margin_percent', 'mean')
            ).reset_index().sort_values('abc_xyz')
            headers = ['Сегмент', 'SKU, шт.', 'Сум. прибыль, ₽', 'Ср. маржа, %']
            for c, h in enumerate(headers):
                ws_abc.write(0, c, h, header_fmt)
            abc_colors = {
                'AX': '#E2EFDA', 'AY': '#E2EFDA', 'AZ': '#FFEB9C',
                'BX': '#DEEAF1', 'BY': '#DEEAF1', 'BZ': '#FFEB9C',
                'CX': '#F2F2F2', 'CY': '#FCE4D6', 'CZ': '#FFC7CE'
            }
            for r_idx, row in enumerate(matrix.itertuples(), 1):
                seg = row.abc_xyz
                bg = abc_colors.get(seg, '#FFFFFF')
                seg_format = workbook.add_format({'bold': True, 'bg_color': bg})
                ws_abc.write(r_idx, 0, seg, seg_format)
                ws_abc.write(r_idx, 1, row.SKU, int_fmt)
                ws_abc.write(r_idx, 2, row.Прибыль, money_fmt)
                ws_abc.write(r_idx, 3, row.Маржа/100, pct_fmt)
            ws_abc.set_column('A:A', 12)
            ws_abc.set_column('B:B', 12)
            ws_abc.set_column('C:C', 18)
            ws_abc.set_column('D:D', 14)
        
        # 5. ПАРАМЕТРЫ
        ws_par = workbook.add_worksheet("⚙️ Параметры")
        params = [
            ("Версия", APP_VERSION),
            ("Дата", datetime.now().strftime("%d.%m.%Y %H:%M:%S")),
            ("Налог", tax_label),
            ("Схема", scheme_label),
            ("Периодичность", payment_frequency),
            ("Всего SKU", len(df)),
            ("Прибыльных", profitable),
            ("Убыточных", loss),
            ("Средняя маржа", f"{avg_margin:.2f}%")
        ]
        for i, (p, v) in enumerate(params):
            ws_par.write(i, 0, p, bold)
            ws_par.write(i, 1, v)
        ws_par.set_column('A:A', 25)
        ws_par.set_column('B:B', 30)
        
        # 6. АНАЛИЗ ЧУВСТВИТЕЛЬНОСТИ (если включено)
        if include_advanced:
            UltimateExcelExporter._add_sensitivity_analysis(workbook, df, money_fmt, pct_fmt, int_fmt, header_fmt)
            UltimateExcelExporter._add_forecast(workbook, df, money_fmt, pct_fmt, int_fmt, header_fmt)
        
        workbook.close()
        output.seek(0)
        return output.getvalue()
    
    @staticmethod
    def _add_sensitivity_analysis(workbook, df, money_fmt, pct_fmt, int_fmt, header_fmt):
        ws = workbook.add_worksheet("🎯 Анализ чувствительности")
        ws.write(0, 0, "Изменение цены, %", header_fmt)
        ws.write(0, 1, "Изменение комиссии, п.п.", header_fmt)
        ws.write(0, 2, "Средняя маржа, %", header_fmt)
        ws.write(0, 3, "Кол-во прибыльных", header_fmt)
        ws.write(0, 4, "Кол-во убыточных", header_fmt)
        ws.set_column('A:A', 18)
        ws.set_column('B:B', 18)
        ws.set_column('C:C', 18)
        ws.set_column('D:D', 18)
        ws.set_column('E:E', 18)
        
        price_changes = [-20, -10, 0, 10, 20]
        commission_changes = [-5, -2, 0, 2, 5]
        
        base_fixed = (df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + 
                      df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost'])
        acq_rate = df['acquiring_transfer_rate'].values if 'acquiring_transfer_rate' in df.columns else np.full(len(df), 0.016)
        acq_sku = df['acquiring_sku_cost'].values if 'acquiring_sku_cost' in df.columns else np.full(len(df), 0.12)
        
        row = 1
        for pc in price_changes:
            for cc in commission_changes:
                new_price = df['selling_price'].values * (1 + pc / 100)
                new_comm_rate = np.maximum(0, df['commission_rate'].values + (cc / 100))
                new_commission = np.maximum(new_price * new_comm_rate, df['min_commission'].values)
                new_acquiring = acq_sku + (new_price * acq_rate)
                new_total_expenses = (base_fixed.values + new_commission + df['delivery_to_customer'].values + 
                                      df['middle_mile_cost'].values + df['sorting_cost'].values + 
                                      new_acquiring + df['return_cost'].values)
                new_gross_profit = new_price - new_total_expenses
                new_margin = np.where(new_price > 0, (new_gross_profit / new_price) * 100, 0.0)
                avg_margin = float(np.mean(new_margin))
                profit_cnt = int(np.sum(new_gross_profit > 0))
                loss_cnt = int(np.sum(new_gross_profit <= 0))
                ws.write(row, 0, pc)
                ws.write(row, 1, cc)
                ws.write(row, 2, avg_margin/100, pct_fmt)
                ws.write(row, 3, profit_cnt, int_fmt)
                ws.write(row, 4, loss_cnt, int_fmt)
                # Условное форматирование не применяем из-за ограничений xlsxwriter
                row += 1
    
    @staticmethod
    def _add_forecast(workbook, df, money_fmt, pct_fmt, int_fmt, header_fmt):
        ws = workbook.add_worksheet("📈 Прогноз")
        ws.write(0, 0, "Показатель", header_fmt)
        for m in range(1, 13):
            ws.write(0, m, f"Месяц {m}", header_fmt)
        base_profit = float(df['gross_profit'].mean()) if 'gross_profit' in df.columns else 0.0
        base_margin = float(df['margin_percent'].mean()) if 'margin_percent' in df.columns else 0.0
        growth = [1.0, 1.05, 1.10, 1.08, 1.12, 1.15, 1.10, 1.05, 1.08, 1.12, 1.15, 1.20]
        
        ws.write(1, 0, "Прогноз роста продаж, %")
        for i, g in enumerate(growth, 1):
            ws.write(1, i, g, pct_fmt)
        
        ws.write(2, 0, "Средняя маржа, %")
        for i in range(1, 13):
            val = (base_margin / 100) * (1 + 0.01 * (i-1))
            ws.write(2, i, val, pct_fmt)
        
        ws.write(3, 0, "Прогноз прибыли, ₽")
        for i in range(1, 13):
            g = growth[i-1]
            margin_factor = 1 + 0.01 * (i-1)
            val = base_profit * g * margin_factor
            ws.write(3, i, val, money_fmt)
        
        # Диаграмма (линейная)
        try:
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': 'Прогноз прибыли',
                'categories': '=📈 Прогноз!$B$1:$M$1',
                'values': '=📈 Прогноз!$B$4:$M$4',
                'line': {'color': '#1F4E79'},
                'marker': {'type': 'circle', 'size': 5}
            })
            chart.set_title({'name': 'Прогноз прибыли на 12 месяцев'})
            chart.set_x_axis({'name': 'Месяц'})
            chart.set_y_axis({'name': 'Прибыль, ₽', 'num_format': '#,##0 ₽'})
            ws.insert_chart('A6', chart)
        except Exception as e:
            logger.warning(f"Ошибка создания диаграммы прогноза: {e}")
    
    @staticmethod
    def _export_summary_only(df, tax_label, scheme_label, payment_frequency):
        agg_df = df.groupby('category', observed=True).agg(
            SKU=('artikul', 'count'),
            Avg_price=('selling_price', 'mean'),
            Avg_margin=('margin_percent', 'mean'),
            Total_profit=('gross_profit', 'sum'),
            Profitable=('gross_profit', lambda x: (x > 0).sum()),
            Loss=('gross_profit', lambda x: (x < 0).sum())
        ).reset_index()
        output = io.BytesIO()
        agg_df.to_csv(output, index=False, encoding='utf-8-sig')
        output.seek(0)
        return output.getvalue()

# ============================================================================
# НОРМАЛИЗАТОР ДАННЫХ С ОПТИМИЗАЦИЕЙ ПАМЯТИ
# ============================================================================
class UniversalDataNormalizer:
    COLUMN_MAPPING = {
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
    NUMERIC_COLS = (
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
        norm_df['artikul'] = norm_df['artikul'].astype(str).str.strip()
        norm_df['category'] = norm_df['category'].astype(str).str.strip().str.lower()
        norm_df = norm_df.drop_duplicates(subset=['artikul'], keep='first')
        norm_df = DtypeOptimizer.optimize(norm_df)
        return norm_df
    
    @classmethod
    def load_file(cls, file_buffer, file_name):
        try:
            if file_name.endswith('.csv'):
                return pd.read_csv(
                    file_buffer,
                    sep=None,
                    engine='python',
                    encoding='utf-8',
                    encoding_errors='replace',
                    low_memory=False
                )
            elif file_name.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file_buffer, dtype={})
            else:
                raise ValueError("Неподдерживаемый формат")
        except UnicodeDecodeError:
            file_buffer.seek(0)
            return pd.read_csv(
                file_buffer,
                sep=None,
                engine='python',
                encoding='cp1251',
                encoding_errors='replace',
                low_memory=False
            )

# ============================================================================
# ОРКЕСТРАТОР ПАЙПЛАЙНА
# ============================================================================
@dataclass
class DataPipeline:
    tax_label: str = "УСН 6% (доходы)"
    scheme_label: str = "FBS (склад продавца)"
    payment_frequency: str = "Еженедельно, 4 нед. (1.6%)"
    use_api: bool = True
    parallel: bool = False
    chunk_size: Optional[int] = None
    
    def process(self, raw_df, tariff_manager, ym_api=None, perf_monitor=None):
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
            if self.parallel and CONCURRENT_AVAILABLE and len(validated_df) > 1000:
                calc_df = ParallelProcessor.process_in_parallel(
                    validated_df,
                    calc_func,
                    chunk_size=self.chunk_size
                )
            else:
                calc_df = calc_func(validated_df)
        return calc_df

# ============================================================================
# ПАРАЛЛЕЛЬНЫЙ ПРОЦЕССОР
# ============================================================================
class ParallelProcessor:
    @staticmethod
    def chunk_dataframe(df, n_chunks):
        if n_chunks <= 1:
            return [df]
        chunk_size = max(1, len(df) // n_chunks)
        return [df.iloc[i:i+chunk_size].copy() for i in range(0, len(df), chunk_size)]
    
    @classmethod
    def process_in_parallel(cls, df, func, n_workers=None, chunk_size=None):
        if not CONCURRENT_AVAILABLE:
            return func(df)
        if n_workers is None:
            n_workers = max(1, mp.cpu_count() - 1)
        if chunk_size is not None:
            chunks = [df.iloc[i:i+chunk_size].copy() for i in range(0, len(df), chunk_size)]
        else:
            chunks = cls.chunk_dataframe(df, n_workers)
        if len(chunks) <= 1:
            return func(df)
        with ThreadPoolExecutor(max_workers=n_workers) as executor:
            results = list(executor.map(func, chunks))
        return pd.concat(results, ignore_index=True)

# ============================================================================
# STREAMLIT UI
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

def render_sidebar():
    with st.sidebar:
        st.header("⚙️ Настройки")
        api_key = st.text_input("API Key Яндекс Маркета", value=st.session_state.api_key, type="password")
        business_id = st.text_input("Business ID", value=st.session_state.business_id)
        st.session_state.api_key = api_key
        st.session_state.business_id = business_id
        
        tax_options = [tax.value.label for tax in TaxSystem]
        tax_label = st.selectbox("Система налогообложения", tax_options, index=0)
        scheme_options = ["FBS (склад продавца)", "FBY (склад Маркета)", "Экспресс", "DBS (доставка продавца)"]
        scheme_label = st.selectbox("Схема работы", scheme_options, index=0)
        payment_options = ["Ежемесячно (1.0%)", "Раз в 2 недели (1.3%)", "Еженедельно, 4 нед. (1.6%)", "Ежедневно (3.3%)"]
        payment_frequency = st.selectbox("Частота выплат", payment_options, index=2)
        use_api = st.checkbox("Использовать API ЯМ для тарифов", value=True)
        st.subheader("⚡ Производительность")
        parallel = st.checkbox("Параллельная обработка (для больших файлов)", value=False)
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

def render_upload_section():
    col1, col2 = st.columns(2)
    tariff_manager = HybridTariffManager()
    main_df = None
    with col1:
        st.subheader("1. Загрузка данных товаров")
        uploaded_file = st.file_uploader("Загрузите Excel или CSV", type=['xlsx', 'xls', 'csv'])
        if uploaded_file is not None:
            try:
                raw_df = UniversalDataNormalizer.load_file(io.BytesIO(uploaded_file.read()), uploaded_file.name)
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
        tariff_file = st.file_uploader("Тарифы (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key="tariff_uploader")
        if tariff_file is not None:
            try:
                t_raw = UniversalDataNormalizer.load_file(io.BytesIO(tariff_file.read()), tariff_file.name)
                t_raw.columns = [str(c).strip().lower() for c in t_raw.columns]
                tariff_manager.load_tariffs_from_file(t_raw)
                st.success(f"Загружено {len(tariff_manager.tariffs)} тарифов")
            except Exception as e:
                st.error(f"Ошибка тарифов: {e}")
    return main_df, tariff_manager

def render_results(df_calc):
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
    display_cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent', 'roi_percent', 'profitability_status', 'source', 'scheme']
    st.dataframe(filtered_df[display_cols], use_container_width=True, height=400)
    return filtered_df

def render_export_section(filtered_df, pipeline):
    st.subheader("4. Экспорт")
    include_advanced = st.checkbox("Включить расширенные листы (анализ чувствительности, прогноз)", value=True)
    if len(filtered_df) > MAX_ROWS_FOR_EXCEL:
        st.warning(f"Данных ({len(filtered_df)} строк) много. Для Excel будет сгенерирована только сводка. Рекомендуется скачать CSV.")
        if st.button("📥 Скачать CSV (полные данные)"):
            csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="Подтвердить скачивание CSV",
                data=csv,
                file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
    excel_data = UltimateExcelExporter.export_to_excel(
        df=filtered_df,
        tax_label=pipeline.tax_label,
        scheme_label=pipeline.scheme_label,
        payment_frequency=pipeline.payment_frequency,
        include_advanced=include_advanced
    )
    if excel_data:
        st.download_button(
            label="📥 Скачать Excel (оптимизированный)",
            data=excel_data,
            file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Не удалось сгенерировать Excel.")

def render_google_sheets_export(df):
    st.subheader("5. Экспорт в Google Sheets")
    with st.expander("🔐 Настройки доступа к Google Sheets"):
        sheets_url = st.text_input("Ссылка на Google Sheets", placeholder="https://docs.google.com/spreadsheets/d/.../edit")
        credentials_file = st.file_uploader("JSON-ключ сервисного аккаунта", type=['json'])
        sheet_name = st.text_input("Название листа", value="Unit Economics")
        if st.button("📤 Экспортировать") and sheets_url and credentials_file:
            st.info("Функция временно отключена для упрощения кода")

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================
def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📈", layout="wide")
    init_session_state()
    st.title(f"📊 {APP_NAME} v{APP_VERSION}")
    st.markdown("Масштабируемый калькулятор юнит-экономики для 300k+ SKU")
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
                # Закрываем сессию API, если она была создана
                if ym_api is not None:
                    ym_api.close()
        if not st.session_state.calc_df.empty:
            filtered_df = render_results(st.session_state.calc_df)
            render_export_section(filtered_df, pipeline)
            render_google_sheets_export(filtered_df)

if __name__ == "__main__":
    main()
